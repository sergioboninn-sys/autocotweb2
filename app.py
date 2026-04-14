import streamlit as st
import pandas as pd
import io
import re
import sqlite3
import hashlib
import os
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from rapidfuzz import fuzz, process

# --- CONFIGURAÇÃO E BANCO DE DADOS ---
st.set_page_config(page_title="PriceBot PRO v3", layout="wide")

DB_NAME = "data_master.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS products 
                 (description TEXT, barcode TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users 
                 (username TEXT PRIMARY KEY, password TEXT, expiry TEXT, role TEXT)''')
    
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        pw_hash = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", 
                  ('admin', pw_hash, '2099-12-31', 'admin'))
    conn.commit()
    conn.close()

init_db()

# --- FUNÇÕES DE APOIO E CORREÇÕES ---
def extra_round(valor):
    if valor is None or pd.isna(valor): return 0.0
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def extract_barcodes(val):
    if pd.isna(val) or val is None: return []
    text = str(val).split('.')[0]
    return re.findall(r'\d{8,14}', text)

def extrair_detalhes(texto):
    """
    MELHORIA: Captura pesos decimais (1,5kg, 1.6L) e padroniza separadores.
    """
    if not texto: return set()
    texto = str(texto).lower().replace(',', '.') 
    # Regex para capturar números decimais ou inteiros + unidade
    padrao = r'(\d+(?:\.\d+)?\s?(?:g|gr|kg|l|lt|ml|mts|und)\b)'
    return set(re.findall(padrao, texto))

def check_login(user, pwd):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    pw_hash = hashlib.sha256(pwd.encode()).hexdigest()
    c.execute("SELECT role, expiry FROM users WHERE username=? AND password=?", (user, pw_hash))
    res = c.fetchone()
    conn.close()
    if res:
        exp = datetime.strptime(res[1], "%Y-%m-%d")
        if datetime.now() <= exp: return res[0]
    return None

# --- AUTH ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Login")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        role = check_login(u, p)
        if role:
            st.session_state.autenticado = True
            st.session_state.user_role = role
            st.rerun()
        else: st.sidebar.error("Acesso negado")
    st.stop()

# --- NAVEGAÇÃO ---
tabs = ["📊 Cotação", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações PRO")
    
    conn = sqlite3.connect(DB_NAME)
    master_db = pd.read_sql("SELECT * FROM products", conn)
    conn.close()

    if master_db.empty:
        st.warning("⚠️ O banco de dados está vazio.")
        st.stop()

    with st.sidebar.expander("🛠️ Parâmetros", expanded=True):
        modo = st.selectbox("Estratégia:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
        sensibilidade = st.slider("Sensibilidade Match (%)", 50, 100, 80)
        discount = st.number_input("Desconto (%)", 0.0)
        arredondar = st.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Planilha de Destino", type=["xlsx"])

    if target_file:
        c1, c2 = st.columns(2)
        header_row = c1.number_input("Linha do Cabeçalho:", 1, 100, 1)
        start_row = c2.number_input("Linha Início Produtos:", 1, 1000, 2)

        cols_detected = pd.read_excel(target_file, header=header_row-1, nrows=0).columns.tolist()
        
        col_m1, col_m2, col_m3 = st.columns(3)
        desc_col = col_m1.selectbox("Coluna Descrição", cols_detected)
        bar_col = col_m2.selectbox("Coluna Barras", cols_detected)
        price_col = col_m3.selectbox("Coluna Preço", cols_detected)

        if st.button("🚀 Processar Planilha"):
            price_map = dict(zip(master_db['barcode'], master_db['price']))
            db_descriptions = master_db['description'].tolist()
            
            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            try:
                header_cells = {str(ws.cell(row=header_row, column=i).value).strip(): i 
                               for i in range(1, ws.max_column + 1)}
                d_idx = header_cells[desc_col.strip()]
                b_idx = header_cells[bar_col.strip()]
                p_idx = header_cells[price_col.strip()]
            except:
                st.error("Erro no mapeamento. Verifique a linha do cabeçalho.")
                st.stop()

            progress = st.progress(0)
            status = st.empty()
            found_count = 0

            for i, r in enumerate(range(int(start_row), ws.max_row + 1)):
                d_val = str(ws.cell(row=r, column=d_idx).value or "")
                b_val = ws.cell(row=r, column=b_idx).value
                found_p = None
                
                # 1. BARRAS
                if "Barras" in modo or "Híbrido" in modo:
                    bcodes = extract_barcodes(b_val)
                    for b in bcodes:
                        if b in price_map:
                            found_p = price_map[b]
                            break
                
                # 2. SIMILARIDADE (Com correção de decimais e pesos)
                if found_p is None and ("Similaridade" in modo or "Híbrido" in modo) and len(d_val) > 3:
                    alvo_detalhes = extrair_detalhes(d_val)
                    
                    # Busca via RapidFuzz pegando o índice
                    match_data = process.extractOne(d_val, db_descriptions, scorer=fuzz.token_set_ratio)
                    
                    if match_data:
                        match_text, score, index = match_data
                        if score >= sensibilidade:
                            db_detalhes = extrair_detalhes(match_text)
                            
                            # Validação Flexível de Pesos
                            conflito = False
                            if alvo_detalhes and db_detalhes:
                                if alvo_detalhes != db_detalhes:
                                    conflito = True # Pesos diferentes = produto diferente
                            
                            if not conflito:
                                found_p = master_db.iloc[index]['price']

                if found_p is not None:
                    final_p = float(found_p) * (1 - (discount / 100))
                    if arredondar: final_p = extra_round(final_p)
                    ws.cell(row=r, column=p_idx).value = final_p
                    found_count += 1

                progress.progress((i + 1) / (ws.max_row - int(start_row) + 1))

            output = io.BytesIO()
            wb.save(output)
            st.success(f"✅ {found_count} itens processados!")
            st.download_button("📥 Baixar Planilha", output.getvalue(), "cotacao_final.xlsx")

# --- ABA 2: GERENCIAR BANCO ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Dados")
    f = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    replace = st.checkbox("Substituir banco existente?")
    
    if f and st.button("💾 Salvar"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['description', 'barcode', 'price']
        df['barcode'] = df['barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
        
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("products", conn, if_exists="replace" if replace else "append", index=False)
        conn.close()
        st.success("Banco atualizado!")

# --- ABA 3: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Usuários")
    with st.form("Novo"):
        nu, np = st.text_input("User"), st.text_input("Pass")
        days = st.number_input("Validade", 1, 365, 30)
        if st.form_submit_button("Criar"):
            conn = sqlite3.connect(DB_NAME)
            pw_h = hashlib.sha256(np.encode()).hexdigest()
            exp = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
            try:
                conn.execute("INSERT INTO users VALUES (?, ?, ?, ?)", (nu, pw_h, exp, 'user'))
                conn.commit()
                st.success("Criado!")
            except: st.error("Erro")
            conn.close()
    
    conn = sqlite3.connect(DB_NAME)
    st.table(pd.read_sql("SELECT username, expiry FROM users", conn))
    conn.close()
