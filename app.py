import streamlit as st
import pandas as pd
import io
import re
import sqlite3
import hashlib
import unicodedata
import openpyxl
from datetime import datetime, timedelta
from rapidfuzz import fuzz, process

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PriceBot PRO V4.5", layout="wide")

# --- FUNÇÕES DE APOIO ---
def normalizar(txt):
    if not txt: return ""
    txt = str(txt).lower().strip()
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    return txt

def extrair_detalhes(texto):
    if not texto: return set()
    texto = str(texto).lower()
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', texto))

# --- REGRA DE BARRAS ORIGINAL ---
def extract_all_barcodes(val):
    if pd.isna(val): return []
    return re.findall(r'\d{8,14}', str(val))

# --- BANCO DE DADOS (USUÁRIOS E PRODUTOS) ---
DB_NAME = "data_master.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS products (description TEXT, barcode TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users 
                 (username TEXT PRIMARY KEY, password TEXT, expiry TEXT, role TEXT)''')
    
    # Criar admin padrão se não existir
    if not c.execute("SELECT * FROM users WHERE username='admin'").fetchone():
        pw = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", 
                  ('admin', pw, '2099-12-31', 'admin'))
    conn.commit()
    conn.close()

init_db()

# --- SISTEMA DE LOGIN ---
if 'auth' not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    st.sidebar.title("🔐 Acesso ao Sistema")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        conn = sqlite3.connect(DB_NAME)
        pw = hashlib.sha256(p.encode()).hexdigest()
        res = conn.execute("SELECT role, expiry FROM users WHERE username=? AND password=?", (u, pw)).fetchone()
        conn.close()
        if res:
            exp_date = datetime.strptime(res[1], '%Y-%m-%d')
            if datetime.now() <= exp_date:
                st.session_state.auth = True
                st.session_state.username = u
                st.session_state.role = res[0]
                st.rerun()
            else: st.sidebar.error("Sua senha expirou!")
        else: st.sidebar.error("Usuário ou senha incorretos")
    st.stop()

# --- MENU DE NAVEGAÇÃO ---
opcoes = ["📊 Processar Cotação", "⚙️ Gerenciar Banco"]
if st.session_state.role == 'admin':
    opcoes.append("👑 Painel Administrativo")

menu = st.sidebar.radio("Navegação", opcoes)

# --- ABA 1: COTAÇÃO ---
if menu == "📊 Processar Cotação":
    st.title("📊 Automatizador de Cotações PRO")

    if 'master_df' not in st.session_state:
        conn = sqlite3.connect(DB_NAME)
        st.session_state.master_df = pd.read_sql("SELECT * FROM products", conn)
        conn.close()

    master_df = st.session_state.master_df

    with st.sidebar:
        st.header("Configurações")
        modo_busca = st.selectbox("Escolha a regra:", 
                                 ["Híbrido (Barras + Similaridade)", 
                                  "Apenas Código de Barras", 
                                  "Apenas Similaridade"])
        sensibilidade = st.slider("Sensibilidade Similaridade (%)", 50, 100, 75)
        desconto = st.number_input("Desconto Global (%)", 0.0)

    file = st.file_uploader("Upload da Cotação", type=["xlsx"])

    if file:
        h_row = st.number_input("Linha do Cabeçalho:", 1, 100, 10)
        df_dest = pd.read_excel(file, header=h_row-1)
        
        c1, c2, c3 = st.columns(3)
        col_desc = c1.selectbox("Coluna Descrição", df_dest.columns)
        col_bar = c2.selectbox("Coluna Cód. Barras", df_dest.columns)
        col_price = c3.selectbox("Coluna Preço", df_dest.columns)

        if st.button("🚀 INICIAR PROCESSAMENTO"):
            price_map = dict(zip(master_df['barcode'].astype(str), master_df['price']))
            db_descs_norm = [normalizar(d) for d in master_df['description']]
            
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            header_map = {str(ws.cell(row=h_row, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            idx_d, idx_b, idx_p = header_map[col_desc], header_map[col_bar], header_map[col_price]

            count = 0
            for r in range(h_row + 1, ws.max_row + 1):
                orig_desc = str(ws.cell(row=r, column=idx_d).value or "")
                norm_desc = normalizar(orig_desc)
                found_p = None

                if "Barras" in modo_busca or "Híbrido" in modo_busca:
                    bcs = extract_all_barcodes(ws.cell(row=r, column=idx_b).value)
                    for b in bcs:
                        if b in price_map:
                            found_p = price_map[b]
                            break

                if found_p is None and ("Similaridade" in modo_busca or "Híbrido" in modo_busca) and len(norm_desc) > 3:
                    alvo_pesos = extrair_detalhes(orig_desc)
                    matches = process.extract(norm_desc, db_descs_norm, scorer=fuzz.token_set_ratio, limit=5)
                    for m_text, score, m_idx in matches:
                        if score >= sensibilidade:
                            item_db = master_df.iloc[m_idx]
                            db_pesos = extrair_detalhes(item_db['description'])
                            if not alvo_pesos or not db_pesos or alvo_pesos == db_pesos:
                                found_p = item_db['price']
                                break

                if found_p:
                    v_final = float(found_p) * (1 - (desconto/100))
                    ws.cell(row=r, column=idx_p).value = round(v_final, 2)
                    count += 1

            output = io.BytesIO()
            wb.save(output)
            st.success(f"Finalizado! {count} itens processados.")
            st.download_button("📥 Baixar Planilha", output.getvalue(), "cotacao_finalizada.xlsx")

# --- ABA 2: BANCO ---
elif menu == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    f_db = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    sub = st.checkbox("Substituir tudo?")
    if f_db and st.button("💾 Sincronizar"):
        df = pd.read_excel(f_db) if f_db.name.endswith('.xlsx') else pd.read_csv(f_db)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['description', 'barcode', 'price']
        df['barcode'] = df['barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("products", conn, if_exists="replace" if sub else "append", index=False)
        conn.close()
        if 'master_df' in st.session_state: del st.session_state['master_df']
        st.success("Banco de Dados Atualizado!")

# --- ABA 3: ADMINISTRAÇÃO (RESTAURADA) ---
elif menu == "👑 Painel Administrativo":
    st.title("👑 Gestão de Usuários")
    
    with st.expander("➕ Criar Novo Usuário"):
        new_u = st.text_input("Nome do Usuário")
        new_p = st.text_input("Senha", type="password", key="new_p")
        validade = st.number_input("Validade da Senha (Dias)", min_value=1, value=30)
        
        if st.button("Confirmar Criação"):
            if new_u and new_p:
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                pw_h = hashlib.sha256(new_p.encode()).hexdigest()
                exp = (datetime.now() + timedelta(days=validade)).strftime('%Y-%m-%d')
                try:
                    c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", (new_u, pw_h, exp, 'user'))
                    conn.commit()
                    st.success(f"Usuário {new_u} criado com sucesso! Válido até {exp}")
                except: st.error("Usuário já existe!")
                conn.close()
            else: st.warning("Preencha todos os campos.")

    st.subheader("👥 Usuários Cadastrados")
    conn = sqlite3.connect(DB_NAME)
    usuarios_df = pd.read_sql("SELECT username, expiry, role FROM users", conn)
    conn.close()
    st.dataframe(usuarios_df, use_container_width=True)
