import streamlit as st
import pandas as pd
import io
import re
import sqlite3
import hashlib
import unicodedata
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from rapidfuzz import fuzz, process

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="PriceBot PRO V4", layout="wide")

def normalizar(txt):
    if not txt: return ""
    txt = str(txt).lower().strip()
    # Remove acentos (Transforma Água em Agua)
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    return txt

def extrair_detalhes(texto):
    if not texto: return set()
    # Padroniza 1,5kg -> 1.5kg e remove espaços
    texto = normalizar(texto).replace(',', '.')
    texto = re.sub(r'(\d+)\s+(g|kg|l|ml|lt|und|mts)', r'\1\2', texto)
    padrao = r'(\d+(?:\.\d+)?\s?(?:g|gr|kg|l|lt|ml|mts|und)\b)'
    return set(re.findall(padrao, texto))

# --- BANCO DE DADOS ---
DB_NAME = "data_master.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS products (description TEXT, barcode TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT, expiry TEXT, role TEXT)''')
    if not c.execute("SELECT * FROM users WHERE username='admin'").fetchone():
        pw = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", ('admin', pw, '2099-12-31', 'admin'))
    conn.commit()
    conn.close()

init_db()

# --- LOGIN ---
if 'auth' not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    st.sidebar.title("🔐 Login")
    u, p = st.sidebar.text_input("Usuário"), st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        conn = sqlite3.connect(DB_NAME)
        pw = hashlib.sha256(p.encode()).hexdigest()
        res = conn.execute("SELECT role FROM users WHERE username=? AND password=?", (u, pw)).fetchone()
        if res:
            st.session_state.auth = True
            st.session_state.role = res[0]
            st.rerun()
        else: st.sidebar.error("Usuário ou senha incorretos")
    st.stop()

# --- INTERFACE ---
aba = st.sidebar.radio("Menu", ["📊 Cotação", "⚙️ Banco de Dados"])

if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações PRO")

    # Carregar banco para evitar re-leitura constante
    if 'master_df' not in st.session_state:
        conn = sqlite3.connect(DB_NAME)
        st.session_state.master_df = pd.read_sql("SELECT * FROM products", conn)
        conn.close()

    master_df = st.session_state.master_df

    if master_df.empty:
        st.warning("⚠️ O banco de dados está vazio! Importe dados na aba 'Banco de Dados'.")
        st.stop()

    with st.sidebar:
        st.header("Configurações")
        sensibilidade = st.slider("Sensibilidade Similaridade (%)", 50, 100, 75)
        discount = st.number_input("Desconto Global (%)", 0.0)
        debug = st.checkbox("🔍 Modo Debug (Ver motivos de erro)")

    file = st.file_uploader("Suba sua planilha de Cotação", type=["xlsx"])

    if file:
        h_row = st.number_input("Linha do Cabeçalho:", 1, 50, 1)
        df_cols = pd.read_excel(file, header=h_row-1, nrows=0).columns.tolist()
        
        c1, c2, c3 = st.columns(3)
        d_col = c1.selectbox("Coluna Descrição", df_cols)
        b_col = c2.selectbox("Coluna Barras", df_cols)
        p_col = c3.selectbox("Coluna Preço", df_cols)

        if st.button("🚀 Iniciar Processamento"):
            price_map = dict(zip(master_df['barcode'].astype(str), master_df['price']))
            db_descs_norm = [normalizar(d) for d in master_df['description']]
            
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            
            # Mapeamento de colunas
            header_map = {str(ws.cell(row=h_row, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            idx_d, idx_b, idx_p = header_map[d_col], header_map[b_col], header_map[p_col]

            count = 0
            logs = []

            for r in range(h_row + 1, ws.max_row + 1):
                desc_orig = str(ws.cell(row=r, column=idx_d).value or "")
                desc_norm = normalizar(desc_orig)
                bar_orig = str(ws.cell(row=r, column=idx_b).value or "").split('.')[0]
                
                found_p = None
                status = "Não localizado"

                # 1. TENTA POR BARRAS (Prioridade 100%)
                if bar_orig in price_map:
                    found_p = price_map[bar_orig]
                    status = "Match: Código de Barras"
                
                # 2. TENTA POR SIMILARIDADE (OLHA O TOP 5 CANDIDATOS)
                if found_p is None and len(desc_norm) > 3:
                    alvo_pesos = extrair_detalhes(desc_orig)
                    
                    # Busca os 5 melhores matches no banco
                    matches = process.extract(desc_norm, db_descs_norm, scorer=fuzz.token_set_ratio, limit=5)
                    
                    for m_text_norm, score, m_idx in matches:
                        if score >= sensibilidade:
                            db_item = master_df.iloc[m_idx]
                            db_pesos = extrair_detalhes(db_item['description'])
                            
                            # VALIDAÇÃO: Só aceita se os pesos forem iguais OU se um deles não tiver peso
                            if not alvo_pesos or not db_pesos or alvo_pesos == db_pesos:
                                found_p = db_item['price']
                                status = f"Match: {score}% similaridade"
                                break
                            else:
                                status = f"Bloqueado: Peso Divergente ({alvo_pesos} vs {db_pesos})"

                if found_p:
                    final_p = float(found_p) * (1 - (discount/100))
                    ws.cell(row=r, column=idx_p).value = round(final_p, 2)
                    count += 1
                
                if debug: logs.append({"Linha": r, "Produto": desc_orig, "Status": status})

            output = io.BytesIO()
            wb.save(output)
            st.success(f"Finalizado! {count} preços preenchidos.")
            st.download_button("📥 Baixar Planilha Pronta", output.getvalue(), "cotacao_final.xlsx")
            if debug: st.table(logs)

elif aba == "⚙️ Banco de Dados":
    st.title("⚙️ Gestão de Preços Mestre")
    f_db = st.file_uploader("Upload Banco (Coluna 1: Desc, 2: Barras, 3: Preço)", type=["xlsx", "csv"])
    replace = st.checkbox("Substituir banco atual?")
    
    if f_db and st.button("💾 Sincronizar Banco"):
        df = pd.read_excel(f_db) if f_db.name.endswith('.xlsx') else pd.read_csv(f_db)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['description', 'barcode', 'price']
        df['barcode'] = df['barcode'].astype(str).apply(lambda x: re.sub(r'\D', '', x.split('.')[0]))
        
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("products", conn, if_exists="replace" if replace else "append", index=False)
        conn.close()
        
        # Limpa cache da sessão
        if 'master_df' in st.session_state: del st.session_state['master_df']
        st.success("Banco de Dados sincronizado com sucesso!")
