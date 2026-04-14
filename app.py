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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema de Cotação Inteligente V4", layout="wide")

# --- FUNÇÕES DE TRATAMENTO DE TEXTO ---
def normalizar(txt):
    if not txt: return ""
    txt = str(txt).lower().strip()
    # Remove acentos para facilitar a comparação (ex: Café -> cafe)
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    return txt

def extrair_detalhes(texto):
    """Extrai pesos e medidas padronizando para comparação (ex: 1,5kg -> 1.5kg)."""
    if not texto: return set()
    texto = normalizar(texto).replace(',', '.')
    # Remove espaços entre número e unidade (ex: 1 kg -> 1kg)
    texto = re.sub(r'(\d+)\s+(g|kg|l|ml|lt|und|mts)', r'\1\2', texto)
    padrao = r'(\d+(?:\.\d+)?\s?(?:g|gr|kg|l|lt|ml|mts|und)\b)'
    return set(re.findall(padrao, texto))

def limpar_barcode(val):
    """Trata códigos de barras evitando erros com células vazias ou decimais do Excel."""
    if pd.isna(val) or val == "" or str(val).strip() == "":
        return ""
    # Remove o .0 que o Excel coloca em números e limpa caracteres não numéricos
    s = str(val).split('.')[0]
    return re.sub(r'\D', '', s)

# --- BANCO DE DADOS (SQLite) ---
DB_NAME = "data_master.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS products 
                 (description TEXT, barcode TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users 
                 (username TEXT PRIMARY KEY, password TEXT, expiry TEXT, role TEXT)''')
    
    # Usuário padrão caso não exista
    if not c.execute("SELECT * FROM users WHERE username='admin'").fetchone():
        pw_hash = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", 
                  ('admin', pw_hash, '2099-12-31', 'admin'))
    conn.commit()
    conn.close()

init_db()

# --- SISTEMA DE LOGIN ---
if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.sidebar.title("🔐 Acesso Restrito")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        conn = sqlite3.connect(DB_NAME)
        pw_input = hashlib.sha256(p.encode()).hexdigest()
        res = conn.execute("SELECT role FROM users WHERE username=? AND password=?", (u, pw_input)).fetchone()
        conn.close()
        if res:
            st.session_state.auth = True
            st.session_state.role = res[0]
            st.rerun()
        else:
            st.sidebar.error("Usuário ou senha incorretos.")
    st.stop()

# --- INTERFACE PRINCIPAL ---
menu = st.sidebar.radio("Navegação", ["📊 Processar Cotação", "⚙️ Gerenciar Banco"])

# ABA 1: PROCESSAMENTO
if menu == "📊 Processar Cotação":
    st.title("📊 Automatizador de Cotações PRO")

    # Carrega banco de dados para a memória da sessão
    if 'master_df' not in st.session_state:
        conn = sqlite3.connect(DB_NAME)
        st.session_state.master_df = pd.read_sql("SELECT * FROM products", conn)
        conn.close()

    master_df = st.session_state.master_df

    if master_df.empty:
        st.warning("⚠️ O banco de dados está vazio. Vá em 'Gerenciar Banco' e faça o upload.")
        st.stop()

    with st.sidebar:
        st.header("Parâmetros de Busca")
        sensibilidade = st.slider("Sensibilidade Similaridade (%)", 50, 100, 75)
        desconto = st.number_input("Desconto Global (%)", 0.0)
        debug = st.checkbox("🔍 Ver Relatório de Erros")

    file = st.file_uploader("Upload da Planilha de Destino (Cotação)", type=["xlsx"])

    if file:
        h_row = st.number_input("Linha do Cabeçalho:", 1, 100, 1)
        df_cols = pd.read_excel(file, header=h_row-1, nrows=0).columns.tolist()
        
        c1, c2, c3 = st.columns(3)
        col_desc = c1.selectbox("Coluna Descrição", df_cols)
        col_bar = c2.selectbox("Coluna Cód. Barras", df_cols)
        col_price = c3.selectbox("Coluna Preço (Destino)", df_cols)

        if st.button("🚀 Iniciar Processamento Inteligente"):
            # Mapeamento rápido de códigos de barras
            price_map = dict(zip(master_df['barcode'].astype(str), master_df['price']))
            # Normaliza descrições do banco para busca veloz
            db_descs_norm = [normalizar(d) for d in master_df['description']]
            
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            
            # Localiza índices das colunas
            header_map = {str(ws.cell(row=h_row, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            idx_d, idx_b, idx_p = header_map[col_desc], header_map[col_bar], header_map[col_price]

            count = 0
            logs = []

            with st.spinner("Processando similaridades..."):
                for r in range(h_row + 1, ws.max_row + 1):
                    orig_desc = str(ws.cell(row=r, column=idx_d).value or "")
                    norm_desc = normalizar(orig_desc)
                    orig_bar = limpar_barcode(ws.cell(row=r, column=idx_b).value)
                    
                    found_p = None
                    status = "Não encontrado"

                    # 1. TENTA POR CÓDIGO DE BARRAS
                    if orig_bar and orig_bar in price_map:
                        found_p = price_map[orig_bar]
                        status = "Match: Código de Barras"
                    
                    # 2. TENTA POR SIMILARIDADE (TOP 5 CANDIDATOS)
                    if found_p is None and len(norm_desc) > 3:
                        alvo_pesos = extrair_detalhes(orig_desc)
                        
                        # Extrai os 5 melhores matches textuais
                        matches = process.extract(norm_desc, db_descs_norm, scorer=fuzz.token_set_ratio, limit=5)
                        
                        for m_text, score, m_idx in matches:
                            if score >= sensibilidade:
                                item_db = master_df.iloc[m_idx]
                                db_pesos = extrair_detalhes(item_db['description'])
                                
                                # VALIDAÇÃO DE PESO: Se ambos têm peso, devem ser iguais.
                                if not alvo_pesos or not db_pesos or alvo_pesos == db_pesos:
                                    found_p = item_db['price']
                                    status = f"Match: {score}% similar"
                                    break
                                else:
                                    status = f"Bloqueado: Peso Divergente ({alvo_pesos} vs {db_pesos})"

                    if found_p:
                        valor_final = float(found_p) * (1 - (desconto/100))
                        ws.cell(row=r, column=idx_p).value = round(valor_final, 2)
                        count += 1
                    
                    if debug: logs.append({"Linha": r, "Produto": orig_desc, "Status": status})

            output = io.BytesIO()
            wb.save(output)
            st.success(f"Finalizado! {count} itens atualizados.")
            st.download_button("📥 Baixar Planilha Processada", output.getvalue(), "cotacao_final.xlsx")
            if debug: st.table(logs)

# ABA 2: GERENCIAMENTO DO BANCO
elif menu == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    
    upload_db = st.file_uploader("Upload Banco Mestre (Descrição, Barras, Preço)", type=["xlsx", "csv"])
    substituir = st.checkbox("Substituir dados existentes?")
    
    if upload_db and st.button("💾 Sincronizar"):
        try:
            df = pd.read_excel(upload_db) if upload_db.name.endswith('.xlsx') else pd.read_csv(upload_db)
            df = df.iloc[:, [0, 1, 2]]
            df.columns = ['description', 'barcode', 'price']
            
            # Limpeza robusta de códigos de barras para evitar o erro de AttributeError
            df['barcode'] = df['barcode'].apply(limpar_barcode)
            df['price'] = pd.to_numeric(df['price'], errors='coerce').fillna(0.0)
            
            conn = sqlite3.connect(DB_NAME)
            df.to_sql("products", conn, if_exists="replace" if substituir else "append", index=False)
            conn.close()
            
            # Limpa o cache para forçar a leitura do novo banco
            if 'master_df' in st.session_state: del st.session_state['master_df']
            st.success("Banco de dados atualizado com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar banco de dados: {e}")
