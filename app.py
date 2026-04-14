import streamlit as st
import pandas as pd
import io
import re
import sqlite3
import hashlib
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from rapidfuzz import fuzz, process
import unicodedata

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="PriceBot PRO Cloud", layout="wide")

# Função para normalizar texto (remove acentos e padroniza)
def normalizar(txt):
    if not txt: return ""
    txt = str(txt).lower().strip()
    return "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

# --- BANCO DE DADOS ---
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
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", ('admin', pw_hash, '2099-12-31', 'admin'))
    conn.commit()
    conn.close()

init_db()

# --- FUNÇÕES DE APOIO ---
def extrair_detalhes(texto):
    if not texto: return set()
    texto = normalizar(texto).replace(',', '.')
    padrao = r'(\d+(?:\.\d+)?\s?(?:g|gr|kg|l|lt|ml|mts|und)\b)'
    return set(re.findall(padrao, texto))

def extra_round(valor):
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acesso")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        pw_h = hashlib.sha256(p.encode()).hexdigest()
        c.execute("SELECT role FROM users WHERE username=? AND password=?", (u, pw_h))
        res = c.fetchone()
        if res:
            st.session_state.autenticado = True
            st.session_state.user_role = res[0]
            st.rerun()
        else: st.sidebar.error("Dados incorretos")
    st.stop()

# --- INTERFACE PRINCIPAL ---
aba = st.sidebar.radio("Menu", ["📊 Cotação", "⚙️ Banco de Dados"])

if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")

    # Carregar banco para a sessão para evitar perdas
    if 'master_df' not in st.session_state:
        conn = sqlite3.connect(DB_NAME)
        st.session_state.master_df = pd.read_sql("SELECT * FROM products", conn)
        conn.close()

    if st.session_state.master_df.empty:
        st.info("O banco de dados está vazio. Vá em 'Banco de Dados' e envie seus preços.")
        st.stop()

    with st.sidebar:
        st.header("Configurações")
        modo = st.selectbox("Regra:", ["Híbrido (Recomendado)", "Apenas Barras", "Apenas Similaridade"])
        sensibilidade = st.slider("Sensibilidade Match (%)", 50, 100, 75)
        discount = st.number_input("Desconto Global (%)", 0.0)
        debug_mode = st.checkbox("Mostrar Log de Processamento (Debug)")

    file = st.file_uploader("Suba a planilha de cotação (XLSX)", type=["xlsx"])

    if file:
        h_row = st.number_input("Linha do Cabeçalho (onde estão os nomes das colunas):", 1, 50, 1)
        
        # Leitura rápida para mapeamento
        df_cols = pd.read_excel(file, header=h_row-1, nrows=0).columns.tolist()
        col1, col2, col3 = st.columns(3)
        d_col = col1.selectbox("Coluna Descrição", df_cols)
        b_col = col2.selectbox("Coluna Barras", df_cols)
        p_col = col3.selectbox("Coluna Preço", df_cols)

        if st.button("🚀 Iniciar"):
            # Preparar dados
            master = st.session_state.master_df
            price_map = dict(zip(master['barcode'].astype(str), master['price']))
            # Criamos uma lista de descrições normalizadas para comparação justa
            db_descs_norm = [normalizar(d) for d in master['description']]
            
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            
            # Mapear índices (A=1, B=2...)
            header_map = {str(ws.cell(row=h_row, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            idx_d, idx_b, idx_p = header_map[d_col], header_map[b_col], header_map[p_col]

            count = 0
            log_debug = []

            for r in range(h_row + 1, ws.max_row + 1):
                desc_orig = str(ws.cell(row=r, column=idx_d).value or "")
                desc_target = normalizar(desc_orig)
                bar_target = str(ws.cell(row=r, column=idx_b).value or "").split('.')[0]
                
                found_p = None
                reason = "Não encontrado"

                # 1. Busca por Barras
                if "Barras" in modo or "Híbrido" in modo:
                    if bar_target in price_map:
                        found_p = price_map[bar_target]
                        reason = "Match por Barras"

                # 2. Busca por Similaridade
                if found_p is None and ("Similaridade" in modo or "Híbrido" in modo) and len(desc_target) > 3:
                    alvo_detalhes = extrair_detalhes(desc_orig)
                    
                    # WRatio é melhor para frases curtas e palavras fora de ordem
                    res = process.extractOne(desc_target, db_descs_norm, scorer=fuzz.WRatio)
                    
                    if res and res[1] >= sensibilidade:
                        match_text_norm, score, idx = res
                        db_item = master.iloc[idx]
                        db_detalhes = extrair_detalhes(db_item['description'])
                        
                        # Verifica se as medidas (kg, ml) batem
                        if not alvo_detalhes or not db_detalhes or alvo_detalhes == db_detalhes:
                            found_p = db_item['price']
                            reason = f"Similaridade: {score}%"
                        else:
                            reason = f"Conflito de Peso ({alvo_detalhes} vs {db_detalhes})"

                if found_p:
                    final_v = float(found_p) * (1 - (discount/100))
                    ws.cell(row=r, column=idx_p).value = extra_round(final_v)
                    count += 1
                
                if debug_mode: log_debug.append({"Linha": r, "Produto": desc_orig, "Resultado": reason})

            output = io.BytesIO()
            wb.save(output)
            st.success(f"Sucesso! {count} itens atualizados.")
            st.download_button("📥 Baixar Resultado", output.getvalue(), "resultado.xlsx")
            
            if debug_mode:
                st.write("### Relatório de Debug")
                st.table(log_debug)

elif aba == "⚙️ Banco de Dados":
    st.title("⚙️ Gerenciar Preços")
    f_db = st.file_uploader("Upload Banco de Dados (Descrição, Barras, Preço)", type=["xlsx", "csv"])
    replace = st.checkbox("Substituir dados antigos?")
    
    if f_db and st.button("💾 Sincronizar"):
        df = pd.read_excel(f_db) if f_db.name.endswith('.xlsx') else pd.read_csv(f_db)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['description', 'barcode', 'price']
        df['barcode'] = df['barcode'].astype(str).apply(lambda x: re.sub(r'\D', '', x.split('.')[0]))
        
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("products", conn, if_exists="replace" if replace else "append", index=False)
        conn.close()
        
        # Forçar atualização da sessão
        st.session_state.pop('master_df', None)
        st.success("Banco de dados atualizado com sucesso!")
