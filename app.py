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
st.set_page_config(page_title="PriceBot PRO v2", layout="wide")

DB_NAME = "data_master.db"

def init_db():
    """Inicializa o banco SQLite e tabelas se não existirem."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # Tabela de Produtos
    c.execute('''CREATE TABLE IF NOT EXISTS products 
                 (description TEXT, barcode TEXT, price REAL)''')
    # Tabela de Usuários
    c.execute('''CREATE TABLE IF NOT EXISTS users 
                 (username TEXT PRIMARY KEY, password TEXT, expiry TEXT, role TEXT)''')
    
    # Criar admin padrão se não existir (senha: admin123)
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        pw_hash = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", 
                  ('admin', pw_hash, '2099-12-31', 'admin'))
    
    conn.commit()
    conn.close()

init_db()

# --- FUNÇÕES DE SEGURANÇA E AUTH ---
def check_login(user, pwd):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    pw_hash = hashlib.sha256(pwd.encode()).hexdigest()
    c.execute("SELECT role, expiry FROM users WHERE username=? AND password=?", (user, pw_hash))
    res = c.fetchone()
    conn.close()
    if res:
        exp = datetime.strptime(res[1], "%Y-%m-%d")
        if datetime.now() <= exp:
            return res[0]
    return None

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if valor is None or pd.isna(valor): return 0.0
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def extract_barcodes(val):
    if pd.isna(val) or val is None: return []
    text = str(val).split('.')[0]
    return re.findall(r'\d{8,14}', text)

def extrair_detalhes(texto):
    """Extrai unidades de medida para evitar falsos positivos (ex: 1kg vs 500g)."""
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', str(texto).lower()))

# --- INTERFACE DE LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Login PRO")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Aceder"):
        role = check_login(u, p)
        if role:
            st.session_state.autenticado = True
            st.session_state.user_role = role
            st.rerun()
        else:
            st.sidebar.error("Dados inválidos ou conta expirada.")
    st.stop()

# --- NAVEGAÇÃO ---
tabs = ["📊 Cotação", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Menu Principal", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações Inteligente")
    
    # Carregar Banco para Memória (Performance)
    conn = sqlite3.connect(DB_NAME)
    master_db = pd.read_sql("SELECT * FROM products", conn)
    conn.close()

    if master_db.empty:
        st.warning("⚠️ Banco de dados vazio. Importe produtos na aba Gerenciar Banco.")
        st.stop()

    with st.sidebar.expander("🛠️ Parâmetros de Busca", expanded=True):
        modo = st.selectbox("Estratégia:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
        sensibilidade = st.slider("Sensibilidade de Match (%)", 50, 100, 85)
        discount = st.number_input("Desconto Global (%)", 0.0)
        arredondar = st.checkbox("Arredondamento Financeiro", value=True)

    target_file = st.file_uploader("Submeter Planilha de Destino (XLSX)", type=["xlsx"])

    if target_file:
        # Preview rápido para mapeamento
        t_df_preview = pd.read_excel(target_file, nrows=10, header=None)
        
        c1, c2 = st.columns(2)
        header_row = c1.number_input("Linha do Cabeçalho:", 1, 100, 1)
        start_row = c2.number_input("Linha de início dos dados:", 1, 1000, 2)

        # Mapeamento de colunas baseado na linha do cabeçalho escolhida
        cols_detected = pd.read_excel(target_file, header=header_row-1, nrows=0).columns.tolist()
        
        col_m1, col_m2, col_m3 = st.columns(3)
        desc_col = col_m1.selectbox("Coluna Descrição", cols_detected)
        bar_col = col_m2.selectbox("Coluna EAN/Barras", cols_detected)
        price_col = col_m3.selectbox("Coluna Preço Alvo", cols_detected)

        if st.button("🚀 Iniciar Processamento PRO"):
            # Preparar Mapas de busca rápida
            price_map = dict(zip(master_db['barcode'], master_db['price']))
            db_descriptions = master_db['description'].tolist()
            
            # Carregar com Openpyxl para preservar estilos
            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            # Identificar índices reais (1-based)
            try:
                # Criar dicionário de nomes de colunas -> index
                header_cells = {str(ws.cell(row=header_row, column=i).value).strip(): i 
                               for i in range(1, ws.max_column + 1)}
                
                d_idx = header_cells[desc_col.strip()]
                b_idx = header_cells[bar_col.strip()]
                p_idx = header_cells[price_col.strip()]
            except Exception as e:
                st.error(f"Erro no mapeamento: {e}. Verifique se o cabeçalho está na linha {header_row}.")
                st.stop()

            # Processamento com Barra de Progresso
            progress_bar = st.progress(0)
            status_text = st.empty()
            found_count = 0
            rows_to_process = range(int(start_row), ws.max_row + 1)
            total_rows = len(rows_to_process)

            for i, r in enumerate(rows_to_process):
                d_val = str(ws.cell(row=r, column=d_idx).value or "")
                b_val = ws.cell(row=r, column=b_idx).value
                
                found_p = None
                
                # 1. Busca por Código de Barras
                if "Barras" in modo or "Híbrido" in modo:
                    bcodes = extract_barcodes(b_val)
                    for b in bcodes:
                        if b in price_map:
                            found_p = price_map[b]
                            break
                
                # 2. Busca por Similaridade (RapidFuzz)
                if found_p is None and ("Similaridade" in modo or "Híbrido" in modo) and len(d_val) > 3:
                    # Filtro de Unidade de Medida (Evita que Leite 1L dê match com Leite 200ml)
                    alvo_detalhes = extrair_detalhes(d_val)
                    
                    # Busca o melhor match textual
                    match = process.extractOne(d_val, db_descriptions, scorer=fuzz.token_set_ratio)
                    
                    if match and match[1] >= sensibilidade:
                        # Verificação extra de detalhes (Ex: Kg, Ml)
                        row_db = master_db[master_db['description'] == match[0]].iloc[0]
                        if alvo_detalhes == extrair_detalhes(row_db['description']):
                            found_p = row_db['price']

                # Aplicar Preço se encontrado
                if found_p is not None:
                    final_p = float(found_p) * (1 - (discount / 100))
                    if arredondar: final_p = extra_round(final_p)
                    ws.cell(row=r, column=p_idx).value = final_p
                    found_count += 1

                # Atualizar UI
                if i % 10 == 0:
                    progress_bar.progress((i + 1) / total_rows)
                    status_text.text(f"Processando linha {r}/{ws.max_row}...")

            # Finalização
            output = io.BytesIO()
            wb.save(output)
            st.success(f"✅ Concluído! {found_count} preços atualizados com sucesso.")
            st.download_button("📥 Baixar Planilha Processada", output.getvalue(), "cotacao_inteligente.xlsx")

# --- ABA 2: GERENCIAR BANCO ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gestão de Dados Mestre")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        f = st.file_uploader("Carregar Novo Banco (Excel/CSV)", type=["xlsx", "csv"])
    
    with col_up2:
        st.info("O arquivo deve conter: Descrição, Código e Preço (nesta ordem).")
        replace_data = st.checkbox("Substituir banco atual?", value=False)

    if f and st.button("💾 Sincronizar com SQLite"):
        try:
            df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
            df = df.iloc[:, [0, 1, 2]]
            df.columns = ['description', 'barcode', 'price']
            
            # Limpeza de dados
            df['barcode'] = df['barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
            df['price'] = pd.to_numeric(df['price'], errors='coerce').fillna(0.0)
            
            conn = sqlite3.connect(DB_NAME)
            if replace_data:
                df.to_sql("products", conn, if_exists="replace", index=False)
            else:
                df.to_sql("products", conn, if_exists="append", index=False)
            conn.close()
            st.success("Dados sincronizados com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {e}")

# --- ABA 3: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Acessos")
    
    with st.form("Criar Novo Utilizador"):
        new_u = st.text_input("Username")
        new_p = st.text_input("Password (Texto simples)")
        new_days = st.number_input("Validade (Dias)", 1, 365, 30)
        new_role = st.selectbox("Perfil", ["user", "admin"])
        
        if st.form_submit_button("Criar Utilizador"):
            if new_u and new_p:
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                pw_h = hashlib.sha256(new_p.encode()).hexdigest()
                exp_date = (datetime.now() + timedelta(days=new_days)).strftime("%Y-%m-%d")
                try:
                    c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", (new_u, pw_h, exp_date, new_role))
                    conn.commit()
                    st.success(f"Utilizador {new_u} criado!")
                except:
                    st.error("Utilizador já existe.")
                conn.close()

    # Listar usuários
    conn = sqlite3.connect(DB_NAME)
    users_df = pd.read_sql("SELECT username, expiry, role FROM users", conn)
    conn.close()
    st.table(users_df)
