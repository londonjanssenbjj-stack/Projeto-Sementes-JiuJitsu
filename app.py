import streamlit as st
import pandas as pd
import datetime
import urllib.parse
from io import BytesIO

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E DESIGN PREMIUM (DARK MODE, AZUL PROFUNDO & DOURADO)
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="wide"
)

# Estilização CSS Customizada de Alta Fidelidade (Identidade Visual Baseada na Logo)
st.markdown("""
<style>
    @import url('https://googleapis.com');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Fundo Geral e Sidebar */
    .stApp { background-color: #0A0F18; color: #E2E…
[23:48, 15/09/2026] Papai London: import streamlit as st
import pandas as pd
import datetime
import urllib.parse
from io import BytesIO

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E DESIGN PREMIUM (DARK MODE, AZUL PROFUNDO & DOURADO)
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="wide"
)

# Estilização CSS Customizada de Alta Fidelidade (Identidade Visual Baseada na Logo)
st.markdown("""
<style>
    @import url('https://googleapis.com');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Fundo Geral e Sidebar */
    .stApp { background-color: #0A0F18; color: #E2E8F0; }
    [data-testid="stSidebar"] { background-color: #111827; border-right: 1px solid #1E293B; }
    
    /* Títulos e Elementos em Dourado Ouro */
    h1, h2, h3, h4, .gold-text { color: #F59E0B !important; font-weight: 700; }
    
    /* Customização dos Cards de Métricas e Conteineres */
    div[data-testid="stMetric"] {
        background-color: #111827;
        border: 1px solid #1E293B;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.4);
    }
    div[data-testid="stMetricLabel"] { color: #94A3B8 !important; font-weight: 600; text-transform: uppercase; font-size: 0.8rem; }
    div[data-testid="stMetricValue"] { color: #F59E0B !important; font-weight: 800; font-size: 2rem; }

    /* Inputs, Selectboxes e Formulários */
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea {
        background-color: #1E293B !important; color: #FFFFFF !important;
        border-radius: 8px !important; border: 1px solid #334155 !important;
    }

    /* Botões Padrão em Dourado Sementes */
    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #D97706 !important; color: #FFFFFF !important;
        border: 1px solid #F59E0B !important; border-radius: 8px !important;
        font-weight: 700 !important; padding: 10px 24px !important;
        box-shadow: 0px 4px 14px rgba(217, 119, 6, 0.2);
        transition: all 0.3s ease !important; width: 100%;
    }
    .stButton>button:hover { background-color: #B45309 !important; transform: translateY(-2px); }

    /* Alertas Personalizados (Azul Profundo IEQ) */
    .stAlert { background-color: #1E1B4B !important; border: 1px solid #3730A3 !important; color: #E0E7FF !important; border-radius: 12px; }
    
    /* Links para WhatsApp Integrados */
    .whatsapp-btn {
        display: inline-flex; align-items: center; justify-content: center;
        padding: 8px 16px; color: #FFFFFF !important; background-color: #059669;
        border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 0.9rem;
        box-shadow: 0px 4px 10px rgba(5, 150, 105, 0.2);
    }
    .whatsapp-btn:hover { background-color: #047857; transform: translateY(-1px); }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

# ==============================================================================
# 2. TRATAMENTO INTELIGENTE DE DADOS DO EXCEL (BACKEND)
# ==============================================================================
def carregar_dados_cadastro():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Ficha de Cadastro", header=0, engine='openpyxl')
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

def processar_frequencia_mensal(nome_aba):
    try:
        df_raw = pd.read_excel(EXCEL_FILE, sheet_name=nome_aba, header=None, engine='openpyxl')
        
        linha_dias_idx = None
        for idx, row in df_raw.iterrows():
            row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if "1" in row_vals and "2" in row_vals and "3" in row_vals:
                linha_dias_idx = idx
                break
        
        if linha_dias_idx is None:
            return pd.DataFrame()
            
        linha_dias = df_raw.iloc[linha_dias_idx]
        col_dias_map = {}
        for col_idx, val in enumerate(linha_dias):
            val_str = str(val).strip().split('.')[0]
            if val_str.isdigit() and 1 <= int(val_str) <= 31:
                col_dias_map[col_idx] = val_str
                
        dados_frequencia = []
        for idx in range(linha_dias_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            nome_aluno = str(row[1]).strip() if pd.notna(row[1]) else ""
            
            if not nome_aluno or nome_aluno.lower() in ["nan", "none", "nome do aluno", "total", "nº"]:
                continue
            if any(term in nome_aluno.lower() for term in ["projeto sementes", "ficha de controle", "mês/ano", "dias de treino"]):
                continue
                
            total_presencas = 0
            for col_idx, _ in col_dias_map.items():
                val_pres = str(row[col_idx]).strip().upper() if pd.notna(row[col_idx]) else ""
                if val_pres == "X":
                    total_presencas += 1
                    
            dados_frequencia.append({
                "Nome do Aluno": nome_aluno,
                "Presenças Registradas": total_presencas
            })
        return pd.DataFrame(dados_frequencia)
    except Exception:
        return pd.DataFrame()

# Carregamento prévio dos dados em cache local
df_cadastro = carregar_dados_cadastro()

def formatar_data_sem_hora(df):
    for col in df.columns:
        if any(kw in str(col).lower() for kw in ["data", "nascimento", "início", "inicio"]):
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').fillna(df[col])
    return df

def gerar_link_whatsapp(numero, mensagem):
    if pd.isna(numero) or str(numero).strip() == "" or str(numero).strip().lower() == "nan":
        return None
    num_limpo = ''.join(filter(str.isdigit, str(numero)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = "55" + num_limpo
    msg_enc = urllib.parse.quote(mensagem)
    return f"https://wa.me/{num_limpo}?text={msg_enc}"

# ==============================================================================
# 3. INTERFACE DE NAVEGAÇÃO E REGRAS DE NEGÓCIO (FRONTEND)
# ==============================================================================
st.sidebar.image("logo_projeto_sementes.png", width=60)
st.sidebar.markdown("<h2 style='text-align: center; color:#F59E0B;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #94A3B8; font-size: 0.8rem;'>IEQ Guaicurus — Jiu-Jitsu Kids</p>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Menu do Aplicativo",
    [
        "🏠 Dashboard dos Pais",
        "📷 Chamada por Biometria IA",
        "📊 Mapa de Abrangência Social",
        "📢 Palestras & Graduações",
        "🌊 Batismo & Integração",
        "📄 Secretaria & Relatórios PDF"
    ]
)

# Trava de Segurança
if df_cadastro.empty:
    st.error("⚠️ ERRO CRÍTICO: Arquivo 'Controle de Presença e Graduação Projeto Sementes.xlsx' ou aba 'Ficha de Cadastro' não encontrados.")
    st.stop()

# ------------------------------------------------------------------------------
# MÓDULO 1: DASHBOARD DOS PAIS
# ------------------------------------------------------------------------------
if menu == "🏠 Dashboard dos Pais":
    st.title("🏠 Painel de Acompanhamento Familiar")
    st.info("👋 Olá, *Débora Maciel da Silva*! Seus filhos cadastrados foram identificados com segurança no sistema.")
    
    filho_selecionado = st.selectbox("Selecione o perfil do aluno para gerenciar:", ["Alvaro Barbosa", "Arthur Barbosa"])
    
    dados_filho = df_cadastro[df_cadastro["Nome do Aluno"].str.strip().str.lower() == filho_selecionado.lower()]
    
    if not dados_filho.empty:
        filho_row = dados_filho.iloc[0]
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric(label="Graduação / Faixa Atual", value=str(filho_row.get("Faixa Graus", "Branca")))
        with c2:
            st.metric(label="Status de Matrícula", value=str(filho_row.get("Pendencia Documento", "Ok")))
        with c3:
            df_ago = processar_frequencia_mensal("Agosto 2026")
            presencas = 0
            if not df_ago.empty:
                aluno_freq = df_ago[df_ago["Nome do Aluno"].str.strip().str.lower() == filho_selecionado.lower()]
                if not aluno_freq.empty:
                    presencas = aluno_freq.iloc[0]["Presenças Registradas"]
            st.metric(label="Presenças em Agosto", value=f"{presencas} Aulas")
            
        saude = str(filho_row.get("Histórico de Saúde", "Não")).strip()
        if saude.lower() != "não" and saude != "":
            st.error(f"🚨 *Atenção Médica Monitorada pelos Instrutores:* {saude}")
        else:
            st.success("🔒 *Histórico de Saúde:* Nenhuma restrição médica ativa. Aluno apto para os treinos.")

        st.markdown("### 👤 Biometria Facial Registrada")
        st.image("logo_projeto_sementes.png", caption=f"Foto Biométrica Oficial de {filho_selecionado}")

# ------------------------------------------------------------------------------
# MÓDULO 2: CHAMADA POR BIOMETRIA FACIAL IA
# ------------------------------------------------------------------------------
elif menu == "📷 Chamada por Biometria IA":
    st.title("📷 Chamada Inteligente por Visão Computacional")
    st.markdown("Tire uma foto coletiva do fim do treino no tatame para registrar as presenças de forma automática.")
    
    foto_upload = st.file_uploader("Capturar ou carregar foto do fim da aula:", type=["jpg", "png", "jpeg"])
