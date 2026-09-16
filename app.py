import streamlit as st
import pandas as pd
import datetime
import urllib.parse

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E DESIGN PREMIUM (DARK MODE, AZUL PROFUNDO & DOURADO)
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS Customizada para Alta Fidelidade com as especificações
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Fundo Geral Azul-marinho quase preto */
    .stApp { background-color: #0A0F18; color: #E2E8F0; }
    [data-testid="stSidebar"] { background-color: #111827; border-right: 1px solid #1E293B; }
    
    /* Títulos e Destaques */
    h1, h2, h3, h4, .gold-text { color: #F59E0B !important; font-weight: 700; }
    
    /* Customização dos Cards de Métricas e Conteineres */
    div[data-testid="stMetric"] {
        background-color: #111827;
        border: 1px solid #D97706;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0px 4px 15px rgba(217, 119, 6, 0.15);
    }
    div[data-testid="stMetricLabel"] { color: #94A3B8 !important; font-weight: 600; text-transform: uppercase; font-size: 0.8rem; }
    div[data-testid="stMetricValue"] { color: #F59E0B !important; font-weight: 800; font-size: 1.8rem; }

    /* Inputs, Selectboxes e Campos */
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea {
        background-color: #1E293B !important; color: #FFFFFF !important;
        border-radius: 8px !important; border: 1px solid #D97706 !important;
    }

    /* Botões Dourados Metálicos */
    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #D97706 !important; color: #FFFFFF !important;
        border: 1px solid #F59E0B !important; border-radius: 8px !important;
        font-weight: 700 !important; padding: 10px 24px !important;
        box-shadow: 0px 4px 14px rgba(217, 119, 6, 0.3);
        transition: all 0.3s ease !important; width: 100%;
    }
    .stButton>button:hover { background-color: #B45309 !important; transform: translateY(-2px); }

    /* Cards e Alertas Personalizados */
    .stAlert { background-color: #1E1B4B !important; border: 1px solid #3730A3 !important; color: #E0E7FF !important; border-radius: 12px; }
    
    .gold-card {
        background-color: #111827;
        border: 1px solid #D97706;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0px 4px 12px rgba(217, 119, 6, 0.1);
    }
    
    .gold-badge {
        background-color: #D97706;
        color: #000000;
        font-weight: bold;
        padding: 4px 12px;
        border-radius: 20px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

# ==============================================================================
# 2. CARREGAMENTO E MANIPULAÇÃO DE DADOS
# ==============================================================================
@st.cache_data
def carregar_dados_cadastro():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Ficha de Cadastro", header=0, engine='openpyxl')
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

df_cadastro = carregar_dados_cadastro()

def gerar_link_whatsapp(numero, mensagem):
    if pd.isna(numero) or str(numero).strip() == "" or str(numero).strip().lower() == "nan":
        return None
    num_limpo = ''.join(filter(str.isdigit, str(numero)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = "55" + num_limpo
    msg_enc = urllib.parse.quote(mensagem)
    return f"https://wa.me/{num_limpo}?text={msg_enc}"

# ==============================================================================
# 3. MENU DE NAVEGAÇÃO LATERAL (SIDEBAR)
# ==============================================================================
st.sidebar.image("logo_projeto_sementes.png", width=80)
st.sidebar.markdown("<h2 style='text-align: center; color:#F59E0B;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #94A3B8; font-size: 0.8rem;'>Cultivando valores, fortalecendo famílias</p>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Navegação nos Painéis:",
    [
        "📱 1. Login",
        "👤 2. Primeiro Acesso & Validação",
        "🏠 3. Dashboard dos Pais",
        "📅 4. Histórico de Frequência",
        "🥋 5. Painel do Professor (Chamada)",
        "📊 6. Secretaria & Mapa Social",
        "📢 7. Palestras Socioeducativas",
        "🌊 8. Batismo & Integração Espiritual"
    ]
)

# ==============================================================================
# 📱 1. PAINEL: TELA DE LOGIN E BOAS-VINDAS
# ==============================================================================
if menu == "📱 1. Login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.image("logo_projeto_sementes.png", width=120)
        st.markdown("<h2 style='text-align: center;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #94A3B8;'>Cultivando valores, fortalecendo famílias</p>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        cpf = st.text_input("CPF do Responsável", placeholder="000.000.000-00")
        senha = st.text_input("Senha", type="password", placeholder="••••••••")
        
        if st.button("ENTRAR"):
            st.success("Login efetuado com sucesso! Redirecionando...")
            
        st.markdown("<br><br><p style='text-align: center; color: #64748B; font-size: 0.8rem;'>Apoio: Igreja do Evangelho Quadrangular – Guaicurus</p>", unsafe_allow_html=True)

# ==============================================================================
# 👤 2. PAINEL: TELA DE PRIMEIRO ACESSO E VALIDAÇÃO
# ==============================================================================
elif menu == "👤 2. Primeiro Acesso & Validação":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## Primeiro Acesso: Confirme seus dados")
        st.caption("Cadastre sua senha segura e confirme os dependentes vinculados.")
        
        nova_senha = st.text_input("Nova Senha Segura", type="password")
        confirma_senha = st.text_input("Confirmar Nova Senha", type="password")
        
        c_whats1, c_whats2 = st.columns([3, 1])
        with c_whats1:
            whatsapp = st.text_input("WhatsApp Cadastrado", value="(67) 99999-0000")
        with c_whats2:
            st.markdown("<br>", unsafe_allow_html=True)
            st.button("Atualizar")
            
        st.markdown("<div class='gold-card'>", unsafe_allow_html=True)
        st.markdown("### 👨‍👩‍👧‍👦 Filhos Vinculados ao CPF")
        st.markdown("✅ **Alvaro Barbosa** — <span style='color:#F59E0B;'>Faixa Amarela</span>", unsafe_allow_html=True)
        st.markdown("✅ **Arthur Barbosa** — <span style='color:#FFFFFF;'>Faixa Branca</span>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.button("CONCLUIR E ENTRAR"):
            st.success("Dados validados com sucesso!")

# ==============================================================================
# 🏠 3. PAINEL: DASHBOARD PRINCIPAL DOS PAIS
# ==============================================================================
elif menu == "🏠 3. Dashboard dos Pais":
    st.markdown("## Olá, Débora! 👋")
    st.info("🚨 **Alerta de Documentação:** Identificamos pendência na entrega da Ficha Médica/Termo de Autorização. Favor regularizar na secretaria.")
    
    st.markdown("### 🥋 Perfil dos Alunos")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='gold-card' style='border-color: #F59E0B;'>", unsafe_allow_html=True)
        st.image("logo_projeto_sementes.png", width=80)
        st.markdown("#### Alvaro Barbosa")
        st.markdown("<span class='gold-badge'>Faixa Amarela</span>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='gold-card' style='border-color: #FFFFFF;'>", unsafe_allow_html=True)
        st.image("logo_projeto_sementes.png", width=80)
        st.markdown("#### Arthur Barbosa")
        st.markdown("<span class='gold-badge' style='background-color:#E2E8F0;'>Faixa Branca</span>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    m1, m2 = st.columns(2)
    with m1:
        st.metric(label="Resumo de Frequência", value="5 Presenças", delta="Agosto 2026")
    with m2:
        st.markdown("<div class='gold-card'>", unsafe_allow_html=True)
        st.markdown("### 🛡️ Saúde Monitorada")
        st.write("Sem restrições graves cadastradas. Apto para atividades de combate.")
        st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# 📅 4. PAINEL: HISTÓRICO COMPLETO DE FREQUÊNCIA
# ==============================================================================
elif menu == "📅 4. Histórico de Frequência":
    st.markdown("## 📅 Frequência & Rotina de Treino")
    
    c_m1, c_m2, c_m3 = st.columns([2, 1, 1])
    with c_m1:
        st.markdown("### <b>&lt; Agosto 2026 &gt;</b>", unsafe_allow_html=True)
    with c_m2:
        st.metric("Presenças", "5 DIIAS")
    with c_m3:
        st.metric("Aulas Realizadas", "9 AULAS")
        
    st.markdown("### Calendário Mensal (Terças e Quintas)")
    
    grid_cal = [
        {"dia": "04/08", "status": "🟡 Presença (X)", "treino": "Devocional: Disciplina / Raspagem de Guarda"},
        {"dia": "06/08", "status": "🟡 Presença (X)", "treino": "Devocional: Respeito / Passagem de Guarda"},
        {"dia": "11/08", "status": "🔴 Falta", "treino": "Treino Tático / Quedas"},
        {"dia": "13/08", "status": "🟡 Presença (X)", "treino": "Devocional: Honestidade / Armlock"},
        {"dia": "18/08", "status": "🟡 Presença (X)", "treino": "Devocional: Amor / Defesa Pessoal"},
        {"dia": "20/08", "status": "🟡 Presença (X)", "treino": "Simulado de Campeonato"},
    ]
    
    df_cal = pd.DataFrame(grid_cal)
    st.dataframe(df_cal, use_container_width=True)
    
    st.markdown("### 🥋 Detalhes do Treino Selecionado")
    st.markdown("""
    - 📖 **Devocional Bíblico:** Fruto do Espírito - Domínio Próprio.
    - 🏃 **Aquecimento:** Polichinelos, rolamentos e fuga de quadril.
    - 🥋 **Técnica de Jiu-Jítsu:** Queda de quadril (Ippon Seoi Nage) + Controle Lateral.
    """)

# ==============================================================================
# 🥋 5. PAINEL: INTERFACE DO PROFESSOR (CHAMADA RÁPIDA)
# ==============================================================================
elif menu == "🥋 5. Painel do Professor (Chamada)":
    st.markdown("## 🥋 Chamada Rápida — Instrutor: London")
    
    f1, f2, f3 = st.columns(3)
    with f1: st.button("Todos")
    with f2: st.button("Faixa Branca")
    with f3: st.button("Faixa Amarela")
    
    st.markdown("---")
    
    alunos_lista = [
        {"nome": "Alvaro Barbosa", "faixa": "Amarela", "alerta": "Nenhum"},
        {"nome": "Arthur Barbosa", "faixa": "Branca", "alerta": "Nenhum"},
        {"nome": "Bernardo Silva", "faixa": "Branca", "alerta": "🚨 Asma (Ativador na mochila)"},
        {"nome": "Gabriel Souza", "faixa": "Cinza", "alerta": "Nenhum"},
    ]
    
    for aluno in alunos_lista:
        ca1, ca2, ca3 = st.columns([3, 2, 1])
        with ca1:
            st.markdown(f"**{aluno['nome']}** ({aluno['faixa']})")
            if aluno['alerta'] != "Nenhum":
                st.caption(f"<span style='color:#EF4444;'>{aluno['alerta']}</span>", unsafe_allow_html=True)
        with ca2:
            st.checkbox("Presença", key=aluno['nome'])
        with ca3:
            st.markdown("🟡")
            
    st.markdown("---")
    b1, b2 = st.columns(2)
    with b1:
        if st.button("📖 Devocional do Dia"):
            st.info("Tema de Hoje: 'A importância da obediência aos pais e mestres'.")
    with b2:
        if st.button("Finalizar Chamada"):
            st.success("Chamada salva e registrada no banco de dados!")

# ==============================================================================
# 📊 6. PAINEL: SECRETARIA E MAPA DE ABRANGÊNCIA SOCIAL
# ==============================================================================
elif menu == "📊 6. Secretaria & Mapa Social":
    st.markdown("## 📊 Painel Administrativo da Secretaria")
    
    btn_pdf = st.button("📄 Emitir & Baixar Relatório Oficial (PDF)")
    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### 🗺️ Mapa Social de Abrangência (Corumbá-MS)")
        mapa_data = pd.DataFrame({
            'Bairro': ['Guaicurus', 'Nova Corumbá', 'Centro', 'Guarani'],
            'Alunos (%)': [45, 30, 15, 10]
        })
        st.bar_chart(mapa_data.set_index('Bairro'))
        
    with col_right:
        st.markdown("### 📋 Pendências de Cadastro")
        pendencias = [
            {"Nome": "Arthur Barbosa", "Pendente": "Falta Documento", "Fone": "5567999990000"},
            {"Nome": "Lucas Mendes", "Pendente": "Falta Endereço", "Fone": "5567999991111"}
        ]
        
        for p in pendencias:
            st.markdown(f"**{p['Nome']}** — <span style='color:#EF4444;'>{p['Pendente']}</span>", unsafe_allow_html=True)
            link_wsp = gerar_link_whatsapp(p['Fone'], f"Olá! Notamos uma pendência de ({p['Pendente']}) do aluno {p['Nome']} no Projeto Sementes.")
            if link_wsp:
                st.markdown(f"[💬 Cobrar via WhatsApp]({link_wsp})")
            st.markdown("---")

# ==============================================================================
# 📢 7. PAINEL: CRONOGRAMA DE PALESTRAS SOCIOEDUCATIVAS
# ==============================================================================
elif menu == "📢 7. Palestras Socioeducativas":
    st.markdown("## 📢 Cronograma de Eventos e Palestras")
    
    aba1, aba2 = st.tabs(["Palestras Socioeducativas", "Exames de Faixa"])
    
    with aba1:
        st.markdown("<div class='gold-card'>", unsafe_allow_html=True)
        st.markdown("### 🧠 Inteligência Emocional e Combate ao Bullying")
        st.write("📅 **Data:** 25/09/2026 — 19:00")
        st.write("👤 **Palestrante Convidado:** Dr. Marco Aurélio (Psicólogo Infantil)")
        st.write("📍 **Local:** Salão Principal — IEQ Guaicurus")
        
        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            if st.button("Sim, Confirmar Presença", key="sim1"):
                st.success("Presença Confirmada!")
        with btn_c2:
            st.button("Não Poderei Ir", key="nao1")
        st.markdown("</div>", unsafe_allow_html=True)

    with aba2:
        st.markdown("<div class='gold-card'>", unsafe_allow_html=True)
        st.markdown("### 🥋 Graduação e Exame de Faixa - 2º Semestre")
        st.write("📅 **Data:** 15/12/2026 — 18:30")
        st.write("📍 **Local:** Tatame Central do Projeto Sementes")
        st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# 🌊 8. PAINEL: BATISMO E INTEGRAÇÃO ESPIRITUAL
# ==============================================================================
elif menu == "🌊 8. Batismo & Integração Espiritual":
    st.markdown("## 🌊 Acompanhamento Eclesiástico & Discipulado")
    
    st.markdown("<div class='gold-card'>", unsafe_allow_html=True)
    st.markdown("### ✝️ Marcos de Fé do Aluno")
    st.markdown("🌊 <span class='gold-badge'>Batizado nas Águas em: 14/06/2026</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("### 🏡 Célula e Integração Familiar")
    st.info("<b>Célula Familiar Ativa:</b> Célula Filhos da Promessa<br><b>Líder Responsável:</b> Diácono Roberto & Mary — IEQ Guaicurus", icon="ℹ️")
    
    if st.button("💬 Quero Visitar uma Célula"):
        link_celula = gerar_link_whatsapp("5567999990000", "Olá! Gostaria de saber mais informações para visitar uma Célula Familiar da IEQ Guaicurus.")
        if link_celula:
            st.markdown(f"[Clique aqui para entrar em contato com a equipe de Recepção]({link_celula})")
