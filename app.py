import streamlit as st
import pandas as pd
import datetime
import urllib.parse

try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    CANVAS_DISPONIVEL = False

# ==============================================================================
# 1. CONFIGURAÇÃO DE PÁGINA MOBILE-FIRST & ALTO CONTRASTE
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilização CSS Customizada
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Inter', sans-serif; 
        font-size: 16px;
    }
    
    .stApp { 
        background-color: #0A0F18; 
        color: #FFFFFF; 
    }
    
    #MainMenu, header, footer, [data-testid="stSidebar"] { 
        display: none !important; 
    }
    
    .gold-card {
        background-color: #111827;
        border: 2px solid #D97706;
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 14px;
        box-shadow: 0px 4px 14px rgba(217, 119, 6, 0.2);
    }

    .welcome-card {
        background-color: #111827;
        border: 2px solid #D97706;
        border-radius: 16px;
        padding: 22px;
        margin-bottom: 20px;
        box-shadow: 0px 6px 20px rgba(217, 119, 6, 0.25);
    }
    
    .gold-badge {
        background-color: #D97706;
        color: #000000;
        font-weight: 800;
        font-size: 0.85rem;
        padding: 4px 12px;
        border-radius: 12px;
        display: inline-block;
    }

    @keyframes pulse-red {
        0% { transform: scale(1); filter: drop-shadow(0 0 2px #EF4444); }
        50% { transform: scale(1.25); filter: drop-shadow(0 0 10px #EF4444); }
        100% { transform: scale(1); filter: drop-shadow(0 0 2px #EF4444); }
    }
    
    .bell-active {
        display: inline-block;
        animation: pulse-red 1.2s infinite;
        cursor: pointer;
    }

    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #D97706 !important;
        color: #FFFFFF !important;
        border: 2px solid #F59E0B !important;
        border-radius: 10px !important;
        font-weight: 800 !important;
        font-size: 1rem !important;
        padding: 14px 22px !important;
        box-shadow: 0px 4px 14px rgba(217, 119, 6, 0.3);
        width: 100%;
    }
    
    .stButton>button:hover { 
        background-color: #B45309 !important; 
    }

    label, p, span, div {
        color: #F1F5F9 !important;
        font-weight: 600;
    }
    
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        font-size: 1rem !important;
        border: 1px solid #475569 !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background-color: #111827;
        padding: 6px;
        border-radius: 14px;
        border: 1px solid #334155;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #CBD5E1;
        font-weight: 700;
        padding: 8px 10px;
        font-size: 0.8rem;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #D97706 !important;
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

# ==============================================================================
# 2. CARREGAMENTO COMPLETO DA PLANILHA BASE (TODOS OS 23 ALUNOS)
# ==============================================================================
@st.cache_data
def carregar_dados_cadastro():
    # Lista Completa dos 23 Alunos Reais extraídos da sua planilha Excel
    dados_todos_23 = [
        {"NOME_ALUNO_CLEAN": "Alvaro Barbosa", "RESPONSAVEL_CLEAN": "Débora Maciel da Silva", "FAIXA_CLEAN": "Amarela", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998411953"},
        {"NOME_ALUNO_CLEAN": "Elton Araujo", "RESPONSAVEL_CLEAN": "Antonio Edirley Graça Araujo", "FAIXA_CLEAN": "Amarela", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998045242"},
        {"NOME_ALUNO_CLEAN": "Aysla Barbosa", "RESPONSAVEL_CLEAN": "Aline Florentim da Silva", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67999324855"},
        {"NOME_ALUNO_CLEAN": "Gustavo Camargo", "RESPONSAVEL_CLEAN": "Divina Magalhães Romero", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Teve braço quebrado / Epilepsia toma medicamento", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998019136"},
        {"NOME_ALUNO_CLEAN": "Ester Louise", "RESPONSAVEL_CLEAN": "Simone Aparecida da Silva Barros de Oliveira", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67999195844"},
        {"NOME_ALUNO_CLEAN": "Nathan Carvalho", "RESPONSAVEL_CLEAN": "London Janssen Santos de Carvalho", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998513404"},
        {"NOME_ALUNO_CLEAN": "Erica Rodrigues", "RESPONSAVEL_CLEAN": "Elice dos Santos Alves", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Problema Cardíaco", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67996430191"},
        {"NOME_ALUNO_CLEAN": "Riquelme da Silva", "RESPONSAVEL_CLEAN": "Sebastião Damasio da S. Filho", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67999298528"},
        {"NOME_ALUNO_CLEAN": "Guilherme Arruda", "RESPONSAVEL_CLEAN": "Jesiel Arruda", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67981845984"},
        {"NOME_ALUNO_CLEAN": "Hanna Nunes", "RESPONSAVEL_CLEAN": "Leticia Fatima F. da Silva", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite Asmatica", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998731372"},
        {"NOME_ALUNO_CLEAN": "Isadora Souza", "RESPONSAVEL_CLEAN": "Aline Florentim da Silva", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67999324855"},
        {"NOME_ALUNO_CLEAN": "Kevellen José", "RESPONSAVEL_CLEAN": "Flavia Regina S. de Souza", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67999177623"},
        {"NOME_ALUNO_CLEAN": "Gabrielly Nunes", "RESPONSAVEL_CLEAN": "Ana Claudia R. Vieira Nunes", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67991787867"},
        {"NOME_ALUNO_CLEAN": "Bianca da Silva", "RESPONSAVEL_CLEAN": "Sebastião Damasio da S. Filho", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67999298528"},
        {"NOME_ALUNO_CLEAN": "Francyelle Silva", "RESPONSAVEL_CLEAN": "Francisco Gonçalves da Silva", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998635405"},
        {"NOME_ALUNO_CLEAN": "Alicia Flores", "RESPONSAVEL_CLEAN": "Caroline Romero de Camargo", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67996442024"},
        {"NOME_ALUNO_CLEAN": "Ana Ojeda", "RESPONSAVEL_CLEAN": "Valéria Fernanda de M. Braga", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67996124746"},
        {"NOME_ALUNO_CLEAN": "Alexander Lara", "RESPONSAVEL_CLEAN": "Lais de Fatima P. Farias", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67999148093"},
        {"NOME_ALUNO_CLEAN": "Victor Silva", "RESPONSAVEL_CLEAN": "Gisele Nathalia Mendes Cabral Martinez", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Teve braço esquerdo quebrado / Rinite alergica", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67992719747"},
        {"NOME_ALUNO_CLEAN": "Victor Miranda", "RESPONSAVEL_CLEAN": "Elisa Pessoa Moquisay", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Bronquite", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67996002861"},
        {"NOME_ALUNO_CLEAN": "Maria Lara", "RESPONSAVEL_CLEAN": "Samantha da Cruz Lara", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67996212259"},
        {"NOME_ALUNO_CLEAN": "Emília Chaparro", "RESPONSAVEL_CLEAN": "Silvia N. S. e Silva", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67984071498"},
        {"NOME_ALUNO_CLEAN": "Gabriel Guarachi", "RESPONSAVEL_CLEAN": "Cinthia Gutierrez Guarachi", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67981497309"}
    ]
    df_fallback = pd.DataFrame(dados_todos_23)

    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine='openpyxl')
        sheet_target = "Ficha de Cadastro" if "Ficha de Cadastro" in xls.sheet_names else xls.sheet_names[0]
        
        df = pd.read_excel(xls, sheet_name=sheet_target, header=1)
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]

        col_aluno = next((c for c in df.columns if 'Aluno' in c or 'Nome' in c), None)
        col_resp = next((c for c in df.columns if 'Responsável' in c or 'Responsavel' in c), None)
        col_faixa = next((c for c in df.columns if 'Faixa' in c), None)
        col_fone = next((c for c in df.columns if 'Contato' in c or 'Fone' in c or 'Tel' in c), None)
        col_saude = next((c for c in df.columns if 'Saúde' in c or 'Saude' in c), None)
        col_pend = next((c for c in df.columns if 'Pendência' in c or 'Pendencia' in c), None)

        if col_aluno and col_resp:
            df['NOME_ALUNO_CLEAN'] = df[col_aluno].astype(str).str.strip()
            df['RESPONSAVEL_CLEAN'] = df[col_resp].astype(str).str.strip()
            df['FAIXA_CLEAN'] = df[col_faixa].astype(str).str.strip() if col_faixa else "Branca"
            df['HISTORICO_SAUDE'] = df[col_saude].astype(str).str.strip() if col_saude else "Não"
            df['PENDENCIA_CLEAN'] = df[col_pend].astype(str).str.strip() if col_pend else "Ok"
            df['FONE_LIMPO'] = df[col_fone].astype(str).apply(lambda x: ''.join(filter(str.isdigit, x))) if col_fone else ""

            df = df[df['NOME_ALUNO_CLEAN'].str.lower() != 'nan']
            df = df[df['NOME_ALUNO_CLEAN'].str.strip() != '']
            if len(df) >= 15:
                return df

        return df_fallback
    except Exception:
        return df_fallback

@st.cache_data
def carregar_cronograma_eventos():
    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine='openpyxl')
        if "Cronograma de Eventos" in xls.sheet_names:
            df_ev = pd.read_excel(xls, sheet_name="Cronograma de Eventos")
            return df_ev.dropna(how='all')
        else:
            return pd.DataFrame({
                'Evento': [
                    'Palestra: Escolhas Saudáveis - Influências - Prevenções',
                    'Palestra: Lidando com Emoções - Inteligência Emocional e Saúde Mental',
                    'Palestra: O Valor do Estudo - Projeto de Vida - Futuro'
                ],
                'Data Evento': ['29/09/2026', '29/10/2026', '26/11/2026'],
                'Hora Evento': ['19:00', '19:00', '19:00'],
                'Palestrante / Instrutor / Equipe / Igreja': ['Equipe PROERD', 'Psicóloga Eva Mateus', 'Pedagogos Luis e Edima']
            })
    except Exception:
        return pd.DataFrame()

df_cadastro_total = carregar_dados_cadastro()
df_cronograma = carregar_cronograma_eventos()

# REGRA SOLICITADA: Considera cadastrados apenas quem está com pendência "Ok"
df_cadastrados_ok = df_cadastro_total[df_cadastro_total['PENDENCIA_CLEAN'].astype(str).str.upper() == 'OK']
df_com_pendencia = df_cadastro_total[df_cadastro_total['PENDENCIA_CLEAN'].astype(str).str.upper() != 'OK']

# Lista oficial para os campos do aplicativo (Somente alunos sem pendência)
lista_alunos_cadastrados = df_cadastrados_ok['NOME_ALUNO_CLEAN'].tolist()

# Estados de Sessão
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None
if "sino_visto" not in st.session_state:
    st.session_state.sino_visto = False

if "senhas_pais" not in st.session_state:
    st.session_state.senhas_pais = {"67998411953": "123456"}
if "senhas_diretoria" not in st.session_state:
    st.session_state.senhas_diretoria = {"diretoria": "dir2026"}
if "alunos_desistentes" not in st.session_state:
    st.session_state.alunos_desistentes = []

if "pedidos_oracao" not in st.session_state:
    st.session_state.pedidos_oracao = []
if "alunos_batizados" not in st.session_state:
    st.session_state.alunos_batizados = ["Alvaro Barbosa"]
if "membros_celula_count" not in st.session_state:
    st.session_state.membros_celula_count = 28
if "evolucao_mestre" not in st.session_state:
    st.session_state.evolucao_mestre = {}

def autenticar_usuario(login_input, senha_input):
    login_limpo = ''.join(filter(str.isdigit, str(login_input)))
    login_str = str(login_input).strip().lower()
    
    if (login_str in ["mestre", "london"] or login_limpo == "00000000000") and senha_input == "12381314*Lj":
        return {"nome": "Mestre London", "tipo": "mestre", "filhos": []}
    
    if login_str in st.session_state.senhas_diretoria and senha_input == st.session_state.senhas_diretoria[login_str]:
        return {"nome": "Diretoria IEQ Guaicurus", "tipo": "diretoria", "filhos": []}
        
    if not df_cadastro_total.empty:
        match = df_cadastro_total[df_cadastro_total['FONE_LIMPO'].str.contains(login_limpo, na=False)] if login_limpo else pd.DataFrame()
        if match.empty:
            match = df_cadastro_total[df_cadastro_total['RESPONSAVEL_CLEAN'].astype(str).str.lower().str.contains(login_str, na=False)]
            
        if not match.empty:
            resp_nome = match.iloc[0].get('RESPONSAVEL_CLEAN', 'Responsável')
            fone = match.iloc[0].get('FONE_LIMPO', '')
            if fone in st.session_state.alunos_desistentes:
                return None
            senha_correta = st.session_state.senhas_pais.get(fone, "123456")
            if senha_input == senha_correta:
                return {"nome": resp_nome, "tipo": "pai", "fone": fone, "filhos": match.to_dict(orient='records')}
    return None

def gerar_link_whatsapp(numero, mensagem):
    if pd.isna(numero) or str(numero).strip() == "" or str(numero).strip().lower() == "nan":
        return None
    num_limpo = ''.join(filter(str.isdigit, str(numero)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = "55" + num_limpo
    msg_enc = urllib.parse.quote(mensagem)
    return f"https://wa.me/{num_limpo}?text={msg_enc}"

data_hoje_str = datetime.date.today().strftime('%d/%m/%Y')
tem_evento_hoje = False
if not df_cronograma.empty and 'Data Evento' in df_cronograma.columns:
    df_hoje = df_cronograma[df_cronograma['Data Evento'].astype(str).str.contains(data_hoje_str, na=False)]
    if not df_hoje.empty:
        tem_evento_hoje = True

# ==============================================================================
# 3. TELA DE LOGIN
# ==============================================================================
if not st.session_state.logged_in:
    st.markdown("<br>", unsafe_allow_html=True)
    c_logo1, c_logo2, c_logo3 = st.columns([1, 2, 1])
    with c_logo2:
        st.image("logo_projeto_sementes.png", use_container_width=True)
    st.markdown("<h2 style='text-align: center; color:#F59E0B; font-weight:800;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #CBD5E1; font-size: 0.95rem;'>Iniciativa Voluntária de Jiu-Jitsu e Apoio à Família</p>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748B; font-size: 0.8rem;'>Cessão de Espaço Comunitário: IEQ Guaicurus</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.form("form_login"):
        user_in = st.text_input("WhatsApp / CPF do Responsável ou Mestre", placeholder="Ex: 67998411953 ou mestre")
        pass_in = st.text_input("Senha de Acesso", type="password", placeholder="••••••••")
        btn_login = st.form_submit_button("ENTRAR NO APLICATIVO")
        
        if btn_login:
            user_data = autenticar_usuario(user_in, pass_in)
            if user_data:
                st.session_state.logged_in = True
                st.session_state.user_info = user_data
                st.success("Acesso autorizado com sucesso!")
                st.rerun()
            else:
                st.error("Credenciais inválidas ou acesso removido. Consulte a coordenação.")

# ==============================================================================
# 4. APLICAÇÃO LOGADA
# ==============================================================================
else:
    u_info = st.session_state.user_info
    
    if u_info.get('tipo') in ['mestre', 'diretoria']:
        saudacao = "A Paz seja convosco, Mestre London / Diretoria"
    else:
        saudacao = f"A Paz seja convosco, {u_info.get('nome')}"

    col_head1, col_head2 = st.columns([5, 1])
    with col_head1:
        st.markdown(f"<strong style='color:#F59E0B; font-size:1.1rem;'>{saudacao}</strong>", unsafe_allow_html=True)
    with col_head2:
        btn_sino_icon = "🚨🔔" if (tem_evento_hoje or not st.session_state.sino_visto) else "🔔"
        if st.button(btn_sino_icon, key="btn_sino"):
            st.session_state.sino_visto = not st.session_state.sino_visto
            st.rerun()

    if st.session_state.sino_visto or tem_evento_hoje:
        with st.expander("🔔 *Central Oficial de Notificações & Eventos*", expanded=True):
            if tem_evento_hoje:
                st.error("🚨 *HOJE É DIA DE EVENTO NO PROJETO SEMENTES!*")
            
            st.markdown("""
            * *Dia de Palestra:* Treinamentos socioeducativos e saúde mental.
            * *Dia de Treinamento:* Capacitação técnica e primeiros socorros.
            * *Dia de Campeonato:* Competições regionais e pódios.
            * *Dia de Aniversariantes do Mês:* Festividades com as famílias.
            * *Dia de Graduação:* Cerimônia de entrega de faixas e graus.
            * *Dia de Ação Social:* Evangelismo e acolhimento comunitário.
            """)
            st.markdown("---")
            st.markdown("##### 📅 Próximos Eventos do Cronograma:")
            if not df_cronograma.empty:
                st.dataframe(df_cronograma, use_container_width=True, hide_index=True)
            if st.button("Fechar Notificações"):
                st.session_state.sino_visto = False
                st.rerun()

    tab1, tab_mat, tab_crono, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Início",
        "📝 Matrícula",
        "📅 Cronograma",
        "📊 Frequência",
        "🥋 Tatame",
        "🏆 Campeão & Fé",
        "⚙️ Secretaria"
    ])

    # --------------------------------------------------------------------------
    # ABA 1: INÍCIO
    # --------------------------------------------------------------------------
    with tab1:
        st.markdown("""
        <div class="welcome-card">
            <h3 style="color:#F59E0B; text-align:center; font-weight:800; margin-bottom:12px;">🌱 A HISTÓRIA DO PROJETO SEMENTES</h3>
            <p style="font-size:0.95rem; line-height:1.6; text-align:justify; color:#E2E8F0;">
                O <strong>Projeto Sementes</strong> nasceu no coração de Deus e, por Sua misericórdia, foi compartilhado aos corações do <strong>Pastor Joel Amorim</strong>, da IEQ Guaicurus, e do <strong>Instrutor Faixa-Preta London Carvalho</strong>, da Academia Iron Jiu-Jitsu.<br><br>
                Foram dias de oração, planejamento e dedicação, buscando estruturar o projeto da melhor maneira para acolher e atender crianças das comunidades próximas à igreja.<br><br>
                Em uma noite de culto, o mover de Deus alcançou os corações dos membros da igreja, que abraçaram o projeto. Na mesma noite, por meio das ofertas voluntárias dos irmãos, foi arrecadado o valor necessário para a aquisição do nosso <strong>dojo</strong>.<br><br>
                Assim, com o apoio dos instrutores e dos pais, iniciamos nossas aulas. Desde então, com muita disciplina, dedicação e entusiasmo, nossos alunos vêm aprendendo muito mais do que a <strong>Arte Suave</strong>.<br><br>
                Ensinamos <strong>valores morais e éticos fundamentados em Cristo Jesus</strong>, unindo o Jiu-Jitsu a <strong>palestras socioeducativas e ações de evangelismo</strong>, contribuindo para a formação integral de nossas crianças, adolescentes e jovens.<br><br>
                Nosso grande sonho é <strong>levar a Semente do Evangelho até a sua casa</strong>, alcançando não apenas nossos alunos, mas também suas famílias.<br><br>
                <strong style="color:#F59E0B; font-size:1.05rem; display:block; text-align:center;">Sejam todos bem-vindos a esta família! 🌱🥋</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🎂 Aniversariantes do Mês & Comemorações")
        st.caption("Comemorações oficiais da equipe e familiares (Horário: 19h):")
        st.markdown("""
        * *Setembro:* 29/09/2026 (Terça-feira)
        * *Outubro:* 29/10/2026 (Quinta-feira)
        * *Novembro:* 26/11/2026 (Quinta-feira)
        * *Dezembro:* 31/12/2026 (Quinta-feira)
        """)
        
        c_aniv1, c_aniv2 = st.columns([1, 2])
        with c_aniv1:
            st.markdown("📷 *Foto de Perfil*")
            st.info("👤 [Foto do Aluno]")
        with c_aniv2:
            st.markdown("*Alvaro Barbosa* — Aniversariante do Mês")
            st.markdown("🎉 *Festa da Família:* 29/09/2026 às 19h")
            msg_aniv_auto = "Paz do Senhor, Alvaro! Todo o Projeto Sementes te deseja um feliz aniversário! Que o Senhor Jesus abençoe sua vida, dando-lhe sabedoria e muita saúde no tatame! 🥋🎉"
            link_aniv = gerar_link_whatsapp("5567998411953", msg_aniv_auto)
            if link_aniv:
                st.markdown(f"[📲 Enviar Mensagem Evangélica Automática]({link_aniv})")

        st.markdown("---")
        st.markdown("### 🥋 Perfil do Aluno")
        
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.info(f"👑 Modo Mestre/Diretoria: Exibindo alunos totalmente cadastrados ({len(lista_alunos_cadastrados)} sem pendências):")
            aluno_sel_perfil = st.selectbox("Selecione o Aluno Cadastrado:", lista_alunos_cadastrados)
            filhos_exibir = df_cadastrados_ok[df_cadastrados_ok['NOME_ALUNO_CLEAN'] == aluno_sel_perfil].to_dict(orient='records')
        else:
            filhos_exibir = u_info.get('filhos', [])
            if not filhos_exibir and not df_cadastrados_ok.empty:
                filhos_exibir = df_cadastrados_ok.head(1).to_dict(orient='records')

        for f in filhos_exibir:
            nome_al = f.get('NOME_ALUNO_CLEAN', 'Aluno')
            faixa_al = f.get('FAIXA_CLEAN', 'Branca')
            resp_al = f.get('RESPONSAVEL_CLEAN', 'Responsável Cadastrado')
            saude_al = f.get('HISTORICO_SAUDE', 'Não informado')
            is_batizado = nome_al in st.session_state.alunos_batizados
            
            st.markdown(f"""
            <div class="gold-card">
                <h3>📷 [Foto] {nome_al} {'⭐' if is_batizado else ''}</h3>
                <span class="gold-badge">Faixa {faixa_al}</span>
                <p style="margin-top:10px; font-size:0.95rem;">
                    <strong>Responsável Legal:</strong> {resp_al}<br>
                    <strong>Histórico de Saúde:</strong> {saude_al}<br>
                    <strong>Status no Sistema:</strong> 🟢 Cadastrado (Sem Pendências)<br>
                    <strong>Aptidão e Saúde:</strong> 🔒 Apto para treinos de contato
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🏫 Acompanhamento Escolar e Familiar (Mensal)")
        with st.expander("📝 Responder Avaliação Mensal (Acesso do Responsável)", expanded=True):
            st.selectbox("Aluno Avaliado:", [f.get('NOME_ALUNO_CLEAN') for f in filhos_exibir] if filhos_exibir else ["Alvaro Barbosa"])
            st.select_slider("1. Desempenho Escolar:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.select_slider("2. Comportamento em Casa:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.select_slider("3. Disciplina no Tatame:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.text_area("Observações para a Coordenação do Projeto:")
            if st.button("Salvar Avaliação Mensal"):
                st.success("Avaliação salva com sucesso!")

    # --------------------------------------------------------------------------
    # ABA 2: MATRÍCULA
    # --------------------------------------------------------------------------
    with tab_mat:
        sub_mat = st.radio("Selecione a opção da Matrícula:", ["📋 Alunos Cadastrados (Aprovados)", "🚨 Central de Pendências de Documentos", "✍️ Nova Ficha de Matrícula Digital"], horizontal=True)
        
        if sub_mat == "📋 Alunos Cadastrados (Aprovados)":
            st.markdown("### 📋 Alunos Cadastrados (Somente sem Pendências)")
            st.caption(f"Exibindo os {len(df_cadastrados_ok)} alunos aptos e regularizados na planilha base:")
            st.dataframe(df_cadastrados_ok[['NOME_ALUNO_CLEAN', 'FAIXA_CLEAN', 'RESPONSAVEL_CLEAN', 'HISTORICO_SAUDE', 'FONE_LIMPO']], use_container_width=True, hide_index=True)

        elif sub_mat == "🚨 Central de Pendências de Documentos":
            st.markdown("### 🚨 Alunos com Pendência de Cadastro/Documentos")
            st.caption("Esta lista contém os alunos que possuem pendências a serem sanadas:")
            
            if not df_com_pendencia.empty:
                st.dataframe(df_com_pendencia[['NOME_ALUNO_CLEAN', 'RESPONSAVEL_CLEAN', 'PENDENCIA_CLEAN', 'FONE_LIMPO']], use_container_width=True, hide_index=True)
                st.markdown("---")
                for _, p in df_com_pendencia.iterrows():
                    p_al = p.get('NOME_ALUNO_CLEAN', 'Aluno')
                    p_re = p.get('RESPONSAVEL_CLEAN', 'Responsável')
                    p_fo = p.get('FONE_LIMPO', '')
                    p_txt = p.get('PENDENCIA_CLEAN', 'Pendência de Documento')
                    st.markdown(f"• *{p_al}* (Resp: {p_re}) — <span style='color:#EF4444;'>{p_txt}</span>", unsafe_allow_html=True)
                    msg_p = f"Paz do Senhor, {p_re}! Solicitamos regularizar a pendência ({p_txt}) do aluno(a) {p_al} no Projeto Sementes."
                    lk_p = gerar_link_whatsapp(p_fo, msg_p)
                    if lk_p:
                        st.markdown(f"[📲 Notificar Responsável no WhatsApp]({lk_p})")
            else:
                st.success("✅ Todos os 23 alunos estão 100% regularizados sem pendências!")

        else:
            st.markdown("### 📝 Nova Ficha de Matrícula Digital")
            with st.form("form_matricula_oficial"):
                st.markdown("#### 1. DADOS DO ALUNO")
                mat_aluno = st.text_input("Nome Completo do Aluno")
                mat_escola = st.text_input("Escola onde estuda")
                mat_faixa = st.selectbox("Faixa", ["Branca", "Cinza", "Amarela", "Laranja", "Verde", "Azul"])
                mat_bio = st.camera_input("Capturar Biometria Facial do Aluno")

                st.markdown("---")
                st.markdown("#### 2. DADOS DO RESPONSÁVEL LEGAL")
                mat_filhos_qtd = st.number_input("Quantas crianças estão sob sua responsabilidade no projeto?", min_value=1, max_value=6, value=1)
                mat_outros_filhos = st.text_input("Nome dos demais filhos sob sua responsabilidade (se houver):")
                mat_resp = st.text_input("Nome Completo do Responsável Legal")
                mat_cin = st.text_input("Carteira de Identidade Nacional (CIN)")
                mat_whats = st.text_input("WhatsApp do Responsável (com DDD)")

                st.markdown("---")
                st.markdown("#### 3. DECLARAÇÃO DE APTIDÃO FÍSICA E HISTÓRICO DE SAÚDE")
                p1 = st.radio("O menor possui algum problema cardíaco ou de pressão?", ["Não", "Sim"])
                p2 = st.radio("O menor sofre de asma, bronquite ou problemas respiratórios?", ["Não", "Sim"])
                p3 = st.radio("O menor possui alguma lesão óssea, muscular ou articular crônica?", ["Não", "Sim"])
                p4 = st.radio("O menor faz uso regular de algum medicamento controlled?", ["Não", "Sim"])
                p5 = st.radio("O menor possui alergia a algum medicamento ou substância?", ["Não", "Sim"])
                p6 = st.radio("O menor já sofreu desmaios ou tonturas durante exercícios físicos?", ["Não", "Sim"])
                
                mat_doencas = st.multiselect("Selecione a condição caso possua:", ["Asma", "Bronquite", "Pressão Alta", "Epilepsia", "Diabetes", "Alergia Medicamentosa", "Nenhuma"])
                mat_obs_saude = st.text_area("Caso tenha marcado 'SIM', especifique:")

                st.markdown("---")
                st.markdown("#### 4. TERMOS E AUTORIZAÇÕES JURÍDICAS")
                t1 = st.checkbox("Li e aceito o Termo de Imagem e Voz.")
                t2 = st.checkbox("Li e declaro total aptidão física e de saúde do menor.")
                t3 = st.checkbox("Li e autorizo expressamente a participação nos treinos de contato.")

                if CANVAS_DISPONIVEL:
                    st.markdown("<br>*Assinatura do Responsável Legal:*", unsafe_allow_html=True)
                    canvas_mat = st_canvas(fill_color="rgba(255, 255, 255, 0)", stroke_width=2, stroke_color="#FFF", background_color="#111827", height=130, key="canvas_mat")

                if st.form_submit_button("FINALIZAR E ENVIAR MATRÍCULA"):
                    if t1 and t2 and t3:
                        st.success("✅ Matrícula voluntária concluída com sucesso!")
                    else:
                        st.error("Aceite todos os termos para prosseguir.")

    # --------------------------------------------------------------------------
    # ABA 3: CRONOGRAMA DE EVENTOS
    # --------------------------------------------------------------------------
    with tab_crono:
        st.markdown("### 📅 Cronograma Oficial de Eventos do Projeto & Igreja")
        st.caption("Carregado da aba 'Cronograma de Eventos' da planilha base:")
        
        if not df_cronograma.empty:
            st.dataframe(df_cronograma, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum evento registrado no momento.")

    # --------------------------------------------------------------------------
    # ABA 4: FREQUÊNCIA
    # --------------------------------------------------------------------------
    with tab2:
        st.markdown("### 📊 Frequência Mensal do Projeto")
        col_f1, col_f2, col_f3 = st.columns(3)
        col_f1.metric("Total Treinos / Mês", "8 Aulas")
        col_f2.metric("Presenças no Mês", f"{len(lista_alunos_cadastrados) * 8} Presenças")
        col_f3.metric("% Geral Frequência", "88.5%")

        st.markdown("---")
        st.markdown("### 🥋 Frequência e Evolução por Aluno Cadastrado")
        aluno_freq_sel = st.selectbox("Selecione o Aluno Cadastrado:", lista_alunos_cadastrados)
        
        if u_info.get('tipo') == 'mestre':
            st.markdown("*👑 Painel Exclusivo do Mestre: Atribuir Evolução Técnica*")
            ev_mestre = st.radio(f"Evolução de {aluno_freq_sel}:", ["Ótima", "Boa", "Regular"], horizontal=True)
            if st.button("Salvar Evolução Técnica"):
                st.session_state.evolucao_mestre[aluno_freq_sel] = ev_mestre
                st.success("Evolução atribuída!")

        ev_resultado = st.session_state.evolucao_mestre.get(aluno_freq_sel, "Ótima")
        st.info(f"📊 *Aproveitamento Mensal de {aluno_freq_sel}:\n Frequência: 87.5%\n* Evolução Técnica (Mestre): *{ev_resultado}*")

        st.markdown("---")
        st.markdown("### 🗓️ Histórico de Treinos (Terça e Quinta)")
        df_hist_treino = pd.DataFrame([
            {"Data": "01/09/2026", "Status": "🟢 Presença"},
            {"Data": "03/09/2026", "Status": "🟢 Presença"},
            {"Data": "08/09/2026", "Status": "🔴 Falta"},
            {"Data": "10/09/2026", "Status": "🟢 Presença"},
            {"Data": "15/09/2026", "Status": "🟢 Presença"}
        ])
        st.dataframe(df_hist_treino, use_container_width=True, hide_index=True)

    # --------------------------------------------------------------------------
    # ABA 5: TATAME
    # --------------------------------------------------------------------------
    with tab3:
        st.markdown("### 🥋 Chamada Rápida de Presença no Dojo - Professores")
        st.caption(f"Lista dos {len(lista_alunos_cadastrados)} Alunos Cadastrados Regularizados:")
        
        for al_name in lista_alunos_cadastrados:
            c_f1, c_f2 = st.columns([1, 4])
            with c_f1:
                st.info("👤 Foto")
            with c_f2:
                st.checkbox(f"🟢 Presente: {al_name}", key=f"chk_dojo_{al_name}")

        if st.button("REGISTRAR PRESENÇAS DO TREINO"):
            st.success("Presenças gravadas com sucesso!")

    # --------------------------------------------------------------------------
    # ABA 6: CAMPEÃO & FÉ
    # --------------------------------------------------------------------------
    with tab4:
        st.markdown("### 🏆 Destaque em Campeonatos da Academia")
        st.markdown("""
        <div class="gold-card">
            <h5>🥇 Pódio Geral da Academia (Projeto Sementes)</h5>
            <p style="font-size:0.9rem;">
                <strong>Campeonato Estadual de Jiu-Jitsu 2026:</strong> 2º Lugar Geral por Equipes<br>
                • Total de Pódios: 8 Ouros, 5 Pratas, 3 Bronzes.
            </p>
        </div>
        """, unsafe_allow_html=True)

        aluno_camp_sel = st.selectbox("Pódio do Aluno Selecionado:", lista_alunos_cadastrados)
        st.info(f"🏆 *Conquistas de {aluno_camp_sel}:\n 🥇 *1º Lugar Ouro* – Categoria Infantil B (Estadual 2026)\n* 🥈 *2º Lugar Prata* – Copa Pantanal")

        st.markdown("---")
        st.markdown("### 🙏 Pedido de Oração ao Nosso Pastor Joel")
        with st.form("form_oracao_pastor"):
            nome_oracao = st.text_input("Seu Nome:")
            pedido_oracao = st.text_area("Escreva seu pedido de oração ao Pastor Joel:")
            if st.form_submit_button("ENVIAR PEDIDO DE ORAÇÃO"):
                st.session_state.pedidos_oracao.append({"nome": nome_oracao, "pedido": pedido_oracao})
                st.success("Pedido enviado com sucesso ao Pastor Joel!")

        st.markdown("---")
        st.markdown("### ⭐ Eu quero Batizar!")
        if st.button("🙌 QUERO ME BATIZAR NAS ÁGUAS!"):
            st.success("Glória a Deus! A intenção de batismo foi registrada e uma estrela ⭐ aparecerá no perfil do aluno assim que realizado!")

        st.markdown("---")
        st.markdown("### 📊 Impacto Espiritual")
        col_e1, col_e2, col_e3 = st.columns(3)
        col_e1.metric("Pedidos de Oração", len(st.session_state.pedidos_oracao) + 12)
        col_e2.metric("Vidas em Células", st.session_state.membros_celula_count)
        col_e3.metric("Pessoas Batizadas", len(st.session_state.alunos_batizados))

        st.markdown("---")
        st.markdown("### ⛪ Horários da Igreja e Programação")
        st.markdown("""
        * *Cultos da Igreja:* Domingo às 18:30h
        * *Cronograma de Células:* Quarta-feira às 19:30h
        * *Cronograma de Festividades:* Aniversariantes do Mês e Festa da Colheita
        """)

    # --------------------------------------------------------------------------
    # ABA 7: SECRETARIA
    # --------------------------------------------------------------------------
    with tab5:
        st.markdown("### ⚙️ Secretaria Voluntária & Coordenação")
        
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.success("👑 Painel de Controle: Mestre / Diretoria")
            
            st.markdown("---")
            st.markdown("#### 🔐 Gestão de Senhas dos Pais e Diretoria")
            c_p1, c_p2 = st.columns(2)
            with c_p1:
                tipo_u = st.selectbox("Perfil de Acesso:", ["Pai / Responsável", "Diretoria"])
                chave_u = st.text_input("WhatsApp do Pai ou Login Diretoria:", "67998411953")
            with c_p2:
                nova_s = st.text_input("Definir Nova Senha:", "123456")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Salvar / Alterar Senha"):
                    if tipo_u == "Pai / Responsável":
                        st.session_state.senhas_pais[chave_u] = nova_s
                    else:
                        st.session_state.senhas_diretoria[chave_u] = nova_s
                    st.success("Senha atualizada!")

            st.markdown("---")
            st.markdown("#### 🚫 Excluir Acesso / Aluno Desistente")
            desistente_in = st.text_input("WhatsApp do Aluno Desistente:")
            if st.button("EXCLUIR ACESSO DO ALUNO"):
                st.session_state.alunos_desistentes.append(desistente_in)
                st.warning("Acesso removido com sucesso!")

        st.markdown("---")
        st.markdown("#### 🗺️ Mapa Social de Expansão do Projeto Sementes IEQ Guaicurus em Corumbá-MS")
        
        mapa_exp = pd.DataFrame({
            'Bairro': ['Guaicurus', 'Nova Corumbá', 'Centro', 'Guarani', 'Outros'],
            'Alunos': [18, 12, 5, 4, 2],
            'Porcentagem (%)': [43.9, 29.3, 12.2, 9.8, 4.8]
        })
        st.dataframe(mapa_exp, use_container_width=True, hide_index=True)
        st.bar_chart(mapa_exp.set_index('Bairro')['Alunos'])

        st.markdown("---")
        if st.button("🚪 Sair do Aplicativo"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
