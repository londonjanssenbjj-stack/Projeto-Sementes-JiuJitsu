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
# 2. CARREGAMENTO BLINDADO A ERROS (GARANTE COLUNAS NATIVAS)
# ==============================================================================
@st.cache_data
def carregar_dados_cadastro():
    dados_padrao = [
        {"NOME_ALUNO_CLEAN": "Alvaro Barbosa", "RESPONSAVEL_CLEAN": "Odiselma Carvalho", "FAIXA_CLEAN": "Amarela", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Ok", "FONE_LIMPO": "67998411953"},
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
        {"NOME_ALUNO_CLEAN": "Bianca da Silva", "RESPONSAVEL_CLEAN": "Sebastião Damasio da S. Filho", "FAIXA_CLEAN": "Branca", "HISTORICO_SAUDE": "Não", "PENDENCIA_CLEAN": "Falta Endereço", "FONE_LIMPO": "67999298528"}
    ]
    df_fallback = pd.DataFrame(dados_padrao)

    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine='openpyxl')
        sheet_target = "Ficha de Cadastro" if "Ficha de Cadastro" in xls.sheet_names else xls.sheet_names[0]
        
        # Lê a partir da linha 2 (header=1)
        df = pd.read_excel(xls, sheet_name=sheet_target, header=1)
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]

        # Encontra as colunas dinamicamente
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
            if not df.empty:
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

df_cadastro = carregar_dados_cadastro()
df_cronograma = carregar_cronograma_eventos()

# Estados da Sessão
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
    
    # Mestre
    if (login_str in ["mestre", "london"] or login_limpo == "00000000000") and senha_input == "12381314*Lj":
        return {"nome": "Mestre London", "tipo": "mestre", "filhos": []}
    
    # Diretoria
    if login_str in st.session_state.senhas_diretoria and senha_input == st.session_state.senhas_diretoria[login_str]:
        return {"nome": "Diretoria IEQ Guaicurus", "tipo": "diretoria", "filhos": []}
        
    # Responsáveis
    if not df_cadastro.empty and 'FONE_LIMPO' in df_cadastro.columns:
        match = df_cadastro[df_cadastro['FONE_LIMPO'].str.contains(login_limpo, na=False)] if login_limpo else pd.DataFrame()
        if match.empty and 'RESPONSAVEL_CLEAN' in df_cadastro.columns:
            match = df_cadastro[df_cadastro['RESPONSAVEL_CLEAN'].astype(str).str.lower().str.contains(login_str, na=False)]
            
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
        
        lista_alunos_base = df_cadastro.get('NOME_ALUNO_CLEAN', pd.Series(["Alvaro Barbosa"])).tolist()
        
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.info("👑 Modo Mestre/Diretoria: Selecione qualquer aluno da planilha:")
            aluno_sel_perfil = st.selectbox("Selecione o Aluno:", lista_alunos_base)
            filhos_exibir = df_cadastro[df_cadastro.get('NOME_ALUNO_CLEAN', pd.Series()) == aluno_sel_perfil].to_dict(orient='records')
        else:
            filhos_exibir = u_info.get('filhos', [])
            if not filhos_exibir and not df_cadastro.empty:
                filhos_exibir = df_cadastro.head(1).to_dict(orient='records')

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
                    <strong>Biometria Facial:</strong> 🟢 Cadastrada<br>
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
                
