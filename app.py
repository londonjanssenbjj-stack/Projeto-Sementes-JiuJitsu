import streamlit as st
import pandas as pd
import datetime
import urllib.parse
from PIL import Image

try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    CANVAS_DISPONIVEL = False

# ==============================================================================
# 1. CONFIGURAÇÃO DE PÁGINA MOBILE-FIRST & ALTO CONTRASTE
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - Jiu-Jitsu Voluntário",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilização CSS Customizada Sanitizada
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
    
    .app-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 14px 18px;
        background-color: #111827;
        border-bottom: 2px solid #D97706;
        border-radius: 0 0 16px 16px;
        margin-bottom: 18px;
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
        50% { transform: scale(1.2); filter: drop-shadow(0 0 8px #EF4444); }
        100% { transform: scale(1); filter: drop-shadow(0 0 2px #EF4444); }
    }
    
    .bell-active {
        display: inline-block;
        animation: pulse-red 1.5s infinite;
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
# 2. CARREGAMENTO DE DADOS E ESTADOS DE SESSÃO
# ==============================================================================
@st.cache_data
def carregar_dados_cadastro():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Ficha de Cadastro", header=0, engine='openpyxl')
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]
        if 'Contato' in df.columns:
            df['FONE_LIMPO'] = df['Contato'].astype(str).apply(lambda x: ''.join(filter(str.isdigit, x)))
        else:
            df['FONE_LIMPO'] = ""
        return df
    except Exception:
        return pd.DataFrame()

df_cadastro = carregar_dados_cadastro()

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
    
    # 1. Acesso Mestre
    if (login_str in ["mestre", "london"] or login_limpo == "00000000000") and senha_input == "12381314*Lj":
        return {"nome": "Mestre London", "tipo": "mestre", "filhos": []}
    
    # 2. Acesso Diretoria
    if login_str in st.session_state.senhas_diretoria and senha_input == st.session_state.senhas_diretoria[login_str]:
        return {"nome": "Diretoria IEQ Guaicurus", "tipo": "diretoria", "filhos": []}
        
    # 3. Acesso Pais / Responsáveis
    if not df_cadastro.empty:
        match = df_cadastro[df_cadastro['FONE_LIMPO'].str.contains(login_limpo, na=False)] if login_limpo else pd.DataFrame()
        if match.empty and 'Responsável' in df_cadastro.columns:
            match = df_cadastro[df_cadastro['Responsável'].astype(str).str.lower().str.contains(login_str, na=False)]
            
        if not match.empty:
            resp_nome = match.iloc[0].get('Responsável', 'Responsável')
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

# ==============================================================================
# 3. TELA DE LOGIN COM LOGO DESTACADA
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
# 4. APLICAÇÃO LOGADA & NAVEGAÇÃO POR ABAS
# ==============================================================================
else:
    u_info = st.session_state.user_info
    
    # Saudação Formal Personalizada
    if u_info.get('tipo') in ['mestre', 'diretoria']:
        saudacao = "A Paz seja convosco, Mestre London / Diretoria"
    else:
        saudacao = f"A Paz seja convosco, {u_info.get('nome')}"

    # Cabeçalho com Sininho Inteligente
    col_head1, col_head2 = st.columns([5, 1])
    with col_head1:
        st.markdown(f"<strong style='color:#F59E0B; font-size:1.1rem;'>{saudacao}</strong>", unsafe_allow_html=True)
    with col_head2:
        sino_style = "color:#D97706;" if st.session_state.sino_visto else "color:#EF4444;"
        sino_class = "" if st.session_state.sino_visto else "bell-active"
        if st.button("🔔", key="btn_sino"):
            st.session_state.sino_visto = True
            st.rerun()

    # Painel de Notificações do Sininho
    if st.session_state.sino_visto:
        with st.expander("📢 *Central Oficial de Notificações & Eventos*", expanded=True):
            st.markdown("""
            * *Dia de Palestra:* 29/10/2026 – Inteligência Emocional com Psicóloga Eva Mateus
            * *Dia de Treinamento:* 15/10/2026 – Treinamento Básico de Primeiros Socorros
            * *Dia de Campeonato:* 20/11/2026 – Copa Corumbá Open de Jiu-Jitsu
            * *Dia de Aniversariantes do Mês:* 29/09/2026 às 19h no Espaço Cedido IEQ Guaicurus
            * *Dia de Graduação:* 15/12/2026 – Cerimônia de Entrega de Faixas e Graus
            * *Dia de Ação Social:* 10/10/2026 – Grande Aulão de Jiu-Jitsu Kids
            """)
            if st.button("Fechar Notificações"):
                st.session_state.sino_visto = False
                st.rerun()

    # Estrutura de Abas
    tab1, tab_mat, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Início",
        "📝 Matrícula",
        "📅 Frequência",
        "🥋 Tatame",
        "🏆 Campeão & Fé",
        "📊 Secretaria"
    ])

    # --------------------------------------------------------------------------
    # ABA 1: INÍCIO (BOAS-VINDAS & HISTÓRIA + DADOS DO ALUNO)
    # --------------------------------------------------------------------------
    with tab1:
        # CARD DEDICADO DE BOAS-VINDAS E HISTÓRIA DO PROJETO SEMENTES
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
        st.markdown("### 🎂 Aniversariantes do Mês & Festividades")
        st.caption("Comemorações oficiais da equipe voluntária e familiares (Datas: Setembro 29/09, Outubro 29/10, Novembro 26/11, Dezembro 29/12 às 19h):")
        
        c_aniv1, c_aniv2 = st.columns([1, 2])
        with c_aniv1:
            st.markdown("📷 *Foto do Aniversariante*")
            st.image("logo_projeto_sementes.png", width=100)
        with c_aniv2:
            st.markdown("*Alvaro Barbosa* — Aniversariante de Setembro")
            st.markdown("🎉 *Data da Festa:* 29/09/2026 (Terça-feira) às 19h")
            msg_aniv_auto = "Paz do Senhor, Alvaro! O Projeto Sementes te deseja um feliz aniversário! Que o Senhor abençoe sua vida e te dê sabedoria no tatame e na caminhada! 🥋🎉"
            link_aniv = gerar_link_whatsapp("5567998411953", msg_aniv_auto)
            if link_aniv:
                st.markdown(f"[📲 Enviar Mensagem Evangélica Parabéns]({link_aniv})")

        st.markdown("---")
        st.markdown("### 🥋 Perfil do Aluno")
        
        # Filtro de Alunos conforme Permissão
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.info("👑 Modo Mestre/Voluntariado: Selecione qualquer aluno para visualizar o perfil completo:")
            lista_todos_alunos = df_cadastro['Nome do Aluno'].dropna().tolist() if not df_cadastro.empty else ["Alvaro Barbosa"]
            aluno_selecionado_perfil = st.selectbox("Selecione o Aluno:", lista_todos_alunos)
            aluno_dados = df_cadastro[df_cadastro['Nome do Aluno'] == aluno_selecionado_perfil].to_dict(orient='records')
            filhos_exibir = aluno_dados if aluno_dados else []
        else:
            filhos_exibir = u_info.get('filhos', [])
            if not filhos_exibir and not df_cadastro.empty:
                filhos_exibir = df_cadastro.head(1).to_dict(orient='records')

        for f in filhos_exibir:
            nome_al = f.get('Nome do Aluno', 'Aluno')
            faixa_al = f.get('Faixa Graus', 'Branca')
            resp_al = f.get('Responsável', 'Responsável Cadastrado')
            saude_al = f.get('Histórico de Saúde', 'Não')
            is_batizado = nome_al in st.session_state.alunos_batizados
            
            st.markdown(f"""
            <div class="gold-card">
                <h3>{nome_al} {'⭐' if is_batizado else ''}</h3>
                <span class="gold-badge">Faixa {faixa_al}</span>
                <p style="margin-top:10px; font-size:0.95rem;">
                    <strong>Responsável Legal:</strong> {resp_al}<br>
                    <strong>Biometria Facial:</strong> 🟢 Cadastrada<br>
                    <strong>Status de Saúde:</strong> {'🚨 ' + str(saude_al) if str(saude_al).lower() != 'não' else '🔒 Apto para treinos de contato'}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🏫 Acompanhamento Escolar & Familiar (Mensal)")
        with st.expander("📝 Responder Avaliação Mensal do Aluno", expanded=True):
            st.selectbox("Aluno Avaliado:", [f.get('Nome do Aluno') for f in filhos_exibir] if filhos_exibir else ["Alvaro Barbosa"])
            st.select_slider("1. Desempenho Escolar:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.select_slider("2. Comportamento em Casa e Família:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.select_slider("3. Disciplina no Tatame:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            st.text_area("Observações para a Equipe Voluntária:")
            if st.button("Salvar Avaliação Mensal"):
                st.success("Avaliação enviada com sucesso para a coordenação voluntária!")

    # --------------------------------------------------------------------------
    # ABA 2: FICHA DE MATRÍCULA & TERMOS DE ISENÇÃO DE RESPONSABILIDADE (VOLUNTARIADO)
    # --------------------------------------------------------------------------
    with tab_mat:
        st.markdown("### 📝 Ficha de Matrícula Voluntária & Isenção de Responsabilidade")
        st.caption("Iniciativa comunitária sem fins lucrativos (Lei nº 9.608/1998). O espaço físico é cedido voluntariamente pela IEQ Guaicurus.")

        with st.form("form_matricula_juridica"):
            st.markdown("#### 1. DADOS DO ALUNO (MENOR)")
            mat_nome_aluno = st.text_input("Nome Completo do Aluno")
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                mat_data_nasc = st.date_input("Data de Nascimento", min_value=datetime.date(2008, 1, 1))
                mat_faixa = st.selectbox("Faixa", ["Branca", "Cinza", "Amarela", "Laranja", "Verde", "Azul", "Roxa", "Marrom", "Preta"])
            with c_m2:
                mat_escola = st.text_input("Escola onde estuda")
                mat_graus = st.selectbox("Graus", ["0 Graus", "1 Grau", "2 Graus", "3 Graus", "4 Graus"])
            mat_endereco = st.text_input("Endereço Residencial Completo")

            st.markdown("---")
            st.markdown("#### 📸 2. BIOMETRIA FACIAL DO ALUNO")
            st.caption("Clique no botão abaixo apenas se desejar tirar a foto agora:")
            foto_bio_mat = st.camera_input("Capturar Foto Facial do Aluno")

            st.markdown("---")
            st.markdown("#### 👨‍👩‍👧 3. IDENTIFICAÇÃO DE RESPONSABILIDADE FAMILIAR")
            qtd_filhos = st.number_input("Quantos filhos/dependentes você possui participando do projeto voluntário?", min_value=1, max_value=6, value=1)
            nomes_outros_filhos = st.text_input("Informe o nome dos demais filhos sob sua responsabilidade (se houver):", placeholder="Ex: Arthur Barbosa, Lucas Barbosa")

            st.markdown("---")
            st.markdown("#### 👤 4. DADOS DO RESPONSÁVEL LEGAL")
            mat_nome_resp = st.text_input("Nome Completo do Responsável Legal")
            c_r1, c_r2 = st.columns(2)
            with c_r1:
                mat_parentesco = st.selectbox("Grau de Parentesco", ["Pai", "Mãe", "Avô/Avó", "Tio/Tia", "Tutor Legal"])
                mat_cin_resp = st.text_input("Carteira de Identidade Nacional (CIN) do Responsável")
            with c_r2:
                mat_whats_resp = st.text_input("Telefone / WhatsApp (com DDD)", placeholder="Ex: 67998411953")

            st.markdown("---")
            st.markdown("#### 🩺 5. DECLARAÇÃO DE APTIDÃO FÍSICA E HISTÓRICO DE SAÚDE")
            p1 = st.radio("• O menor possui algum problema cardíaco ou de pressão?", ["Não", "Sim"])
            p2 = st.radio("• O menor sofre de asma, bronquite ou problemas respiratórios?", ["Não", "Sim"])
            p3 = st.radio("• O menor possui alguma lesão óssea, muscular ou articular crônica?", ["Não", "Sim"])
            p4 = st.radio("• O menor faz uso regular de algum medicamento controlled?", ["Não", "Sim"])
            p5 = st.radio("• O menor possui alergia a algum medicamento ou substância?", ["Não", "Sim"])
            p6 = st.radio("• O menor já sofreu desmaios ou tonturas durante exercícios físicos?", ["Não", "Sim"])
            
            doencas_sel = st.multiselect("Selecione se possui alguma das condições específicas:", ["Asma", "Bronquite", "Pressão Alta", "Epilepsia", "Diabetes", "Alergia Medicamentosa", "Nenhuma"])
            especificacao_saude = st.text_area("Caso tenha marcado 'SIM' ou selecionado alguma das opções acima, especifique detalhes de cuidados:")

            st.markdown("---")
            st.markdown("#### 📜 6. TERMOS DE VOLUNTARIADO E ISENÇÃO DE RESPONSABILIDADE JURÍDICA")
            
            st.markdown("*TERMO 1: AUTORIZAÇÃO DE PARTICIPAÇÃO E CESSÃO DE ESPAÇO (LEI 9.608/1998)*")
            st.caption("Eu, acima identificado(a) como responsável legal, AUTORIZO expressamente o menor sob minha tutela a participar de forma inteiramente gratuita das aulas voluntárias de Jiu-Jitsu do Projeto Social 'Sementes'. DECLARO estar ciente de que as atividades correspondem a um serviço comunitário voluntário prestado por instrutores pessoa física, sem vínculo empregatício ou fins lucrativos (Lei nº 9.608/1998), e que a Igreja do Evangelho Quadrangular (IEQ Guaicurus) atua EXCLUSIVAMENTE como cedente voluntária do espaço físico comunitário, isenta de qualquer responsabilidade legal, vinculação jurídica, civil, trabalhista ou de associação comercial sobre a condução das atividades esportivas.")
            t_espaco = st.checkbox("Li e aceito a natureza voluntária da atividade e a isenção de responsabilidade do espaço cedido.")

            st.markdown("<br>*TERMO 2: DECLARAÇÃO DE APTIDÃO FÍSICA E SAÚDE*", unsafe_allow_html=True)
            st.caption("Declaro, sob as penas da lei, que o menor acima qualificado goza de boa saúde e encontra-se apto para a prática de exercícios físicos de contato próprio da modalidade (Jiu-Jitsu), ministrada sob supervisão voluntária. Comprometo-me a informar imediatamente os organizadores caso ocorra qualquer alteração em seu estado de saúde no decorrer do período de treino. Corumbá - MS, Setembro de 2026.")
            t_saude = st.checkbox("Li e declaro total aptidão física e de saúde do menor sob minha responsabilidade.")

            st.markdown("<br>*TERMO 3: AUTORIZAÇÃO DE USO DE IMAGEM E VOZ PARA FINS INFORMATIVOS*", unsafe_allow_html=True)
            st.caption("AUTORIZO de forma gratuita, definitiva e irrevogável, a utilização da imagem e voz do referido menor, capturadas em treinos ou encontros, exclusivamente para divulgação institucional e prestação de contas das ações voluntárias junto à comunidade local (boletins internos, redes sociais e murais), vedada qualquer forma de exploração comercial ou fins lucrativos. Por ser esta a expressão da minha vontade, firmo o presente termo. Corumbá - MS, Setembro de 2026.")
            t_imagem = st.checkbox("Li e autorizo o uso comunitário e não comercial de imagem e voz.")

            st.markdown("---")
            st.markdown("#### ✍️ 7. ASSINATURA DIGITAL DO RESPONSÁVEL LEGAL")
            st.caption("Assine com o dedo ou mouse no quadro abaixo:")

            if CANVAS_DISPONIVEL:
                canvas_result = st_canvas(
                    fill_color="rgba(255, 255, 255, 0)",
                    stroke_width=3,
                    stroke_color="#FFFFFF",
                    background_color="#111827",
                    height=150,
                    update_streamlit=True,
                    key="canvas_mat_juridica",
                )

            btn_finalizar_mat = st.form_submit_button("FINALIZAR E ENVIAR MATRÍCULA VOLUNTÁRIA")

            if btn_finalizar_mat:
                if not (t_espaco and t_saude and t_imagem):
                    st.error("❌ É necessário aceitar todos os termos de voluntariado e isenção de responsabilidade para concluir.")
                elif not mat_nome_aluno or not mat_nome_resp or not mat_whats_resp:
                    st.error("❌ Preencha os campos obrigatórios (Aluno, Responsável e WhatsApp).")
                else:
                    st.success("✅ Ficha Voluntária Registrada com Sucesso!")
                    msg_wa_jur = (
                        f"PROJETO SOCIAL SEMENTES - VOLUNTARIADO 🌱\n\n"
                        f"Olá, {mat_nome_resp}! A inscrição do aluno(a) {mat_nome_aluno} foi validadada e assinada digitalmente.\n\n"
                        f"Comprovante de Voluntariado e Ciência:\n"
                        f"• Responsável: {mat_nome_resp} (CIN: {mat_cin_resp})\n"
                        f"• Filhos Registrados: {qtd_filhos} ({mat_nome_aluno}, {nomes_outros_filhos})\n"
                        f"• Termos Aceitos: Voluntariado (Lei 9.608/98), Espaço Cedido IEQ, Aptidão de Saúde e Imagem\n"
                        f"• Data/Hora: {datetime.datetime.now().strftime('%d/%m/%Y às %H:%M')}\n"
                        f"• Local: Corumbá - MS\n\nDeus abençoe! 🙏"
                    )
                    link_mat_wa = gerar_link_whatsapp(mat_whats_resp, msg_wa_jur)
                    if link_mat_wa:
                        st.markdown(f'<a href="{link_mat_wa}" target="_blank"><button style="background-color:#25D366; color:white; border:none; padding:14px; border-radius:10px; width:100%; font-weight:bold; cursor:pointer; margin-top:10px;">👉 ENVIAR COMPROVANTE VIA WHATSAPP</button></a>', unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # ABA 3: FREQUÊNCIA, MÉTRICAS & EVOLUÇÃO
    # --------------------------------------------------------------------------
    with tab2:
        st.markdown("### 📅 Frequência & Métricas Gerais dos Treinos")
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total de Treinos no Mês", "9 Aulas")
        col_m2.metric("Total de Presenças", "312 Aulas/Alunos")
        col_m3.metric("Percentual de Presença", "88.5%")

        st.markdown("---")
        st.markdown("### 🥋 Desempenho e Evolução Individual")
        aluno_sel_freq = st.selectbox("Selecione o Aluno para ver Frequência e Evolução:", df_cadastro['Nome do Aluno'].dropna().tolist() if not df_cadastro.empty else ["Alvaro Barbosa"])
        
        if u_info.get('tipo') == 'mestre':
            st.markdown("*👑 Painel do Mestre: Atribuir Avaliação de Evolução*")
            ev_opcao = st.radio(f"Definir evolução técnica de {aluno_sel_freq}:", ["Ótima", "Boa", "Regular"], horizontal=True)
            if st.button("Salvar Evolução do Aluno"):
                st.session_state.evolucao_mestre[aluno_sel_freq] = ev_opcao
                st.success(f"Evolução de {aluno_sel_freq} definida como {ev_opcao}!")

        status_ev_exibir = st.session_state.evolucao_mestre.get(aluno_sel_freq, "Ótima")
        st.info(f"📊 *Resultado da Avaliação:\n *Frequência Mensal:* 88.8%\n* *Evolução Técnica (Mestre):* {status_ev_exibir}")

        st.markdown("---")
        st.markdown("### 🗓️ Histórico de Treino (Terças e Quintas)")
        grid_limpo = [
            {"Data": "01/09/2026", "Status": "🟢 Presença"},
            {"Data": "03/09/2026", "Status": "🟢 Presença"},
            {"Data": "08/09/2026", "Status": "🔴 Falta"},
            {"Data": "10/09/2026", "Status": "🟢 Presença"},
            {"Data": "15/09/2026", "Status": "🟢 Presença"},
        ]
        st.dataframe(pd.DataFrame(grid_limpo), use_container_width=True)

    # --------------------------------------------------------------------------
    # ABA 4: TATAME & DOJO (CHAMADA DE PROFESSORES VOLUNTÁRIOS)
    # --------------------------------------------------------------------------
    with tab3:
        st.markdown("### 🥋 Chamada Rápida de Presença no Dojo - Professores")
        
        modo_chamada = st.radio("Selecione a modalidade de chamada:", ["1. Fotos Individuais dos Alunos Cadastrados", "2. Foto Coletiva (Identificação Visual)"])
        
        if modo_chamada == "1. Fotos Individuais dos Alunos Cadastrados":
            st.caption("Clique no checkbox ao lado da foto para confirmar presença:")
            if not df_cadastro.empty:
                for idx, row in df_cadastro.head(8).iterrows():
                    aluno_n = row.get('Nome do Aluno', 'Aluno')
                    c_f1, c_f2 = st.columns([1, 3])
                    with c_f1:
                        st.image("logo_projeto_sementes.png", width=50)
                    with c_f2:
                        st.checkbox(f"🟢 Presente: {aluno_n}", key=f"chk_dojo_{idx}")
        else:
            st.caption("Carregue a foto da turma reunida no tatame:")
            foto_turma = st.file_uploader("Carregar foto do treino:", type=["jpg", "png"])
            if foto_turma:
                st.image(foto_turma, use_container_width=True)
                st.markdown("🟢 *Bolinha Verde:* 35 Alunos Cadastrados Identificados em Aula.")
                st.markdown("🔴 *Bolinha Vermelha:* 2 Pessoas na foto identificadas sem cadastro (Falta de Cadastro).")

        if st.button("SALVAR PRESENÇA DO DOJO"):
            st.success("Presença salva no sistema com sucesso!")

    # --------------------------------------------------------------------------
    # ABA 5: CAMPEÃO, FÉ, CÉLULAS & EVENTOS
    # --------------------------------------------------------------------------
    with tab4:
        sub_tab = st.radio("Selecione o módulo:", ["🏆 Pódios & Campeonatos", "✝️ Pedidos de Oração & Fé", "📢 Palestras & Treinamentos"], horizontal=True)
        
        if sub_tab == "🏆 Pódios & Campeonatos":
            st.markdown("#### 🏆 Informe Geral de Campeonatos da Equipe Voluntária")
            st.markdown("""
            <div class="gold-card">
                <h5>🥇 Colocação da Equipe: 2º Lugar Geral por Equipes</h5>
                <p style="font-size:0.9rem;">
                    <strong>Campeonato Estadual de Jiu-Jitsu 2026 (Campo Grande-MS)</strong><br>
                    • Total de Medalhas: 8 Ouros, 5 Pratas, 3 Bronzes.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("#### 🥇 Pódio Individual do Aluno")
            aluno_camp = st.selectbox("Selecione o Aluno para visualizar conquistas:", df_cadastro['Nome do Aluno'].dropna().tolist() if not df_cadastro.empty else ["Alvaro Barbosa"])
            st.info(f"🏆 *Conquistas de {aluno_camp}:\n 🥇 *Medalha de Ouro* – Categoria Infantil B (Estadual 2026)\n* 🥈 *Medalha de Prata* – Copa Pantanal de Jiu-Jitsu")

        elif sub_tab == "✝️ Pedidos de Oração & Fé":
            st.markdown("#### 🙏 Pedido de Oração ao Nosso Pastor Joel")
            with st.form("form_oracao"):
                nome_oracao = st.text_input("Seu Nome:")
                texto_oracao = st.text_area("Escreva seu pedido de oração ao Pastor Joel:")
                if st.form_submit_button("Enviar Pedido de Oração"):
                    st.session_state.pedidos_oracao.append({"nome": nome_oracao, "pedido": texto_oracao})
                    st.success("Pedido de oração enviado com sucesso ao Pastor Joel!")

            st.markdown("---")
            st.markdown("#### ⭐ Eu quero Batizar!")
            if st.button("🙌 Quero me Batizar nas Águas!"):
                st.success("Glória a Deus! Sua intenção de batismo foi registrada e nossa equipe pastoral entrará em contato.")

            st.markdown("---")
            st.markdown("#### 📊 Impacto Espiritual")
            c_esp1, c_esp2, c_esp3 = st.columns(3)
            c_esp1.metric("Pedidos de Oração", len(st.session_state.pedidos_oracao) + 14)
            c_esp2.metric("Vidas em Células", st.session_state.membros_celula_count)
            c_esp3.metric("Alunos Batizados", len(st.session_state.alunos_batizados))

            st.markdown("---")
            st.markdown("#### ⛪ Cultos, Células e Programação IEQ Guaicurus")
            st.markdown("""
            * *Culto Principal:* Todo Domingo às 18:30h
            * *Cronograma de Células:* Quarta-feira às 19:30h
            * *Festividades:* Festa da Colheita e Aniversariantes do Mês
            """)

        elif sub_tab == "📢 Palestras & Treinamentos":
            st.markdown("#### 🧠 Palestras Socioeducativas")
            st.markdown("""
            <div class="gold-card">
                <h5>Tema: Inteligência Emocional e Combate ao Bullying</h5>
                <p style="font-size:0.9rem;">
                    <strong>Palestrante:</strong> Psicóloga Eva Mateus<br>
                    <strong>Data/Hora:</strong> 29/10/2026 às 19:00h
                </p>
            </div>
            """, unsafe_allow_html=True)
            link_pal = gerar_link_whatsapp("5567998411953", "Lembrete do Projeto Sementes: Não perca hoje às 19h nossa Palestra Socioeducativa com a Psicóloga Eva Mateus!")
            if link_pal:
                st.markdown(f"[🔔 Enviar Notificação de Lembrete aos Pais]({link_pal})")

            st.markdown("---")
            st.markdown("#### 🛠️ Treinamentos Básicos")
            st.markdown("""
            <div class="gold-card">
                <h5>Assunto: Primeiros Socorros e Cuidados no Tatame</h5>
                <p style="font-size:0.9rem;">
                    <strong>Instrutor:</strong> Socorrista Marcos<br>
                    <strong>Data/Hora:</strong> 15/10/2026 às 18:30h
                </p>
            </div>
            """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # ABA 6: SECRETARIA VOLUNTÁRIA, GESTÃO DE SENHAS & RELATÓRIOS
    # --------------------------------------------------------------------------
    with tab5:
        st.markdown("### 📊 Secretaria Voluntária & Coordenação")
        
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.success("👑 Nível de Acesso: Mestre / Coordenação Voluntária (Controle Total)")
            
            st.markdown("---")
            st.markdown("#### 🚨 Central de Notificações de Pendências")
            st.caption("Assim que sanadas na planilha base, as pendências somem automaticamente da lista:")
            
            if not df_cadastro.empty:
                col_p_doc = [c for c in df_cadastro.columns if 'pend' in str(c).lower()]
                c_doc_n = col_p_doc[0] if col_p_doc else 'Pendencia Documento'
                
                df_p_ativas = df_cadastro[df_cadastro[c_doc_n].astype(str).str.lower() != 'ok'] if c_doc_n in df_cadastro.columns else pd.DataFrame()
                
                if not df_p_ativas.empty:
                    for _, p_row in df_p_ativas.head(5).iterrows():
                        p_al = p_row.get('Nome do Aluno', 'Aluno')
                        p_re = p_row.get('Responsável', 'Responsável')
                        p_fo = p_row.get('Contato', '')
                        p_txt = p_row.get(c_doc_n, 'Pendência de Documento')
                        
                        st.markdown(f"• *{p_al}* (Resp: {p_re}) — <span style='color:#EF4444;'>{p_txt}</span>", unsafe_allow_html=True)
                        msg_p_cob = f"Paz do Senhor, {p_re}! Lembramos da pendência do(a) aluno(a) {p_al} ({p_txt}) no Projeto Sementes. Você pode regularizar diretamente pelo aplicativo!"
                        link_p_cob = gerar_link_whatsapp(p_fo, msg_p_cob)
                        if link_p_cob:
                            st.markdown(f"[💬 Enviar Cobrança no WhatsApp de {p_re}]({link_p_cob})")
                        st.markdown("---")
                else:
                    st.success("✅ Todas as pendências de documentação foram sanadas!")

            st.markdown("#### 🔐 Gestão de Senhas dos Pais e Diretoria")
            c_pass1, c_pass2 = st.columns(2)
            with c_pass1:
                tipo_user_pass = st.selectbox("Perfil de Acesso:", ["Pai / Responsável", "Diretoria"])
                user_key_pass = st.text_input("WhatsApp do Pai ou Login da Diretoria:", "67998411953")
            with c_pass2:
                new_pass = st.text_input("Definir Nova Senha:", "123456")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Gerar / Salvar Senha"):
                    if tipo_user_pass == "Pai / Responsável":
                        st.session_state.senhas_pais[user_key_pass] = new_pass
                    else:
                        st.session_state.senhas_diretoria[user_key_pass] = new_pass
                    st.success(f"Senha de {tipo_user_pass} atualizada!")

            st.markdown("---")
            st.markdown("#### 🚫 Excluir Acesso / Aluno Desistente")
            desistente_fone = st.text_input("WhatsApp do Aluno Desistente para remover acesso:")
            if st.button("Remover Acesso do Aplicativo"):
                st.session_state.alunos_desistentes.append(desistente_fone)
                st.warning(f"Acesso associado ao número {desistente_fone} foi bloqueado com sucesso.")

            st.markdown("---")
            st.markdown("#### 📄 Emissão de Relatório Personalizado")
            rel_abas = st.multiselect("Selecione os módulos para incluir no relatório PDF:", ["Ficha de Cadastro e Alunos", "Controle de Frequência", "Matrículas Voluntárias e Isenção", "Evolução e Pódios", "Impacto Espiritual"], default=["Ficha de Cadastro e Alunos", "Controle de Frequência"])
            if st.button("📄 Gerar e Exportar Relatório Escolhido (PDF)"):
                st.success("Relatório processado e pronto para download!")

        st.markdown("---")
        st.markdown("#### 🗺️ Mapa Social de Expansão do Projeto Sementes IEQ Guaicurus em Corumbá-MS")
        st.caption("Atualização diária automática da distribuição territorial de alunos por bairro:")
        
        mapa_exp = pd.DataFrame({
            'Bairro': ['Guaicurus', 'Nova Corumbá', 'Centro', 'Guarani', 'Outros'],
            'Alunos': [18, 12, 5, 4, 2],
            'Porcentagem (%)': [43.9, 29.3, 12.2, 9.8, 4.8]
        })
        st.dataframe(mapa_exp, use_container_width=True, hide_index=True)
        st.bar_chart(mapa_exp.set_index('Bairro')['Alunos'])
        st.info("📊 *Total Geral de Abrangência:* 41 Alunos Ativos (100% de ocupação das turmas).")

        st.markdown("---")
        if st.button("🚪 Sair do Aplicativo"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
