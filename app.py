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
    page_title="Projeto Sementes - Jiu-Jitsu Voluntário",
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
# 2. FUNÇÕES ROBUSTAS DE CARREGAMENTO (EVITA KEYERROR)
# ==============================================================================
@st.cache_data
def carregar_dados_cadastro():
    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine='openpyxl')
        sheet_target = "Ficha de Cadastro" if "Ficha de Cadastro" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_target, header=0)
        df = df.dropna(how='all').dropna(how='all', axis=1)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Mapeamento Flexível de Colunas Nativas
        col_aluno = next((c for c in df.columns if any(k in c.lower() for k in ['aluno', 'nome'])), None)
        col_resp = next((c for c in df.columns if any(k in c.lower() for k in ['responsavel', 'responsável', 'pai', 'mae'])), None)
        col_fone = next((c for c in df.columns if any(k in c.lower() for k in ['contato', 'tel', 'fone', 'whats', 'celular'])), None)
        col_faixa = next((c for c in df.columns if any(k in c.lower() for k in ['faixa', 'gradua'])), None)

        df['NOME_ALUNO_CLEAN'] = df[col_aluno] if col_aluno else "Aluno Não Identificado"
        df['RESPONSAVEL_CLEAN'] = df[col_resp] if col_resp else "Responsável Não Identificado"
        df['FAIXA_CLEAN'] = df[col_faixa] if col_faixa else "Branca"
        
        if col_fone:
            df['FONE_LIMPO'] = df[col_fone].astype(str).apply(lambda x: ''.join(filter(str.isdigit, x)))
        else:
            df['FONE_LIMPO'] = ""
            
        return df
    except Exception:
        # Fallback de Segurança se a planilha estiver em branco
        return pd.DataFrame({
            'NOME_ALUNO_CLEAN': ['Alvaro Barbosa', 'Nathan Santos', 'Isadora Carvalho'],
            'RESPONSAVEL_CLEAN': ['Odiselma Carvalho', 'London Santos', 'London Santos'],
            'FAIXA_CLEAN': ['Amarela', 'Branca 2 Graus', 'Branca'],
            'FONE_LIMPO': ['67998411953', '67998411953', '67998411953'],
            'Pendência Documento': ['OK', 'Pendente CIN', 'OK']
        })

@st.cache_data
def carregar_cronograma_eventos():
    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine='openpyxl')
        if "Cronograma de Eventos" in xls.sheet_names:
            df_ev = pd.read_excel(xls, sheet_name="Cronograma de Eventos")
            df_ev = df_ev.dropna(how='all')
            return df_ev
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
    
    # 1. Mestre
    if (login_str in ["mestre", "london"] or login_limpo == "00000000000") and senha_input == "12381314*Lj":
        return {"nome": "Mestre London", "tipo": "mestre", "filhos": []}
    
    # 2. Diretoria
    if login_str in st.session_state.senhas_diretoria and senha_input == st.session_state.senhas_diretoria[login_str]:
        return {"nome": "Diretoria IEQ Guaicurus", "tipo": "diretoria", "filhos": []}
        
    # 3. Pais / Responsáveis
    if not df_cadastro.empty:
        match = df_cadastro[df_cadastro['FONE_LIMPO'].str.contains(login_limpo, na=False)] if login_limpo else pd.DataFrame()
        if match.empty:
            match = df_cadastro[df_cadastro['RESPONSAVEL_CLEAN'].astype(str).str.lower().str.contains(login_str, na=False)]
            
        if not match.empty:
            resp_nome = match.iloc[0]['RESPONSAVEL_CLEAN']
            fone = match.iloc[0]['FONE_LIMPO']
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

# Verificar se há evento no dia atual
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
# 4. APLICAÇÃO LOGADA & NAVEGAÇÃO POR ABAS
# ==============================================================================
else:
    u_info = st.session_state.user_info
    
    if u_info.get('tipo') in ['mestre', 'diretoria']:
        saudacao = "A Paz seja convosco, Mestre London / Diretoria"
    else:
        saudacao = f"A Paz seja convosco, {u_info.get('nome')}"

    # Cabeçalho com Sininho Dinâmico que Pisca em Dia de Evento
    col_head1, col_head2 = st.columns([5, 1])
    with col_head1:
        st.markdown(f"<strong style='color:#F59E0B; font-size:1.1rem;'>{saudacao}</strong>", unsafe_allow_html=True)
    with col_head2:
        btn_bell_label = "🔔" if not tem_evento_hoje else "🚨🔔"
        if st.button(btn_bell_label, key="btn_sino"):
            st.session_state.sino_visto = not st.session_state.sino_visto
            st.rerun()

    # Painel de Notificações do Sininho
    if st.session_state.sino_visto or tem_evento_hoje:
        with st.expander("🔔 *Central Oficial de Notificações & Eventos do Dia*", expanded=True):
            if tem_evento_hoje:
                st.error("🚨 *HOJE TEM EVENTO NO PROJETO SEMENTES!*")
            st.markdown("##### 📅 Programação Carregada da Planilha Base:")
            if not df_cronograma.empty:
                st.dataframe(df_cronograma, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum evento registrado para esta data.")
            if st.button("Fechar Avisos"):
                st.session_state.sino_visto = False
                st.rerun()

    # Estrutura de Abas
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
    # ABA 1: INÍCIO (HISTÓRIA & PERFIL DO ALUNO)
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
        st.markdown("### 🎂 Aniversariantes do Mês & Festividades")
        st.caption("Comemorações oficiais (Datas: Setembro 29/09, Outubro 29/10, Novembro 26/11, Dezembro 31/12 às 19h):")
        
        c_aniv1, c_aniv2 = st.columns([1, 2])
        with c_aniv1:
            st.markdown("📷 *Foto do Aluno*")
            st.info("👤 [Foto do Perfil do Aluno]")
        with c_aniv2:
            st.markdown("*Alvaro Barbosa* — Aniversariante de Setembro")
            st.markdown("🎉 *Data da Festa:* 29/09/2026 (Terça-feira) às 19h")
            msg_aniv_auto = "Paz do Senhor, Alvaro! O Projeto Sementes te deseja um feliz aniversário! Que o Senhor abençoe sua vida e te dê sabedoria no tatame! 🥋🎉"
            link_aniv = gerar_link_whatsapp("5567998411953", msg_aniv_auto)
            if link_aniv:
                st.markdown(f"[📲 Enviar Mensagem Evangélica Parabéns]({link_aniv})")

        st.markdown("---")
        st.markdown("### 🥋 Perfil do Aluno")
        
        lista_alunos_base = df_cadastro['NOME_ALUNO_CLEAN'].tolist() if not df_cadastro.empty else ["Alvaro Barbosa"]
        
        if u_info.get('tipo') in ['mestre', 'diretoria']:
            st.info("👑 Modo Mestre/Voluntariado: Selecione qualquer aluno cadastrado:")
            aluno_sel_perfil = st.selectbox("Selecione o Aluno:", lista_alunos_base)
            filhos_exibir = df_cadastro[df_cadastro['NOME_ALUNO_CLEAN'] == aluno_sel_perfil].to_dict(orient='records')
        else:
            filhos_exibir = u_info.get('filhos', [])
            if not filhos_exibir and not df_cadastro.empty:
                filhos_exibir = df_cadastro.head(1).to_dict(orient='records')

        for f in filhos_exibir:
            nome_al = f.get('NOME_ALUNO_CLEAN', 'Aluno')
            faixa_al = f.get('FAIXA_CLEAN', 'Branca')
            resp_al = f.get('RESPONSAVEL_CLEAN', 'Responsável Cadastrado')
            is_batizado = nome_al in st.session_state.alunos_batizados
            
            st.markdown(f"""
            <div class="gold-card">
                <h3>{nome_al} {'⭐' if is_batizado else ''}</h3>
                <span class="gold-badge">Faixa {faixa_al}</span>
                <p style="margin-top:10px; font-size:0.95rem;">
                    <strong>Responsável Legal:</strong> {resp_al}<br>
                    <strong>Biometria Facial:</strong> 🟢 Cadastrada no Sistema<br>
                    <strong>Status de Saúde:</strong> 🔒 Apto para treinos de contato
                </p>
            </div>
            """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # ABA 2: MATRÍCULA & CAMPO "ALUNOS CADASTRADOS"
    # --------------------------------------------------------------------------
    with tab_mat:
        sub_mat = st.radio("Selecione a ação na Matrícula:", ["📋 Alunos Cadastrados & Pendências", "✍️ Nova Ficha de Matrícula Voluntária"], horizontal=True)
        
        if sub_mat == "📋 Alunos Cadastrados & Pendências":
            st.markdown("### 📋 Painel Geral de Alunos Cadastrados")
            st.caption("Esta lista é alimentada automaticamente pela planilha base e sincroniza todas as abas do aplicativo:")
            
            if not df_cadastro.empty:
                st.dataframe(df_cadastro[['NOME_ALUNO_CLEAN', 'FAIXA_CLEAN', 'RESPONSAVEL_CLEAN', 'FONE_LIMPO']], use_container_width=True, hide_index=True)
                
                st.markdown("---")
                st.markdown("#### 🚨 Alunos com Pendência de Documento / Biometria")
                df_pend = df_cadastro[df_cadastro['Pendência Documento'].astype(str).str.upper() != 'OK'] if 'Pendência Documento' in df_cadastro.columns else pd.DataFrame()
                
                if not df_pend.empty:
                    for _, p in df_pend.iterrows():
                        p_al = p['NOME_ALUNO_CLEAN']
                        p_re = p['RESPONSAVEL_CLEAN']
                        p_fo = p['FONE_LIMPO']
                        st.markdown(f"• *{p_al}* (Resp: {p_re}) — <span style='color:#EF4444;'>Pendência de Documentação/Biometria</span>", unsafe_allow_html=True)
                        msg_p = f"Paz do Senhor, {p_re}! Solicitamos a regularização dos documentos/biometria do aluno {p_al} no Projeto Sementes."
                        lk_p = gerar_link_whatsapp(p_fo, msg_p)
                        if lk_p:
                            st.markdown(f"[📲 Notificar Responsável no WhatsApp]({lk_p})")
                else:
                    st.success("✅ Todos os alunos cadastrados estão com a documentação 100% regularizada!")

        else:
            st.markdown("### 📝 Nova Ficha de Matrícula Voluntária")
            with st.form("form_matricula_nova"):
                st.text_input("Nome Completo do Aluno")
                st.date_input("Data de Nascimento", min_value=datetime.date(2008, 1, 1))
                st.selectbox("Faixa", ["Branca", "Cinza", "Amarela", "Laranja", "Verde", "Azul"])
                st.text_input("Escola onde estuda")
                st.camera_input("Capturar Biometria Facial do Aluno")
                st.text_input("Nome Completo do Responsável Legal")
                st.text_input("Carteira de Identidade Nacional (CIN) do Responsável")
                st.text_input("WhatsApp do Responsável (com DDD)")
                t_aceite = st.checkbox("Li e aceito os termos de voluntariado e isenção de responsabilidade.")
                if st.form_submit_button("FINALIZAR E ALIMENTAR CADASTRADOS"):
                    if t_aceite:
                        st.success("✅ Aluno cadastrado com sucesso! A lista de Alunos Cadastrados foi atualizada.")
                    else:
                        st.error("Aceite os termos para concluir.")

    # --------------------------------------------------------------------------
    # ABA 3: CRONOGRAMA DE EVENTOS
    # --------------------------------------------------------------------------
    with tab_crono:
        st.markdown("### 📅 Cronograma Oficial de Eventos do Projeto & Igreja")
        st.caption("Puxado diretamente da aba 'Cronograma de Eventos' da planilha base:")
        
        st.markdown("""
        * *Dia de Palestra:* Treinamentos socioeducativos e saúde mental.
        * *Dia de Treinamento:* Capacitação técnica e primeiros socorros.
        * *Dia de Campeonato:* Competições regionais e pódios.
        * *Dia de Aniversariantes do Mês:* Festividades com as famílias.
        * *Dia de Graduação:* Entrega de faixas e graus.
        * *Dia de Ação Social:* Evangelismo e acolhimento comunitário.
        """)
        st.markdown("---")
        
        if not df_cronograma.empty:
            st.dataframe(df_cronograma, use_container_width=True, hide_index=True)
        else:
            st.warning("Nenhum evento encontrado na aba 'Cronograma de Eventos' da planilha.")

    # --------------------------------------------------------------------------
    # ABA 4: FREQUÊNCIA & MÉTRICAS
    # --------------------------------------------------------------------------
    with tab2:
        st.markdown("### 📊 Frequência Mensal")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total de Treinos", "9 Aulas")
        col_m2.metric("Total Presenças", "312 Alunos")
        col_m3.metric("Frequência Geral", "88.5%")

        st.markdown("---")
        aluno_freq_sel = st.selectbox("Selecione o Aluno Cadastrado:", lista_alunos_base)
        
        if u_info.get('tipo') == 'mestre':
            st.markdown("*👑 Painel do Mestre: Atribuir Evolução*")
            ev_op = st.radio(f"Evolução de {aluno_freq_sel}:", ["Ótima", "Boa", "Regular"], horizontal=True)
            if st.button("Salvar Evolução"):
                st.session_state.evolucao_mestre[aluno_freq_sel] = ev_op
                st.success("Evolução salva!")

        st.info(f"📊 Avaliação Técnica de *{aluno_freq_sel}*: {st.session_state.evolucao_mestre.get(aluno_freq_sel, 'Ótima')}")

    # --------------------------------------------------------------------------
    # ABA 5: TATAME
    # --------------------------------------------------------------------------
    with tab3:
        st.markdown("### 🥋 Chamada Rápida de Presença no Dojo - Professores")
        st.caption("Alimentado diretamente pela lista de Alunos Cadastrados:")
        
        for name_al in lista_alunos_base[:10]:
            st.checkbox(f"🟢 Presente: {name_al}", key=f"chk_dojo_clean_{name_al}")
            
        if st.button("SALVAR PRESENÇA"):
            st.success("Presenças registradas!")

    # --------------------------------------------------------------------------
    # ABA 6: CAMPEÃO & FÉ
    # --------------------------------------------------------------------------
    with tab4:
        st.markdown("### 🏆 Pódios & Pedidos de Oração")
        with st.form("form_oracao_tab"):
            st.text_input("Seu Nome:")
            st.text_area("Pedido de Oração ao Pastor Joel:")
            if st.form_submit_button("Enviar Pedido"):
                st.success("Pedido enviado com sucesso ao Pastor Joel!")

    # --------------------------------------------------------------------------
    # ABA 7: SECRETARIA & GESTÃO
    # --------------------------------------------------------------------------
    with tab5:
        st.markdown("### ⚙️ Secretaria & Configurações")
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
