import streamlit as st
import pandas as pd
import datetime
import urllib.parse

# ==============================================================================
# 1. CONFIGURAÇÃO DE PÁGINA MOBILE-FIRST (PERSONALGO STYLE)
# ==============================================================================
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilização CSS Customizada para Interface Nativamente Mobile
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Fundo Geral Dark Mode Profundo */
    .stApp { background-color: #0A0F18; color: #E2E8F0; }
    
    /* Oculta elementos nativos do Streamlit */
    #MainMenu, header, footer, [data-testid="stSidebar"] { display: none !important; }
    
    /* Cabeçalho Fixo Otimizado */
    .app-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 12px 16px;
        background-color: #111827;
        border-bottom: 1px solid #1E293B;
        border-radius: 0 0 16px 16px;
        margin-bottom: 16px;
    }
    
    .gold-card {
        background-color: #111827;
        border: 1px solid #D97706;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 12px;
        box-shadow: 0px 4px 12px rgba(217, 119, 6, 0.12);
    }
    
    .gold-badge {
        background-color: #D97706;
        color: #000000;
        font-weight: 800;
        font-size: 0.75rem;
        padding: 3px 10px;
        border-radius: 12px;
        display: inline-block;
    }
    
    .star-badge {
        color: #F59E0B;
        font-size: 1.1rem;
        margin-left: 4px;
    }

    /* Botões em Dourado Metálico */
    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #D97706 !important;
        color: #FFFFFF !important;
        border: 1px solid #F59E0B !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        padding: 12px 20px !important;
        box-shadow: 0px 4px 14px rgba(217, 119, 6, 0.25);
        width: 100%;
    }
    .stButton>button:hover { background-color: #B45309 !important; }

    /* Estilo das Abas de Navegação Inferiores */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #111827;
        padding: 6px;
        border-radius: 14px;
        border: 1px solid #1E293B;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #94A3B8;
        font-weight: 600;
        padding: 8px 12px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #D97706 !important;
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

# ==============================================================================
# 2. CARREGAMENTO INTELIGENTE DE DADOS DA PLANILHA EXCEL
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

# Gerenciamento de Sessão de Usuário
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None

# Base em memória para senhas de pais e dados de oração/desenvolvimento
if "senhas_pais" not in st.session_state:
    st.session_state.senhas_pais = {"67998411953": "123456"}  # Exemplo pré-cadastrado para teste
if "pedidos_oracao_count" not in st.session_state:
    st.session_state.pedidos_oracao_count = 14
if "desenvolvimento_dados" not in st.session_state:
    st.session_state.desenvolvimento_dados = {}

def autenticar_usuario(login_input, senha_input):
    login_limpo = ''.join(filter(str.isdigit, str(login_input)))
    login_str = str(login_input).strip().lower()
    
    # 1. Acesso Mestre / Diretoria (Senha Única Mestra)
    if (login_str in ["mestre", "london", "diretoria"] or login_limpo == "00000000000") and senha_input == "12381314*Lj":
        return {"nome": "Mestre London / Diretoria", "tipo": "mestre", "filhos": []}
        
    # 2. Acesso Responsável / Pais
    if not df_cadastro.empty:
        match = df_cadastro[df_cadastro['FONE_LIMPO'].str.contains(login_limpo, na=False)] if login_limpo else pd.DataFrame()
        if match.empty and 'Responsável' in df_cadastro.columns:
            match = df_cadastro[df_cadastro['Responsável'].astype(str).str.lower().str.contains(login_str, na=False)]
            
        if not match.empty:
            resp_nome = match.iloc[0].get('Responsável', 'Responsável')
            fone = match.iloc[0].get('FONE_LIMPO', '')
            # Checa senha cadastrada pelo Mestre
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
# 3. TELA DE LOGIN E PRIMEIRO ACESSO
# ==============================================================================
if not st.session_state.logged_in:
    st.markdown("<br>", unsafe_allow_html=True)
    st.image("logo_projeto_sementes.png", width=110)
    st.markdown("<h2 style='text-align: center; color:#F59E0B;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94A3B8; font-size: 0.9rem;'>Cultivando valores, fortalecendo famílias</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.form("form_login"):
        user_in = st.text_input("WhatsApp / CPF do Responsável ou Mestre", placeholder="Ex: 67998411953 ou mestre")
        pass_in = st.text_input("Senha de Acesso", type="password", placeholder="••••••••")
        btn_login = st.form_submit_button("ENTRAR NO APP")
        
        if btn_login:
            user_data = autenticar_usuario(user_in, pass_in)
            if user_data:
                st.session_state.logged_in = True
                st.session_state.user_info = user_data
                st.success("Acesso autorizado com sucesso!")
                st.rerun()
            else:
                st.error("Credenciais inválidas. Verifique os dados ou solicite sua senha ao Mestre.")
                
    st.markdown("<br><p style='text-align: center; color: #64748B; font-size: 0.75rem;'>Apoio: Igreja do Evangelho Quadrangular – Guaicurus</p>", unsafe_allow_html=True)

# ==============================================================================
# 4. APLICAÇÃO PRINCIPAL LOGADA (ESTRUTURA DE NAVEGAÇÃO POR ABAS)
# ==============================================================================
else:
    # Top Bar / Cabeçalho Fixo
    u_info = st.session_state.user_info
    st.markdown(f"""
    <div class="app-header">
        <div>
            <span style="font-size:0.75rem; color:#94A3B8;">Paz do Senhor,</span><br>
            <strong style="color:#F59E0B;">{u_info.get('nome')}</strong>
        </div>
        <div style="font-size: 1.4rem;">🔔<span style="color:#EF4444; font-size:0.8rem;">●</span></div>
    </div>
    """, unsafe_allow_html=True)
    
    # Abas Inferiores de Navegação
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Início",
        "📅 Frequência",
        "🥋 Tatame",
        "🏆 Campeão & Fé",
        "📊 Secretaria"
    ])

    # --------------------------------------------------------------------------
    # ABA 1: INÍCIO (DASHBOARD DOS PAIS & DESENVOLVIMENTO)
    # --------------------------------------------------------------------------
    with tab1:
        st.markdown("### 🎂 Aniversariante do Dia & Mês")
        st.info("🎉 *Hoje é aniversário do aluno Alvaro Barbosa!* (30/08)")
        link_bday = gerar_link_whatsapp("5567998411953", "Parabéns Alvaro! Que o Senhor te abençoe grandemente e te dê muita sabedoria no tatame e na vida! 🎉🌱")
        if link_bday:
            st.markdown(f"[🥳 Enviar Mensagem de Parabéns no WhatsApp]({link_bday})")
            
        st.markdown("---")
        st.markdown("### 🥋 Perfil dos Seus Filhos")
        filhos = u_info.get('filhos', [])
        if not filhos and not df_cadastro.empty:
            filhos = df_cadastro.head(2).to_dict(orient='records')
            
        for f in filhos:
            nome_aluno = f.get('Nome do Aluno', 'Aluno')
            faixa = f.get('Faixa Graus', 'Branca')
            saude = f.get('Histórico de Saúde', 'Não')
            batizado = True if "14" in str(f.get('Nº', '')) or idx_aluno % 2 == 0 else False
            
            st.markdown(f"""
            <div class="gold-card">
                <h4>{nome_aluno} {'⭐' if batizado else ''}</h4>
                <span class="gold-badge">Faixa {faixa}</span>
                <p style="font-size:0.85rem; margin-top:8px;">
                    <strong>Biometria Facial:</strong> 🟢 Cadastrada<br>
                    <strong>Saúde:</strong> {'🚨 ' + str(saude) if str(saude).lower() != 'não' else '🔒 Apto para os treinos'}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🏫 Acompanhamento Escolar & Familiar (Mensal)")
        with st.expander("📝 Preencher Avaliação Mensal do Filho"):
            aluno_sel = st.selectbox("Selecione o Filho:", [f.get('Nome do Aluno') for f in filhos] if filhos else ["Alvaro Barbosa"])
            esc = st.select_slider("1. Desenvolvimento Escolar:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            fam = st.select_slider("2. Desenvolvimento Familiar:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Ótimo")
            soc = st.select_slider("3. Desenvolvimento Social:", options=["Ruim", "Regular", "Bom", "Ótimo"], value="Bom")
            obs = st.text_area("Observações para o Mestre:")
            if st.button("Salvar Avaliação Privada"):
                st.success("Avaliação salva! Apenas você e o Mestre possuem acesso a estas informações.")

    # --------------------------------------------------------------------------
    # ABA 2: FREQUÊNCIA & PREVISÃO DE GRADUAÇÃO
    # --------------------------------------------------------------------------
    with tab2:
        st.markdown("### 📅 Frequência — Agosto 2026")
        c1, c2 = st.columns(2)
        c1.metric("Presenças Registradas", "5 Dias")
        c2.metric("Total de Treinos", "9 Aulas")
        
        st.markdown("#### Histórico de Treino (Terças e Quintas)")
        grid_cal = [
            {"Data": "04/08", "Status": "🟡 Presença (X)", "Devocional": "Disciplina"},
            {"Data": "06/08", "Status": "🟡 Presença (X)", "Devocional": "Respeito"},
            {"Data": "11/08", "Status": "🔴 Falta", "Devocional": "Honestidade"},
            {"Data": "13/08", "Status": "🟡 Presença (X)", "Devocional": "Amor"},
        ]
        st.dataframe(pd.DataFrame(grid_cal), use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 🥋 Previsão de Graduação (Sugestão do Mestre)")
        if u_info.get('tipo') == 'mestre':
            st.date_input("Mestre: Definir Próxima Data de Exame de Faixa:", datetime.date(2026, 12, 15))
            st.button("Atualizar Calendário da Academia")
        else:
            st.info("🎯 *Próxima Cerimônia de Graduação e Grau:* 15/12/2026\n\n*Frequência mínima requerida: 80% dos treinos.*")

    # --------------------------------------------------------------------------
    # ABA 3: TATAME, BIOMETRIA & DUPLA CHAMADA
    # --------------------------------------------------------------------------
    with tab3:
        st.markdown("### 🥋 Chamada Rápida do Tatame — Instrutor London")
        
        modo_chamada = st.radio("Escolha a forma de chamada:", ["1. Toque na Foto do Aluno", "2. Foto Coletiva (Visão Computacional IA)"])
        
        if modo_chamada == "1. Toque na Foto do Aluno":
            st.caption("Toque na foto ou caixa do aluno para registrar a presença instantânea:")
            if not df_cadastro.empty:
                for idx, row in df_cadastro.head(10).iterrows():
                    aluno_n = row.get('Nome do Aluno', 'Aluno')
                    saude_aluno = str(row.get('Histórico de Saúde', 'Não'))
                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        st.markdown(f"*{aluno_n}*")
                        if saude_aluno.lower() != 'não' and saude_aluno != '':
                            st.caption(f"<span style='color:#EF4444;'>🚨 Alerta Saúde: {saude_aluno}</span>", unsafe_allow_html=True)
                    with col_b:
                        st.checkbox("Presente", key=f"p_box_{idx}")
        else:
            st.caption("Carregue a foto coletiva tirada ao final da aula:")
            f_aula = st.file_uploader("Capturar ou subir foto do tatame:", type=["jpg", "png"])
            if f_aula:
                st.image(f_aula, use_container_width=True)
                st.warning("⚠️ *IA Notifica:* 2 alunos na foto ainda não possuem Biometria Facial cadastrada. Clique abaixo para vincular.")
                st.button("📸 Cadastrar Biometria dos Alunos Incompletos")
                
        if st.button("FINALIZAR E SALVAR CHAMADA"):
            st.success("Chamada salva com sucesso na planilha oficial!")

    # --------------------------------------------------------------------------
    # ABA 4: CAMPEÃO, EVANGELISMO & PALESTRAS
    # --------------------------------------------------------------------------
    with tab4:
        sub_aba = st.radio("Selecione a área:", ["🏆 Área do Campeão", "✝️ Evangelismo & Células", "📢 Palestras Socioeducativas"], horizontal=True)
        
        if sub_aba == "🏆 Área do Campeão":
            st.markdown("#### 🏆 Histórico de Campeonatos e Pódios")
            st.markdown("""
            <div class="gold-card">
                <h5>🥇 Campeonato Estadual de Jiu-Jitsu 2026</h5>
                <p style="font-size:0.85rem;">
                    <strong>Alvaro Barbosa:</strong> Medalha de Ouro 🥇 (Infantil B)<br>
                    <strong>Arthur Barbosa:</strong> Medalha de Prata 🥈 (Mirim)<br>
                    <strong>Classificação Geral da Academia:</strong> 2º Lugar Geral por Equipes 🏆
                </p>
            </div>
            """, unsafe_allow_html=True)
            st.info("📅 *Próximo Campeonato:* Copa Corumbá Open de Jiu-Jitsu — 20/11/2026")

        elif sub_aba == "✝️ Evangelismo & Células":
            st.markdown("#### 💬 Células Familiares — IEQ Guaicurus")
            celulas = [
                {"nome": "Célula Filhos da Promessa", "lider": "Diácono Roberto", "fone": "5567999990000"},
                {"nome": "Célula Sementes de Fé", "lider": "Líder Mary", "fone": "5567999990001"},
                {"nome": "Célula Amigos do Tatame", "lider": "Odiselma", "fone": "5567999990002"},
                {"nome": "Célula Guerreiros da Luz", "lider": "Marcos", "fone": "5567999990003"},
                {"nome": "Célula Graça e Vida", "lider": "Luciana", "fone": "5567999990004"},
            ]
            for c in celulas:
                st.markdown(f"*{c['nome']}* (Líder: {c['lider']})")
                link_c = gerar_link_whatsapp(c['fone'], f"Olá {c['lider']}, paz do Senhor! Gostaria de saber mais e participar da sua Célula!")
                if link_c:
                    st.markdown(f"[💬 Quero Participar desta Célula]({link_c})")
                st.markdown("---")
                
            st.markdown("#### 📊 Impacto Espiritual (Anônimo)")
            st.metric("Total de Pedidos de Oração e Vidas Encaminhadas", f"{st.session_state.pedidos_oracao_count} Vidas")

        elif sub_aba == "📢 Palestras Socioeducativas":
            st.markdown("#### 🧠 Palestras Realizadas & Treinamentos")
            st.markdown("""
            <div class="gold-card">
                <h5>Inteligência Emocional e Combate ao Bullying</h5>
                <p style="font-size:0.85rem;">
                    <strong>Palestrante:</strong> Dr. Marco Aurélio<br>
                    <strong>Impacto:</strong> 88% de adesão das famílias cadastradas
                </p>
            </div>
            """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # ABA 5: SECRETARIA, MAPA SOCIAL & GESTÃO DE SENHAS
    # --------------------------------------------------------------------------
    with tab5:
        st.markdown("### 📊 Secretaria & Diretoria")
        
        # Acesso Restrito ao Mestre e Diretoria para Emissão de Relatório e Senhas
        if u_info.get('tipo') == 'mestre':
            st.success("👑 *Nível de Permissão do Usuário:* Mestre / Diretoria (Acesso Total)")
            
            st.markdown("---")
            st.markdown("#### 🔐 Gestão de Senhas dos Pais (Exclusivo Mestre)")
            c_p1, c_p2, c_p3 = st.columns([2, 2, 1])
            with c_p1:
                fone_pai = st.text_input("WhatsApp do Pai (Apenas números):", "67998411953")
            with c_p2:
                nova_senha_pai = st.text_input("Definir Nova Senha:", "123456")
            with c_p3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Salvar Senha"):
                    st.session_state.senhas_pais[fone_pai] = nova_senha_pai
                    st.success("Senha atualizada!")

            st.markdown("---")
            st.markdown("#### 📄 Emissão de Relatórios Oficiais")
            st.button("📄 Exportar Relatório Completo da Academia (PDF)")
            
        else:
            st.info("ℹ️ Dados gerais da secretaria e mapa social de Corumbá-MS:")

        st.markdown("---")
        st.markdown("#### 🗺️ Mapa Social de Abrangência Territorial (Corumbá-MS)")
        mapa_df = pd.DataFrame({
            'Bairro': ['Guaicurus', 'Nova Corumbá', 'Centro', 'Guarani'],
            'Porcentagem (%)': [45, 30, 15, 10]
        })
        st.bar_chart(mapa_df.set_index('Bairro'))
        
        st.markdown("---")
        if st.button("🚪 Sair do Aplicativo"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
