# ==========================================
# TELA DE AUTENTICAÇÃO / LOGIN
# ==========================================
if not st.session_state.get("authenticated", False):
    st.markdown("<h1 style='text-align: center;'>PROJETO SEMENTES</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #aaaaaa;'>Iniciativa Voluntária de Jiu-Jitsu e Apoio à Família<br>Cessão de Espaço Comunitário: IEQ Guaicurus</p>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 1ª CORREÇÃO: Mudança do rótulo para 'Login do Responsável'
    usuario_input = st.text_input("Login do Responsável", value="London", key="login_usuario")
    
    # 1ª CORREÇÃO: Injeção de CSS para garantir contraste visual na revelação da senha
    st.markdown("""
        <style>
        /* Ajuste de cor do texto digitado no campo de senha */
        input[type="password"], input[type="text"] {
            color: #FFFFFF !important;
            background-color: #1E293B !important;
        }
        /* Ajuste do botão de alternar visibilidade/olho no Streamlit */
        button[aria-label="Show password"], button[aria-label="Hide password"] {
            color: #F59E0B !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    senha_input = st.text_input("Senha de Acesso", type="password", key="login_senha")
    
    if st.button("ENTRAR NO APLICATIVO", use_container_width=True):
        # Validação de credenciais (ajuste conforme suas chaves salvas)
        if senha_input == "1234":  # Altere para a sua senha real ou variável
            st.session_state["authenticated"] = True
            st.session_state["usuario"] = usuario_input
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")
    st.stop()


# ==========================================
# ESTRUTURA FIXA COM NAVEGAÇÃO E CABEÇALHO
# ==========================================

# 4ª CORREÇÃO: CSS para congelar e fixar a barra superior e abas no topo da tela durante a rolagem
st.markdown("""
    <style>
    /* Fixar cabeçalho e abas no topo */
    div[data-testid="stHeader"] {
        z-index: 999;
    }
    div[data-testid="stTabs"] > div:first-child {
        position: sticky;
        top: 3.5rem;
        background-color: #0E1117;
        z-index: 998;
        padding-top: 10px;
        padding-bottom: 10px;
        border-bottom: 2px solid #333333;
    }
    /* Estilo para a caixa do Mural de Avisos (3ª Correção) */
    .mural-avisos {
        background: linear-gradient(135deg, #FF416C 0%, #FF4B2B 100%);
        color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0px 4px 15px rgba(255, 65, 108, 0.4);
        margin-top: 15px;
        margin-bottom: 20px;
    }
    .mural-avisos h3 {
        color: #FFFFFF !important;
        margin-top: 0;
    }
    </style>
""", unsafe_allow_html=True)

# 2ª CORREÇÃO: Data e Hora atualizada de acesso na tela inicial
agora = datetime.now()
data_hora_formatada = agora.strftime("%d/%m/%Y - %H:%M:%S")

st.markdown(f"### A Paz seja convosco, Mestre {st.session_state.get('usuario', 'London')} / Diretoria")
st.caption(f"📅 *Acesso em:* {data_hora_formatada}")

# Definição das Abas do Aplicativo
aba1, aba2, aba3, aba4 = st.tabs(["🏠 Início", "📋 Lista de Presença", "🥋 Graduação", "📅 Cronograma"])

# ==========================================
# ABA 1: INÍCIO
# ==========================================
with aba1:
    # Leitura do Cronograma de Eventos a partir do Excel
    caminho_excel = "Controle de Presença e Graduação Projeto Sementes.xlsx"
    
    eventos_hoje = []
    data_hoje_str = agora.strftime("%d/%m/%Y")
    
    try:
        df_cronograma = pd.read_excel(caminho_excel, sheet_name="Cronograma de Eventos")
        # Garantir tratamento das datas da coluna 'Data Evento'
        if 'Data Evento' in df_cronograma.columns:
            df_cronograma['Data_Fmt'] = pd.to_datetime(df_cronograma['Data Evento'], errors='coerce').dt.strftime("%d/%m/%Y")
            # Filtrar eventos que coincidem com a data de hoje
            df_hoje = df_cronograma[df_cronograma['Data_Fmt'] == data_hoje_str]
            eventos_hoje = df_hoje.to_dict('records')
    except Exception as e:
        st.warning("Não foi possível carregar o Cronograma de Eventos diretamente da planilha.")

    # Estado da notificação do sino
    if "visualizou_notificacao" not in st.session_state:
        st.session_state["visualizou_notificacao"] = False

    tem_evento_hoje = len(eventos_hoje) > 0

    # 3ª CORREÇÃO: Lógica do Sino Dourado vs Vermelho Piscando
    st.markdown("""
        <style>
        @keyframes piscar {
            0% { opacity: 1.0; transform: scale(1); }
            50% { opacity: 0.4; transform: scale(1.1); }
            100% { opacity: 1.0; transform: scale(1); }
        }
        .sino-vermelho {
            color: #FF0000;
            font-size: 28px;
            animation: piscar 1s infinite;
            cursor: pointer;
            border: none;
            background: none;
        }
        .sino-dourado {
            color: #FFD700;
            font-size: 28px;
            cursor: pointer;
            border: none;
            background: none;
        }
        </style>
    """, unsafe_allow_html=True)

    # Definir ícone e estado visual
    if tem_evento_hoje and not st.session_state["visualizou_notificacao"]:
        label_sino = "🔔 (Novo Evento Hoje!)"
        estilo_sino = "sino-vermelho"
    else:
        label_sino = "🔔"
        estilo_sino = "sino-dourado"

    col_sino, col_vazia = st.columns([1, 5])
    with col_sino:
        if st.button(label_sino, key="btn_sino_notificacao"):
            st.session_state["visualizou_notificacao"] = True
            st.session_state["exibir_mural"] = True
            st.rerun()

    # 3ª CORREÇÃO: Mural de avisos chamativo exibido ao clicar no sino
    if st.session_state.get("exibir_mural", False):
        st.markdown("""
            <div class='mural-avisos'>
                <h2>📢 MURAL DE AVISOS - EVENTOS DE HOJE</h2>
            </div>
        """, unsafe_allow_html=True)
        
        if tem_evento_hoje:
            for ev in eventos_hoje:
                st.info(f"📌 *Evento:* {ev.get('Evento', 'N/A')}\n\n"
                        f"⏰ *Horário:* {ev.get('Hora Evento', 'N/A')}\n\n"
                        f"👤 *Responsável/Instrutor:* {ev.get('Palestrante / Instrutor / Equipe / Igreja', 'N/A')}")
        else:
            st.success("🎉 Não há eventos agendados para a data de hoje no Cronograma!")
            
        if st.button("Fechar Mural de Avisos"):
            st.session_state["exibir_mural"] = False
            st.rerun()

    st.markdown("---")

    # Informações institucionais e resumo
    st.markdown("""
    * *Dia de Palestra:* Treinamentos socioeducativos e saúde mental.
    * *Dia de Treinamento:* Capacitação técnica e primeiros socorros.
    * *Dia de Campeonato:* Competições regionais e pódios.
    * *Dia de Aniversariantes do Mês:* Festividades com as famílias.
    * *Dia de Graduação:* Cerimônia de entrega de faixas e graus.
    * *Dia de Ação Social:* Evangelismo e acolhimento comunitário.
    """)

    st.subheader("📅 Próximos Eventos do Cronograma")
    try:
        st.dataframe(df_cronograma[['Evento', 'Data Evento', 'Hora Evento', 'Palestrante / Instrutor / Equipe / Igreja']], use_container_width=True)
    except Exception:
        st.write("Carregando eventos...")
