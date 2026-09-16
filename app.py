import streamlit as st
import pandas as pd
import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🥋",
    layout="wide"
)

# --- CSS CUSTOMIZADO (Alta Contraste, Fontes Maiores e Estilo do Sininho) ---
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-size: 18px !important;
        color: #111111 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .stButton>button {
        font-size: 18px !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        background-color: #0d6efd !important;
        color: white !important;
    }
    .notification-bell-red {
        font-size: 28px;
        color: #dc3545;
        cursor: pointer;
        animation: blinker 1s linear infinite;
    }
    .notification-bell-gold {
        font-size: 28px;
        color: #ffc107;
        cursor: pointer;
    }
    @keyframes blinker {
        50% { opacity: 0; }
    }
    .card-box {
        background-color: #f8f9fa;
        border: 2px solid #dee2e6;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# --- INICIALIZAÇÃO DO ESTADO DA SESSÃO ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'has_notification' not in st.session_state:
    st.session_state.has_notification = True
if 'notification_seen' not in st.session_state:
    st.session_state.notification_seen = False

# --- TELA DE LOGIN ---
if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Exibição da Logo na Tela de Login
        try:
            st.image("logo_projeto_sementes.png", width=220)
        except:
            st.warning("Logo 'logo_projeto_sementes.png' não encontrada na raiz do projeto.")
        
        st.markdown("<h2 style='text-align: center;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center;'>IEQ Guaicurus - Corumbá/MS</h4>", unsafe_allow_html=True)
        
        user_input = st.text_input("Usuário / Responsável")
        password_input = st.text_input("Senha", type="password")
        
        if st.button("Acessar Sistema"):
            if password_input == "mestre123":
                st.session_state.logged_in = True
                st.session_state.user_role = "Mestre/Diretoria"
                st.rerun()
            elif password_input == "pais123":
                st.session_state.logged_in = True
                st.session_state.user_role = "Responsável"
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    st.stop()

# --- CABEÇALHO SUPERIOR E NOTIFICAÇÕES ---
top_col1, top_col2 = st.columns([8, 2])

with top_col1:
    if st.session_state.user_role == "Mestre/Diretoria":
        st.markdown("## *A Paz seja convosco, Mestre London / Diretoria*")
    else:
        st.markdown("## *A Paz seja convosco, Responsável*")

with top_col2:
    # Lógica do Sininho Dourado / Vermelho
    if not st.session_state.notification_seen:
        if st.button("🔔 Notificação Pendente", key="bell_red"):
            st.session_state.notification_seen = True
            st.rerun()
    else:
        if st.button("🔔 Notificações", key="bell_gold"):
            st.session_state.notification_seen = False
            st.rerun()

if st.session_state.notification_seen:
    st.info("📌 *Mensagem do Sistema:* Lembramos todos da nossa próxima Palestra Socioeducativa com a Equipe PROERD no dia 29/09/2026 às 19:00h no Dojo!")

st.divider()

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.title("Navegação Principal")
menu = st.sidebar.radio("Selecione a Aba:", [
    "Início / Perfil",
    "Matrícula",
    "Frequência",
    "Chamada no Dojo",
    "Campeonato e Pódio",
    "Campeão e Fé",
    "Palestras & Treinamentos",
    "Secretaria & Gestão",
    "Relatórios",
    "Mapa Social de Expansão"
])

if st.sidebar.button("Sair do Sistema"):
    st.session_state.logged_in = False
    st.rerun()

# --- 1. ABA INÍCIO / PERFIL ---
if menu == "Início / Perfil":
    st.header("📌 Painel Inicial")
    
    tab_perfil, tab_aniv, tab_escola = st.tabs(["Perfil do Aluno", "Aniversariantes do Mês", "Acompanhamento Escolar/Familiar"])
    
    with tab_perfil:
        st.subheader("Perfil do Aluno")
        if st.session_state.user_role == "Mestre/Diretoria":
            aluno_sel = st.selectbox("Selecione o Aluno para visualizar:", ["Alvaro Barbosa", "Elton Araujo", "Aysla Barbosa", "Nathan Carvalho", "Gustavo Camargo"])
        else:
            aluno_sel = "Nathan Carvalho" # Exemplo de vinculo para responsável
            st.info(f"Visualizando dados do dependente: *{aluno_sel}*")
            
        c1, c2 = st.columns([1, 3])
        with c1:
            st.image("https://via.placeholder.com/150", caption=f"Foto de {aluno_sel}")
            st.markdown("⭐ *Status Espiritual:* Batizado")
        with c2:
            st.write(f"*Nome do Aluno:* {aluno_sel}")
            st.write("*Faixa:* Branca (1º Grau)")
            st.write("*Escola onde estuda:* E.M. Pedro Paulo de Medeiros")
            st.write("*Responsável Legal:* London Janssen Santos de Carvalho")
            st.write("*Outros Irmãos no Projeto:* Isadora Souza")
            st.write("*Documentação:* OK")

    with tab_aniv:
        st.subheader("🎂 Aniversariantes do Mês")
        st.markdown("*Datas de Comemoração em Grupo (às 19h):*")
        st.write("• *Setembro:* 29/09/2026 (Terça-feira)")
        st.write("• *Outubro:* 29/10/2026 (Quinta-feira)")
        st.write("• *Novembro:* 26/11/2026 (Quinta-feira)")
        st.write("• *Dezembro:* 31/12/2026 (Quinta-feira)")
        st.divider()
        
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            st.image("https://via.placeholder.com/100", width=100)
            st.write("*Alvaro Barbosa* - Aniversário: 30/08")
            st.success("✉️ Mensagem Evangélica do Projeto Sementes enviada automaticamente!")
        with col_a2:
            st.image("https://via.placeholder.com/100", width=100)
            st.write("*Nathan Carvalho* - Aniversário: 04/03")

    with tab_escola:
        st.subheader("📝 Acompanhamento Escolar e Familiar (Mensal)")
        with st.form("form_escola"):
            st.selectbox("Selecione o Aluno:", ["Nathan Carvalho", "Alvaro Barbosa"])
            st.select_slider("Desempenho Escolar neste Mês:", options=["Precisa Melhorar", "Regular", "Bom", "Excelente"])
            st.select_slider("Comportamento em Casa / Respeito aos Pais:", options=["Precisa Melhorar", "Regular", "Bom", "Excelente"])
            st.text_area("Observações da Família ou Notas Escolares:")
            st.form_submit_button("Enviar Acompanhamento Mensal")

# --- 2. ABA MATRÍCULA ---
elif menu == "Matrícula":
    st.header("📋 Ficha de Matrícula e Termos")
    
    with st.form("form_matricula"):
        st.subheader("1. Dados do Aluno")
        st.text_input("Nome Completo do Aluno")
        st.text_input("Escola onde estuda")
        st.multiselect("Outras crianças/irmãos sob responsabilidade deste responsável:", ["Alvaro Barbosa", "Aysla Barbosa", "Nathan Carvalho", "Isadora Souza"])
        
        st.subheader("2. Dados do Pai / Responsável")
        st.text_input("Nome do Pai / Responsável")
        st.text_input("Carteira de Identidade Nacional (CIN)")
        st.text_input("Telefone de Contato (WhatsApp)")
        
        st.subheader("3. Declaração de Aptidão Física e Histórico de Saúde")
        st.write("Marque as condições clínicas aplicáveis ao menor:")
        q1 = st.checkbox("O menor possui algum problema cardíaco ou de pressão?")
        q2 = st.checkbox("O menor sofre de asma, bronquite ou problemas respiratórios?")
        q3 = st.checkbox("O menor possui alguma lesão óssea, muscular ou articular crônica?")
        q4 = st.checkbox("O menor faz uso regular de algum medicamento controlado?")
        q5 = st.checkbox("O menor possui alergia a algum medicamento ou substância?")
        q6 = st.checkbox("O menor já sofreu desmaios ou tonturas durante exercícios físicos?")
        
        st.multiselect("Selecione as condições identificadas:", ["Pressão/Coração", "Asma/Bronquite", "Lesão Articular", "Medicamento Contínuo", "Alergias"])
        st.text_area("Caso tenha marcado SIM em alguma das opções, especifique aqui:")
        
        st.info("""
        *DECLARAÇÃO:* Declaro, sob as penas da lei, que o menor acima qualificado goza de boa saúde e encontra-se apto para a prática de exercícios físicos e treinos de artes marciais (Jiu-Jitsu). Comprometo-me a informar imediatamente a coordenação do projeto caso ocorra qualquer alteração em seu estado de saúde no decorrer do período letivo.
        
        Corumbá - MS, __ de _______ de 2026.
        """)
        
        st.subheader("4. Termo de Autorização de Uso de Imagem e Voz")
        st.write("""
        Eu, responsável legal pelo menor matriculado, AUTORIZO de forma inteiramente gratuita, definitiva e irrevogável, a Igreja do Evangelho Quadrangular a utilizar a imagem e voz do referido menor, capturadas em fotos ou vídeos durante as atividades do projeto "Sementes".
        Esta autorização é concessiva para fins de divulgação institucional e prestação de contas das ações da igreja junto à comunidade, podendo ser veiculada em boletins internos, redes sociais oficiais da igreja, murais e apresentações em cultos locais, vedada qualquer forma de exploração comercial ou fins lucrativos.
        """)
        
        st.subheader("5. Termo de Autorização de Participação")
        st.write("""
        Eu, acima identificado(a) como responsável legal, AUTORIZO expressamente o menor sob minha tutela a participar de forma gratuita das aulas de Jiu-Jitsu do Projeto Social "Sementes", promovido por esta Igreja local.
        Declaro estar ciente de que as atividades envolvem treinos físicos de contato próprio da modalidade e que serão supervisionadas por membros habilitados da comunidade eclesial.
        
        Corumbá - MS, __ de _______ de 2026.
        """)
        
        st.checkbox("Li e aceito expressamente todos os termos e autorizações acima.")
        st.form_submit_button("Finalizar Matrícula / Enviar Assinatura")

# --- 3. ABA FREQUÊNCIA ---
elif menu == "Frequência":
    st.header("📊 Painel Geral de Frequência e Frequência Individual")
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Treinos (Mês)", "12 Aulas")
    m2.metric("Total de Alunos Presentes", "21/23")
    m3.metric("Frequência Média", "91.3%")
    
    st.divider()
    st.subheader("Acompanhamento por Aluno")
    aluno_freq = st.selectbox("Selecione o Aluno:", ["Alvaro Barbosa", "Elton Araujo", "Nathan Carvalho"])
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.write(f"*Frequência Mensal de {aluno_freq}:* 95%")
        if st.session_state.user_role == "Mestre/Diretoria":
            evolucao = st.radio("Evolução do Aluno (Definida pelo Mestre):", ["Ótima", "Bom", "Regular"])
            st.button("Atualizar Evolução")
        else:
            st.write("*Evolução Técnica:* Ótima")
            
    with col_f2:
        st.write("*Histórico de Treinos (Dias de Aula):*")
        df_hist = pd.DataFrame({
            "Data (Dia/Mês/Ano)": ["01/09/2026", "03/09/2026", "08/09/2026", "10/09/2026", "15/09/2026"],
            "Status": ["🟢 Presença", "🟢 Presença", "🔴 Falta", "🟢 Presença", "🟢 Presença"]
        })
        st.table(df_hist)

# --- 4. ABA CHAMADA NO DOJO ---
elif menu == "Chamada no Dojo":
    st.header("🥋 Chamada Rápida de Presença no Dojo - Professores")
    
    opcao_chamada = st.radio("Selecione o método de chamada:", ["1. Chamada Individual por Fotos", "2. Foto Coletiva (Reconhecimento)"])
    
    if "1." in opcao_chamada:
        st.subheader("Clique na foto do aluno para registrar presença no treino de hoje:")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.image("https://via.placeholder.com/120", width=120)
            st.button("🟢 Alvaro Barbosa")
        with c2:
            st.image("https://via.placeholder.com/120", width=120)
            st.button("🟢 Elton Araujo")
        with c3:
            st.image("https://via.placeholder.com/120", width=120)
            st.button("🟢 Nathan Carvalho")
            
    else:
        st.subheader("Foto Coletiva do Tatame")
        st.info("Faça o upload da foto do treino coletivo para registrar automaticalmente:")
        st.file_uploader("Enviar Foto da Turma", type=["jpg", "png", "jpeg"])
        st.write("*Resultado do Escaneamento:*")
        st.write("🟢 *21 Alunos Cadastrados Detectados (Presença Registrada)*")
        st.write("🔴 *2 Alunos com Falta de Cadastro Detectados (Pendência)*")

# --- 5. ABA CAMPEONATO E PÓDIO ---
elif menu == "Campeonato e Pódio":
    st.header("🏆 Histórico de Campeonatos e Pódios")
    
    st.markdown("""
    <div class='card-box'>
        <h3>🥇 Destaque Geral do Projeto Sementes</h3>
        <p><b>Campeonato Estadual de Jiu-Jitsu MS 2026:</b> 2º Lugar Geral por Equipes!</p>
        <p><b>Copa Corumbá de Artes Martiais:</b> 1º Lugar Geral no Infantil!</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.subheader("Pódio Individual do Aluno")
    aluno_camp = st.selectbox("Selecione o Aluno para consultar o histórico:", ["Nathan Carvalho", "Alvaro Barbosa"])
    
    df_podio = pd.DataFrame({
        "Campeonato": ["Copa Corumbá 2026", "Estadual MS 2026"],
        "Categoria": ["Infantil B- Pesado", "Infantil B - Pesado"],
        "Colocação": ["🥇 1º Lugar (Ouro)", "🥈 2º Lugar (Prata)"],
        "Equipe Representada": ["Projeto Sementes", "Projeto Sementes"]
    })
    st.table(df_podio)

# --- 6. ABA CAMPEÃO E FÉ ---
elif menu == "Campeão e Fé":
    st.header("🙏 Campeão e Fé")
    
    tab_oracao, tab_batismo, tab_agenda = st.tabs(["Pedido de Oração", "Impacto Espiritual", "Agenda da Igreja & Células"])
    
    with tab_oracao:
        st.subheader("Pedido de Oração ao Nosso Pastor Joel")
        with st.form("form_oracao"):
            st.text_input("Seu Nome (Opcional):")
            st.text_area("Escreva aqui seu pedido de oração para o Pastor Joel:")
            st.form_submit_button("Enviar Pedido ao Pastor Joel")
            
    with tab_batismo:
        st.subheader("Impacto Espiritual")
        st.write("Decisão de Fé:")
        if st.button("✝️ Eu quero Batizar!"):
            st.success("Glória a Deus! Sua decisão foi enviada para a liderança pastoral.")
            
        st.divider()
        st.markdown("### *Métricas de Impacto Espiritual (Geral)*")
        st.caption("Dados consolidados sem exibição de nomes por privacidade.")
        e1, e2 = st.columns(2)
        e1.metric("Pessoas Integradas em Células", "38")
        e2.metric("Pessoas Batizadas através do Projeto", "14")

    with tab_agenda:
        st.subheader("⛪ Programação IEQ Guaicurus")
        st.write("• *Cultos Principais:* Todo Domingo às 18:30h")
        st.write("• *Cronograma de Células:* Terças e Quintas às 20:00h")
        st.write("• *Festividades da Igreja:* Redes sociais e avisos internos")

# --- 7. ABA PALESTRAS & TREINAMENTOS ---
elif menu == "Palestras & Treinamentos":
    st.header("📢 Palestras Socioeducativas & Treinamentos Básicos")
    
    col_p1, col_p2 = st.columns(2)
    
    with col_p1:
        st.subheader("🎤 Palestras Socioeducativas")
        st.write("*29/09/2026 às 19h:* Escolhas Saudáveis - PROERD")
        st.write("*29/10/2026 às 19h:* Inteligência Emocional - Psic. Eva Mateus")
        st.write("*26/11/2026 às 19h:* O Valor do Estudo - Ped. Luis e Edima")
        if st.button("📲 Notificar Pais sobre Próxima Palestra"):
            st.success("Notificação enviada aos responsáveis com sucesso!")

    with col_p2:
        st.subheader("🥋 Treinamentos Básicos")
        st.write("*Assunto:* Regras do Jiu-Jitsu e Conduta no Dojo")
        st.write("*Instrutor:* Mestre London / Antonio")
        st.write("*Data e Hora:* 22/09/2026 às 18:00h")
        if st.button("📲 Notificar Pais sobre Treinamento"):
            st.success("Lembrete de Treinamento enviado com sucesso!")

# --- 8. ABA SECRETARIA & GESTÃO ---
elif menu == "Secretaria & Gestão":
    if st.session_state.user_role != "Mestre/Diretoria":
        st.warning("Acesso restrito à Diretoria e ao Mestre.")
    else:
        st.header("⚙️ Central de Secretaria e Senhas")
        
        st.subheader("1. Central de Notificações de Pendências Documentais")
        df_pend = pd.DataFrame({
            "Aluno": ["Erica Rodrigues", "Guilherme Arruda", "Bianca da Silva"],
            "Pendência": ["Falta Endereço", "Falta Documento", "Falta CIN"],
            "Ação": ["Enviar Lembrete WhatsApp", "Enviar Lembrete WhatsApp", "Enviar Lembrete WhatsApp"]
        })
        st.table(df_pend)
        if st.button("Marcar Pendência como Sanada"):
            st.success("Lista atualizada automaticamente.")
            
        st.divider()
        st.subheader("2. Gestão de Senhas e Acessos")
        st.write("Gerenciar e gerar senhas individuais para Pais e Diretoria:")
        st.text_input("Usuário do Aluno/Pai:")
        st.button("Gerar Nova Senha Automática")
        
        st.divider()
        st.subheader("3. Excluir Acesso de Aluno Desistente")
        aluno_excluir = st.selectbox("Selecione o aluno para remover acesso:", ["Alvaro Barbosa", "Elton Araujo"])
        if st.button("❌ Remover Acesso do Aplicativo"):
            st.error(f"Acesso de {aluno_excluir} removido do sistema.")

# --- 9. ABA RELATÓRIOS ---
elif menu == "Relatórios":
    st.header("📄 Emissão de Relatórios Customizados")
    
    st.selectbox("Escolha o tipo de relatório para gerar:", [
        "Relatório Completo de Alunos e Matrículas",
        "Relatório de Frequência Mensal",
        "Relatório de Pendências Documentais",
        "Relatório de Desempenho Escolar e Familiar",
        "Relatório de Impacto Espiritual e Batismos",
        "Relatório do Cronograma de Eventos e Palestras",
        "Relatório de Pódios e Campeonatos"
    ])
    
    if st.button("📥 Gerar e Baixar Relatório (PDF/Excel)"):
        st.success("Relatório gerado com sucesso! Clique para realizar o download.")

# --- 10. ABA MAPA SOCIAL DE EXPANSÃO ---
elif menu == "Mapa Social de Expansão":
    st.header("📍 Mapa Social de Expansão do Projeto Sementes IEQ Guaicurus em Corumbá-MS")
    st.caption("Atualização Diária Automática")
    
    col_m1, col_m2 = st.columns([1, 2])
    
    with col_m1:
        st.write("*Distribuição por Bairro:*")
        df_bairro = pd.DataFrame({
            "Bairro": ["Guaicurus", "Centro", "Nova Corumbá", "Cristo Redentor"],
            "Nº Alunos": [12, 5, 4, 2],
            "Porcentagem": ["52.2%", "21.7%", "17.4%", "8.7%"]
        })
        st.table(df_bairro)
        st.markdown("*Total de Alunos Mapeados:* 23 (100%)")
        
    with col_m2:
        st.subheader("Gráfico de Cobertura de Bairros")
        st.bar_chart(df_bairro.set_index("Bairro")["Nº Alunos"])
