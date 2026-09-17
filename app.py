import streamlit as st
import pandas as pd
import datetime
import urllib.parse
from zoneinfo import ZoneInfo

try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    CANVAS_DISPONIVEL = False


# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==============================================================================

st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)


# ==============================================================================
# 2. ESTILIZAÇÃO
# DESIGN ORIGINAL MANTIDO
# ==============================================================================

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

    #MainMenu,
    header,
    footer,
    [data-testid="stSidebar"] {
        display: none !important;
    }


    /* ==========================================================
       CARTÕES
       ========================================================== */

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


    /* ==========================================================
       SINO VERMELHO PISCANDO
       ========================================================== */

    @keyframes pulse-red {

        0% {
            transform: scale(1);
            filter: drop-shadow(0 0 2px #EF4444);
        }

        50% {
            transform: scale(1.25);
            filter: drop-shadow(0 0 12px #EF4444);
        }

        100% {
            transform: scale(1);
            filter: drop-shadow(0 0 2px #EF4444);
        }

    }

    .sino-vermelho {
        display: inline-block;
        color: #EF4444 !important;
        animation: pulse-red 1.1s infinite;
        font-size: 1.55rem;
        font-weight: 900;
        text-align: center;
    }


    /* ==========================================================
       SINO DOURADO
       ========================================================== */

    .sino-dourado {
        display: inline-block;
        color: #F59E0B !important;
        font-size: 1.55rem;
        font-weight: 900;
        text-align: center;
    }


    /* ==========================================================
       MURAL DE AVISOS
       ========================================================== */

    .mural-eventos {
        background: linear-gradient(
            135deg,
            #7F1D1D 0%,
            #DC2626 45%,
            #F59E0B 100%
        );

        border: 3px solid #FBBF24;
        border-radius: 18px;
        padding: 20px;
        margin: 10px 0 18px 0;

        box-shadow:
            0px 8px 25px rgba(239, 68, 68, 0.35);
    }

    .mural-titulo {
        color: #FFFFFF !important;
        text-align: center;
        font-size: 1.25rem;
        font-weight: 900;
        margin-bottom: 6px;
    }

    .mural-subtitulo {
        color: #FFFFFF !important;
        text-align: center;
        font-size: 0.9rem;
        font-weight: 700;
        margin-bottom: 16px;
    }

    .evento-dia {
        background-color: rgba(255,255,255,0.96);
        color: #111827 !important;
        border-radius: 14px;
        padding: 15px;
        margin-top: 12px;
        box-shadow: 0px 4px 12px rgba(0,0,0,0.20);
    }

    .evento-dia strong {
        color: #111827 !important;
    }


    /* ==========================================================
       ÍCONE DO CAMPO DE SENHA
       ========================================================== */

    input[type="password"] {
        padding-right: 45px !important;
    }

    div[data-baseweb="input"] button {
        color: #D97706 !important;
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }

    div[data-baseweb="input"] button svg {
        fill: #D97706 !important;
        color: #D97706 !important;
        stroke: #D97706 !important;
    }


    /* ==========================================================
       BOTÕES
       ========================================================== */

    .stButton > button,
    div[data-testid="stFormSubmitButton"] > button {

        background-color: #D97706 !important;
        color: #FFFFFF !important;

        border: 2px solid #F59E0B !important;
        border-radius: 10px !important;

        font-weight: 800 !important;
        font-size: 1rem !important;

        padding: 14px 22px !important;

        box-shadow:
            0px 4px 14px rgba(217, 119, 6, 0.3);

        width: 100%;
    }

    .stButton > button:hover {
        background-color: #B45309 !important;
    }


    /* ==========================================================
       CAMPOS
       ========================================================== */

    label,
    p,
    span,
    div {
        color: #F1F5F9 !important;
        font-weight: 600;
    }

    .stTextInput input,
    .stSelectbox div[data-baseweb="select"],
    .stTextArea textarea {

        background-color: #1E293B !important;
        color: #FFFFFF !important;

        font-size: 1rem !important;

        border: 1px solid #475569 !important;
    }


    /* ==========================================================
       ABAS FIXAS / CONGELADAS
       ========================================================== */

    .stTabs [data-baseweb="tab-list"] {

        gap: 4px;

        background-color: #111827;

        padding: 6px;

        border-radius: 14px;

        border: 1px solid #334155;

        position: sticky !important;

        top: 0 !important;

        z-index: 9999 !important;

        box-shadow:
            0px 5px 15px rgba(0,0,0,0.35);
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


# ==============================================================================
# 3. ARQUIVO EXCEL
# ==============================================================================

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"


# ==============================================================================
# 4. CARREGAMENTO DOS CADASTROS
# ==============================================================================

@st.cache_data
def carregar_dados_cadastro():

    dados_padrao = [

        {
            "NOME_ALUNO_CLEAN": "Alvaro Barbosa",
            "RESPONSAVEL_CLEAN": "Odiselma Carvalho",
            "FAIXA_CLEAN": "Amarela",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67998411953"
        },

        {
            "NOME_ALUNO_CLEAN": "Elton Araujo",
            "RESPONSAVEL_CLEAN": "Antonio Edirley Graça Araujo",
            "FAIXA_CLEAN": "Amarela",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67998045242"
        },

        {
            "NOME_ALUNO_CLEAN": "Aysla Barbosa",
            "RESPONSAVEL_CLEAN": "Aline Florentim da Silva",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67999324855"
        },

        {
            "NOME_ALUNO_CLEAN": "Gustavo Camargo",
            "RESPONSAVEL_CLEAN": "Divina Magalhães Romero",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Teve braço quebrado / Epilepsia toma medicamento",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67998019136"
        },

        {
            "NOME_ALUNO_CLEAN": "Ester Louise",
            "RESPONSAVEL_CLEAN": "Simone Aparecida da Silva Barros de Oliveira",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Bronquite",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67999195844"
        },

        {
            "NOME_ALUNO_CLEAN": "Nathan Carvalho",
            "RESPONSAVEL_CLEAN": "London Janssen Santos de Carvalho",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67998513404"
        },

        {
            "NOME_ALUNO_CLEAN": "Erica Rodrigues",
            "RESPONSAVEL_CLEAN": "Elice dos Santos Alves",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Problema Cardíaco",
            "PENDENCIA_CLEAN": "Falta Endereço",
            "FONE_LIMPO": "67996430191"
        },

        {
            "NOME_ALUNO_CLEAN": "Riquelme da Silva",
            "RESPONSAVEL_CLEAN": "Sebastião Damasio da S. Filho",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67999298528"
        },

        {
            "NOME_ALUNO_CLEAN": "Guilherme Arruda",
            "RESPONSAVEL_CLEAN": "Jesiel Arruda",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Bronquite",
            "PENDENCIA_CLEAN": "Falta Endereço",
            "FONE_LIMPO": "67981845984"
        },

        {
            "NOME_ALUNO_CLEAN": "Hanna Nunes",
            "RESPONSAVEL_CLEAN": "Leticia Fatima F. da Silva",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Bronquite Asmatica",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67998731372"
        },

        {
            "NOME_ALUNO_CLEAN": "Isadora Souza",
            "RESPONSAVEL_CLEAN": "Aline Florentim da Silva",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67999324855"
        },

        {
            "NOME_ALUNO_CLEAN": "Kevellen José",
            "RESPONSAVEL_CLEAN": "Flavia Regina S. de Souza",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67999177623"
        },

        {
            "NOME_ALUNO_CLEAN": "Gabrielly Nunes",
            "RESPONSAVEL_CLEAN": "Ana Claudia R. Vieira Nunes",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Ok",
            "FONE_LIMPO": "67991787867"
        },

        {
            "NOME_ALUNO_CLEAN": "Bianca da Silva",
            "RESPONSAVEL_CLEAN": "Sebastião Damasio da S. Filho",
            "FAIXA_CLEAN": "Branca",
            "HISTORICO_SAUDE": "Não",
            "PENDENCIA_CLEAN": "Falta Endereço",
            "FONE_LIMPO": "67999298528"
        }
    ]

    df_fallback = pd.DataFrame(dados_padrao)

    try:

        xls = pd.ExcelFile(
            EXCEL_FILE,
            engine="openpyxl"
        )

        sheet_target = (
            "Ficha de Cadastro"
            if "Ficha de Cadastro" in xls.sheet_names
            else xls.sheet_names[0]
        )

        df = pd.read_excel(
            xls,
            sheet_name=sheet_target,
            header=1
        )

        df = df.dropna(
            how="all"
        ).dropna(
            how="all",
            axis=1
        )

        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        col_aluno = next(
            (
                c for c in df.columns
                if "Aluno" in c or "Nome" in c
            ),
            None
        )

        col_resp = next(
            (
                c for c in df.columns
                if "Responsável" in c
                or "Responsavel" in c
            ),
            None
        )

        col_faixa = next(
            (
                c for c in df.columns
                if "Faixa" in c
            ),
            None
        )

        col_fone = next(
            (
                c for c in df.columns
                if "Contato" in c
                or "Fone" in c
                or "Tel" in c
            ),
            None
        )

        col_saude = next(
            (
                c for c in df.columns
                if "Saúde" in c
                or "Saude" in c
            ),
            None
        )

        col_pend = next(
            (
                c for c in df.columns
                if "Pendência" in c
                or "Pendencia" in c
            ),
            None
        )

        if col_aluno and col_resp:

            df["NOME_ALUNO_CLEAN"] = (
                df[col_aluno]
                .astype(str)
                .str.strip()
            )

            df["RESPONSAVEL_CLEAN"] = (
                df[col_resp]
                .astype(str)
                .str.strip()
            )

            df["FAIXA_CLEAN"] = (
                df[col_faixa]
                .astype(str)
                .str.strip()
                if col_faixa
                else "Branca"
            )

            df["HISTORICO_SAUDE"] = (
                df[col_saude]
                .astype(str)
                .str.strip()
                if col_saude
                else "Não"
            )

            df["PENDENCIA_CLEAN"] = (
                df[col_pend]
                .astype(str)
                .str.strip()
                if col_pend
                else "Ok"
            )

            df["FONE_LIMPO"] = (
                df[col_fone]
                .astype(str)
                .apply(
                    lambda x:
                    "".join(
                        filter(
                            str.isdigit,
                            x
                        )
                    )
                )
                if col_fone
                else ""
            )

            df = df[
                df["NOME_ALUNO_CLEAN"]
                .str.lower() != "nan"
            ]

            df = df[
                df["NOME_ALUNO_CLEAN"]
                .str.strip() != ""
            ]

            if not df.empty:
                return df

        return df_fallback

    except Exception:

        return df_fallback


# ==============================================================================
# 5. CARREGAMENTO DO CRONOGRAMA
# ==============================================================================

@st.cache_data
def carregar_cronograma_eventos():

    try:

        xls = pd.ExcelFile(
            EXCEL_FILE,
            engine="openpyxl"
        )

        if "Cronograma de Eventos" in xls.sheet_names:

            df_ev = pd.read_excel(
                xls,
                sheet_name="Cronograma de Eventos"
            )

            return df_ev.dropna(
                how="all"
            )

        else:

            return pd.DataFrame({

                "Evento": [
                    "Palestra: Escolhas Saudáveis - Influências - Prevenções",
                    "Palestra: Lidando com Emoções - Inteligência Emocional e Saúde Mental",
                    "Palestra: O Valor do Estudo - Projeto de Vida - Futuro"
                ],

                "Data Evento": [
                    "29/09/2026",
                    "29/10/2026",
                    "26/11/2026"
                ],

                "Hora Evento": [
                    "19:00",
                    "19:00",
                    "19:00"
                ],

                "Palestrante / Instrutor / Equipe / Igreja": [
                    "Equipe PROERD",
                    "Psicóloga Eva Mateus",
                    "Pedagogos Luis e Edima"
                ]
            })

    except Exception:

        return pd.DataFrame()


df_cadastro = carregar_dados_cadastro()

df_cronograma = carregar_cronograma_eventos()


# ==============================================================================
# 6. ESTADOS DA SESSÃO
# ==============================================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "user_info" not in st.session_state:
    st.session_state.user_info = None

if "sino_visto" not in st.session_state:
    st.session_state.sino_visto = False

if "senhas_pais" not in st.session_state:

    st.session_state.senhas_pais = {
        "67998411953": "123456"
    }

if "senhas_diretoria" not in st.session_state:

    st.session_state.senhas_diretoria = {
        "diretoria": "dir2026"
    }

if "alunos_desistentes" not in st.session_state:
    st.session_state.alunos_desistentes = []

if "pedidos_oracao" not in st.session_state:
    st.session_state.pedidos_oracao = []

if "alunos_batizados" not in st.session_state:

    st.session_state.alunos_batizados = [
        "Alvaro Barbosa"
    ]

if "membros_celula_count" not in st.session_state:
    st.session_state.membros_celula_count = 28

if "evolucao_mestre" not in st.session_state:
    st.session_state.evolucao_mestre = {}


# ==============================================================================
# 7. AUTENTICAÇÃO
# ==============================================================================

def autenticar_usuario(
    login_input,
    senha_input
):

    login_str = (
        str(login_input)
        .strip()
        .lower()
    )

    login_limpo = "".join(
        filter(
            str.isdigit,
            str(login_input)
        )
    )


    # --------------------------------------------------------------------------
    # ACESSO DO MIKHAEL / MESTRE
    # --------------------------------------------------------------------------
    #
    # Login: Mikhael
    # Senha: 12381314*Lj
    #
    # --------------------------------------------------------------------------

    if (
        login_str == "mikhael"
        and senha_input == "12381314*Lj"
    ):

        return {

            "nome": "Mikhael",

            "tipo": "mestre",

            "filhos": []
        }


    # --------------------------------------------------------------------------
    # DIRETORIA
    # --------------------------------------------------------------------------

    if (
        login_str in st.session_state.senhas_diretoria
        and senha_input
        == st.session_state.senhas_diretoria[
            login_str
        ]
    ):

        return {

            "nome": "Diretoria IEQ Guaicurus",

            "tipo": "diretoria",

            "filhos": []
        }


    # --------------------------------------------------------------------------
    # RESPONSÁVEIS
    # --------------------------------------------------------------------------

    if (
        not df_cadastro.empty
        and "FONE_LIMPO"
        in df_cadastro.columns
    ):

        match = (
            df_cadastro[
                df_cadastro[
                    "FONE_LIMPO"
                ].str.contains(
                    login_limpo,
                    na=False
                )
            ]
            if login_limpo
            else pd.DataFrame()
        )


        if (
            match.empty
            and "RESPONSAVEL_CLEAN"
            in df_cadastro.columns
        ):

            match = df_cadastro[
                df_cadastro[
                    "RESPONSAVEL_CLEAN"
                ]
                .astype(str)
                .str.lower()
                .str.contains(
                    login_str,
                    na=False
                )
            ]


        if not match.empty:

            resp_nome = match.iloc[0].get(
                "RESPONSAVEL_CLEAN",
                "Responsável"
            )

            fone = match.iloc[0].get(
                "FONE_LIMPO",
                ""
            )


            if (
                fone
                in st.session_state.alunos_desistentes
            ):
                return None


            senha_correta = (
                st.session_state.senhas_pais.get(
                    fone,
                    "123456"
                )
            )


            if senha_input == senha_correta:

                return {

                    "nome": resp_nome,

                    "tipo": "pai",

                    "fone": fone,

                    "filhos":
                        match.to_dict(
                            orient="records"
                        )
                }


    return None


# ==============================================================================
# 8. LINK DO WHATSAPP
# ==============================================================================

def gerar_link_whatsapp(
    numero,
    mensagem
):

    if (
        pd.isna(numero)
        or str(numero).strip() == ""
        or str(numero).strip().lower() == "nan"
    ):

        return None


    num_limpo = "".join(
        filter(
            str.isdigit,
            str(numero)
        )
    )


    if (
        not num_limpo.startswith("55")
        and len(num_limpo) in [10, 11]
    ):

        num_limpo = "55" + num_limpo


    msg_enc = urllib.parse.quote(
        mensagem
    )


    return (
        f"https://wa.me/"
        f"{num_limpo}?text={msg_enc}"
    )


# ==============================================================================
# 9. DATA E HORA ATUAL
# HORÁRIO DE CAMPO GRANDE / MS
# ==============================================================================

def obter_data_hora_local():

    return datetime.datetime.now(
        ZoneInfo(
            "America/Campo_Grande"
        )
    )


agora_local = obter_data_hora_local()

data_hoje = agora_local.date()

data_hoje_str = data_hoje.strftime(
    "%d/%m/%Y"
)


# ==============================================================================
# 10. IDENTIFICAR EVENTOS DO DIA
# ==============================================================================

def obter_eventos_hoje(
    df_eventos
):

    if df_eventos.empty:
        return pd.DataFrame()


    if "Data Evento" not in df_eventos.columns:
        return pd.DataFrame()


    df_temp = df_eventos.copy()


    try:

        df_temp["_DATA_CONVERTIDA"] = pd.to_datetime(
            df_temp["Data Evento"],
            errors="coerce",
            dayfirst=True
        )


        df_temp = df_temp[
            df_temp[
                "_DATA_CONVERTIDA"
            ].dt.date
            == data_hoje
        ]


        return df_temp


    except Exception:

        return pd.DataFrame()


df_eventos_hoje = obter_eventos_hoje(
    df_cronograma
)

tem_evento_hoje = (
    not df_eventos_hoje.empty
)


# ==============================================================================
# 11. TELA DE LOGIN
# ==============================================================================

if not st.session_state.logged_in:

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )


    c_logo1, c_logo2, c_logo3 = st.columns(
        [1, 2, 1]
    )


    with c_logo2:

        st.image(
            "logo_projeto_sementes.png",
            use_container_width=True
        )


    st.markdown(
        """
        <h2 style="
            text-align: center;
            color:#F59E0B;
            font-weight:800;
        ">
            PROJETO SEMENTES
        </h2>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <p style="
            text-align: center;
            color: #CBD5E1;
            font-size: 0.95rem;
        ">
            Iniciativa Voluntária de Jiu-Jitsu e Apoio à Família
        </p>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <p style="
            text-align: center;
            color: #64748B;
            font-size: 0.8rem;
        ">
            Cessão de Espaço Comunitário: IEQ Guaicurus
        </p>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )


    # --------------------------------------------------------------------------
    # FORMULÁRIO DE LOGIN
    # --------------------------------------------------------------------------

    with st.form("form_login"):


        # APARECE SOMENTE:
        # LOGIN DO RESPONSÁVEL

        user_in = st.text_input(
            "Login do Responsável",
            placeholder="Digite seu login"
        )


        pass_in = st.text_input(
            "Senha de Acesso",
            type="password",
            placeholder="••••••••"
        )


        btn_login = st.form_submit_button(
            "ENTRAR NO APLICATIVO"
        )


        if btn_login:

            user_data = autenticar_usuario(
                user_in,
                pass_in
            )


            if user_data:

                st.session_state.logged_in = True

                st.session_state.user_info = user_data

                st.session_state.sino_visto = False

                st.success(
                    "Acesso autorizado com sucesso!"
                )

                st.rerun()


            else:

                st.error(
                    "Credenciais inválidas ou acesso removido. Consulte a coordenação."
                )


# ==============================================================================
# 12. APLICAÇÃO LOGADA
# ==============================================================================

else:

    u_info = st.session_state.user_info


    # --------------------------------------------------------------------------
    # SAUDAÇÃO
    # --------------------------------------------------------------------------

    if u_info.get("tipo") in [
        "mestre",
        "diretoria"
    ]:

        saudacao = (
            "A Paz seja convosco, Mestre Mikhael / Diretoria"
        )

    else:

        saudacao = (
            f"A Paz seja convosco, "
            f"{u_info.get('nome')}"
        )


    # --------------------------------------------------------------------------
    # CABEÇALHO
    # --------------------------------------------------------------------------

    col_head1, col_head2 = st.columns(
        [5, 1]
    )


    with col_head1:

        st.markdown(
            f"""
            <strong style="
                color:#F59E0B;
                font-size:1.1rem;
            ">
                {saudacao}
            </strong>
            """,
            unsafe_allow_html=True
        )


        # ----------------------------------------------------------------------
        # DATA E HORA ATUALIZADA
        # ----------------------------------------------------------------------

        st.markdown(
            f"""
            <div style="
                margin-top:6px;
                color:#CBD5E1;
                font-size:0.82rem;
                font-weight:600;
            ">
                🕐 Acesso:
                {agora_local.strftime('%d/%m/%Y')}
                •
                {agora_local.strftime('%H:%M:%S')}
            </div>
            """,
            unsafe_allow_html=True
        )


    # --------------------------------------------------------------------------
    # SINO
    # --------------------------------------------------------------------------

    with col_head2:


        if (
            tem_evento_hoje
            and not st.session_state.sino_visto
        ):

            # SOMENTE SINO VERMELHO
            # PISCANDO

            st.markdown(
                """
                <div class="sino-vermelho">
                    🔔
                </div>
                """,
                unsafe_allow_html=True
            )


            btn_sino = st.button(
                "🔔",
                key="btn_sino_evento"
            )


        else:

            # SOMENTE SINO DOURADO

            st.markdown(
                """
                <div class="sino-dourado">
                    🔔
                </div>
                """,
                unsafe_allow_html=True
            )


            btn_sino = st.button(
                "🔔",
                key="btn_sino_normal"
            )


        if btn_sino:

            if tem_evento_hoje:

                st.session_state.sino_visto = True

            else:

                st.session_state.sino_visto = (
                    not st.session_state.sino_visto
                )

            st.rerun()


    # ==============================================================================
    # 13. MURAL DE AVISOS
    # SOMENTE EVENTOS DO DIA
    # ==============================================================================

    if (
        st.session_state.sino_visto
        and tem_evento_hoje
    ):

        st.markdown(
            """
            <div class="mural-eventos">

                <div class="mural-titulo">
                    🔔 MURAL DE AVISOS
                </div>

                <div class="mural-subtitulo">
                    EVENTOS DO PROJETO SEMENTES — HOJE
                </div>

            """,
            unsafe_allow_html=True
        )


        for _, evento in df_eventos_hoje.iterrows():

            nome_evento = evento.get(
                "Evento",
                "Evento do Projeto Sementes"
            )


            data_evento = evento.get(
                "Data Evento",
                data_hoje_str
            )


            hora_evento = evento.get(
                "Hora Evento",
                ""
            )


            responsavel_evento = evento.get(
                "Palestrante / Instrutor / Equipe / Igreja",
                ""
            )


            if pd.isna(data_evento):
                data_evento = data_hoje_str


            if pd.isna(hora_evento):
                hora_evento = ""


            if pd.isna(responsavel_evento):
                responsavel_evento = ""


            st.markdown(
                f"""
                <div class="evento-dia">

                    <div style="
                        font-size:1.08rem;
                        font-weight:900;
                        margin-bottom:8px;
                    ">
                        📢 {nome_evento}
                    </div>

                    <div style="
                        font-size:0.92rem;
                        margin-top:5px;
                    ">
                        📅
                        <strong>Data:</strong>
                        {data_evento}
                    </div>

                    <div style="
                        font-size:0.92rem;
                        margin-top:5px;
                    ">
                        🕐
                        <strong>Horário:</strong>
                        {hora_evento}
                    </div>

                    <div style="
                        font-size:0.92rem;
                        margin-top:5px;
                    ">
                        👤
                        <strong>Responsável:</strong>
                        {responsavel_evento}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )


        # ----------------------------------------------------------------------
        # BOTÃO PARA MARCAR COMO VISUALIZADO
        # ----------------------------------------------------------------------

        if st.button(
            "✓ MARCAR AVISOS COMO VISUALIZADOS",
            key="fechar_mural_eventos"
        ):

            st.session_state.sino_visto = True

            st.rerun()


    # ==============================================================================
    # 14. ABAS
    # FICAM FIXAS DURANTE A ROLAGEM
    # ==============================================================================

    tab1, tab_mat, tab_crono, tab2, tab3, tab4, tab5 = st.tabs([

        "🏠 Início",

        "📝 Matrícula",

        "📅 Cronograma",

        "📊 Frequência",

        "🥋 Tatame",

        "🏆 Campeão & Fé",

        "⚙️ Secretaria"

    ])


    # ==============================================================================
    # ABA 1 — INÍCIO
    # ==============================================================================

    with tab1:

        st.markdown(
            """
            <div class="welcome-card">

                <h3 style="
                    color:#F59E0B;
                    text-align:center;
                    font-weight:800;
                    margin-bottom:12px;
                ">
                    🌱 A HISTÓRIA DO PROJETO SEMENTES
                </h3>

                <p style="
                    font-size:0.95rem;
                    line-height:1.6;
                    text-align:justify;
                    color:#E2E8F0;
                ">

                    O <strong>Projeto Sementes</strong> nasceu no coração de Deus
                    e, por Sua misericórdia, foi compartilhado aos corações do
                    <strong>Pastor Joel Amorim</strong>, da IEQ Guaicurus, e do
                    <strong>Instrutor Faixa-Preta London Carvalho</strong>,
                    da Academia Iron Jiu-Jitsu.

                    <br><br>

                    Foram dias de oração, planejamento e dedicação, buscando
                    estruturar o projeto da melhor maneira para acolher e atender
                    crianças das comunidades próximas à igreja.

                    <br><br>

                    Em uma noite de culto, o mover de Deus alcançou os corações
                    dos membros da igreja, que abraçaram o projeto. Na mesma noite,
                    por meio das ofertas voluntárias dos irmãos, foi arrecadado o
                    valor necessário para a aquisição do nosso <strong>dojo</strong>.

                    <br><br>

                    Assim, com o apoio dos instrutores e dos pais, iniciamos nossas
                    aulas. Desde então, com muita disciplina, dedicação e entusiasmo,
                    nossos alunos vêm aprendendo muito mais do que a
                    <strong>Arte Suave</strong>.

                    <br><br>

                    Ensinamos <strong>valores morais e éticos fundamentados em
                    Cristo Jesus</strong>, unindo o Jiu-Jitsu a
                    <strong>palestras socioeducativas e ações de evangelismo</strong>,
                    contribuindo para a formação integral de nossas crianças,
                    adolescentes e jovens.

                    <br><br>

                    Nosso grande sonho é <strong>levar a Semente do Evangelho
                    até a sua casa</strong>, alcançando não apenas nossos alunos,
                    mas também suas famílias.

                    <br><br>

                    <strong style="
                        color:#F59E0B;
                        font-size:1.05rem;
                        display:block;
                        text-align:center;
                    ">
                        Sejam todos bem-vindos a esta família! 🌱🥋
                    </strong>

                </p>

            </div>
            """,
            unsafe_allow_html=True
        )


        st.markdown("---")


        st.markdown(
            "### 🎂 Aniversariantes do Mês & Comemorações"
        )


        st.caption(
            "Comemorações oficiais da equipe e familiares (Horário: 19h):"
        )


        st.markdown(
            """
            * *Setembro:* 29/09/2026 (Terça-feira)
            * *Outubro:* 29/10/2026 (Quinta-feira)
            * *Novembro:* 26/11/2026 (Quinta-feira)
            * *Dezembro:* 31/12/2026 (Quinta-feira)
            """
        )


        c_aniv1, c_aniv2 = st.columns(
            [1, 2]
        )


        with c_aniv1:

            st.markdown(
                "📷 *Foto de Perfil*"
            )

            st.info(
                "👤 [Foto do Aluno]"
            )


        with c_aniv2:

            st.markdown(
                "*Alvaro Barbosa* — Aniversariante do Mês"
            )

            st.markdown(
                "🎉 *Festa da Família:* 29/09/2026 às 19h"
            )


            msg_aniv_auto = (
                "Paz do Senhor, Alvaro! "
                "Todo o Projeto Sementes te deseja um feliz aniversário! "
                "Que o Senhor Jesus abençoe sua vida, dando-lhe sabedoria "
                "e muita saúde no tatame! 🥋🎉"
            )


            link_aniv = gerar_link_whatsapp(
                "5567998411953",
                msg_aniv_auto
            )


            if link_aniv:

                st.markdown(
                    f"[📲 Enviar Mensagem Evangélica Automática]({link_aniv})"
                )


        st.markdown("---")


        st.markdown(
            "### 🥋 Perfil do Aluno"
        )


        lista_alunos_base = df_cadastro.get(
            "NOME_ALUNO_CLEAN",
            pd.Series(
                ["Alvaro Barbosa"]
            )
        ).tolist()


        if u_info.get("tipo") in [
            "mestre",
            "diretoria"
        ]:

            st.info(
                "👑 Modo Mestre/Diretoria: Selecione qualquer aluno da planilha:"
            )


            aluno_sel_perfil = st.selectbox(
                "Selecione o Aluno:",
                lista_alunos_base
            )


            filhos_exibir = df_cadastro[
                df_cadastro.get(
                    "NOME_ALUNO_CLEAN",
                    pd.Series()
                )
                == aluno_sel_perfil
            ].to_dict(
                orient="records"
            )


        else:

            filhos_exibir = u_info.get(
                "filhos",
                []
            )


            if (
                not filhos_exibir
                and not df_cadastro.empty
            ):

                filhos_exibir = (
                    df_cadastro
                    .head(1)
                    .to_dict(
                        orient="records"
                    )
                )


        for f in filhos_exibir:

            nome_al = f.get(
                "NOME_ALUNO_CLEAN",
                "Aluno"
            )


            faixa_al = f.get(
                "FAIXA_CLEAN",
                "Branca"
            )


            resp_al = f.get(
                "RESPONSAVEL_CLEAN",
                "Responsável Cadastrado"
            )


            saude_al = f.get(
                "HISTORICO_SAUDE",
                "Não informado"
            )


            is_batizado = (
                nome_al
                in st.session_state.alunos_batizados
            )


            st.markdown(
                f"""
                <div class="gold-card">

                    <h3>
                        📷 [Foto] {nome_al}
                        {'⭐' if is_batizado else ''}
                    </h3>

                    <span class="gold-badge">
                        Faixa {faixa_al}
                    </span>

                    <p style="
                        margin-top:10px;
                        font-size:0.95rem;
                    ">

                        <strong>Responsável Legal:</strong>
                        {resp_al}
                        <br>

                        <strong>Histórico de Saúde:</strong>
                        {saude_al}
                        <br>

                        <strong>Biometria Facial:</strong>
                        🟢 Cadastrada
                        <br>

                        <strong>Aptidão e Saúde:</strong>
                        🔒 Apto para treinos de contato

                    </p>

                </div>
                """,
                unsafe_allow_html=True
            )


        st.markdown("---")


        st.markdown(
            "### 🏫 Acompanhamento Escolar e Familiar (Mensal)"
        )


        with st.expander(
            "📝 Responder Avaliação Mensal (Acesso do Responsável)",
            expanded=True
        ):

            st.selectbox(
                "Aluno Avaliado:",
                [
                    f.get(
                        "NOME_ALUNO_CLEAN"
                    )
                    for f in filhos_exibir
                ]
                if filhos_exibir
                else [
                    "Alvaro Barbosa"
                ]
            )


            st.select_slider(
                "1. Desempenho Escolar:",
                options=[
                    "Ruim",
                    "Regular",
                    "Bom",
                    "Ótimo"
                ],
                value="Ótimo"
            )


            st.select_slider(
                "2. Comportamento em Casa:",
                options=[
                    "Ruim",
                    "Regular",
                    "Bom",
                    "Ótimo"
                ],
                value="Ótimo"
            )


            st.select_slider(
                "3. Disciplina no Tatame:",
                options=[
                    "Ruim",
                    "Regular",
                    "Bom",
                    "Ótimo"
                ],
                value="Ótimo"
            )


            st.text_area(
                "Observações para a Coordenação do Projeto:"
            )


            if st.button(
                "Salvar Avaliação Mensal"
            ):

                st.success(
                    "Avaliação salva com sucesso!"
                )
