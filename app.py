import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import re
from pathlib import Path

try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    CANVAS_DISPONIVEL = False


# ==============================================================================
# 1. CONFIGURAÇÃO
# ==============================================================================

st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"


# ==============================================================================
# 2. CSS
# ==============================================================================

st.markdown("""
<style>

html, body, [class*="css"] {
    font-family: Arial, sans-serif;
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

/* ------------------------------------------------------------
   CAMPOS
------------------------------------------------------------ */

.stTextInput input,
.stTextArea textarea,
.stNumberInput input {
    background-color: #1E293B !important;
    color: #FFFFFF !important;
    border: 2px solid #475569 !important;
    border-radius: 10px !important;
}

/* ------------------------------------------------------------
   CORREÇÃO DO ÍCONE DE SENHA
------------------------------------------------------------ */

.stTextInput [data-testid="stInputRootElement"] button {
    color: #F59E0B !important;
    background: transparent !important;
}

.stTextInput [data-testid="stInputRootElement"] button svg {
    color: #F59E0B !important;
    fill: none !important;
    stroke: #F59E0B !important;
}

/* ------------------------------------------------------------
   CARTÕES
------------------------------------------------------------ */

.gold-card {
    background-color: #111827;
    border: 2px solid #D97706;
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 14px;
    box-shadow: 0 4px 14px rgba(217,119,6,.20);
}

.welcome-card {
    background-color: #111827;
    border: 2px solid #D97706;
    border-radius: 16px;
    padding: 22px;
    margin-bottom: 20px;
    box-shadow: 0 6px 20px rgba(217,119,6,.25);
}

/* ------------------------------------------------------------
   STATUS
------------------------------------------------------------ */

.status-amarelo {
    display: inline-block;
    width: 17px;
    height: 17px;
    background: #FACC15;
    border-radius: 50%;
    box-shadow: 0 0 8px rgba(250,204,21,.8);
    vertical-align: middle;
}

.status-verde {
    display: inline-block;
    width: 17px;
    height: 17px;
    background: #22C55E;
    border-radius: 50%;
    box-shadow: 0 0 8px rgba(34,197,94,.8);
    vertical-align: middle;
}

/* ------------------------------------------------------------
   SINO
------------------------------------------------------------ */

.bell-gold {
    color: #F59E0B;
    font-size: 34px;
    text-align: center;
}

.bell-red {
    color: #EF4444;
    font-size: 34px;
    text-align: center;
    animation: pulse-red 1s infinite;
}

@keyframes pulse-red {
    0% {
        transform: scale(1);
        filter: drop-shadow(0 0 2px #EF4444);
    }

    50% {
        transform: scale(1.25);
        filter: drop-shadow(0 0 13px #EF4444);
    }

    100% {
        transform: scale(1);
        filter: drop-shadow(0 0 2px #EF4444);
    }
}

/* ------------------------------------------------------------
   MURAL
------------------------------------------------------------ */

.event-wall {
    background: linear-gradient(
        135deg,
        #7F1D1D,
        #DC2626,
        #F97316
    );

    border: 3px solid #FACC15;
    border-radius: 20px;
    padding: 22px;
    margin: 12px 0 20px 0;
    box-shadow: 0 0 25px rgba(239,68,68,.45);
}

.event-wall h2 {
    color: #FFFFFF !important;
    text-align: center;
    font-weight: 900;
}

.event-wall p {
    color: #FFFFFF !important;
    font-size: 1rem;
    line-height: 1.7;
}

/* ------------------------------------------------------------
   BOTÕES
------------------------------------------------------------ */

.stButton > button,
div[data-testid="stFormSubmitButton"] > button {
    background-color: #D97706 !important;
    color: #FFFFFF !important;
    border: 2px solid #F59E0B !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    min-height: 45px;
}

.stButton > button:hover {
    background-color: #B45309 !important;
}

/* ------------------------------------------------------------
   ABAS FIXAS
------------------------------------------------------------ */

.stTabs [data-baseweb="tab-list"] {
    position: sticky !important;
    top: 0 !important;
    z-index: 999 !important;

    display: flex;
    gap: 4px;

    background-color: #111827 !important;
    padding: 6px;

    border-radius: 14px;
    border: 1px solid #334155;

    overflow-x: auto;
}

.stTabs [data-baseweb="tab"] {
    flex-shrink: 0;
    border-radius: 8px;
    color: #CBD5E1 !important;
    font-weight: 700;
    padding: 9px 10px;
    font-size: .78rem;
}

.stTabs [aria-selected="true"] {
    background-color: #D97706 !important;
    color: #FFFFFF !important;
}

/* ------------------------------------------------------------
   TEXTOS
------------------------------------------------------------ */

label,
p,
span,
div {
    color: #F1F5F9;
}

h1, h2, h3, h4, h5 {
    color: #FFFFFF;
}

.metric-card {
    background: #111827;
    border: 1px solid #475569;
    border-radius: 12px;
    padding: 12px;
}

</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 3. FUNÇÕES AUXILIARES
# ==============================================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()

    substituicoes = {
        "á": "a",
        "à": "a",
        "ã": "a",
        "â": "a",
        "ä": "a",
        "é": "e",
        "ê": "e",
        "ë": "e",
        "í": "i",
        "ï": "i",
        "ó": "o",
        "ô": "o",
        "õ": "o",
        "ö": "o",
        "ú": "u",
        "ü": "u",
        "ç": "c"
    }

    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def limpar_numero(valor):
    if pd.isna(valor):
        return ""

    return "".join(filter(str.isdigit, str(valor)))


def encontrar_coluna(df, possibilidades):
    """
    Procura uma coluna por correspondência aproximada.
    """

    if df is None or df.empty:
        return None

    colunas = list(df.columns)

    normalizadas = {
        coluna: normalizar_texto(coluna)
        for coluna in colunas
    }

    # Primeiro tenta correspondência exata
    for possibilidade in possibilidades:

        alvo = normalizar_texto(possibilidade)

        for coluna, normalizada in normalizadas.items():

            if normalizada == alvo:
                return coluna

    # Depois procura se o termo está contido
    for possibilidade in possibilidades:

        alvo = normalizar_texto(possibilidade)

        for coluna, normalizada in normalizadas.items():

            if alvo and alvo in normalizada:
                return coluna

    return None


def converter_data(valor):

    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        return valor.date()

    if isinstance(valor, datetime.datetime):
        return valor.date()

    if isinstance(valor, datetime.date):
        return valor

    texto = str(valor).strip()

    formatos = [
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y"
    ]

    for formato in formatos:

        try:
            return datetime.datetime.strptime(
                texto,
                formato
            ).date()

        except ValueError:
            continue

    try:
        data = pd.to_datetime(
            valor,
            dayfirst=True,
            errors="coerce"
        )

        if pd.notna(data):
            return data.date()

    except Exception:
        pass

    return None


def formatar_data(valor):

    data = converter_data(valor)

    if data:
        return data.strftime("%d/%m/%Y")

    return str(valor) if not pd.isna(valor) else ""


def valor_preenchido(valor):

    if pd.isna(valor):
        return False

    texto = str(valor).strip().lower()

    if texto in [
        "",
        "nan",
        "none",
        "null",
        "-",
        "--"
    ]:
        return False

    return True


def status_biometria(valor):

    texto = normalizar_texto(valor)

    positivos = [
        "sim",
        "cadastrado",
        "ativo",
        "realizada",
        "realizado",
        "ok",
        "verde",
        "true",
        "1",
        "biometria cadastrada"
    ]

    return texto in positivos


def gerar_link_whatsapp(numero, mensagem=""):

    numero_limpo = limpar_numero(numero)

    if not numero_limpo:
        return None

    if not numero_limpo.startswith("55"):

        if len(numero_limpo) in [10, 11]:
            numero_limpo = "55" + numero_limpo

    mensagem_codificada = urllib.parse.quote(mensagem)

    return f"https://wa.me/{numero_limpo}?text={mensagem_codificada}"


# ==============================================================================
# 4. LEITURA DA PLANILHA
# ==============================================================================

def carregar_excel():

    caminho = Path(EXCEL_FILE)

    if not caminho.exists():
        st.error(
            f"❌ A planilha '{EXCEL_FILE}' não foi encontrada."
        )

        st.stop()

    try:

        return pd.ExcelFile(
            caminho,
            engine="openpyxl"
        )

    except Exception as erro:

        st.error(
            f"Erro ao abrir a planilha: {erro}"
        )

        st.stop()


@st.cache_data(ttl=30)
def ler_aba(nome_aba):

    try:

        xls = pd.ExcelFile(
            EXCEL_FILE,
            engine="openpyxl"
        )

        if nome_aba not in xls.sheet_names:
            return pd.DataFrame()

        # Primeiro tenta header 0
        df = pd.read_excel(
            xls,
            sheet_name=nome_aba,
            header=0
        )

        # Se não encontrar estrutura útil, tenta header 1
        if df.empty or len(df.columns) <= 1:

            df = pd.read_excel(
                xls,
                sheet_name=nome_aba,
                header=1
            )

        df = df.dropna(
            how="all"
        )

        df = df.dropna(
            how="all",
            axis=1
        )

        df.columns = [
            str(coluna).strip()
            for coluna in df.columns
        ]

        return df

    except Exception:
        return pd.DataFrame()


# ==============================================================================
# 5. CARREGAR TODAS AS ABAS
# ==============================================================================

xls = carregar_excel()

NOMES_ABAS = xls.sheet_names

df_cadastro = ler_aba("Ficha de Cadastro")

df_cronograma = ler_aba("Cronograma de Eventos")

df_agosto = ler_aba("Lista de Presença Agosto")

df_setembro = ler_aba("Lista de Presença Setembro")

df_campeao = ler_aba("Campeão")

df_fe = ler_aba("Fé")


# ==============================================================================
# 6. PREPARAR CADASTRO
# ==============================================================================

COL_ALUNO = encontrar_coluna(
    df_cadastro,
    [
        "Nome do Aluno",
        "Nome Aluno",
        "Aluno",
        "Nome"
    ]
)

COL_RESPONSAVEL = encontrar_coluna(
    df_cadastro,
    [
        "Responsável",
        "Responsavel",
        "Nome do Responsável",
        "Responsável Legal"
    ]
)

COL_FAIXA = encontrar_coluna(
    df_cadastro,
    [
        "Faixa",
        "Faixa Atual"
    ]
)

COL_ENDERECO = encontrar_coluna(
    df_cadastro,
    [
        "Endereço",
        "Endereco",
        "Endereço Completo"
    ]
)

COL_BAIRRO = encontrar_coluna(
    df_cadastro,
    [
        "Bairro"
    ]
)

COL_DOCUMENTOS = encontrar_coluna(
    df_cadastro,
    [
        "Documentos",
        "Documento",
        "Pendência",
        "Pendencia",
        "Status"
    ]
)

COL_BIOMETRIA = encontrar_coluna(
    df_cadastro,
    [
        "Biometria Fácil",
        "Biometria Facil",
        "Biometria",
        "Biometria Facial",
        "Status Biometria",
        "Biometria Cadastrada"
    ]
)

COL_TELEFONE = encontrar_coluna(
    df_cadastro,
    [
        "WhatsApp",
        "Whatsapp",
        "Contato",
        "Telefone",
        "Celular",
        "Fone"
    ]
)


def preparar_cadastro(df):

    if df.empty:
        return pd.DataFrame()

    dados = df.copy()

    if COL_ALUNO:

        dados["__ALUNO"] = (
            dados[COL_ALUNO]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    else:

        dados["__ALUNO"] = ""

    if COL_RESPONSAVEL:
        dados["__RESPONSAVEL"] = (
            dados[COL_RESPONSAVEL]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        dados["__RESPONSAVEL"] = ""

    if COL_FAIXA:
        dados["__FAIXA"] = (
            dados[COL_FAIXA]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        dados["__FAIXA"] = ""

    if COL_ENDERECO:
        dados["__ENDERECO"] = (
            dados[COL_ENDERECO]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        dados["__ENDERECO"] = ""

    if COL_BAIRRO:
        dados["__BAIRRO"] = (
            dados[COL_BAIRRO]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        dados["__BAIRRO"] = ""

    if COL_TELEFONE:
        dados["__TELEFONE"] = dados[COL_TELEFONE].apply(
            limpar_numero
        )
    else:
        dados["__TELEFONE"] = ""

    if COL_DOCUMENTOS:
        dados["__PENDENCIA"] = (
            dados[COL_DOCUMENTOS]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        dados["__PENDENCIA"] = ""

    if COL_BIOMETRIA:
        dados["__BIOMETRIA"] = dados[COL_BIOMETRIA]
    else:
        dados["__BIOMETRIA"] = ""

    # Remove linhas sem nome
    dados = dados[
        dados["__ALUNO"].str.strip() != ""
    ]

    return dados


df_alunos = preparar_cadastro(df_cadastro)


# ==============================================================================
# 7. DETERMINAR STATUS DO ALUNO
# ==============================================================================

def aluno_status(linha):

    nome = str(
        linha.get("__ALUNO", "")
    ).strip()

    pendencia = normalizar_texto(
        linha.get("__PENDENCIA", "")
    )

    endereco = linha.get(
        "__ENDERECO",
        ""
    )

    responsavel = linha.get(
        "__RESPONSAVEL",
        ""
    )

    faixa = linha.get(
        "__FAIXA",
        ""
    )

    biometria = linha.get(
        "__BIOMETRIA",
        ""
    )

    # ----------------------------------------------------------
    # PENDÊNCIAS EXPRESSAS
    # ----------------------------------------------------------

    possui_pendencia = (
        "falta endereco" in pendencia
        or
        "falta documentos" in pendencia
        or
        "falta documento" in pendencia
    )

    # ----------------------------------------------------------
    # CAMPOS OBRIGATÓRIOS
    # ----------------------------------------------------------

    cadastro_completo = all([
        valor_preenchido(nome),
        valor_preenchido(responsavel),
        valor_preenchido(endereco),
        valor_preenchido(faixa)
    ])

    # ----------------------------------------------------------
    # BIOMETRIA
    # ----------------------------------------------------------

    biometria_ok = status_biometria(
        biometria
    )

    # ----------------------------------------------------------
    # VERDE SOMENTE SE TUDO ESTIVER COMPLETO
    # ----------------------------------------------------------

    if (
        cadastro_completo
        and
        not possui_pendencia
        and
        biometria_ok
    ):

        return "verde"

    return "amarelo"


if not df_alunos.empty:

    df_alunos["__STATUS"] = df_alunos.apply(
        aluno_status,
        axis=1
    )

else:

    df_alunos["__STATUS"] = pd.Series(
        dtype=str
    )


# ==============================================================================
# 8. ESTADOS DO APLICATIVO
# ==============================================================================

defaults = {

    "logged_in": False,

    "user_info": None,

    "sino_visualizado": False,

    "data_notificacao_visualizada": "",

    "senhas_pais": {
        "67998411953": "123456"
    },

    "senhas_diretoria": {
        "diretoria": "dir2026"
    },

    "alunos_desistentes": [],

    "pedidos_oracao": [],

    "presencas_tatame": {},

    "evolucao_mestre": {},

    "alunos_batizados": [],

    "membros_celula_count": 0

}


for chave, valor in defaults.items():

    if chave not in st.session_state:

        st.session_state[chave] = valor


# ==============================================================================
# 9. AUTENTICAÇÃO
# ==============================================================================

def autenticar_usuario(login_input, senha_input):

    login_original = str(
        login_input
    ).strip()

    login_limpo = limpar_numero(
        login_original
    )

    login_str = normalizar_texto(
        login_original
    )

    # ----------------------------------------------------------
    # MESTRE
    # ----------------------------------------------------------

    if (
        login_str in ["mestre", "london"]
        or
        login_limpo == "00000000000"
    ):

        if senha_input == "12381314*Lj":

            return {
                "nome": "Mestre London",
                "tipo": "mestre",
                "filhos": []
            }

    # ----------------------------------------------------------
    # DIRETORIA
    # ----------------------------------------------------------

    if (
        login_str in st.session_state.senhas_diretoria
        and
        senha_input ==
        st.session_state.senhas_diretoria[
            login_str
        ]
    ):

        return {
            "nome": "Diretoria IEQ Guaicurus",
            "tipo": "diretoria",
            "filhos": []
        }

    # ----------------------------------------------------------
    # RESPONSÁVEL
    # ----------------------------------------------------------

    if not df_alunos.empty:

        match = pd.DataFrame()

        if login_limpo:

            match = df_alunos[
                df_alunos[
                    "__TELEFONE"
                ].astype(str)
                == login_limpo
            ]

        if match.empty and login_str:

            match = df_alunos[
                df_alunos[
                    "__RESPONSAVEL"
                ]
                .astype(str)
                .apply(normalizar_texto)
                .str.contains(
                    login_str,
                    na=False
                )
            ]

        if not match.empty:

            primeiro = match.iloc[0]

            telefone = primeiro.get(
                "__TELEFONE",
                ""
            )

            if telefone in st.session_state.alunos_desistentes:
                return None

            senha_correta = (
                st.session_state.senhas_pais
                .get(
                    telefone,
                    "123456"
                )
            )

            if senha_input == senha_correta:

                return {
                    "nome": primeiro.get(
                        "__RESPONSAVEL",
                        "Responsável"
                    ),

                    "tipo": "pai",

                    "fone": telefone,

                    "filhos": match.to_dict(
                        orient="records"
                    )
                }

    return None


# ==============================================================================
# 10. CRONOGRAMA / EVENTOS DO DIA
# ==============================================================================

def preparar_eventos(df):

    if df.empty:
        return df.copy()

    dados = df.copy()

    col_data = encontrar_coluna(
        dados,
        [
            "Data Evento",
            "Data do Evento",
            "Data",
            "Data Evento "
        ]
    )

    if col_data:

        dados["__DATA_EVENTO"] = dados[
            col_data
        ].apply(
            converter_data
        )

    else:

        dados["__DATA_EVENTO"] = None

    return dados


df_eventos = preparar_eventos(
    df_cronograma
)

data_hoje = datetime.date.today()


if not df_eventos.empty:

    df_eventos_hoje = df_eventos[
        df_eventos[
            "__DATA_EVENTO"
        ].apply(
            lambda x: x == data_hoje
        )
    ].copy()

else:

    df_eventos_hoje = pd.DataFrame()


tem_evento_hoje = (
    not df_eventos_hoje.empty
)


# ==============================================================================
# 11. VERIFICAÇÃO DE NOTIFICAÇÃO
# ==============================================================================

data_hoje_codigo = data_hoje.strftime(
    "%d/%m/%Y"
)

notificacao_ja_visualizada = (
    st.session_state.data_notificacao_visualizada
    ==
    data_hoje_codigo
)

sino_vermelho = (
    tem_evento_hoje
    and
    not notificacao_ja_visualizada
)


# ==============================================================================
# 12. TELA DE LOGIN
# ==============================================================================

if not st.session_state.logged_in:

    st.markdown(
        "<br><br>",
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns(
        [1, 2, 1]
    )

    with c2:

        logo = Path(
            "logo_projeto_sementes.png"
        )

        if logo.exists():

            st.image(
                str(logo),
                use_container_width=True
            )

    st.markdown(
        """
        <h2 style="
        text-align:center;
        color:#F59E0B;
        font-weight:900;">
        PROJETO SEMENTES
        </h2>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style="
        text-align:center;
        color:#CBD5E1;">
        Iniciativa Voluntária de Jiu-Jitsu
        e Apoio à Família
        </p>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style="
        text-align:center;
        color:#64748B;">
        Cessão de Espaço Comunitário:
        IEQ Guaicurus
        </p>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)

    with st.form("form_login"):

        login = st.text_input(
            "Login do Responsável",
            placeholder="Digite seu WhatsApp ou CPF"
        )

        senha = st.text_input(
            "Senha de Acesso",
            type="password",
            placeholder="Digite sua senha"
        )

        entrar = st.form_submit_button(
            "ENTRAR NO APLICATIVO"
        )

        if entrar:

            usuario = autenticar_usuario(
                login,
                senha
            )

            if usuario:

                st.session_state.logged_in = True

                st.session_state.user_info = usuario

                st.session_state.sino_visualizado = False

                st.rerun()

            else:

                st.error(
                    "Credenciais inválidas ou acesso removido. Consulte a coordenação."
                )

    st.stop()


# ==============================================================================
# 13. USUÁRIO LOGADO
# ==============================================================================

usuario = st.session_state.user_info

if usuario.get("tipo") in [
    "mestre",
    "diretoria"
]:

    saudacao = (
        "A Paz seja convosco, "
        "Mestre London / Diretoria"
    )

else:

    saudacao = (
        "A Paz seja convosco, "
        +
        str(
            usuario.get(
                "nome",
                "Responsável"
            )
        )
    )


# ==============================================================================
# 14. CABEÇALHO — SAUDAÇÃO + DATA/HORA + SINO
# ==============================================================================

col1, col2 = st.columns(
    [4, 1]
)

with col1:

    st.markdown(
        f"""
        <div style="
        padding:8px 0;">
        
        <strong style="
        color:#F59E0B;
        font-size:1.05rem;">
        {saudacao}
        </strong>

        <br>

        <span style="
        color:#CBD5E1;
        font-size:.9rem;">
        📅 {data_hoje.strftime("%d/%m/%Y")}
        &nbsp;&nbsp;
        🕐 {datetime.datetime.now().strftime("%H:%M:%S")}
        </span>

        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    if sino_vermelho:

        st.markdown(
            """
            <div class="bell-red">
            🔔
            </div>
            """,
            unsafe_allow_html=True
        )

    else:

        st.markdown(
            """
            <div class="bell-gold">
            🔔
            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 15. BOTÃO DO SINO
# ==============================================================================

if sino_vermelho:

    if st.button(
        "🔔",
        key="abrir_sino",
        help="Ver evento de hoje"
    ):

        st.session_state.data_notificacao_visualizada = (
            data_hoje_codigo
        )

        st.session_state.sino_visualizado = True

        st.rerun()

elif st.session_state.get(
    "sino_visualizado",
    False
):

    if st.button(
        "🔔",
        key="abrir_sino_visualizado",
        help="Ver evento de hoje"
    ):

        st.session_state.sino_visualizado = False

        st.rerun()


# ==============================================================================
# 16. MURAL DO EVENTO DO DIA
# ==============================================================================

mostrar_mural = (
    st.session_state.get(
        "sino_visualizado",
        False
    )
    and
    not df_eventos_hoje.empty
)


if mostrar_mural:

    st.markdown(
        """
        <div class="event-wall">

        <h2>
        🔔 EVENTO DE HOJE
        </h2>

        <p style="
        text-align:center;
        font-weight:900;">
        PROJETO SEMENTES
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    for _, evento in df_eventos_hoje.iterrows():

        st.markdown(
            '<div class="gold-card">',
            unsafe_allow_html=True
        )

        for coluna in df_cronograma.columns:

            if str(coluna).startswith("__"):
                continue

            valor = evento.get(
                coluna,
                ""
            )

            if not valor_preenchido(valor):
                continue

            if (
                "data" in normalizar_texto(coluna)
            ):

                valor = formatar_data(
                    valor
                )

            st.markdown(
                f"""
                <p>
                <strong>{coluna}:</strong>
                {valor}
                </p>
                """,
                unsafe_allow_html=True
            )

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    if st.button(
        "FECHAR AVISO",
        key="fechar_mural"
    ):

        st.session_state.sino_visualizado = False

        st.rerun()


# ==============================================================================
# 17. ABAS
# ==============================================================================

tab_inicio, \
tab_matricula, \
tab_cronograma, \
tab_frequencia, \
tab_tatame, \
tab_campeao, \
tab_secretaria = st.tabs(
    [
        "🏠 Início",
        "📝 Matrícula / Cadastrados",
        "📅 Cronograma",
        "📊 Frequência",
        "🥋 Tatame",
        "🏆 Campeão & Fé",
        "⚙️ Secretaria"
    ]
)


# ==============================================================================
# 18. ABA INÍCIO
# ==============================================================================

with tab_inicio:

    st.markdown(
        """
        <div class="welcome-card">

        <h3 style="
        color:#F59E0B;
        text-align:center;">
        🌱 A HISTÓRIA DO PROJETO SEMENTES
        </h3>

        <p style="
        line-height:1.7;
        text-align:justify;">

        O <strong>Projeto Sementes</strong> nasceu no coração
        de Deus e, por Sua misericórdia, foi compartilhado aos
        corações do <strong>Pastor Joel Amorim</strong>, da
        IEQ Guaicurus, e do <strong>Instrutor Faixa-Preta
        London Carvalho</strong>, da Academia Iron Jiu-Jitsu.

        <br><br>

        Foram dias de oração, planejamento e dedicação,
        buscando estruturar o projeto da melhor maneira para
        acolher e atender crianças das comunidades próximas
        à igreja.

        <br><br>

        Assim, com o apoio dos instrutores e dos pais,
        iniciamos nossas aulas. Desde então, nossos alunos
        vêm aprendendo muito mais do que a Arte Suave.

        <br><br>

        Ensinamos valores morais e éticos fundamentados
        em Cristo Jesus, unindo o Jiu-Jitsu a palestras
        socioeducativas e ações de evangelismo.

        <br><br>

        Nosso grande sonho é levar a Semente do Evangelho
        até a sua casa, alcançando não apenas nossos alunos,
        mas também suas famílias.

        </p>

        <strong style="
        color:#F59E0B;
        display:block;
        text-align:center;">
        Sejam todos bem-vindos a esta família! 🌱🥋
        </strong>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")

    st.markdown(
        "### 🥋 Perfil do Aluno"
    )

    if usuario.get("tipo") in [
        "mestre",
        "diretoria"
    ]:

        if df_alunos.empty:

            st.info(
                "Nenhum aluno encontrado na Ficha de Cadastro."
            )

        else:

            nomes = sorted(
                df_alunos[
                    "__ALUNO"
                ].unique().tolist()
            )

            aluno_selecionado = st.selectbox(
                "Selecione o aluno:",
                nomes
            )

            registros = df_alunos[
                df_alunos[
                    "__ALUNO"
                ] == aluno_selecionado
            ]

            filhos = registros.to_dict(
                orient="records"
            )

    else:

        filhos = usuario.get(
            "filhos",
            []
        )

    for aluno in filhos:

        nome = aluno.get(
            "__ALUNO",
            "Aluno"
        )

        faixa = aluno.get(
            "__FAIXA",
            ""
        )

        responsavel = aluno.get(
            "__RESPONSAVEL",
            ""
        )

        endereco = aluno.get(
            "__ENDERECO",
            ""
        )

        status = aluno_status(
            aluno
        )

        if status == "verde":

            bolinha = (
                '<span class="status-verde"></span>'
            )

            texto_status = (
                "Cadastro completo + Biometria"
            )

        else:

            bolinha = (
                '<span class="status-amarelo"></span>'
            )

            texto_status = (
                "Cadastro pendente / aguardando biometria"
            )

        st.markdown(
            f"""
            <div class="gold-card">

            <h3>
            {bolinha}
            &nbsp;
            {nome}
            </h3>

            <p>
            <strong>Faixa:</strong>
            {faixa or "Não informado"}
            <br>

            <strong>Responsável:</strong>
            {responsavel or "Não informado"}
            <br>

            <strong>Endereço:</strong>
            {endereco or "Não informado"}
            <br>

            <strong>Status:</strong>
            {texto_status}
            </p>

            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 19. ABA MATRÍCULA / CADASTRADOS
# ==============================================================================

with tab_matricula:

    st.markdown(
        "### 📝 Matrícula / Cadastrados"
    )

    st.caption(
        "Dados carregados automaticamente da aba Ficha de Cadastro."
    )

    if df_alunos.empty:

        st.warning(
            "Nenhum aluno encontrado na aba Ficha de Cadastro."
        )

    else:

        col_m1, col_m2 = st.columns(2)

        total_alunos = len(df_alunos)

        total_verdes = len(
            df_alunos[
                df_alunos["__STATUS"]
                == "verde"
            ]
        )

        total_amarelos = (
            total_alunos
            -
            total_verdes
        )

        with col_m1:

            st.metric(
                "Total de alunos",
                total_alunos
            )

        with col_m2:

            st.metric(
                "Cadastro completo + biometria",
                total_verdes
            )

        st.markdown("---")

        for _, aluno in df_alunos.iterrows():

            nome = aluno.get(
                "__ALUNO",
                "Aluno"
            )

            status = aluno.get(
                "__STATUS",
                "amarelo"
            )

            pendencia = aluno.get(
                "__PENDENCIA",
                ""
            )

            if status == "verde":

                bolinha = (
                    '<span class="status-verde"></span>'
                )

                descricao = (
                    "Cadastro completo e biometria cadastrada."
                )

            else:

                bolinha = (
                    '<span class="status-amarelo"></span>'
                )

                descricao = (
                    "Cadastro incompleto ou aguardando biometria."
                )

            # --------------------------------------------------
            # PENDÊNCIAS
            # --------------------------------------------------

            pendencia_normalizada = normalizar_texto(
                pendencia
            )

            if (
                "falta endereco"
                in pendencia_normalizada
            ):

                descricao = (
                    "Falta Endereço. "
                    "Preencher os demais dados posteriormente."
                )

            elif (
                "falta documentos"
                in pendencia_normalizada
                or
                "falta documento"
                in pendencia_normalizada
            ):

                descricao = (
                    "Falta Documentos. "
                    "Preencher os demais dados posteriormente."
                )

            st.markdown(
                f"""
                <div class="gold-card">

                <h3>
                {bolinha}
                &nbsp;
                {nome}
                </h3>

                <p>
                <strong>Status:</strong>
                {descricao}
                </p>

                </div>
                """,
                unsafe_allow_html=True
            )


# ==============================================================================
# 20. ABA CRONOGRAMA
# ==============================================================================

with tab_cronograma:

    st.markdown(
        "### 📅 Cronograma Oficial de Eventos"
    )

    st.caption(
        "Fonte: aba Cronograma de Eventos da planilha base."
    )

    if df_cronograma.empty:

        st.info(
            "Nenhum evento encontrado."
        )

    else:

        mostrar = df_cronograma.copy()

        if "__DATA_EVENTO" in mostrar.columns:

            mostrar = mostrar.drop(
                columns=["__DATA_EVENTO"]
            )

        st.dataframe(
            mostrar,
            use_container_width=True,
            hide_index=True
        )


# ==============================================================================
# 21. FUNÇÃO DE FREQUÊNCIA
# ==============================================================================

def preparar_frequencia(df, nome_mes):

    if df.empty:
        return pd.DataFrame()

    dados = df.copy()

    coluna_nome = encontrar_coluna(
        dados,
        [
            "Nome do Aluno",
            "Nome Aluno",
            "Aluno",
            "Nome"
        ]
    )

    if not coluna_nome:
        return pd.DataFrame()

    resultado = []

    for _, linha in dados.iterrows():

        nome = linha.get(
            coluna_nome,
            ""
        )

        if not valor_preenchido(nome):
            continue

        total_dias = 0
        presencas = 0

        detalhes = []

        for coluna in dados.columns:

            if coluna == coluna_nome:
                continue

            texto_coluna = str(
                coluna
            ).strip()

            # Identifica colunas que representam dias
            match = re.search(
                r"\b(\d{1,2})\b",
                texto_coluna
            )

            if not match:
                continue

            dia = int(
                match.group(1)
            )

            if dia < 1 or dia > 31:
                continue

            valor = linha.get(
                coluna,
                ""
            )

            total_dias += 1

            presente = (
                normalizar_texto(
                    valor
                ) == "x"
                or
                str(valor).strip().upper()
                == "X"
            )

            if presente:

                presencas += 1

                detalhes.append(
                    {
                        "Dia": dia,
                        "Status": "🟢 Presente"
                    }
                )

            else:

                detalhes.append(
                    {
                        "Dia": dia,
                        "Status": "🔴 Falta"
                    }
                )

        percentual = (
            (presencas / total_dias) * 100
            if total_dias
            else 0
        )

        resultado.append(
            {
                "Nome do Aluno": str(nome),
                "Mês": nome_mes,
                "Presenças": presencas,
                "Dias Registrados": total_dias,
                "Frequência (%)": round(
                    percentual,
                    1
                ),
                "__DETALHES": detalhes
            }
        )

    return pd.DataFrame(
        resultado
    )


df_freq_agosto = preparar_frequencia(
    df_agosto,
    "Agosto"
)

df_freq_setembro = preparar_frequencia(
    df_setembro,
    "Setembro"
)


# ==============================================================================
# 22. ABA FREQUÊNCIA
# ==============================================================================

with tab_frequencia:

    st.markdown(
        "### 📊 Frequência"
    )

    st.caption(
        "Dados das abas Lista de Presença Agosto e Lista de Presença Setembro."
    )

    meses_disponiveis = []

    if not df_freq_agosto.empty:
        meses_disponiveis.append(
            "Agosto"
        )

    if not df_freq_setembro.empty:
        meses_disponiveis.append(
            "Setembro"
        )

    if not meses_disponiveis:

        st.info(
            "Nenhum registro de frequência encontrado."
        )

    else:

        mes_escolhido = st.selectbox(
            "Selecione o mês:",
            meses_disponiveis
        )

        if mes_escolhido == "Agosto":

            df_freq = df_freq_agosto

        else:

            df_freq = df_freq_setembro

        if df_freq.empty:

            st.info(
                "Nenhuma frequência encontrada para este mês."
            )

        else:

            total_presencas = int(
                df_freq[
                    "Presenças"
                ].sum()
            )

            media = (
                df_freq[
                    "Frequência (%)"
                ].mean()
            )

            c1, c2 = st.columns(2)

            with c1:

                st.metric(
                    "Presenças registradas",
                    total_presencas
                )

            with c2:

                st.metric(
                    "Frequência média",
                    f"{media:.1f}%"
                )

            st.markdown("---")

            st.dataframe(
                df_freq.drop(
                    columns=["__DETALHES"],
                    errors="ignore"
                ),
                use_container_width=True,
                hide_index=True
            )

            st.markdown(
                "### 🗓️ Detalhamento por aluno"
            )

            nomes = df_freq[
                "Nome do Aluno"
            ].tolist()

            aluno_freq = st.selectbox(
                "Selecione o aluno:",
                nomes
            )

            registro = df_freq[
                df_freq[
                    "Nome do Aluno"
                ] == aluno_freq
            ].iloc[0]

            detalhes = registro[
                "__DETALHES"
            ]

            if detalhes:

                st.dataframe(
                    pd.DataFrame(
                        detalhes
                    ),
                    use_container_width=True,
                    hide_index=True
                )


# ==============================================================================
# 23. ABA TATAME
# ==============================================================================

with tab_tatame:

    st.markdown(
        "### 🥋 Tatame — Chamada do Treino"
    )

    st.caption(
        f"Data do treino: {data_hoje.strftime('%d/%m/%Y')}"
    )

    if df_alunos.empty:

        st.warning(
            "Nenhum aluno cadastrado."
        )

    else:

        nomes_alunos = sorted(
            df_alunos[
                "__ALUNO"
            ].unique().tolist()
        )

        for nome in nomes_alunos:

            chave = (
                f"{data_hoje.isoformat()}_"
                f"{normalizar_texto(nome)}"
            )

            atual = st.session_state.presencas_tatame.get(
                chave,
                False
            )

            marcado = st.checkbox(
                f"🟢 Presente: {nome}",
                value=atual,
                key=f"tatame_{chave}"
            )

            st.session_state.presencas_tatame[
                chave
            ] = marcado

        st.markdown("---")

        if st.button(
            "REGISTRAR PRESENÇAS DO TREINO",
            key="registrar_tatame"
        ):

            presentes = []

            for nome in nomes_alunos:

                chave = (
                    f"{data_hoje.isoformat()}_"
                    f"{normalizar_texto(nome)}"
                )

                if st.session_state.presencas_tatame.get(
                    chave,
                    False
                ):

                    presentes.append(
                        nome
                    )

            st.success(
                f"{len(presentes)} presença(s) registrada(s) para "
                f"{data_hoje.strftime('%d/%m/%Y')}."
            )

            if presentes:

                st.dataframe(
                    pd.DataFrame(
                        {
                            "Nome do Aluno":
                                presentes,
                            "Data":
                                [
                                    data_hoje.strftime(
                                        "%d/%m/%Y"
                                    )
                                    for _ in presentes
                                ],
                            "Status":
                                [
                                    "🟢 Presente"
                                    for _ in presentes
                                ]
                        }
                    ),
                    use_container_width=True,
                    hide_index=True
                )

            st.info(
                "As presenças registradas no Tatame ficam disponíveis "
                "para alimentar a frequência do aplicativo. "
                "A gravação permanente na planilha depende de uma rotina "
                "de escrita no Excel."
            )


# ==============================================================================
# 24. ABA CAMPEÃO & FÉ
# ==============================================================================

with tab_campeao:

    st.markdown(
        "### 🏆 Campeão & Fé"
    )

    # --------------------------------------------------------------------------
    # CAMPEÃO
    # --------------------------------------------------------------------------

    st.markdown(
        "## 🏆 Campeonatos"
    )

    if df_campeao.empty:

        st.info(
            "Nenhuma informação encontrada na aba Campeão."
        )

    else:

        st.dataframe(
            df_campeao,
            use_container_width=True,
            hide_index=True
        )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # FÉ
    # --------------------------------------------------------------------------

    st.markdown(
        "## ⛪ Fé — Células"
    )

    if df_fe.empty:

        st.info(
            "Nenhuma informação encontrada na aba Fé."
        )

    else:

        col_setor = encontrar_coluna(
            df_fe,
            [
                "Setor"
            ]
        )

        col_nome = encontrar_coluna(
            df_fe,
            [
                "Nomes",
                "Nome",
                "Líder",
                "Lider"
            ]
        )

        col_contato = encontrar_coluna(
            df_fe,
            [
                "Contato",
                "WhatsApp",
                "Whatsapp",
                "Telefone"
            ]
        )

        if col_setor and col_nome:

            dados_celulas = df_fe[
                [
                    col_setor,
                    col_nome
                ]
                +
                (
                    [col_contato]
                    if col_contato
                    else []
                )
            ].copy()

            # Somente registros de célula
            mascara = dados_celulas[
                col_setor
            ].astype(str).str.contains(
                "celula|célula",
                case=False,
                na=False
            )

            dados_celulas = dados_celulas[
                mascara
            ]

            # Limita visualmente às cinco células solicitadas
            dados_celulas = dados_celulas.head(
                5
            )

            if dados_celulas.empty:

                st.info(
                    "Nenhuma célula cadastrada na aba Fé."
                )

            else:

                for _, celula in dados_celulas.iterrows():

                    setor = celula.get(
                        col_setor,
                        ""
                    )

                    lider = celula.get(
                        col_nome,
                        ""
                    )

                    contato = (
                        celula.get(
                            col_contato,
                            ""
                        )
                        if col_contato
                        else ""
                    )

                    st.markdown(
                        f"""
                        <div class="gold-card">

                        <h3>
                        ⛪ {setor}
                        </h3>

                        <p>
                        👤 <strong>Líder:</strong>
                        {lider}
                        </p>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    link = gerar_link_whatsapp(
                        contato,
                        f"Paz do Senhor, {lider}!"
                    )

                    if link:

                        st.markdown(
                            f"""
                            <a href="{link}"
                            target="_blank"
                            style="
                            display:block;
                            text-align:center;
                            background:#16A34A;
                            color:white;
                            padding:12px;
                            border-radius:10px;
                            font-weight:800;
                            text-decoration:none;
                            margin-bottom:15px;">
                            📲 Falar com o Líder no WhatsApp
                            </a>
                            """,
                            unsafe_allow_html=True
                        )

        else:

            st.warning(
                "Não foi possível identificar as colunas Setor e Nomes da aba Fé."
            )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # PASTOR JOEL — SEPARADO DAS CÉLULAS
    # --------------------------------------------------------------------------

    st.markdown(
        "## 🙏 Pedido de Oração ao Pastor Joel"
    )

    contato_pastor = ""

    nome_pastor = "Pastor Joel"

    if not df_fe.empty:

        col_setor_fe = encontrar_coluna(
            df_fe,
            [
                "Setor"
            ]
        )

        col_nome_fe = encontrar_coluna(
            df_fe,
            [
                "Nomes",
                "Nome",
                "Líder",
                "Lider"
            ]
        )

        col_contato_fe = encontrar_coluna(
            df_fe,
            [
                "Contato",
                "WhatsApp",
                "Whatsapp",
                "Telefone"
            ]
        )

        if col_nome_fe:

            for _, linha in df_fe.iterrows():

                texto = normalizar_texto(
                    linha.get(
                        col_nome_fe,
                        ""
                    )
                )

                if (
                    "pastor joel" in texto
                    or
                    texto == "joel"
                ):

                    nome_pastor = str(
                        linha.get(
                            col_nome_fe,
                            "Pastor Joel"
                        )
                    )

                    if col_contato_fe:

                        contato_pastor = linha.get(
                            col_contato_fe,
                            ""
                        )

                    break

    with st.form(
        "form_pedido_oracao"
    ):

        nome_oracao = st.text_input(
            "Seu Nome"
        )

        pedido = st.text_area(
            "Escreva seu pedido de oração"
        )

        enviar_oracao = st.form_submit_button(
            "ENVIAR PEDIDO DE ORAÇÃO"
        )

        if enviar_oracao:

            if not pedido.strip():

                st.warning(
                    "Digite seu pedido de oração."
                )

            else:

                st.session_state.pedidos_oracao.append(
                    {
                        "nome":
                            nome_oracao,
                        "pedido":
                            pedido,
                        "data":
                            data_hoje.strftime(
                                "%d/%m/%Y"
                            )
                    }
                )

                st.success(
                    "Pedido de oração registrado com sucesso."
                )

    link_pastor = gerar_link_whatsapp(
        contato_pastor,
        "Paz do Senhor, Pastor Joel. Gostaria de enviar um pedido de oração."
    )

    if link_pastor:

        st.markdown(
            f"""
            <a href="{link_pastor}"
            target="_blank"
            style="
            display:block;
            text-align:center;
            background:#16A34A;
            color:white;
            padding:12px;
            border-radius:10px;
            font-weight:800;
            text-decoration:none;
            margin-top:10px;">
            📲 Falar com {nome_pastor} pelo WhatsApp
            </a>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # BATISMO
    # --------------------------------------------------------------------------

    st.markdown(
        "## ⭐ Eu quero Batizar!"
    )

    if st.button(
        "🙌 QUERO ME BATIZAR NAS ÁGUAS!",
        key="botao_batismo"
    ):

        st.success(
            "Glória a Deus! Sua intenção de batismo foi registrada."
        )


# ==============================================================================
# 25. MAPA DE BAIRROS
# ==============================================================================

def obter_bairros(df):

    if df.empty:
        return pd.DataFrame()

    dados = df.copy()

    bairros = []

    for _, linha in dados.iterrows():

        bairro = linha.get(
            "__BAIRRO",
            ""
        )

        endereco = linha.get(
            "__ENDERECO",
            ""
        )

        # Se houver bairro específico, utilizar
        if valor_preenchido(bairro):

            bairro_final = str(
                bairro
            ).strip()

        elif valor_preenchido(endereco):

            # Tenta identificar bairro no endereço
            texto = str(
                endereco
            ).strip()

            encontrado = re.search(
                r"bairro\s+([^,\-]+)",
                texto,
                re.IGNORECASE
            )

            if encontrado:

                bairro_final = (
                    encontrado
                    .group(1)
                    .strip()
                )

            else:

                bairro_final = (
                    "Endereço informado "
                    "(bairro não identificado)"
                )

        else:

            bairro_final = (
                "Endereço não informado"
            )

        bairros.append(
            bairro_final
        )

    dados["Bairro Calculado"] = bairros

    resumo = (
        dados[
            "Bairro Calculado"
        ]
        .value_counts()
        .reset_index()
    )

    resumo.columns = [
        "Bairro",
        "Alunos"
    ]

    total = resumo[
        "Alunos"
    ].sum()

    if total:

        resumo["Porcentagem (%)"] = (
            resumo["Alunos"]
            / total
            * 100
        ).round(1)

    else:

        resumo["Porcentagem (%)"] = 0

    return resumo


df_mapa = obter_bairros(
    df_alunos
)


# ==============================================================================
# 26. ABA SECRETARIA
# ==============================================================================

with tab_secretaria:

    st.markdown(
        "### ⚙️ Secretaria"
    )

    if usuario.get("tipo") in [
        "mestre",
        "diretoria"
    ]:

        st.success(
            "👑 Painel de Controle — Mestre / Diretoria"
        )

        st.markdown("---")

        # ------------------------------------------------------
        # SENHAS
        # ------------------------------------------------------

        st.markdown(
            "#### 🔐 Gestão de Senhas"
        )

        tipo_usuario = st.selectbox(
            "Perfil",
            [
                "Pai / Responsável",
                "Diretoria"
            ]
        )

        chave_usuario = st.text_input(
            "WhatsApp ou login",
            value=(
                "67998411953"
                if tipo_usuario ==
                "Pai / Responsável"
                else "diretoria"
            )
        )

        nova_senha = st.text_input(
            "Nova senha",
            value="123456"
        )

        if st.button(
            "SALVAR / ALTERAR SENHA",
            key="alterar_senha"
        ):

            if tipo_usuario == "Pai / Responsável":

                st.session_state.senhas_pais[
                    limpar_numero(
                        chave_usuario
                    )
                ] = nova_senha

            else:

                st.session_state.senhas_diretoria[
                    normalizar_texto(
                        chave_usuario
                    )
                ] = nova_senha

            st.success(
                "Senha atualizada."
            )

        st.markdown("---")

        # ------------------------------------------------------
        # EXCLUSÃO DE ACESSO
        # ------------------------------------------------------

        st.markdown(
            "#### 🚫 Excluir Acesso"
        )

        telefone_exclusao = st.text_input(
            "WhatsApp do responsável",
            key="telefone_exclusao"
        )

        if st.button(
            "EXCLUIR ACESSO",
            key="excluir_acesso"
        ):

            numero = limpar_numero(
                telefone_exclusao
            )

            if numero:

                if numero not in st.session_state.alunos_desistentes:

                    st.session_state.alunos_desistentes.append(
                        numero
                    )

                st.warning(
                    "Acesso removido."
                )

    st.markdown("---")

    # ==========================================================================
    # MAPA DE ALCANCE
    # ==========================================================================

    st.markdown(
        "### 🗺️ Mapa de Alcance por Bairro"
    )

    st.caption(
        "Os dados são calculados automaticamente a partir dos endereços "
        "existentes em Matrícula / Cadastrados."
    )

    if df_mapa.empty:

        st.info(
            "Nenhum endereço cadastrado."
        )

    else:

        st.dataframe(
            df_mapa,
            use_container_width=True,
            hide_index=True
        )

        try:

            grafico = df_mapa.set_index(
                "Bairro"
            )[
                "Alunos"
            ]

            st.bar_chart(
                grafico
            )

        except Exception:
            pass

    st.markdown("---")

    # ==========================================================================
    # ATUALIZAR PLANILHA
    # ==========================================================================

    if st.button(
        "🔄 ATUALIZAR DADOS DA PLANILHA",
        key="atualizar_planilha"
    ):

        st.cache_data.clear()

        st.rerun()

    st.markdown("---")

    # ==========================================================================
    # SAIR
    # ==========================================================================

    if st.button(
        "🚪 SAIR DO APLICATIVO",
        key="sair_aplicativo"
    ):

        st.session_state.logged_in = False

        st.session_state.user_info = None

        st.rerun()
