import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import os
import json
import base64
import re
from pathlib import Path
from io import BytesIO

try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    CANVAS_DISPONIVEL = False

try:
    import face_recognition
    FACE_RECOGNITION_DISPONIVEL = True
except ImportError:
    FACE_RECOGNITION_DISPONIVEL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.lib.utils import ImageReader
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False


# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA MOBILE-FIRST & ALTO CONTRASTE
# ==============================================================================

st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

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
    
    .stTextInput input, 
    .stSelectbox div[data-baseweb="select"], 
    .stTextArea textarea,
    .stDateInput input,
    .stNumberInput input {
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

    .perfil-nome {
        color: #FFFFFF !important;
        font-size: 1.25rem !important;
        font-weight: 900 !important;
        text-shadow: 0 0 8px rgba(255,255,255,0.25);
    }

    .status-verde {
        color: #22C55E !important;
        font-weight: 900 !important;
    }

    .status-vermelho {
        color: #EF4444 !important;
        font-weight: 900 !important;
    }

    .status-amarelo {
        color: #F59E0B !important;
        font-weight: 900 !important;
    }

    .public-card {
        background-color: #111827;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 14px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 2. CONFIGURAÇÕES GERAIS
# ==============================================================================

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

DATA_DIR = Path("dados_aplicativo")
FOTOS_DIR = DATA_DIR / "fotos_alunos"
DOCS_DIR = DATA_DIR / "documentos_assinados"

DATA_DIR.mkdir(exist_ok=True)
FOTOS_DIR.mkdir(exist_ok=True)
DOCS_DIR.mkdir(exist_ok=True)

JSON_FILE = DATA_DIR / "dados_aplicativo.json"

LOGIN_MESTRE = "Mikhael"
SENHA_MESTRE = "12381314"

DIAS_COMEMORACAO = {
    9: {
        "data": datetime.date(2026, 9, 29),
        "texto": "29/09/2026 (Terça-feira)"
    },
    10: {
        "data": datetime.date(2026, 10, 29),
        "texto": "29/10/2026 (Quinta-feira)"
    },
    11: {
        "data": datetime.date(2026, 11, 26),
        "texto": "26/11/2026 (Quinta-feira)"
    },
    12: {
        "data": datetime.date(2026, 12, 31),
        "texto": "31/12/2026 (Quinta-feira)"
    }
}

DOENCAS_OPCOES = [
    "Problema cardíaco ou de pressão",
    "Asma, bronquite ou problema respiratório",
    "Lesão óssea, muscular ou articular crônica",
    "Uso regular de medicamento controlado",
    "Alergia a medicamento ou substância",
    "Desmaios ou tonturas durante exercícios físicos",
    "Nenhuma das opções acima"
]

DESENVOLVIMENTO_OPCOES = [
    "Ótimo",
    "Bom",
    "Regular",
    "Dificuldade"
]

DESEMPENHO_MESTRE = [
    "Ótimo",
    "Regular",
    "Dificuldade"
]


# ==============================================================================
# 3. DADOS LOCAIS DO APLICATIVO
# ==============================================================================

def carregar_json():
    if JSON_FILE.exists():
        try:
            with open(JSON_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    return {
        "alunos": {},
        "avaliacoes_familia": {},
        "desempenho_mestre": {},
        "campeonatos": {},
        "frequencias": {},
        "presencas_tatame": {},
        "pedidos_oracao": [],
        "batizados": [],
        "celulas": [],
        "assinaturas": {},
        "biometrias": {},
        "senhas_pais": {},
        "senhas_diretoria": {
            "diretoria": "dir2026"
        }
    }


def salvar_json():
    try:
        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(st.session_state.dados_app, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


if "dados_app" not in st.session_state:
    st.session_state.dados_app = carregar_json()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "user_info" not in st.session_state:
    st.session_state.user_info = None

if "sino_visto" not in st.session_state:
    st.session_state.sino_visto = False


# ==============================================================================
# 4. FUNÇÕES AUXILIARES
# ==============================================================================

def limpar_numero(valor):
    return ''.join(filter(str.isdigit, str(valor)))


def normalizar_nome(nome):
    return re.sub(r"\s+", " ", str(nome).strip()).lower()


def gerar_link_whatsapp(numero, mensagem):
    if numero is None:
        return None

    numero_limpo = limpar_numero(numero)

    if not numero_limpo:
        return None

    if not numero_limpo.startswith("55") and len(numero_limpo) in [10, 11]:
        numero_limpo = "55" + numero_limpo

    msg = urllib.parse.quote(mensagem)

    return f"https://wa.me/{numero_limpo}?text={msg}"


def enviar_whatsapp(numero, mensagem, arquivo=None):
    """
    Estrutura preparada para WhatsApp.

    Sem API oficial configurada:
    retorna link do WhatsApp.

    Com API oficial configurada por variáveis de ambiente:
    poderá ser integrado posteriormente ao envio automático.
    """

    link = gerar_link_whatsapp(numero, mensagem)

    return link


def salvar_foto_aluno(nome, imagem_bytes):
    if not imagem_bytes:
        return None

    nome_arquivo = re.sub(
        r"[^a-zA-Z0-9_-]",
        "_",
        normalizar_nome(nome)
    )

    caminho = FOTOS_DIR / f"{nome_arquivo}.jpg"

    with open(caminho, "wb") as f:
        f.write(imagem_bytes)

    return str(caminho)


def foto_aluno(nome):
    dados = st.session_state.dados_app["alunos"].get(normalizar_nome(nome), {})

    caminho = dados.get("foto")

    if caminho and Path(caminho).exists():
        return caminho

    nome_arquivo = re.sub(
        r"[^a-zA-Z0-9_-]",
        "_",
        normalizar_nome(nome)
    )

    caminho_padrao = FOTOS_DIR / f"{nome_arquivo}.jpg"

    if caminho_padrao.exists():
        return str(caminho_padrao)

    return None


def salvar_assinatura(nome, termo, imagem):
    if imagem is None:
        return None

    chave = f"{normalizar_nome(nome)}_{termo}"

    arquivo = DOCS_DIR / f"{chave}.png"

    imagem.save(arquivo)

    return str(arquivo)


def gerar_documento_assinado(nome, dados, termos):
    if not REPORTLAB_DISPONIVEL:
        return None

    arquivo = DOCS_DIR / (
        re.sub(r"[^a-zA-Z0-9_-]", "_", normalizar_nome(nome))
        + "_documentacao.pdf"
    )

    c = pdf_canvas.Canvas(str(arquivo), pagesize=A4)

    largura, altura = A4

    y = altura - 45

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(
        largura / 2,
        y,
        "PROJETO SEMENTES - FICHA DE INSCRIÇÃO"
    )

    y -= 35

    c.setFont("Helvetica", 9)

    linhas = [
        f"Aluno: {nome}",
        f"Data de nascimento: {dados.get('data_nascimento', '')}",
        f"Escola: {dados.get('escola', '')}",
        f"Endereço: {dados.get('endereco', '')}",
        f"Responsável: {dados.get('responsavel', '')}",
        f"Grau de parentesco: {dados.get('parentesco', '')}",
        f"Contato: {dados.get('telefone', '')}",
        f"CIN: {dados.get('cin', '')}",
    ]

    for linha in linhas:
        c.drawString(40, y, linha)
        y -= 16

    y -= 15

    for titulo, texto in termos.items():

        if y < 120:
            c.showPage()
            y = altura - 45
            c.setFont("Helvetica", 9)

        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, titulo)

        y -= 18

        c.setFont("Helvetica", 8)

        palavras = texto.split()
        linha = ""

        for palavra in palavras:
            teste = linha + " " + palavra

            if c.stringWidth(teste, "Helvetica", 8) > 510:
                c.drawString(40, y, linha.strip())
                y -= 12
                linha = palavra
            else:
                linha = teste

        if linha:
            c.drawString(40, y, linha.strip())
            y -= 15

        y -= 10

    c.drawString(
        40,
        55,
        f"Documento gerado em {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    c.save()

    return str(arquivo)


def obter_aniversariantes_mes(df):
    if df.empty:
        return pd.DataFrame()

    coluna_data = next(
        (
            c for c in df.columns
            if any(
                termo in str(c).lower()
                for termo in [
                    "nascimento",
                    "data nasc",
                    "data de nascimento"
                ]
            )
        ),
        None
    )

    coluna_nome = "NOME_ALUNO_CLEAN"

    if not coluna_data or coluna_nome not in df.columns:
        return pd.DataFrame()

    temp = df.copy()

    temp["DATA_NASCIMENTO_CALC"] = pd.to_datetime(
        temp[coluna_data],
        errors="coerce",
        dayfirst=True
    )

    temp = temp.dropna(subset=["DATA_NASCIMENTO_CALC"])

    return temp


# ==============================================================================
# 5. CARREGAMENTO DO EXCEL
# ==============================================================================

@st.cache_data(ttl=30)
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

        df = df.dropna(how="all")
        df = df.dropna(how="all", axis=1)

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
                or "WhatsApp" in c
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
                else "Não informado"
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
                .apply(limpar_numero)
                if col_fone
                else ""
            )

            df = df[
                df["NOME_ALUNO_CLEAN"].str.lower() != "nan"
            ]

            df = df[
                df["NOME_ALUNO_CLEAN"].str.strip() != ""
            ]

            if not df.empty:
                return df

        return df_fallback

    except Exception:
        return df_fallback


@st.cache_data(ttl=30)
def carregar_cronograma_eventos():

    try:

        xls = pd.ExcelFile(
            EXCEL_FILE,
            engine="openpyxl"
        )

        nome_aba = next(
            (
                s for s in xls.sheet_names
                if "cronograma" in s.lower()
                or "evento" in s.lower()
            ),
            None
        )

        if nome_aba:

            df_ev = pd.read_excel(
                xls,
                sheet_name=nome_aba
            )

            return df_ev.dropna(how="all")

    except Exception:
        pass

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


@st.cache_data(ttl=30)
def carregar_frequencia_excel():

    try:

        xls = pd.ExcelFile(
            EXCEL_FILE,
            engine="openpyxl"
        )

        aba = next(
            (
                s for s in xls.sheet_names
                if "frequ" in s.lower()
                or "presen" in s.lower()
                or "chamada" in s.lower()
            ),
            None
        )

        if aba:

            df = pd.read_excel(
                xls,
                sheet_name=aba
            )

            return df.dropna(how="all")

    except Exception:
        pass

    return pd.DataFrame()


df_cadastro = carregar_dados_cadastro()
df_cronograma = carregar_cronograma_eventos()
df_frequencia_excel = carregar_frequencia_excel()


# ==============================================================================
# 6. AUTENTICAÇÃO
# ==============================================================================

def autenticar_usuario(login_input, senha_input):

    login = str(login_input).strip()
    senha = str(senha_input).strip()

    # NOVO LOGIN DO MESTRE
    if login.lower() == LOGIN_MESTRE.lower() and senha == SENHA_MESTRE:

        return {
            "nome": "Mestre Mikhael",
            "tipo": "mestre",
            "filhos": []
        }

    # DIRETORIA
    senhas_diretoria = (
        st.session_state.dados_app
        .get("senhas_diretoria", {})
    )

    login_diretoria = login.lower()

    if (
        login_diretoria in senhas_diretoria
        and senha == senhas_diretoria[login_diretoria]
    ):

        return {
            "nome": "Diretoria IEQ Guaicurus",
            "tipo": "diretoria",
            "filhos": []
        }

    # RESPONSÁVEIS
    login_numero = limpar_numero(login)

    if not df_cadastro.empty:

        if login_numero:

            match = df_cadastro[
                df_cadastro["FONE_LIMPO"]
                .astype(str)
                .str.contains(
                    login_numero,
                    na=False
                )
            ]

        else:

            match = df_cadastro[
                df_cadastro["RESPONSAVEL_CLEAN"]
                .astype(str)
                .str.lower()
                .str.contains(
                    login.lower(),
                    na=False
                )
            ]

        if not match.empty:

            resp_nome = match.iloc[0].get(
                "RESPONSAVEL_CLEAN",
                "Responsável"
            )

            fone = str(
                match.iloc[0].get(
                    "FONE_LIMPO",
                    ""
                )
            )

            senhas_pais = (
                st.session_state.dados_app
                .get("senhas_pais", {})
            )

            senha_correta = senhas_pais.get(
                fone,
                "123456"
            )

            if senha == senha_correta:

                return {
                    "nome": resp_nome,
                    "tipo": "pai",
                    "fone": fone,
                    "filhos": match.to_dict(
                        orient="records"
                    )
                }

    return None


# ==============================================================================
# 7. IDENTIFICAÇÃO DE ALUNOS DO RESPONSÁVEL
# ==============================================================================

def filhos_do_responsavel(usuario):

    if usuario.get("tipo") in [
        "mestre",
        "diretoria"
    ]:
        return df_cadastro.to_dict(
            orient="records"
        )

    nome_responsavel = normalizar_nome(
        usuario.get("nome", "")
    )

    if df_cadastro.empty:
        return []

    mask = (
        df_cadastro["RESPONSAVEL_CLEAN"]
        .astype(str)
        .apply(normalizar_nome)
        .eq(nome_responsavel)
    )

    return df_cadastro[mask].to_dict(
        orient="records"
    )


# ==============================================================================
# 8. NOTIFICAÇÃO DE ANIVERSARIANTES
# ==============================================================================

def aniversario_do_dia():

    hoje = datetime.date.today()

    resultado = []

    for mes, dados in DIAS_COMEMORACAO.items():

        if dados["data"] == hoje:

            resultado.append(dados)

    return resultado


def mensagem_aniversario(nome):

    return (
        f"Paz do Senhor, {nome}! 🌱🎉\n\n"
        f"Todo o Projeto Sementes deseja a você "
        f"um feliz aniversário!\n\n"
        f"Que Deus abençoe sua vida, sua família, "
        f"seus estudos e seus sonhos. "
        f"Que o Senhor Jesus continue guiando "
        f"cada passo da sua caminhada.\n\n"
        f"Receba o carinho de toda a família "
        f"Projeto Sementes! ❤️🥋🌱"
    )


# ==============================================================================
# 9. LOGIN
# ==============================================================================

if not st.session_state.logged_in:

    st.markdown("<br>", unsafe_allow_html=True)

    c_logo1, c_logo2, c_logo3 = st.columns([1, 2, 1])

    with c_logo2:

        if Path("logo_projeto_sementes.png").exists():

            st.image(
                "logo_projeto_sementes.png",
                use_container_width=True
            )

    st.markdown(
        """
        <h2 style='text-align:center;color:#F59E0B;font-weight:800;'>
        PROJETO SEMENTES
        </h2>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style='text-align:center;color:#CBD5E1;font-size:0.95rem;'>
        Iniciativa Voluntária de Jiu-Jitsu e Apoio à Família
        </p>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style='text-align:center;color:#64748B;font-size:0.8rem;'>
        Cessão de Espaço Comunitário: IEQ Guaicurus
        </p>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)

    with st.form("form_login"):

        user_in = st.text_input(
            "Login do Responsável",
            placeholder="Informe seu login"
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

                st.success(
                    "Acesso autorizado com sucesso!"
                )

                st.rerun()

            else:

                st.error(
                    "Credenciais inválidas ou acesso não autorizado."
                )


# ==============================================================================
# 10. APLICAÇÃO LOGADA
# ==============================================================================

else:

    u_info = st.session_state.user_info

    if u_info.get("tipo") in [
        "mestre",
        "diretoria"
    ]:

        saudacao = (
            "A Paz seja convosco, "
            f"{u_info.get('nome')}"
        )

    else:

        saudacao = (
            "A Paz seja convosco, "
            f"{u_info.get('nome')}"
        )

    col_head1, col_head2 = st.columns([5, 1])

    with col_head1:

        st.markdown(
            f"""
            <strong style='color:#F59E0B;font-size:1.1rem;'>
            {saudacao}
            </strong>
            """,
            unsafe_allow_html=True
        )

    # --------------------------------------------------------------------------
    # SININHO ÚNICO
    # --------------------------------------------------------------------------

    eventos_hoje = aniversario_do_dia()

    tem_evento_hoje = False

    hoje_str = datetime.date.today().strftime(
        "%d/%m/%Y"
    )

    if not df_cronograma.empty:

        for coluna in df_cronograma.columns:

            if (
                "data" in str(coluna).lower()
                or "dia" in str(coluna).lower()
            ):

                if (
                    df_cronograma[coluna]
                    .astype(str)
                    .str.contains(
                        hoje_str,
                        na=False
                    )
                    .any()
                ):

                    tem_evento_hoje = True

    if eventos_hoje:
        tem_evento_hoje = True

    with col_head2:

        icone_sino = (
            "🚨🔔"
            if tem_evento_hoje
            else "🔔"
        )

        if st.button(
            icone_sino,
            key="btn_sino"
        ):

            st.session_state.sino_visto = True

    if (
        st.session_state.sino_visto
        or tem_evento_hoje
    ):

        with st.expander(
            "🔔 Central Oficial de Notificações & Eventos",
            expanded=True
        ):

            if tem_evento_hoje:

                st.error(
                    "🚨 EXISTE UMA NOTIFICAÇÃO PARA HOJE!"
                )

            if eventos_hoje:

                st.warning(
                    "🎂 HOJE É DIA DA COMEMORAÇÃO "
                    "DOS ANIVERSARIANTES!"
                )

                st.markdown(
                    "Mensagem evangelística preparada "
                    "para o aniversariante."
                )

            st.markdown("---")

            st.markdown(
                "### 📅 Cronograma"
            )

            if not df_cronograma.empty:

                st.dataframe(
                    df_cronograma,
                    use_container_width=True,
                    hide_index=True
                )

            if st.button(
                "Fechar Notificações",
                key="fechar_notificacoes"
            ):

                st.session_state.sino_visto = False
                st.rerun()


    # ==========================================================================
    # ABAS
    # ==========================================================================

    tab1, tab_mat, tab_crono, tab_freq, tab_tatame, tab_campeao, tab_fe, tab_sec = st.tabs(
        [
            "🏠 Início",
            "📝 Matrícula",
            "📅 Cronograma",
            "📊 Frequência",
            "🥋 Tatame",
            "🏆 Campeão",
            "🙏 Fé",
            "⚙️ Secretaria"
        ]
    )


    # ==========================================================================
    # ABA INÍCIO
    # ==========================================================================

    with tab1:

        st.markdown(
            """
            <div class="welcome-card">
                <h3 style="color:#F59E0B;text-align:center;font-weight:800;">
                🌱 A HISTÓRIA DO PROJETO SEMENTES
                </h3>

                <p style="font-size:0.95rem;line-height:1.6;
                text-align:justify;color:#E2E8F0;">

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
                dos membros da igreja, que abraçaram o projeto.

                <br><br>

                Assim, com o apoio dos instrutores e dos pais, iniciamos
                nossas aulas.

                <br><br>

                Ensinamos <strong>valores morais e éticos fundamentados
                em Cristo Jesus</strong>, unindo o Jiu-Jitsu a
                <strong>palestras socioeducativas e ações de evangelismo</strong>,
                contribuindo para a formação integral de nossas crianças,
                adolescentes e jovens.

                <br><br>

                Nosso grande sonho é
                <strong>levar a Semente do Evangelho até a sua casa</strong>,
                alcançando não apenas nossos alunos, mas também suas famílias.

                <br><br>

                <strong style="color:#F59E0B;font-size:1.05rem;
                display:block;text-align:center;">
                Sejam todos bem-vindos a esta família! 🌱🥋
                </strong>

                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        # ----------------------------------------------------------------------
        # ANIVERSARIANTES
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### 🎂 Aniversariantes do Mês & Comemorações"
        )

        st.caption(
            "Comemorações oficiais da equipe e familiares — 19h"
        )

        for mes, info in DIAS_COMEMORACAO.items():

            nome_mes = {
                9: "Setembro",
                10: "Outubro",
                11: "Novembro",
                12: "Dezembro"
            }[mes]

            st.markdown(
                f"🎉 **{nome_mes}: {info['texto']} às 19h**"
            )

        # ----------------------------------------------------------------------
        # ANIVERSARIANTES REAIS CADASTRADOS
        # ----------------------------------------------------------------------

        st.markdown(
            "#### 👥 Aniversariantes cadastrados"
        )

        aniversariantes_df = obter_aniversariantes_mes(
            df_cadastro
        )

        mes_atual = datetime.date.today().month

        if not aniversariantes_df.empty:

            aniversariantes_mes = aniversariantes_df[
                aniversariantes_df[
                    "DATA_NASCIMENTO_CALC"
                ].dt.month == mes_atual
            ]

            if aniversariantes_mes.empty:

                st.info(
                    "Não há aniversariantes cadastrados "
                    "para o mês atual."
                )

            else:

                for _, aluno in aniversariantes_mes.iterrows():

                    nome = aluno["NOME_ALUNO_CLEAN"]

                    foto = foto_aluno(nome)

                    c1, c2 = st.columns([1, 3])

                    with c1:

                        if foto:
                            st.image(
                                foto,
                                width=100
                            )
                        else:
                            st.info("👤 Foto")

                    with c2:

                        data_nasc = aluno[
                            "DATA_NASCIMENTO_CALC"
                        ].strftime("%d/%m")

                        st.markdown(
                            f"""
                            <div class="gold-card">
                            <div class="perfil-nome">
                            🎂 {nome}
                            </div>
                            <p>
                            Aniversário: <strong>{data_nasc}</strong>
                            </p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

        # ----------------------------------------------------------------------
        # PERFIL DO ALUNO
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### 🥋 Perfil do Aluno"
        )

        filhos_exibir = filhos_do_responsavel(
            u_info
        )

        lista_alunos = [
            f.get(
                "NOME_ALUNO_CLEAN",
                "Aluno"
            )
            for f in filhos_exibir
        ]

        if u_info.get("tipo") in [
            "mestre",
            "diretoria"
        ]:

            aluno_sel_perfil = st.selectbox(
                "Selecione o Aluno:",
                lista_alunos
                if lista_alunos
                else ["Nenhum aluno"]
            )

            filhos_exibir = [
                f for f in filhos_exibir
                if f.get("NOME_ALUNO_CLEAN")
                == aluno_sel_perfil
            ]

        for f in filhos_exibir:

            nome_al = f.get(
                "NOME_ALUNO_CLEAN",
                "Aluno"
            )

            dados_aluno = (
                st.session_state
                .dados_app["alunos"]
                .get(
                    normalizar_nome(nome_al),
                    {}
                )
            )

            faixa_al = f.get(
                "FAIXA_CLEAN",
                "Branca"
            )

            resp_al = f.get(
                "RESPONSAVEL_CLEAN",
                "Responsável cadastrado"
            )

            saude_al = dados_aluno.get(
                "saude_detalhada",
                f.get(
                    "HISTORICO_SAUDE",
                    "Não informado"
                )
            )

            foto = foto_aluno(nome_al)

            if foto:

                c_foto, c_info = st.columns(
                    [1, 3]
                )

                with c_foto:

                    st.image(
                        foto,
                        width=130
                    )

                with c_info:

                    st.markdown(
                        f"""
                        <div class="perfil-nome">
                        {nome_al}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f"""
                        <span class="gold-badge">
                        Faixa {faixa_al}
                        </span>
                        """,
                        unsafe_allow_html=True
                    )

            else:

                st.markdown(
                    f"""
                    <div class="perfil-nome">
                    👤 {nome_al}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            st.markdown(
                f"""
                <div class="gold-card">

                <strong>Nome completo:</strong>
                {dados_aluno.get('nome', nome_al)}

                <br><br>

                <strong>Data de nascimento:</strong>
                {dados_aluno.get('data_nascimento', 'Não informado')}

                <br><br>

                <strong>Escola:</strong>
                {dados_aluno.get('escola', 'Não informado')}

                <br><br>

                <strong>Endereço:</strong>
                {dados_aluno.get('endereco', 'Não informado')}

                <br><br>

                <strong>Responsável:</strong>
                {resp_al}

                <br><br>

                <strong>Grau de parentesco:</strong>
                {dados_aluno.get('parentesco', 'Não informado')}

                <br><br>

                <strong>Contato:</strong>
                {dados_aluno.get('telefone', f.get('FONE_LIMPO', 'Não informado'))}

                <br><br>

                <strong>CIN:</strong>
                {dados_aluno.get('cin', 'Não informado')}

                <br><br>

                <strong>Faixa:</strong>
                {faixa_al}

                <br><br>

                <strong>Biometria Facial:</strong>
                {'🟢 Cadastrada' if dados_aluno.get('biometria_cadastrada') else '🔴 Pendente'}

                <br><br>

                <strong>Saúde informada:</strong>
                {saude_al}

                </div>
                """,
                unsafe_allow_html=True
            )

            # ------------------------------------------------------------------
            # CAMPEONATOS
            # ------------------------------------------------------------------

            st.markdown(
                "#### 🏆 Campeonatos"
            )

            campeonatos = (
                st.session_state
                .dados_app["campeonatos"]
                .get(
                    normalizar_nome(nome_al),
                    []
                )
            )

            if campeonatos:

                st.dataframe(
                    pd.DataFrame(campeonatos),
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "Nenhum campeonato registrado."
                )

            # ------------------------------------------------------------------
            # FREQUÊNCIA
            # ------------------------------------------------------------------

            st.markdown(
                "#### 📊 Frequência"
            )

            registros_freq = (
                st.session_state
                .dados_app["frequencias"]
                .get(
                    normalizar_nome(nome_al),
                    {}
                )
            )

            mes_nome = datetime.date.today().strftime(
                "%B/%Y"
            )

            freq_mes = registros_freq.get(
                str(datetime.date.today().month),
                {}
            )

            st.write(
                f"Frequência no mês atual: "
                f"**{freq_mes.get('quantidade', 0)} aulas**"
            )

            st.write(
                f"Percentual: "
                f"**{freq_mes.get('percentual', 0):.1f}%**"
            )

            desempenho = (
                st.session_state
                .dados_app["desempenho_mestre"]
                .get(
                    normalizar_nome(nome_al),
                    "Não informado"
                )
            )

            st.write(
                f"Desempenho informado pelo Mestre: "
                f"**{desempenho}**"
            )

            # ------------------------------------------------------------------
            # DESENVOLVIMENTO EM CASA / ESCOLA / SOCIAL
            # ------------------------------------------------------------------

            st.markdown(
                "#### 🏫 Desenvolvimento em Casa, Escola e Social"
            )

            avaliacao = (
                st.session_state
                .dados_app["avaliacoes_familia"]
                .get(
                    normalizar_nome(nome_al),
                    {}
                )
            )

            if u_info.get("tipo") in [
                "pai",
                "mestre"
            ]:

                if u_info.get("tipo") == "pai":

                    with st.expander(
                        "📝 Informar desenvolvimento do filho",
                        expanded=False
                    ):

                        casa = st.selectbox(
                            "Casa:",
                            DESENVOLVIMENTO_OPCOES,
                            key=f"casa_{normalizar_nome(nome_al)}"
                        )

                        escola = st.selectbox(
                            "Escola:",
                            DESENVOLVIMENTO_OPCOES,
                            key=f"escola_{normalizar_nome(nome_al)}"
                        )

                        social = st.selectbox(
                            "Socialmente:",
                            DESENVOLVIMENTO_OPCOES,
                            key=f"social_{normalizar_nome(nome_al)}"
                        )

                        obs = st.text_area(
                            "Observações:",
                            key=f"obs_{normalizar_nome(nome_al)}"
                        )

                        if st.button(
                            "Salvar Desenvolvimento",
                            key=f"salvar_dev_{normalizar_nome(nome_al)}"
                        ):

                            st.session_state.dados_app[
                                "avaliacoes_familia"
                            ][
                                normalizar_nome(nome_al)
                            ] = {
                                "casa": casa,
                                "escola": escola,
                                "social": social,
                                "observacao": obs,
                                "data": datetime.date.today().strftime(
                                    "%d/%m/%Y"
                                )
                            }

                            salvar_json()

                            st.success(
                                "Informações salvas com sucesso."
                            )

                avaliacao = (
                    st.session_state
                    .dados_app["avaliacoes_familia"]
                    .get(
                        normalizar_nome(nome_al),
                        {}
                    )
                )

                st.write(
                    f"Casa: **{avaliacao.get('casa', 'Não informado')}**"
                )

                st.write(
                    f"Escola: **{avaliacao.get('escola', 'Não informado')}**"
                )

                st.write(
                    f"Social: **{avaliacao.get('social', 'Não informado')}**"
                )


    # ==========================================================================
    # ABA MATRÍCULA
    # ==========================================================================

    with tab_mat:

        st.markdown(
            "### 📝 Matrícula"
        )

        st.info(
            "O responsável poderá cadastrar o aluno pelo celular. "
            "A biometria facial será realizada no final do cadastro."
        )

        sub_cadastro, sub_cadastrados = st.tabs(
            [
                "📝 Matricular",
                "👥 Cadastrados"
            ]
        )

        # ----------------------------------------------------------------------
        # MATRICULAR
        # ----------------------------------------------------------------------

        with sub_cadastro:

            st.markdown(
                "## FICHA DE INSCRIÇÃO E AUTORIZAÇÃO DE PARTICIPAÇÃO"
            )

            with st.form(
                "form_matricula"
            ):

                st.markdown(
                    "### 1. DADOS DO ALUNO (MENOR)"
                )

                nome = st.text_input(
                    "Nome Completo:"
                )

                data_nascimento = st.date_input(
                    "Data de Nascimento:",
                    value=datetime.date(
                        2015,
                        1,
                        1
                    ),
                    min_value=datetime.date(
                        1950,
                        1,
                        1
                    ),
                    max_value=datetime.date.today()
                )

                escola = st.text_input(
                    "Escola onde estuda:"
                )

                endereco = st.text_area(
                    "Endereço Residencial Completo:"
                )

                st.markdown(
                    "### 2. DADOS DO RESPONSÁVEL LEGAL"
                )

                responsavel = st.text_input(
                    "Nome do Responsável:"
                )

                parentesco = st.text_input(
                    "Grau de Parentesco:"
                )

                telefone = st.text_input(
                    "Contato / WhatsApp:"
                )

                cin = st.text_input(
                    "Carteira de Identidade Nacional (CIN):"
                )

                st.markdown(
                    "### 3. HISTÓRICO DE SAÚDE"
                )

                doencas = st.multiselect(
                    "Selecione as condições que se aplicam:",
                    DOENCAS_OPCOES
                )

                observacao_saude = st.text_area(
                    "Observações de Saúde / Especificações:"
                )

                st.markdown(
                    "### 4. AUTORIZAÇÃO DE PARTICIPAÇÃO"
                )

                autorizacao_participacao = st.checkbox(
                    "Li e concordo com o Termo de Autorização de Participação."
                )

                st.markdown(
                    "### 5. AUTORIZAÇÃO DE USO DE IMAGEM E VOZ"
                )

                autorizacao_imagem = st.checkbox(
                    "Li e concordo especificamente com o Termo de Uso de Imagem e Voz."
                )

                st.markdown(
                    "### 6. TRATAMENTO DE DADOS DE SAÚDE E BIOMETRIA"
                )

                autorizacao_dados_sensiveis = st.checkbox(
                    "Declaro estar ciente e autorizo, de forma específica, o tratamento dos dados de saúde e biometria facial para as finalidades informadas no aplicativo."
                )

                enviar = st.form_submit_button(
                    "SALVAR MATRÍCULA E CONTINUAR"
                )

            if enviar:

                erros = []

                if not nome.strip():
                    erros.append(
                        "Informe o nome completo do aluno."
                    )

                if not responsavel.strip():
                    erros.append(
                        "Informe o responsável legal."
                    )

                if not cin.strip():
                    erros.append(
                        "Informe a CIN."
                    )

                if not endereco.strip():
                    erros.append(
                        "Informe o endereço completo."
                    )

                if not telefone.strip():
                    erros.append(
                        "Informe o contato."
                    )

                if not autorizacao_participacao:
                    erros.append(
                        "É necessário aceitar a autorização de participação."
                    )

                if not autorizacao_imagem:
                    erros.append(
                        "É necessário registrar a decisão sobre uso de imagem."
                    )

                if not autorizacao_dados_sensiveis:
                    erros.append(
                        "É necessário registrar o consentimento específico para dados sensíveis."
                    )

                if erros:

                    for erro in erros:
                        st.error(erro)

                else:

                    chave = normalizar_nome(
                        nome
                    )

                    st.session_state.dados_app[
                        "alunos"
                    ][chave] = {

                        "nome": nome,

                        "data_nascimento":
                            data_nascimento.strftime(
                                "%d/%m/%Y"
                            ),

                        "escola": escola,

                        "endereco": endereco,

                        "responsavel": responsavel,

                        "parentesco": parentesco,

                        "telefone": telefone,

                        "cin": cin,

                        "doencas": doencas,

                        "saude_detalhada":
                            observacao_saude
                            if observacao_saude
                            else "Não informado",

                        "biometria_cadastrada": False,

                        "data_matricula":
                            datetime.date.today().strftime(
                                "%d/%m/%Y"
                            ),

                        "autorizacao_participacao": True,

                        "autorizacao_imagem":
                            True,

                        "autorizacao_dados_sensiveis":
                            True,

                        "documentacao_completa":
                            True,

                        "pendencia":
                            "Biometria Facial Pendente"
                    }

                    salvar_json()

                    st.success(
                        "Matrícula cadastrada. "
                        "Agora realize a Biometria Facial para finalizar."
                    )

                    st.rerun()


            # ------------------------------------------------------------------
            # BIOMETRIA APÓS CADASTRO
            # ------------------------------------------------------------------

            st.markdown("---")

            st.markdown(
                "### 📷 Biometria Facial"
            )

            nome_biometria = st.text_input(
                "Digite o nome completo do aluno já matriculado:"
            )

            if nome_biometria:

                chave_bio = normalizar_nome(
                    nome_biometria
                )

                aluno_bio = (
                    st.session_state
                    .dados_app["alunos"]
                    .get(chave_bio)
                )

                if aluno_bio:

                    st.success(
                        "Aluno encontrado no cadastro."
                    )

                    foto = st.camera_input(
                        "Tire a foto do rosto do aluno para cadastrar a biometria:"
                    )

                    if foto:

                        caminho = salvar_foto_aluno(
                            aluno_bio["nome"],
                            foto.getvalue()
                        )

                        aluno_bio["foto"] = caminho
                        aluno_bio[
                            "biometria_cadastrada"
                        ] = True

                        aluno_bio[
                            "data_biometria"
                        ] = datetime.date.today().strftime(
                            "%d/%m/%Y"
                        )

                        aluno_bio[
                            "pendencia"
                        ] = "Ok"

                        salvar_json()

                        st.success(
                            "🟢 Biometria facial cadastrada. "
                            "Cadastro finalizado."
                        )

                        st.rerun()

                else:

                    st.warning(
                        "Aluno não encontrado. "
                        "Faça primeiro a matrícula."
                    )


        # ----------------------------------------------------------------------
        # CADASTRADOS
        # ----------------------------------------------------------------------

        with sub_cadastrados:

            st.markdown(
                "### 👥 Alunos Cadastrados"
            )

            alunos_local = (
                st.session_state
                .dados_app["alunos"]
            )

            if not alunos_local:

                st.info(
                    "Nenhuma matrícula nova cadastrada no aplicativo."
                )

            else:

                for chave, aluno in alunos_local.items():

                    bio_ok = aluno.get(
                        "biometria_cadastrada",
                        False
                    )

                    pendencia = (
                        "Ok"
                        if bio_ok
                        else "Biometria Facial Pendente"
                    )

                    st.markdown(
                        f"""
                        <div class="gold-card">

                        <div class="perfil-nome">
                        {aluno.get('nome')}
                        </div>

                        <p>
                        Responsável:
                        {aluno.get('responsavel')}
                        </p>

                        <p>
                        CIN:
                        {aluno.get('cin')}
                        </p>

                        <p>
                        Status:
                        <span class="{'status-verde' if bio_ok else 'status-vermelho'}">
                        {'🟢 Cadastro completo' if bio_ok else '🔴 Pendente'}
                        </span>
                        </p>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )


    # ==========================================================================
    # ABA CRONOGRAMA
    # ==========================================================================

    with tab_crono:

        st.markdown(
            "### 📅 Cronograma"
        )

        st.caption(
            "Informações puxadas automaticamente da planilha base."
        )

        if not df_cronograma.empty:

            st.dataframe(
                df_cronograma,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "Nenhum evento encontrado."
            )


    # ==========================================================================
    # ABA FREQUÊNCIA
    # ==========================================================================

    with tab_freq:

        st.markdown(
            "### 📊 Frequência"
        )

        hoje = datetime.date.today()

        mes_atual = hoje.month
        ano_atual = hoje.year

        total_alunos = len(
            df_cadastro
        )

        total_frequentes = 0
        total_presencas = 0

        for chave, registro in (
            st.session_state
            .dados_app["frequencias"]
            .items()
        ):

            mes = registro.get(
                str(mes_atual),
                {}
            )

            quantidade = mes.get(
                "quantidade",
                0
            )

            if quantidade > 0:
                total_frequentes += 1
                total_presencas += quantidade

        percentual_geral = (
            total_frequentes / total_alunos * 100
            if total_alunos
            else 0
        )

        c1, c2 = st.columns(2)

        with c1:

            st.metric(
                "Total de Alunos Frequentes no Mês",
                total_frequentes
            )

        with c2:

            st.metric(
                "Percentual de Frequência do Mês",
                f"{percentual_geral:.1f}%"
            )

        st.markdown("---")

        lista = [
            x.get(
                "NOME_ALUNO_CLEAN"
            )
            for x in df_cadastro.to_dict(
                orient="records"
            )
        ]

        if lista:

            aluno_freq = st.selectbox(
                "Selecione o aluno:",
                lista
            )

            chave_freq = normalizar_nome(
                aluno_freq
            )

            registro = (
                st.session_state
                .dados_app["frequencias"]
                .get(
                    chave_freq,
                    {}
                )
            )

            mes_registro = registro.get(
                str(mes_atual),
                {}
            )

            quantidade = mes_registro.get(
                "quantidade",
                0
            )

            percentual = mes_registro.get(
                "percentual",
                0
            )

            st.markdown(
                f"""
                <div class="gold-card">

                <div class="perfil-nome">
                {aluno_freq}
                </div>

                <p>
                Frequência:
                <strong>{quantidade} aulas</strong>
                </p>

                <p>
                Percentual:
                <strong>{percentual:.1f}%</strong>
                </p>

                </div>
                """,
                unsafe_allow_html=True
            )

            if u_info.get("tipo") == "mestre":

                desempenho = st.selectbox(
                    "Desempenho do Mestre:",
                    DESEMPENHO_MESTRE
                )

                if st.button(
                    "Salvar Desempenho",
                    key="salvar_desempenho_freq"
                ):

                    st.session_state.dados_app[
                        "desempenho_mestre"
                    ][
                        chave_freq
                    ] = desempenho

                    salvar_json()

                    st.success(
                        "Desempenho atualizado."
                    )

        # ----------------------------------------------------------------------
        # IMPORTAÇÃO DA BASE
        # ----------------------------------------------------------------------

        if not df_frequencia_excel.empty:

            st.markdown(
                "### 📥 Informações da Planilha Base"
            )

            st.dataframe(
                df_frequencia_excel,
                use_container_width=True,
                hide_index=True
            )


    # ==========================================================================
    # ABA TATAME
    # ==========================================================================

    with tab_tatame:

        st.markdown(
            "### 🥋 Tatame"
        )

        st.info(
            "A chamada normal continua disponível. "
            "A chamada por foto utiliza reconhecimento facial "
            "quando o recurso estiver instalado."
        )

        st.markdown(
            "#### 📷 Chamada por Foto"
        )

        foto_turma = st.camera_input(
            "Tire uma foto de todos os alunos juntos:"
        )

        if foto_turma:

            arquivo_temp = DATA_DIR / "foto_turma.jpg"

            with open(
                arquivo_temp,
                "wb"
            ) as f:
                f.write(
                    foto_turma.getvalue()
                )

            if not FACE_RECOGNITION_DISPONIVEL:

                st.warning(
                    "O reconhecimento facial não está instalado. "
                    "A foto foi recebida, mas não será possível "
                    "identificar automaticamente os alunos."
                )

                st.code(
                    "pip install face-recognition"
                )

            else:

                try:

                    imagem_turma = face_recognition.load_image_file(
                        str(arquivo_temp)
                    )

                    locais = face_recognition.face_locations(
                        imagem_turma
                    )

                    codificacoes = (
                        face_recognition.face_encodings(
                            imagem_turma,
                            locais
                        )
                    )

                    conhecidos = []
                    nomes_conhecidos = []

                    for chave, aluno in (
                        st.session_state
                        .dados_app["alunos"]
                        .items()
                    ):

                        caminho = aluno.get(
                          "foto"
                        )

                        if (
                            caminho
                            and Path(caminho).exists()
                        ):

                            imagem_aluno = (
                                face_recognition
                                .load_image_file(caminho)
                            )

                            enc = (
                                face_recognition
                                .face_encodings(
                                    imagem_aluno
                                )
                            )

                            if enc:
                                conhecidos.append(
                                    enc[0]
                                )

                                nomes_conhecidos.append(
                                    aluno.get("nome")
                                )

                    encontrados = []

                    for cod in codificacoes:

                        resultados = (
                            face_recognition
                            .compare_faces(
                                conhecidos,
                                cod
                            )
                        )
                                
                                tolerance=0.48
                            )
                        )

                        if True in resultados:

                            indice = resultados.index(
                                True
                            )

                            encontrados.append(
                                nomes_conhecidos[indice]
                            )

                    data_presenca = datetime.date.today().strftime(
                        "%d/%m/%Y"
                    )

                    st.markdown(
                        "### Resultado da Chamada"
                    )

                    if encontrados:

                        for nome in encontrados:

                            st.success(
                                f"🟢 {nome} — Presença registrada."
                            )

                            chave = normalizar_nome(
                                nome
                            )

                            if chave not in (
                                st.session_state
                                .dados_app[
                                    "presencas_tatame"
                                ]
                            ):
                                st.session_state.dados_app[
                                    "presencas_tatame"
                                ][chave] = []

                            if data_presenca not in (
                                st.session_state
                                .dados_app[
                                    "presencas_tatame"
                                ][chave]
                            ):

                                st.session_state.dados_app[
                                    "presencas_tatame"
                                ][chave].append(
                                    data_presenca
                                )

                            salvar_json()

                    quantidade_faces = len(
                        codificacoes
                    )

                    desconhecidos = (
                        quantidade_faces
                        - len(encontrados)
                    )

                    if desconhecidos > 0:

                        st.error(
                            f"🔴 {desconhecidos} pessoa(s) "
                            "foram identificadas na foto, "
                            "mas não possuem cadastro facial."
                        )

                except Exception as e:

                    st.error(
                        "Erro durante o reconhecimento facial."
                    )

                    st.code(
                        str(e)
                    )

        st.markdown("---")

        st.markdown(
            "### 📋 Chamada Manual"
        )

        nomes_tatame = [
            x.get(
                "NOME_ALUNO_CLEAN"
            )
            for x in df_cadastro.to_dict(
                orient="records"
            )
        ]

        if nomes_tatame:

            aluno_tatame = st.selectbox(
                "Aluno presente:",
                nomes_tatame
            )

            if st.button(
                "MARCAR PRESENÇA",
                key="marcar_presenca_manual"
            ):

                chave = normalizar_nome(
                    aluno_tatame
                )

                data_presenca = datetime.date.today().strftime(
                    "%d/%m/%Y"
                )

                st.session_state.dados_app[
                    "presencas_tatame"
                ].setdefault(
                    chave,
                    []
                )

                if data_presenca not in (
                    st.session_state
                    .dados_app[
                        "presencas_tatame"
                    ][chave]
                ):

                    st.session_state.dados_app[
                        "presencas_tatame"
                    ][chave].append(
                        data_presenca
                    )

                    salvar_json()

                    st.success(
                        "Presença registrada."
                    )


    # ==========================================================================
    # ABA CAMPEÃO
    # ==========================================================================

    with tab_campeao:

        st.markdown(
            "### 🏆 Campeão"
        )

        st.markdown(
            "### 🥋 Campeonatos Participados"
        )

        if u_info.get("tipo") == "mestre":

            lista_campeao = [
                x.get(
                    "NOME_ALUNO_CLEAN"
                )
                for x in df_cadastro.to_dict(
                    orient="records"
                )
            ]

            if lista_campeao:

                aluno_campeao = st.selectbox(
                    "Aluno:",
                    lista_campeao,
                    key="aluno_campeao"
                )

                chave = normalizar_nome(
                    aluno_campeao
                )

                campeonato_nome = st.text_input(
                    "Nome do Campeonato:"
                )

                campeonato_data = st.date_input(
                    "Data do Campeonato:",
                    key="data_campeonato"
                )

                classificacao = st.text_input(
                    "Classificação do aluno:"
                )

                if st.button(
                    "Adicionar Campeonato"
                ):

                    st.session_state.dados_app[
                        "campeonatos"
                    ].setdefault(
                        chave,
                        []
                    )

                    st.session_state.dados_app[
                        "campeonatos"
                    ][chave].append(
                        {
                            "Campeonato":
                                campeonato_nome,

                            "Data":
                                campeonato_data.strftime(
                                    "%d/%m/%Y"
                                ),

                            "Classificação":
                                classificacao
                        }
                    )

                    salvar_json()

                    st.success(
                        "Campeonato registrado."
                    )

        # ----------------------------------------------------------------------
        # CAMPEONATOS PÚBLICOS
        # ----------------------------------------------------------------------

        st.markdown(
            "### 🏅 Resultados registrados"
        )

        todos_campeonatos = []

        for chave, lista in (
            st.session_state
            .dados_app["campeonatos"]
            .items()
        ):

            for item in lista:

                todos_campeonatos.append(
                    {
                        "Aluno": chave,
                        **item
                    }
                )

        if todos_campeonatos:

            st.dataframe(
                pd.DataFrame(
                    todos_campeonatos
                ),
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "Nenhum campeonato registrado."
            )


    # ==========================================================================
    # ABA FÉ
    # ==========================================================================

    with tab_fe:

        st.markdown(
            "### 🙏 Fé"
        )

        st.markdown(
            "## Impacto Espiritual"
        )

        total_celulas = len(
            st.session_state
            .dados_app
            .get(
                "celulas",
                []
            )
        )

        total_batizados = len(
            st.session_state
            .dados_app
            .get(
                "batizados",
                []
            )
        )

        # SOMENTE NÚMEROS
        c1, c2 = st.columns(2)

        with c1:

            st.metric(
                "Pessoas que entraram em células",
                total_celulas
            )

        with c2:

            st.metric(
                "Pessoas batizadas",
                total_batizados
            )

        # GRÁFICO IMPACTO ESPIRITUAL

        grafico = pd.DataFrame(
            {
                "Indicador": [
                    "Entraram em células",
                    "Batizados"
                ],
                "Quantidade": [
                    total_celulas,
                    total_batizados
                ]
            }
        )

        st.bar_chart(
            grafico.set_index(
                "Indicador"
            )
        )

        st.caption(
            "Impacto Espiritual — números consolidados, "
            "sem exposição dos nomes."
        )

        # ----------------------------------------------------------------------
        # PEDIDOS DE ORAÇÃO
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### 🙏 Pedidos de Oração"
        )

        pedido = st.text_area(
            "Digite seu pedido de oração:"
        )

        if st.button(
            "Enviar Pedido de Oração"
        ):

            if pedido.strip():

                st.session_state.dados_app[
                    "pedidos_oracao"
                ].append(
                    {
                        "texto": pedido,
                        "data":
                            datetime.date.today().strftime(
                                "%d/%m/%Y"
                            )
                    }
                )

                salvar_json()

                st.success(
                    "Pedido de oração enviado."
                )

        # ----------------------------------------------------------------------
        # CULTOS
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### ⛪ Cultos da Igreja"
        )

        st.markdown(
            """
            <div class="gold-card">
            <div class="perfil-nome">
            Domingo — 18:30h
            </div>
            <p>
            Culto da Igreja
            </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        # ----------------------------------------------------------------------
        # CONTROLE DO MESTRE
        # ----------------------------------------------------------------------

        if u_info.get("tipo") == "mestre":

            st.markdown("---")

            st.markdown(
                "### 🔐 Atualizar Impacto Espiritual"
            )

            qtd_celulas = st.number_input(
                "Quantidade de pessoas que entraram em células:",
                min_value=0,
                value=total_celulas,
                step=1
            )

            qtd_batizados = st.number_input(
                "Quantidade de pessoas batizadas:",
                min_value=0,
                value=total_batizados,
                step=1
            )

            if st.button(
                "Atualizar Números"
            ):

                # Mantém somente registros anônimos
                st.session_state.dados_app[
                    "celulas"
                ] = [
                    {}
                    for _ in range(qtd_celulas)
                ]

                st.session_state.dados_app[
                    "batizados"
                ] = [
                    {}
                    for _ in range(qtd_batizados)
                ]

                salvar_json()

                st.success(
                    "Impacto Espiritual atualizado."
                )

        # ----------------------------------------------------------------------
        # ACESSO A PEDIDOS
        # ----------------------------------------------------------------------

        if u_info.get("tipo") in [
            "mestre",
            "diretoria"
        ]:

            st.markdown(
                "### 🔒 Pedidos de Oração"
            )

            pedidos = (
                st.session_state
                .dados_app
                .get(
                    "pedidos_oracao",
                    []
                )
            )

            if pedidos:

                for item in pedidos:

                    st.markdown(
                        f"""
                        <div class="gold-card">
                        🙏 {item.get('texto')}
                        <br>
                        <small>
                        {item.get('data')}
                        </small>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

            else:

                st.info(
                    "Nenhum pedido de oração registrado."
                )


    # ==========================================================================
    # ABA SECRETARIA
    # ==========================================================================

    with tab_sec:

        st.markdown(
            "### ⚙️ Secretaria"
        )

        if u_info.get("tipo") not in [
            "mestre",
            "diretoria"
        ]:

            st.info(
                "O responsável possui acesso somente "
                "aos gráficos públicos desta área."
            )

            st.markdown(
                "### 📊 Gráficos"
            )

            st.metric(
                "Alunos cadastrados",
                len(df_cadastro)
            )

            st.stop()

        # ----------------------------------------------------------------------
        # GERAR SENHAS
        # ----------------------------------------------------------------------

        st.markdown(
            "### 🔐 Gerar Senha"
        )

        col_senha1, col_senha2 = st.columns(2)

        with col_senha1:

            nome_resp_senha = st.text_input(
                "Responsável:"
            )

            login_resp_senha = st.text_input(
                "Login do Responsável:"
            )

            nova_senha_resp = st.text_input(
                "Nova senha:",
                type="password"
            )

            if st.button(
                "Gerar Senha do Responsável"
            ):

                if (
                    login_resp_senha
                    and nova_senha_resp
                ):

                    st.session_state.dados_app[
                        "senhas_pais"
                    ][
                        limpar_numero(
                            login_resp_senha
                        )
                        or login_resp_senha.lower()
                    ] = nova_senha_resp

                    salvar_json()

                    st.success(
                        "Senha do responsável criada."
                    )

        with col_senha2:

            login_diretoria = st.text_input(
                "Login da Diretoria:"
            )

            senha_diretoria = st.text_input(
                "Nova senha da Diretoria:",
                type="password"
            )

            if st.button(
                "Gerar Senha da Diretoria"
            ):

                if (
                    login_diretoria
                    and senha_diretoria
                ):

                    st.session_state.dados_app[
                        "senhas_diretoria"
                    ][
                        login_diretoria.lower()
                    ] = senha_diretoria

                    salvar_json()

                    st.success(
                        "Senha da diretoria criada."
                    )

        # ----------------------------------------------------------------------
        # RELATÓRIOS
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### 📄 Gerar Relatório"
        )

        tipo_relatorio = st.selectbox(
            "Selecione o relatório:",
            [
                "Cadastro dos Alunos",
                "Frequência",
                "Desempenho",
                "Campeonatos",
                "Desenvolvimento Familiar",
                "Saúde",
                "Biometria Facial",
                "Pedidos de Oração",
                "Impacto Espiritual",
                "Relatório Geral"
            ]
        )

        if st.button(
            "GERAR RELATÓRIO"
        ):

            dados_relatorio = []

            if tipo_relatorio == "Cadastro dos Alunos":

                for aluno in (
                    st.session_state
                    .dados_app["alunos"]
                    .values()
                ):

                    dados_relatorio.append(
                        aluno
                    )

            elif tipo_relatorio == "Saúde":

                for aluno in (
                    st.session_state
                    .dados_app["alunos"]
                    .values()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno":
                                aluno.get("nome"),

                            "Doenças":
                                ", ".join(
                                    aluno.get(
                                        "doencas",
                                        []
                                    )
                                ),

                            "Observação":
                                aluno.get(
                                    "saude_detalhada",
                                    ""
                                )
                        }
                    )

            elif tipo_relatorio == "Biometria Facial":

                for aluno in (
                    st.session_state
                    .dados_app["alunos"]
                    .values()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno":
                                aluno.get("nome"),

                            "Biometria":
                                "Cadastrada"
                                if aluno.get(
                                    "biometria_cadastrada"
                                )
                                else "Pendente"
                        }
                    )

            elif tipo_relatorio == "Desempenho":

                for chave, valor in (
                    st.session_state
                    .dados_app["desempenho_mestre"]
                    .items()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno": chave,
                            "Desempenho": valor
                        }
                    )

            elif tipo_relatorio == "Campeonatos":

                for chave, lista in (
                    st.session_state
                    .dados_app["campeonatos"]
                    .items()
                ):

                    for item in lista:

                        dados_relatorio.append(
                            {
                                "Aluno": chave,
                                **item
                            }
                        )

            elif tipo_relatorio == "Desenvolvimento Familiar":

                for chave, valor in (
                    st.session_state
                    .dados_app["avaliacoes_familia"]
                    .items()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno": chave,
                            **valor
                        }
                    )

            elif tipo_relatorio == "Impacto Espiritual":

                dados_relatorio = [
                    {
                        "Indicador":
                            "Pessoas em células",
                        "Quantidade":
                            len(
                                st.session_state
                                .dados_app
                                .get(
                                    "celulas",
                                    []
                                )
                            )
                    },
                    {
                        "Indicador":
                            "Pessoas batizadas",
                        "Quantidade":
                            len(
                                st.session_state
                                .dados_app
                                .get(
                                    "batizados",
                                    []
                                )
                            )
                    }
                ]

            elif tipo_relatorio == "Pedidos de Oração":

                dados_relatorio = (
                    st.session_state
                    .dados_app
                    .get(
                        "pedidos_oracao",
                        []
                    )
                )

            elif tipo_relatorio == "Frequência":

                for chave, registro in (
                    st.session_state
                    .dados_app["frequencias"]
                    .items()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno": chave,
                            "Registro": registro
                        }
                    )

            elif tipo_relatorio == "Relatório Geral":

                for aluno in (
                    st.session_state
                    .dados_app["alunos"]
                    .values()
                ):

                    dados_relatorio.append(
                        {
                            "Aluno":
                                aluno.get("nome"),

                            "Responsável":
                                aluno.get(
                                    "responsavel"
                                ),

                            "Telefone":
                                aluno.get(
                                    "telefone"
                                ),

                            "Biometria":
                                aluno.get(
                                    "biometria_cadastrada"
                                ),

                            "Pendência":
                                aluno.get(
                                    "pendencia"
                                )
                        }
                    )

            df_relatorio = pd.DataFrame(
                dados_relatorio
            )

            if df_relatorio.empty:

                st.warning(
                    "Não existem dados para este relatório."
                )

            else:

                st.dataframe(
                    df_relatorio,
                    use_container_width=True,
                    hide_index=True
                )

                csv = df_relatorio.to_csv(
                    index=False
                ).encode(
                    "utf-8-sig"
                )

                st.download_button(
                    "⬇️ Baixar Relatório",
                    data=csv,
                    file_name=(
                        "relatorio_projeto_sementes.csv"
                    ),
                    mime="text/csv"
                )

        # ----------------------------------------------------------------------
        # GRÁFICOS PARA DIRETORIA
        # ----------------------------------------------------------------------

        st.markdown("---")

        st.markdown(
            "### 📊 Gráficos Administrativos"
        )

        st.metric(
            "Total de alunos",
            len(df_cadastro)
        )

        total_bio = sum(
            1
            for aluno in (
                st.session_state
                .dados_app["alunos"]
                .values()
            )
            if aluno.get(
                "biometria_cadastrada"
            )
        )

        st.metric(
            "Biometrias cadastradas",
            total_bio
        )

        st.metric(
            "Pedidos de oração",
            len(
                st.session_state
                .dados_app
                .get(
                    "pedidos_oracao",
                    []
                )
            )
        )

        st.markdown("---")

        # ----------------------------------------------------------------------
        # REGRA DE PERMISSÃO
        # ----------------------------------------------------------------------

        st.markdown(
            """
            <div class="gold-card">

            <div class="perfil-nome">
            🔐 Controle de Acesso
            </div>

            <p>
            🥋 <strong>Mestre:</strong>
            acesso total e permissão para inserir e alterar informações.
            </p>

            <p>
            👑 <strong>Diretoria:</strong>
            acesso administrativo às informações autorizadas e gráficos.
            </p>

            <p>
            👨‍👩‍👧 <strong>Responsável:</strong>
            acesso somente aos seus próprios filhos,
            preenchimento de cadastro,
            desenvolvimento familiar e pedidos de oração.
            </p>

            <p>
            🔒 Um responsável não consegue visualizar
            o perfil de outro aluno.
            </p>

            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 11. ENCERRAMENTO / LOGOUT
# ==============================================================================

if st.session_state.logged_in:

    st.markdown("---")

    if st.button(
        "SAIR DO APLICATIVO",
        key="logout_final"
    ):

        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.session_state.sino_visto = False

        st.rerun()
        

                      
