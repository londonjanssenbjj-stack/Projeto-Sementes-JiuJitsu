import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import base64
import json
import re
from pathlib import Path

# ============================================================
# IMPORTAÇÃO OPCIONAL DO STREAMLIT CANVAS E FACE RECOGNITION
# ============================================================
try:
    from streamlit_canvas import st_canvas
    CANVAS_DISPONIVEL = True
except ImportError:
    st_canvas = None
    CANVAS_DISPONIVEL = False

try:
    import face_recognition
    FACE_RECOGNITION_DISPONIVEL = True
except ImportError:
    face_recognition = None
    FACE_RECOGNITION_DISPONIVEL = False

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Projeto Sementes - Jiu-Jitsu",
    page_icon="🥋",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"
DATA_DIR = Path("dados_aplicativo")
FOTOS_DIR = DATA_DIR / "fotos_alunos"

DATA_DIR.mkdir(exist_ok=True)
FOTOS_DIR.mkdir(exist_ok=True)

JSON_FILE = DATA_DIR / "dados_aplicativo.json"

LOGIN_MESTRE = "Mikhael"
SENHA_MESTRE = "12381314*Lj"

# ==============================================================================
# ESTILOS CSS
# ==============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: Inter, sans-serif; font-size: 17px; }
.stApp { background:#0A0F18; color:#FFFFFF; }
#MainMenu, header, footer, [data-testid="stSidebar"] { display:none !important; }
.gold-card,.welcome-card,.banner-card {
background:#111827; border:2px solid #D97706; border-radius:16px;
padding:20px; margin:12px 0; box-shadow:0 5px 18px rgba(217,119,6,.18);
}
.banner-card { border-color:#F59E0B; text-align:center; }
.banner-title { color:#F59E0B !important; font-size:1.25rem; font-weight:800; }
.banner-text { color:#FFFFFF !important; font-size:1rem; line-height:1.65; }
.perfil-nome { color:#FFFFFF !important; font-size:1.25rem !important; font-weight:900 !important; }
.status-verde { color:#22C55E !important; font-weight:900 !important; }
.status-vermelho { color:#EF4444 !important; font-weight:900 !important; }
.status-amarelo { color:#F59E0B !important; font-weight:900 !important; }
label, p, span, div { color:#F1F5F9 !important; font-weight:600; }
.stTextInput input,.stTextArea textarea,.stDateInput input,.stNumberInput input,
.stSelectbox div[data-baseweb="select"], .stMultiSelect div[data-baseweb="select"] {
background:#1E293B !important; color:#FFFFFF !important;
font-size:1.08rem !important; border:2px solid #475569 !important;
}
.stSelectbox div[data-baseweb="select"] *,.stMultiSelect div[data-baseweb="select"] * {
color:#FFFFFF !important;
}
.stButton>button, div[data-testid="stFormSubmitButton"]>button {
background:#D97706 !important; color:#FFFFFF !important;
border:2px solid #F59E0B !important; border-radius:11px !important;
font-weight:800 !important; font-size:1.02rem !important; min-height:48px;
}
.stButton>button:hover { background:#B45309 !important; }
.stTabs [data-baseweb="tab-list"] {
gap:4px; background:#111827; padding:6px; border-radius:14px;
border:1px solid #334155;
}
.stTabs [data-baseweb="tab"] {
border-radius:8px; color:#FFFFFF !important; font-weight:800;
padding:9px 10px; font-size:.82rem;
}
.stTabs [aria-selected="true"] { background:#D97706 !important; color:#FFFFFF !important; }
div[data-testid="stMetricValue"] { color:#FFFFFF !important; }
div[data-testid="stMetricLabel"] { color:#F59E0B !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. ESTADO LOCAL
# ==============================================================================
def carregar_json():
    base = {
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
        "senhas_diretoria": {"diretoria": "dir2026"},
        "presencas_nao_cadastrado": [],
        "evento_visualizado": ""
    }
    if JSON_FILE.exists():
        try:
            with open(JSON_FILE, "r", encoding="utf-8") as f:
                existente = json.load(f)
                for k, v in base.items():
                    existente.setdefault(k, v)
                return existente
        except Exception:
            pass
    return base

def salvar_json():
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(st.session_state.dados_app, f, ensure_ascii=False, indent=2)

if "dados_app" not in st.session_state:
    st.session_state.dados_app = carregar_json()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "user_info" not in st.session_state:
    st.session_state.user_info = None

if "notificacao_aberta" not in st.session_state:
    st.session_state.notificacao_aberta = False

if "camera_habilitada" not in st.session_state:
    st.session_state.camera_habilitada = False

# ==============================================================================
# 3. FUNÇÕES UTILITÁRIAS
# ==============================================================================
def limpar_numero(valor):
    return "".join(filter(str.isdigit, str(valor)))

def normalizar_nome(nome):
    return re.sub(r"\s+", " ", str(nome).strip()).lower()

def texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def encontrar_coluna(df, termos):
    for c in df.columns:
        lc = str(c).strip().lower()
        if any(t.lower() in lc for t in termos):
            return c
    return None

def gerar_link_whatsapp(numero, mensagem):
    numero = limpar_numero(numero)
    if not numero:
        return None
    if len(numero) in (10, 11) and not numero.startswith("55"):
        numero = "55" + numero
    return f"https://wa.me/{numero}?text={urllib.parse.quote(mensagem)}"

def valor_por_coluna(row, termos, padrao=""):
    for c in row.index:
        lc = str(c).lower()
        if any(t.lower() in lc for t in termos):
            v = texto(row[c])
            if v:
                return v
    return padrao

def salvar_foto(nome, imagem_bytes):
    if not imagem_bytes:
        return None
    arq = FOTOS_DIR / (re.sub(r"[^a-zA-Z0-9_-]", "_", normalizar_nome(nome)) + ".jpg")
    arq.write_bytes(imagem_bytes)
    return str(arq)

def foto_aluno(nome):
    a = st.session_state.dados_app["alunos"].get(normalizar_nome(nome), {})
    caminho = a.get("foto")
    if caminho and Path(caminho).exists():
        return caminho
    padrao = FOTOS_DIR / (re.sub(r"[^a-zA-Z0-9_-]", "_", normalizar_nome(nome)) + ".jpg")
    return str(padrao) if padrao.exists() else None

def data_evento(row):
    c = encontrar_coluna(pd.DataFrame([row]), ["data", "dia"])
    if not c:
        return None
    v = row[c]
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    if pd.isna(d):
        return None
    return d.date()

def banner_evento(row, destaque=False):
    evento = valor_por_coluna(row, ["evento", "atividade", "programa"], "Evento")
    data = valor_por_coluna(row, ["data", "dia"], "")
    hora = valor_por_coluna(row, ["hora"], "")
    local = valor_por_coluna(row, ["local"], "")
    responsavel = valor_por_coluna(row, ["palestrante", "instrutor", "equipe", "igreja"], "")
    st.markdown(f"""
    <div class="banner-card">
    <div class="banner-title">📅 {evento}</div>
    <div class="banner-text">
    <strong>Data:</strong> {data or "Não informada"}<br>
    {f"<strong>Horário:</strong> {hora}<br>" if hora else ""}
    {f"<strong>Local:</strong> {local}<br>" if local else ""}
    {f"<strong>Responsável:</strong> {responsavel}" if responsavel else ""}
    </div>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# 4. LEITURA DAS ABAS DA PLANILHA
# ==============================================================================
@st.cache_data(ttl=30)
def carregar_aba(nome_aba, header=0):
    try:
        xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
        if nome_aba not in xls.sheet_names:
            return pd.DataFrame()
        df = pd.read_excel(xls, sheet_name=nome_aba, header=header)
        df = df.dropna(how="all").dropna(how="all", axis=1)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=30)
def nomes_abas():
    try:
        return pd.ExcelFile(EXCEL_FILE, engine="openpyxl").sheet_names
    except Exception:
        return []

@st.cache_data(ttl=30)
def carregar_cadastro():
    for header in (0, 1, 2):
        df = carregar_aba("Ficha de Cadastro", header)
        if not df.empty:
            col = encontrar_coluna(df, ["aluno", "nome"])
            if col:
                df["NOME_ALUNO_CLEAN"] = df[col].map(texto)
                df = df[df["NOME_ALUNO_CLEAN"].ne("")]
                return df
    return pd.DataFrame()

df_cadastro = carregar_cadastro()
ABAS = nomes_abas()

# ==============================================================================
# 5. EVENTOS / CRONOGRAMA
# ==============================================================================
def carregar_cronograma():
    candidatos = [x for x in ABAS if "cronograma" in x.lower()]
    if candidatos:
        return carregar_aba(candidatos[0], 0)
    return pd.DataFrame()

df_cronograma = carregar_cronograma()

def eventos_de_hoje():
    hoje = datetime.date.today()
    achados = []
    if df_cronograma.empty:
        return achados
    for _, row in df_cronograma.iterrows():
        d = data_evento(row)
        if d == hoje:
            achados.append(row.to_dict())
    return achados

# ==============================================================================
# 6. NORMALIZAÇÃO DOS ALUNOS DO EXCEL
# ==============================================================================
def alunos_cadastrados():
    if df_cadastro.empty:
        return []
    resultado = []
    for _, row in df_cadastro.iterrows():
        nome = texto(row.get("NOME_ALUNO_CLEAN", ""))
        if not nome:
            continue
        resultado.append({
            "nome": nome,
            "data_inicio": valor_por_coluna(row, ["data início", "data inicio", "início", "inicio"]),
            "faixa": valor_por_coluna(row, ["faixa"], "Branca"),
            "data_nascimento": valor_por_coluna(row, ["data de nascimento", "data nascimento", "nascimento"]),
            "responsavel": valor_por_coluna(row, ["responsável", "responsavel"]),
            "contato": valor_por_coluna(row, ["contato", "telefone", "whatsapp", "fone"]),
            "historico_saude": valor_por_coluna(row, ["histórico de saúde", "historico de saude", "saúde", "saude"]),
            "endereco": valor_por_coluna(row, ["endereço", "endereco"]),
            "escola": valor_por_coluna(row, ["escola"]),
            "pendencia": valor_por_coluna(row, ["pendência documento", "pendencia documento", "pendência", "pendencia"], "Ok"),
            "_row": row
        })
    return resultado

def aluno_excel(nome):
    alvo = normalizar_nome(nome)
    for a in alunos_cadastrados():
        if normalizar_nome(a["nome"]) == alvo:
            return a
    return None

def biometria_ok(nome):
    a = st.session_state.dados_app["alunos"].get(normalizar_nome(nome), {})
    return bool(a.get("biometria_cadastrada")) or bool(
        st.session_state.dados_app.get("biometrias", {}).get(normalizar_nome(nome))
    )

def status_aluno(a):
    p = normalizar_nome(a["pendencia"])
    if "falta endereço" in p or "falta endereco" in p or "falta documento" in p:
        return "🔴", "vermelho"
    if p == "ok":
        return ("🟢", "verde") if biometria_ok(a["nome"]) else ("🟡", "amarelo")
    return "🔴", "vermelho"

# ==============================================================================
# 7. ANIVERSARIANTES
# ==============================================================================
def parse_data_nascimento(valor):
    d = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.date()

def aniversariantes_mes():
    hoje = datetime.date.today()
    return [a for a in alunos_cadastrados()
            if (d := parse_data_nascimento(a["data_nascimento"])) and d.month == hoje.month]

def aniversariantes_dia():
    hoje = datetime.date.today()
    return [a for a in aniversariantes_mes()
            if (d := parse_data_nascimento(a["data_nascimento"])) and d.day == hoje.day]

def mensagem_aniversario(nome):
    return (
        f"Paz do Senhor, {nome}! 🎉🌱\n\n"
        "Toda a família Projeto Sementes deseja a você um feliz aniversário! "
        "Que Deus abençoe sua vida, sua família, seus estudos e seus sonhos. "
        "Que Jesus continue guiando sua caminhada. 🙏🥋"
    )

# ==============================================================================
# 8. FREQUÊNCIA
# ==============================================================================
def carregar_abas_presenca():
    frames = []
    for aba in ("Lista de Presença Agosto", "Lista de Presença Setembro"):
        df = carregar_aba(aba, 0)
        if not df.empty:
            df["_ABA_ORIGEM"] = aba
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

df_presencas_excel = carregar_abas_presenca()

def presencas_locais(nome):
    return st.session_state.dados_app["presencas_tatame"].get(normalizar_nome(nome), [])

def quantidade_presencas(nome):
    locais = presencas_locais(nome)
    return len(locais)

# ==============================================================================
# 9. CAMPEONATOS
# ==============================================================================
def carregar_campeonatos_excel():
    abas = [x for x in ABAS if x.lower() == "campeão" or x.lower() == "campeao" or "campeão" in x.lower() or "campeao" in x.lower()]
    frames = []
    for aba in abas:
        df = carregar_aba(aba, 0)
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

df_campeao_excel = carregar_campeonatos_excel()

def campeonatos_do_aluno(nome):
    chave = normalizar_nome(nome)
    local = st.session_state.dados_app["campeonatos"].get(chave, [])
    if df_campeao_excel.empty:
        return pd.DataFrame(local)
    nome_col = encontrar_coluna(df_campeao_excel, ["aluno", "nome"])
    if not nome_col:
        return pd.DataFrame(local)
    df = df_campeao_excel[df_campeao_excel[nome_col].map(normalizar_nome).eq(chave)].copy()
    if not local:
        return df
    extra = pd.DataFrame(local)
    return pd.concat([df, extra], ignore_index=True)

# ==============================================================================
# 10. FÉ
# ==============================================================================
def carregar_fe():
    abas = [x for x in ABAS if x.lower() == "fé" or x.lower() == "fe" or "fé" in x.lower() or "fe" in x.lower()]
    frames = []
    for aba in abas:
        df = carregar_aba(aba, 0)
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

df_fe = carregar_fe()

def celulas_fe():
    if df_fe.empty:
        return []
    setor_col = encontrar_coluna(df_fe, ["setor"])
    nome_col = encontrar_coluna(df_fe, ["nomes", "nome"])
    contato_col = encontrar_coluna(df_fe, ["contato", "telefone", "whatsapp"])
    if not setor_col or not nome_col:
        return []
    saida = []
    for _, r in df_fe.iterrows():
        setor = texto(r[setor_col])
        nome = texto(r[nome_col])
        contato = texto(r[contato_col]) if contato_col else ""
        if setor or nome:
            saida.append({"setor": setor, "responsavel": nome, "contato": contato})
    return saida[:7]

def contato_pastor_joel():
    if df_fe.empty:
        return ""
    nome_col = encontrar_coluna(df_fe, ["nomes", "nome"])
    contato_col = encontrar_coluna(df_fe, ["contato", "telefone", "whatsapp"])
    if not nome_col or not contato_col:
        return ""
    for _, r in df_fe.iterrows():
        nome = texto(r[nome_col]).lower()
        if "joel" in nome:
            return texto(r[contato_col])
    return ""

# ==============================================================================
# 11. LOGIN
# ==============================================================================
def autenticar(login, senha):
    login = texto(login)
    senha = texto(senha)
    if login.lower() == LOGIN_MESTRE.lower() and senha == SENHA_MESTRE:
        return {"nome": "Mestre Mikhael", "tipo": "mestre"}
    dirs = st.session_state.dados_app.get("senhas_diretoria", {})
    if login.lower() in dirs and senha == dirs[login.lower()]:
        return {"nome": "Diretoria IEQ Guaicurus", "tipo": "diretoria"}
    numero = limpar_numero(login)
    if not df_cadastro.empty:
        fone_col = encontrar_coluna(df_cadastro, ["contato", "telefone", "whatsapp", "fone"])
        resp_col = encontrar_coluna(df_cadastro, ["responsável", "responsavel"])
        if fone_col and numero:
            mask = df_cadastro[fone_col].map(limpar_numero).eq(numero)
        elif resp_col:
            mask = df_cadastro[resp_col].map(normalizar_nome).eq(normalizar_nome(login))
        else:
            mask = pd.Series(False, index=df_cadastro.index)
        match = df_cadastro[mask]
        if not match.empty:
            fone = limpar_numero(match.iloc[0][fone_col]) if fone_col else numero
            senhas = st.session_state.dados_app.get("senhas_pais", {})
            senha_correta = senhas.get(fone, "123456")
            if senha == senha_correta:
                return {"nome": texto(match.iloc[0][resp_col]) if resp_col else "Responsável",
                        "tipo": "pai", "fone": fone}
    return None

# ==============================================================================
# 12. LOGIN SCREEN — SEM CÂMERA
# ==============================================================================
if not st.session_state.logged_in:
    if Path("logo_projeto_sementes.png").exists():
        st.image("logo_projeto_sementes.png", use_container_width=True)
    st.markdown("<h2 style='text-align:center;color:#F59E0B !important;'>PROJETO SEMENTES</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;'>Iniciativa Voluntária de Jiu-Jitsu e Apoio à Família</p>", unsafe_allow_html=True)
    with st.form("login"):
        login = st.text_input("Login do Responsável", placeholder="Informe seu login")
        senha = st.text_input("Senha de Acesso", type="password", placeholder="Digite sua senha")
        entrar = st.form_submit_button("ENTRAR NO APLICATIVO")
        if entrar:
            u = autenticar(login, senha)
            if u:
                st.session_state.logged_in = True
                st.session_state.user_info = u
                st.rerun()
            else:
                st.error("Credenciais inválidas ou acesso não autorizado.")
    st.stop()

# ==============================================================================
# 13. CABEÇALHO + SINO ÚNICO
# ==============================================================================
u = st.session_state.user_info
hoje = datetime.date.today()
eventos_hoje = eventos_de_hoje()
tem_evento = bool(eventos_hoje)

h1, h2 = st.columns([5, 1])
with h1:
    st.markdown(f"<div class='perfil-nome'>A Paz seja convosco, {u['nome']}</div>", unsafe_allow_html=True)
with h2:
    sino = "🔴🔔" if tem_evento else "🔔"
    if st.button(sino, key="sino_unico"):
        st.session_state.notificacao_aberta = True

if st.session_state.notificacao_aberta:
    st.markdown("### 🔔 Evento de Hoje")
    if eventos_hoje:
        for ev in eventos_hoje:
            banner_evento(ev, True)
        if st.button("Visualizar e fechar notificação", key="fechar_sino"):
            st.session_state.notificacao_aberta = False
            st.session_state.dados_app["evento_visualizado"] = hoje.isoformat()
            salvar_json()
            st.rerun()
    else:
        st.info("Não há evento no Cronograma para hoje.")
        if st.button("Fechar", key="fechar_sino_sem_evento"):
            st.session_state.notificacao_aberta = False
            st.rerun()

# ==============================================================================
# 14. ABAS
# ==============================================================================
tab_inicio, tab_mat, tab_alunos, tab_crono, tab_freq, tab_tatame, tab_campeao, tab_fe, tab_sec = st.tabs([
    "🏠 Início", "📝 Matrícula", "👥 Alunos Cadastrados", "📅 Cronograma",
    "📊 Frequência", "🥋 Tatame", "🏆 Campeão", "🙏 Fé", "⚙️ Secretaria"
])

# ==============================================================================
# INÍCIO
# ==============================================================================
with tab_inicio:
    st.markdown("""
    <div class="welcome-card">
    <div class="banner-title">🌱 BEM-VINDOS AO PROJETO SEMENTES</div>
    <div class="banner-text" style="text-align:justify;">
    O <strong>Projeto Sementes</strong> nasceu no coração de Deus e, por Sua misericórdia,
    foi compartilhado aos corações do Pastor Joel Amorim, da IEQ Guaicurus, e do
    Instrutor Faixa-Preta London Carvalho, da Academia Iron Jiu-Jitsu.<br><br>
    Foram dias de oração, planejamento e dedicação para acolher crianças, adolescentes
    e jovens das comunidades próximas à igreja. Com o apoio dos instrutores e das famílias,
    iniciamos nossas aulas, unindo o Jiu-Jitsu a palestras socioeducativas e ações de
    evangelismo.<br><br>
    Trabalhamos valores morais e éticos fundamentados em Cristo Jesus, buscando contribuir
    para a formação integral dos alunos e alcançar também seus pais e famílias.
    <br><br><strong>Sejam todos bem-vindos a esta família! 🌱🥋</strong>
    </div></div>
    """, unsafe_allow_html=True)

    st.markdown("### 🎂 Aniversariante do Dia")
    dia = aniversariantes_dia()
    if not dia:
        st.info("Não há aniversariante cadastrado hoje.")
    else:
        for a in dia:
            foto = foto_aluno(a["nome"])
            c1, c2 = st.columns([1, 3])
            with c1:
                if foto: st.image(foto, width=110)
                else: st.markdown("### 👤")
            with c2:
                st.markdown(f"*🎉 {a['nome']}*")
                link = gerar_link_whatsapp(a["contato"], mensagem_aniversario(a["nome"]))
                if link:
                    st.markdown(f"[💬 Enviar mensagem pelo WhatsApp]({link})")

    st.markdown("### 🎂 Aniversariantes do Mês")
    mes = aniversariantes_mes()
    if mes:
        for a in mes:
            d = parse_data_nascimento(a["data_nascimento"])
            c1, c2, c3 = st.columns([1, 3, 2])
            with c1:
                foto = foto_aluno(a["nome"])
                if foto: st.image(foto, width=75)
                else: st.write("👤")
            with c2:
                st.markdown(f"*{a['nome']}*")
                st.write(d.strftime("%d/%m") if d else "Data não informada")
            with c3:
                link = gerar_link_whatsapp(a["contato"], mensagem_aniversario(a["nome"]))
                if link:
                    st.markdown(f"[💬 WhatsApp]({link})")
    else:
        st.info("Não há aniversariantes cadastrados neste mês.")

# ==============================================================================
# MATRÍCULA
# ==============================================================================
with tab_mat:
    st.markdown("### 📝 Matrícula")
    st.info("A assinatura deve ser realizada antes da biometria facial. A câmera não é aberta ao entrar nesta aba.")
    with st.form("form_matricula"):
        st.markdown("### 1. Dados do Aluno")
        nome = st.text_input("Nome Completo")
        nascimento = st.date_input("Data de Nascimento", value=datetime.date(2015,1,1),
                                  min_value=datetime.date(1950,1,1), max_value=datetime.date.today())
        escola = st.text_input("Escola")
        endereco = st.text_area("Endereço Completo")

        st.markdown("### 2. Responsável")
        responsavel = st.text_input("Nome do Responsável")
        parentesco = st.text_input("Grau de Parentesco")
        contato = st.text_input("Contato / WhatsApp")
        cin = st.text_input("CIN")

        st.markdown("### 3. Histórico de Saúde")
        saude = st.text_area("Histórico de Saúde / Observações")

        st.markdown("### 4. Autorizações")
        autorizacao = st.checkbox("Li e concordo com o Termo de Autorização de Participação.")
        imagem = st.checkbox("Li e concordo com o Termo de Uso de Imagem e Voz.")
        dados = st.checkbox("Li e concordo com o tratamento específico dos dados de saúde e biometria.")

        salvar = st.form_submit_button("SALVAR MATRÍCULA E IR PARA ASSINATURA")

        if salvar:
            erros = []
            if not nome.strip(): erros.append("Informe o nome completo.")
            if not responsavel.strip(): erros.append("Informe o responsável.")
            if not endereco.strip(): erros.append("Informe o endereço.")
            if not contato.strip(): erros.append("Informe o contato.")
            if not cin.strip(): erros.append("Informe a CIN.")
            if not autorizacao: erros.append("Aceite a autorização de participação.")
            if not imagem: erros.append("Registre a decisão sobre uso de imagem.")
            if not dados: erros.append("Aceite o tratamento específico dos dados sensíveis.")

            if erros:
                for e in erros: st.error(e)
            else:
                chave = normalizar_nome(nome)
                st.session_state.dados_app["alunos"][chave] = {
                    "nome": nome, "data_nascimento": nascimento.strftime("%d/%m/%Y"),
                    "escola": escola, "endereco": endereco, "responsavel": responsavel,
                    "parentesco": parentesco, "telefone": contato, "cin": cin,
                    "saude_detalhada": saude or "Não informado",
                    "biometria_cadastrada": False, "assinatura_cadastrada": False,
                    "data_matricula": hoje.strftime("%d/%m/%Y"), "pendencia": "Assinatura Pendente"
                }
                salvar_json()
                st.success("Matrícula salva. Agora realize a assinatura antes da biometria.")
                st.rerun()

    st.markdown("---")
    st.markdown("### ✍️ Assinatura do Responsável")
    nomes_local = [x.get("nome") for x in st.session_state.dados_app["alunos"].values()]
    if nomes_local:
        aluno_ass = st.selectbox("Aluno para assinatura", nomes_local, key="aluno_assinatura")
        if not st.session_state.dados_app["alunos"][normalizar_nome(aluno_ass)].get("assinatura_cadastrada"):
            if CANVAS_DISPONIVEL:
                assinatura = st_canvas(
                    fill_color="rgba(255,255,255,0)",
                    stroke_width=3, stroke_color="#000000",
                    background_color="#FFFFFF", height=180, width=700,
                    drawing_mode="freedraw", key="canvas_assinatura"
                )
                if st.button("CONFIRMAR ASSINATURA", key="confirmar_assinatura"):
                    if assinatura.image_data is None:
                        st.error("Faça a assinatura na tela antes de confirmar.")
                    else:
                        from PIL import Image
                        img = Image.fromarray(assinatura.image_data.astype("uint8"))
                        arq = DATA_DIR / ("assinatura_" + re.sub(r"[^a-zA-Z0-9_-]", "_", normalizar_nome(aluno_ass)) + ".png")
                        img.save(arq)
                        aluno_obj = st.session_state.dados_app["alunos"][normalizar_nome(aluno_ass)]
                        aluno_obj["assinatura_cadastrada"] = True
                        aluno_obj["data_assinatura"] = hoje.strftime("%d/%m/%Y")
                        aluno_obj["pendencia"] = "Biometria Facial Pendente"
                        st.session_state.dados_app["assinaturas"][normalizar_nome(aluno_ass)] = str(arq)
                        salvar_json()
                        st.success("🟢 Assinatura registrada. A biometria está liberada.")
                        st.rerun()
            else:
                st.warning("Instale streamlit-drawable-canvas/streamlit-canvas conforme o componente disponível no seu ambiente para assinatura na tela.")
        else:
            st.success("🟢 Assinatura já registrada.")

    st.markdown("---")
    st.markdown("### 📷 Biometria Facial")
    st.info("A câmera NÃO é aberta automaticamente. Ela só é solicitada depois que você clicar para iniciar a biometria e somente no aparelho da própria pessoa.")

    nomes_bio = [x.get("nome") for x in st.session_state.dados_app["alunos"].values()
                 if x.get("assinatura_cadastrada") and not x.get("biometria_cadastrada")]

    if nomes_bio:
        aluno_bio = st.selectbox("Aluno autorizado para biometria", nomes_bio, key="aluno_bio")
        if st.button("📷 INICIAR BIOMETRIA", key="iniciar_biometria"):
            st.session_state.camera_habilitada = True

        if st.session_state.camera_habilitada:
            foto = st.camera_input("A câmera será usada somente para capturar a foto deste aluno.", key="camera_biometria")
            if foto:
                caminho = salvar_foto(aluno_bio, foto.getvalue())
                obj = st.session_state.dados_app["alunos"][normalizar_nome(aluno_bio)]
                obj["foto"] = caminho
                obj["biometria_cadastrada"] = True
                obj["data_biometria"] = hoje.strftime("%d/%m/%Y")
                obj["pendencia"] = "Ok"
                st.session_state.dados_app["biometrias"][normalizar_nome(aluno_bio)] = True
                st.session_state.camera_habilitada = False
                salvar_json()
                st.success("🟢 Biometria cadastrada.")
                st.rerun()
    else:
        st.info("Nenhum aluno com assinatura pendente de biometria.")

# ==============================================================================
# ALUNOS CADASTRADOS
# ==============================================================================
with tab_alunos:
    st.markdown("### 👥 Alunos Cadastrados")
    alunos = alunos_cadastrados()
    if not alunos:
        st.warning("A aba Ficha de Cadastro não foi encontrada ou não possui alunos.")
    else:
        for a in alunos:
            bolinha, classe = status_aluno(a)
            st.markdown(f"""
            <div class="gold-card">
            <div class="perfil-nome">{bolinha} {a['nome']}</div>
            <p><strong>Data Início:</strong> {a['data_inicio'] or "Não informado"}</p>
            <p><strong>Faixa:</strong> {a['faixa']}</p>
            <p><strong>Data de Nascimento:</strong> {a['data_nascimento'] or "Não informado"}</p>
            <p><strong>Responsável:</strong> {a['responsavel'] or "Não informado"}</p>
            <p><strong>Contato:</strong> {a['contato'] or "Não informado"}</p>
            <p><strong>Histórico de Saúde:</strong> {a['historico_saude'] or "Não informado"}</p>
            <p><strong>Endereço:</strong> {a['endereco'] or "Não informado"}</p>
            <p><strong>Escola:</strong> {a['escola'] or "Não informado"}</p>
            </div>
            """, unsafe_allow_html=True)

# ==============================================================================
# CRONOGRAMA — BANNERS
# ==============================================================================
with tab_crono:
    st.markdown("### 📅 Cronograma")
    if df_cronograma.empty:
        st.info("Nenhum evento encontrado.")
    else:
        for _, row in df_cronograma.iterrows():
            banner_evento(row)

# ==============================================================================
# FREQUÊNCIA
# ==============================================================================
with tab_freq:
    st.markdown("### 📊 Frequência")
    st.caption("Desempenho do Mestre foi renomeado para Avaliação do Mestre.")
    nomes = [a["nome"] for a in alunos_cadastrados()]
    if nomes:
        aluno = st.selectbox("Selecione o aluno", nomes, key="freq_aluno")
        chave = normalizar_nome(aluno)
        registros_excel = []
        if not df_presencas_excel.empty:
            nome_col = encontrar_coluna(df_presencas_excel, ["aluno", "nome"])
            if nome_col:
                registros_excel = df_presencas_excel[
                    df_presencas_excel[nome_col].map(normalizar_nome).eq(chave)
                ].to_dict("records")
        locais = presencas_locais(aluno)
        st.metric("Presenças registradas no aplicativo", len(locais))
        if locais:
            st.write(", ".join(locais))
        if registros_excel:
            st.markdown("#### Registros das listas de presença")
            st.dataframe(pd.DataFrame(registros_excel), use_container_width=True, hide_index=True)

        if u["tipo"] == "mestre":
            atual = st.session_state.dados_app["desempenho_mestre"].get(chave, "Não informado")
            avaliacao = st.selectbox("Avaliação do Mestre", ["Ótimo", "Regular", "Dificuldade"], index=["Ótimo","Regular","Dificuldade"].index(atual) if atual in ["Ótimo","Regular","Dificuldade"] else 0)
            if st.button("Salvar Avaliação do Mestre", key="salvar_avaliacao"):
                st.session_state.dados_app["desempenho_mestre"][chave] = avaliacao
                salvar_json()
                st.success("Avaliação atualizada.")

# ==============================================================================
# TATAME
# ==============================================================================
with tab_tatame:
    st.markdown("### 🥋 Tatame")
    st.markdown("#### 📷 Chamada por Foto")
    st.info("A câmera não é aberta ao entrar nesta aba. Ela só será solicitada após o botão abaixo.")

    if st.button("📷 INICIAR CHAMADA POR FOTO", key="iniciar_chamada_foto"):
        st.session_state.camera_habilitada = True

    if st.session_state.camera_habilitada:
        foto_turma = st.camera_input("Capture a turma. A foto fica disponível somente nesta sessão para identificação.", key="camera_turma")
        if foto_turma:
            if not FACE_RECOGNITION_DISPONIVEL:
                st.error("O pacote face-recognition não está instalado. Instale-o no ambiente do aplicativo.")
            else:
                try:
                    arquivo = DATA_DIR / "foto_turma.jpg"
                    arquivo.write_bytes(foto_turma.getvalue())
                    img = face_recognition.load_image_file(str(arquivo))
                    locs = face_recognition.face_locations(img)
                    encs = face_recognition.face_encodings(img, locs)
                    conhecidos, nomes_conhecidos = [], []
                    for nome in [a["nome"] for a in alunos_cadastrados()]:
                        if not biometria_ok(nome):
                            continue
                        foto = foto_aluno(nome)
                        if foto and Path(foto).exists():
                            base = face_recognition.load_image_file(foto)
                            e = face_recognition.face_encodings(base)
                            if e:
                                conhecidos.append(e[0])
                                nomes_conhecidos.append(nome)
                    encontrados = []
                    for enc in encs:
                        resultados = face_recognition.compare_faces(conhecidos, enc, tolerance=0.48)
                        if True in resultados:
                            n = nomes_conhecidos[resultados.index(True)]
                            if n not in encontrados:
                                encontrados.append(n)
                    for nome in encontrados:
                        chave = normalizar_nome(nome)
                        st.session_state.dados_app["presencas_tatame"].setdefault(chave, [])
                        data_s = hoje.strftime("%d/%m/%Y")
                        if data_s not in st.session_state.dados_app["presencas_tatame"][chave]:
                            st.session_state.dados_app["presencas_tatame"][chave].append(data_s)
                    salvar_json()
                    for n in encontrados:
                        st.success(f"🟢 {n} — presença registrada.")
                    if len(encs) > len(encontrados):
                        st.warning("Há rostos não identificados. Use Presença Não Cadastrado quando necessário.")
                    st.session_state.camera_habilitada = False
                except Exception as e:
                    st.error("Erro na chamada por foto.")
                    st.code(str(e))

    st.markdown("---")
    st.markdown("#### ⚡ Chamada de Presença Rápida")
    nomes_t = [a["nome"] for a in alunos_cadastrados()]
    if nomes_t:
        for i, nome in enumerate(nomes_t):
            c1, c2, c3 = st.columns([1, 3, 2])
            with c1:
                foto = foto_aluno(nome)
                if foto: st.image(foto, width=70)
                else: st.write("👤")
            with c2:
                st.markdown(f"*{nome}*")
            with c3:
                if st.button("PRESENTE", key=f"pres_{i}"):
                    chave = normalizar_nome(nome)
                    data_s = hoje.strftime("%d/%m/%Y")
                    st.session_state.dados_app["presencas_tatame"].setdefault(chave, [])
                    if data_s not in st.session_state.dados_app["presencas_tatame"][chave]:
                        st.session_state.dados_app["presencas_tatame"][chave].append(data_s)
                    salvar_json()
                    st.success(f"Presença registrada: {nome}")

    st.markdown("---")
    st.markdown("#### ➕ Presença Não Cadastrado")
    with st.form("presenca_nao_cadastrado"):
        nc_nome = st.text_input("Nome Completo")
        nc_inicio = st.date_input("Data Início", value=hoje)
        nc_salvar = st.form_submit_button("Registrar e enviar para Alunos Cadastrados")
        if nc_salvar:
            if not nc_nome.strip():
                st.error("Informe o nome completo.")
            else:
                chave = normalizar_nome(nc_nome)
                if not any(normalizar_nome(a["nome"]) == chave for a in alunos_cadastrados()):
                    st.session_state.dados_app["presencas_nao_cadastrado"].append({
                        "nome": nc_nome, "data_inicio": nc_inicio.strftime("%d/%m/%Y"),
                        "data": hoje.strftime("%d/%m/%Y")
                    })
                    st.session_state.dados_app["alunos"][chave] = {
                        "nome": nc_nome, "data_inicio": nc_inicio.strftime("%d/%m/%Y"),
                        "biometria_cadastrada": False, "assinatura_cadastrada": False,
                        "pendencia": "Falta Documento"
                    }
                    salvar_json()
                    st.success("Aluno registrado e sinalizado com 🔴 em Alunos Cadastrados.")
                    st.rerun()
                else:
                    st.warning("Este aluno já existe no cadastro.")

# ==============================================================================
# CAMPEÃO
# ==============================================================================
with tab_campeao:
    st.markdown("### 🏆 Campeão")
    nomes = [a["nome"] for a in alunos_cadastrados()]
    if nomes:
        aluno = st.selectbox("Selecione o Aluno", nomes, key="campeao_aluno")
        dfc = campeonatos_do_aluno(aluno)
        if not dfc.empty:
            st.markdown("### 🥋 Campeonatos e classificações")
            st.dataframe(dfc, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum campeonato encontrado para este aluno.")

        if u["tipo"] == "mestre":
            st.markdown("---")
            with st.form("novo_campeonato"):
                cn = st.text_input("Campeonato")
                cd = st.date_input("Data", value=hoje)
                ca = st.text_input("Classificação do Aluno")
                ce = st.text_input("Classificação da Equipe")
                add = st.form_submit_button("Adicionar Campeonato")
                if add and cn.strip():
                    st.session_state.dados_app["campeonatos"].setdefault(normalizar_nome(aluno), []).append({
                        "Campeonato": cn, "Data": cd.strftime("%d/%m/%Y"),
                        "Classificação do Aluno": ca, "Classificação da Equipe": ce
                    })
                    salvar_json()
                    st.success("Campeonato registrado.")
                    st.rerun()

# ==============================================================================
# FÉ
# ==============================================================================
with tab_fe:
    st.markdown("### 🙏 Fé")
    st.markdown("## 🌱 Células")
    celulas = celulas_fe()
    if celulas:
        for i, c in enumerate(celulas):
            st.markdown(f"""
            <div class="gold-card">
            <div class="banner-title">🌱 {c['setor'] or f'Célula {i+1}'}</div>
            <div class="banner-text"><strong>Responsável:</strong> {c['responsavel']}</div>
            <div class="banner-text">Eu quero participar da Célula</div>
            </div>
            """, unsafe_allow_html=True)
            link = gerar_link_whatsapp(
                c["contato"],
                f"Paz do Senhor, {c['responsavel']}! Eu quero participar da Célula {c['setor']} do Projeto Sementes."
            )
            if link:
                st.markdown(f"[💬 Eu quero participar da Célula]({link})")
    else:
        st.info("Não foram encontradas as informações das células na aba Fé.")

    st.markdown("""
    <div class="banner-card">
    <div class="banner-title">⛪ CULTO — IEQ GUAICURUS</div>
    <div class="banner-text">
    <strong>Domingo às 18h</strong><br><br>
    Sua presença é muito importante para nós!<br>
    Venha cultuar ao Senhor conosco!
    </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🙏 Pedido de oração ao Pastor Joel")
    pedido = st.text_area("Escreva seu pedido de oração", key="pedido_joel")
    contato_joel = contato_pastor_joel()
    if contato_joel:
        link_oracao = gerar_link_whatsapp(
            contato_joel,
            f"Paz do Senhor, Pastor Joel. Gostaria de compartilhar um pedido de oração:\n\n{pedido}"
        )
        if pedido.strip():
            st.markdown(f"[🙏 Enviar pedido ao Pastor Joel pelo WhatsApp]({link_oracao})")
    else:
        st.warning("Contato do Pastor Joel não encontrado na aba Fé.")

    st.markdown("""
    <div class="gold-card" style="text-align:center;">
    <div class="banner-title">❤️ João 3:16</div>
    <div class="banner-text">
    “Porque Deus amou o mundo de tal maneira que deu o seu Filho unigênito,
    para que todo aquele que nele crê não pereça, mas tenha a vida eterna.”
    </div>
    </div>
    """, unsafe_allow_html=True)

    if contato_joel:
        link_batismo = gerar_link_whatsapp(
            contato_joel,
            "Paz do Senhor, Pastor Joel! Eu quero me Batizar!"
        )
        st.markdown(f"[💧 EU QUERO ME BATIZAR! — falar com o Pastor Joel]({link_batismo})")

    st.markdown("---")
    st.markdown("## 🗺️ Mapa de Impacto Espiritual")
    total_celulas = len(st.session_state.dados_app.get("celulas", []))
    total_batizados = len(st.session_state.dados_app.get("batizados", []))

    st.metric("Pessoas que entraram em células", total_celulas)
    st.metric("Pessoas que quiseram ser batizadas", total_batizados)

    st.markdown("### 📍 Mapa Social de Abrangência Territorial")
    st.markdown("*Projeto Sementes — Igreja IEQ Guaicurus (Corumbá - MS)*")
    enderecos = []
    for a in alunos_cadastrados():
        if a["endereco"]:
            enderecos.append(a["endereco"])
    for a in st.session_state.dados_app["alunos"].values():
        if a.get("endereco"):
            enderecos.append(a["endereco"])

    if enderecos:
        st.markdown(f"*{len(enderecos)} endereço(s) cadastrado(s)*")
        for endereco in sorted(set(enderecos)):
            st.markdown(f"📍 {endereco}")
        st.caption("O aplicativo lista os endereços cadastrados. Geocodificação automática não é feita sem um serviço de mapas configurado.")
    else:
        st.info("Nenhum endereço encontrado nos cadastros.")

# ==============================================================================
# SECRETARIA
# ==============================================================================
with tab_sec:
    st.markdown("### ⚙️ Secretaria")
    if u["tipo"] not in ("mestre", "diretoria"):
        st.info("Área administrativa restrita ao Mestre e à Diretoria.")
    else:
        st.markdown("### 🔐 Gerar Senha")
        st.markdown("#### 👨‍👩‍👧 Responsável")
        login_resp = st.text_input("Login do Responsável", key="sec_login_resp")
        nova_resp = st.text_input("Nova Senha", type="password", key="sec_senha_resp")
        if st.button("GERAR SENHA DO RESPONSÁVEL", key="gerar_resp"):
            if login_resp and nova_resp:
                chave = limpar_numero(login_resp) or login_resp.lower()
                st.session_state.dados_app["senhas_pais"][chave] = nova_resp
                salvar_json()
                st.success("Senha do responsável criada.")

        st.markdown("#### 👑 Diretoria")
        login_dir = st.text_input("Login Diretoria", key="sec_login_dir")
        nova_dir = st.text_input("Nova Senha", type="password", key="sec_senha_dir")
        if st.button("GERAR SENHA DA DIRETORIA", key="gerar_dir"):
            if login_dir and nova_dir:
                st.session_state.dados_app["senhas_diretoria"][login_dir.lower()] = nova_dir
                salvar_json()
                st.success("Senha da diretoria criada.")

        st.markdown("---")
        st.markdown("### 📊 Resumo Administrativo")
        st.metric("Alunos na Ficha de Cadastro", len(alunos_cadastrados()))
        st.metric("Biometrias cadastradas", sum(biometria_ok(a["nome"]) for a in alunos_cadastrados()))
        st.metric("Pedidos de oração registrados", len(st.session_state.dados_app.get("pedidos_oracao", [])))

        st.markdown("---")
        st.markdown("### 🗺️ MAPA SOCIAL DE ABRANGÊNCIA TERRITORIAL")
        st.markdown("*PROJETO SEMENTES — Igreja IEQ Guaicurus (Corumbá - MS)*")
        enderecos = [a["endereco"] for a in alunos_cadastrados() if a["endereco"]]
        if enderecos:
            for e in sorted(set(enderecos)):
                st.markdown(f"📍 {e}")
        else:
            st.info("Nenhum endereço cadastrado.")

# ==============================================================================
# LOGOUT
# ==============================================================================
st.markdown("---")
if st.button("SAIR DO APLICATIVO", key="logout"):
    st.session_state.logged_in = False
    st.session_state.user_info = None
    st.session_state.notificacao_aberta = False
    st.session_state.camera_habilitada = False
    st.rerun()
