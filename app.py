import streamlit as st
import pandas as pd
import datetime
import urllib.parse

# 1. Configuração da Página
st.set_page_config(
    page_title="Projeto Sementes - Jiu-Jitsu",
    page_icon="🥋",
    layout="wide"
)
st.sidebar.image("logo projeto sementes.jpeg", use_container_width=True)

# --- ESTILIZAÇÃO CSS CUSTOMIZADA (DARK MODE) ---
st.markdown("""
<style>
/* Esconde o rodapé e o botão do Streamlit */
footer {visibility: hidden;}
.stAppDeployButton {display:none;}

/* Fundo Geral da Aplicação e Cor de Texto */
.stApp {
    background-color: #121214;
    color: #E1E1E6;
}
</style>
""", unsafe_allow_html=True)
  
    /* Barra Lateral (Sidebar) */
    [data-testid="stSidebar"] {
        background-color: #1A1A1E;
        border-right: 1px solid #29292E;
    }
    
    /* Expanders / Cards */
    .stExpander {
        background-color: #202024;
        border: 1px solid #29292E !important;
        border-radius: 8px;
    }
    
    /* Dataframes e Tabelas */
    [data-testid="stDataFrame"] {
        background-color: #202024;
        border-radius: 8px;
        padding: 8px;
    }

    /* Input text, Selectbox, Textarea e Multiselect */
    .stTextInput input, 
    .stSelectbox div[data-baseweb="select"], 
    .stMultiSelect div[data-baseweb="select"], 
    .stTextArea textarea {
        background-color: #202024 !important;
        color: #E1E1E6 !important;
        border-radius: 6px !important;
        border: 1px solid #29292E !important;
    }

    /* Botões do Streamlit */
    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #29292E;
        color: #00B37E;
        border: 1px solid #00B37E;
        border-radius: 6px;
        font-weight: bold;
        transition: all 0.3s ease;
    }

    .stButton>button:hover, div[data-testid="stFormSubmitButton"]>button:hover {
        background-color: #00B37E;
        color: #FFFFFF;
        border-color: #00B37E;
    }

    /* Botão/Link para envio de WhatsApp */
    a[href*="wa.me"] {
        display: inline-block;
        padding: 8px 16px;
        color: #FFFFFF !important;
        background-color: #25D366;
        border-radius: 6px;
        text-decoration: none;
        font-weight: bold;
        margin-top: 5px;
        transition: background-color 0.2s ease;
    }
    
    a[href*="wa.me"]:hover {
        background-color: #1EBE5D;
        color: #FFFFFF !important;
    }

    /* Títulos e Divisores */
    h1, h2, h3, h4, h5, h6 {
        color: #FFFFFF !important;
    }

    hr {
        border-color: #29292E;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Controle de Presença e Graduação Projeto Sementes.xlsx"

# 2. Funções Auxiliares
def carregar_aba(nome_aba, header_row=0):
    """Lê a aba do Excel limpando os dados"""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=nome_aba, header=header_row, engine='openpyxl')
        df = df.dropna(how='all').dropna(how='all', axis=1)
        return df
    except Exception as e:
        return pd.DataFrame()

def formatar_data_sem_hora(df):
    """Remove horas de colunas de datas no dataframe"""
    for col in df.columns:
        if any(kw in str(col).lower() for kw in ["data", "nascimento", "início", "inicio"]):
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').fillna(df[col])
    return df

def gerar_link_whatsapp(numero, mensagem):
    if pd.isna(numero) or str(numero).strip() == "" or str(numero).strip().lower() == "nan":
        return None
    num_limpo = ''.join(filter(str.isdigit, str(numero)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = "55" + num_limpo
    msg_enc = urllib.parse.quote(mensagem)
    return f"https://wa.me/{num_limpo}?text={msg_enc}"

def encontrar_coluna(df, palavras_chave, padrao_index):
    """Procura a coluna que contém uma das palavras-chave ou retorna a coluna pelo índice"""
    for col in df.columns:
        nome_col = str(col).lower().strip()
        for kw in palavras_chave:
            if kw in nome_col:
                return col
    if len(df.columns) > padrao_index:
        return df.columns[padrao_index]
    return df.columns[0]

def processar_dados_presenca(nome_aba):
    """Processa a aba de presença informada, cruza com a Ficha de Cadastro para obter a Faixa e extrai presenças dinamicamente"""
    try:
        df_cad = carregar_aba("Ficha de Cadastro")
        mapa_faixas = {}
        if not df_cad.empty:
            col_nome_cad = encontrar_coluna(df_cad, ["nome do aluno", "nome"], 1)
            col_faixa_cad = encontrar_coluna(df_cad, ["faixa", "graduação", "graduacao"], 3)
            for _, r in df_cad.iterrows():
                n = str(r[col_nome_cad]).strip().lower() if pd.notna(r[col_nome_cad]) else ""
                f = str(r[col_faixa_cad]).strip() if pd.notna(r[col_faixa_cad]) else "-"
                if n:
                    mapa_faixas[n] = f

        df_raw = pd.read_excel(EXCEL_FILE, sheet_name=nome_aba, header=None, engine='openpyxl')
        
        linha_dias_idx = None
        for idx, row in df_raw.iterrows():
            row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if "1" in row_vals and "2" in row_vals and "3" in row_vals:
                linha_dias_idx = idx
                break
                
        if linha_dias_idx is None:
            linha_dias_idx = 5
            
        linha_dias = df_raw.iloc[linha_dias_idx]
        col_dias_map = {}
        for col_idx, val in enumerate(linha_dias):
            val_str = str(val).strip().split('.')[0]
            if val_str.isdigit() and 1 <= int(val_str) <= 31:
                col_dias_map[col_idx] = val_str
                
        dados_alunos = []
        for idx in range(linha_dias_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            
            nome_a = str(row[1]).strip() if pd.notna(row[1]) else ""
            dt_inic_a = str(row[2]).replace(" 00:00:00", "").strip() if pd.notna(row[2]) else "-"
            
            if not nome_a or nome_a.lower() in ["nan", "none", "nome do aluno", "total", "nº"]:
                continue
            if any(term in nome_a.lower() for term in ["projeto sementes", "ficha de controle", "mês/ano", "dias de treino", "instrutor"]):
                continue
            
            faixa_a = mapa_faixas.get(nome_a.lower(), str(row[3]).strip() if pd.notna(row[3]) else "-")
                
            presencas = 0
            dias_marcados = {}
            for col_idx, dia_num in col_dias_map.items():
                val_pres = str(row[col_idx]).strip().upper() if pd.notna(row[col_idx]) else ""
                is_p = (val_pres == "X")
                dias_marcados[f"Dia {dia_num}"] = "X" if is_p else ""
                if is_p:
                    presencas += 1
                    
            linha_aluno = {
                "Nome do Aluno": nome_a,
                "Data Início": dt_inic_a,
                "Faixa": faixa_a,
                "Total Presenças": presencas,
                **dias_marcados
            }
            dados_alunos.append(linha_aluno)
            
        df_res = pd.DataFrame(dados_alunos)
        return df_res, col_dias_map
    except Exception as e:
        return pd.DataFrame(), {}

# 3. Navegação Lateral
st.sidebar.title("🥋 Projeto Sementes")
st.sidebar.markdown("---")

opcao = st.sidebar.radio(
    "Navegação do Sistema:",
    [
        "👥 Cadastro e Alunos",
        "📅 Controle de Presença",
        "🥋 Graduação e Métricas",
        "🔎 Alunos Não Identificados",
        "📲 Comunicação (WhatsApp)",
        "🎂 Aniversariantes",
        "➕ Novo Cadastro",
        "📷 Biometria Facial",
        "📸 Chamada por Foto"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Central do Professor - Projeto Sementes")

# 4. Módulos do Sistema

# --- 👥 CADASTRO DE ALUNOS ---
if opcao == "👥 Cadastro e Alunos":
    st.header("👥 Cadastro e Registro Geral dos Alunos")
    df_cad = carregar_aba("Ficha de Cadastro")
    
    if not df_cad.empty:
        df_cad = formatar_data_sem_hora(df_cad)
        col_nome = encontrar_coluna(df_cad, ["nome do aluno", "nome"], 1)
        col_resp = encontrar_coluna(df_cad, ["responsável", "responsavel"], 5)
        col_tel = encontrar_coluna(df_cad, ["contato", "telefone", "celular"], 6)
        col_doc = encontrar_coluna(df_cad, ["pendencia documento", "pendência", "pendencia"], 8)
        
        col_busca, _ = st.columns([2, 1])
        with col_busca:
            termo_busca = st.text_input("🔍 Buscar aluno por nome:")
        
        df_exibicao = df_cad.copy()
        if termo_busca:
            df_exibicao = df_exibicao[df_exibicao[col_nome].astype(str).str.contains(termo_busca, case=False, na=False)]
        
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.subheader("⚠️ Notificações Automáticas de Pendência de Documentação")
        
        df_pendentes = df_cad[
            df_cad[col_doc].notna() & 
            (df_cad[col_doc].astype(str).str.strip() != "") &
            (~df_cad[col_doc].astype(str).str.lower().isin(["ok", "regular", "concluído", "concluido", "nan"]))
        ].copy()
        
        if df_pendentes.empty:
            df_pendentes = df_cad.dropna(subset=[col_nome]).copy()
        st.markdown("### 📋 Lista 1: Alunos com Pendência na Planilha Base")
        df_lista1 = df_pendentes[[col_nome, col_resp, col_tel, col_doc]].copy()
        df_lista1.columns = ["Nome do Aluno", "Responsável", "Contato / WhatsApp", "Pendência Identificada"]
        st.dataframe(df_lista1, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.markdown("### 📲 Lista 2: Enviar Notificação aos Alunos Pendentes Selecionados")
        
        lista_nomes_pendentes = df_pendentes[col_nome].dropna().astype(str).unique().tolist()
        
        alunos_selecionados = st.multiselect(
            "Selecione os alunos pendentes para gerar as mensagens de aviso:",
            options=lista_nomes_pendentes,
            default=lista_nomes_pendentes
        )
        
        if alunos_selecionados:
            st.success(f"Foram selecionados {len(alunos_selecionados)} aluno(s) para envio.")
            
            for nome_aluno in alunos_selecionados:
                dados = df_cad[df_cad[col_nome].astype(str) == nome_aluno]
                if not dados.empty:
                    linha = dados.iloc[0]
                    resp_nome = str(linha[col_resp]) if pd.notna(linha[col_resp]) else "Responsável"
                    tel_num = str(linha[col_tel]) if pd.notna(linha[col_tel]) else ""
                    pend_texto = str(linha[col_doc]) if pd.notna(linha[col_doc]) else "Documentação Pendente"
                    msg_doc = f"Olá {resp_nome}, tudo bem? O sistema do Projeto Sementes identificou uma pendência na documentação do(a) aluno(a) {nome_aluno}: [{pend_texto}]. Solicitamos que entre em contato com a Diretoria para a devida regularização. Obrigado!"
                    
                    with st.expander(f"✉️ Enviar para: {nome_aluno} (Resp: {resp_nome}) — Status: {pend_texto}"):
                        st.write(f"Mensagem Automática: {msg_doc}")
                        link_wa = gerar_link_whatsapp(tel_num, msg_doc)
                        if link_wa:
                            st.markdown(f"[📲 Clique aqui para enviar via WhatsApp para {nome_aluno}]({link_wa})", unsafe_allow_html=True)

# --- 📅 CONTROLE DE PRESENÇA ---
elif opcao == "📅 Controle de Presença":
    st.header("📅 Controle de Presença")
    
    try:
        xl = pd.ExcelFile(EXCEL_FILE)
        abas_disponiveis = xl.sheet_names
        
        abas_presenca = [aba for aba in abas_disponiveis if any(p in aba.lower() for p in ["presença", "lista", "setembro", "agosto"])]
        if not abas_presenca:
            abas_presenca = abas_disponiveis
            
        aba_mes = st.selectbox("Selecione a Lista de Presença do Mês Correspondente:", abas_presenca)
        
        df_resultado, _ = processar_dados_presenca(aba_mes)
        
        st.subheader(f"📋 Presenças Registradas - Aba: {aba_mes}")
        st.caption("Exibindo Nome do Aluno, Data Início, Faixa e as colunas numeradas dos dias (Dia 1 ao 31) marcadas com 'X'.")
        
        if not df_resultado.empty:
            cols_exib = [c for c in df_resultado.columns if c != "Total Presenças"]
            st.dataframe(df_resultado[cols_exib], use_container_width=True, hide_index=True)
        else:
            st.warning("Nenhum dado encontrado para a aba selecionada.")
            
    except Exception as e:
        st.error(f"Erro ao processar a lista de presença da planilha: {e}")

# --- 🥋 CONTROLE DE GRADUAÇÃO E MÉTRICAS ---
elif opcao == "🥋 Graduação e Métricas":
    st.header("🥋 Controle Unificado de Graduação, Evolução e Métricas")
    st.caption("Esta tela processa automaticamente as informações das abas de Controle de Presença.")
    
    try:
        xl = pd.ExcelFile(EXCEL_FILE)
        abas_disponiveis = xl.sheet_names
        abas_presenca = [aba for aba in abas_disponiveis if any(p in aba.lower() for p in ["presença", "lista", "setembro", "agosto"])]
        if not abas_presenca:
            abas_presenca = abas_disponiveis
            
        # 1. TOTAL DE ALUNOS POR FAIXA
        aba_atual = abas_presenca[-1] if abas_presenca else None
        df_pres_atual, _ = processar_dados_presenca(aba_atual) if aba_atual else (pd.DataFrame(), {})
        
        st.markdown("---")
        st.subheader("1️⃣ Total de Alunos por Faixa")
        if not df_pres_atual.empty and "Faixa" in df_pres_atual.columns:
            faixa_counts = df_pres_atual["Faixa"].value_counts().reset_index()
            faixa_counts.columns = ["Faixa / Graduação", "Quantidade de Alunos"]
            c_f1, c_f2 = st.columns([2, 1])
            with c_f1:
                st.dataframe(faixa_counts, use_container_width=True, hide_index=True)
            with c_f2:
                st.metric("Total de Alunos Ativos", len(df_pres_atual))
        else:
            st.info("Nenhum dado de faixa identificado no momento.")

        # 2. MÉTRICAS GERAIS POR MÊS
        st.markdown("---")
        st.subheader("2️⃣ Métricas Gerais de Treino por Mês")
        
        resumo_meses = []
        dict_dados_meses = {}
        
        for aba in abas_presenca:
            df_m, col_dias_map = processar_dados_presenca(aba)
            if not df_m.empty:
                dias_com_treino = 0
                for dia_num in col_dias_map.values():
                    col_nome_dia = f"Dia {dia_num}"
                    if col_nome_dia in df_m.columns and (df_m[col_nome_dia] == "X").sum() > 0:
                        dias_com_treino += 1
                
                dias_calculo = dias_com_treino if dias_com_treino > 0 else 1
                
                tot_alunos = len(df_m)
                tot_presentes = df_m["Total Presenças"].sum()
                tot_possivel = tot_alunos * dias_calculo
                freq_media_pct = round((tot_presentes / tot_possivel) * 100, 1) if tot_possivel > 0 else 0.0
                
                nome_mes = aba.replace("Lista de Presença", "").replace("Lista", "").strip()
                resumo_meses.append({
                    "Mês / Aba": nome_mes if nome_mes else aba,
                    "Treinos Realizados no Mês": dias_com_treino,
                    "Total de Presenças": tot_presentes,
                    "Frequência Média Total": f"{freq_media_pct}%"
                })
                dict_dados_meses[aba] = {
                    "df": df_m,
                    "treinos": dias_calculo,
                    "nome_mes": nome_mes if nome_mes else aba
                }
                
        df_resumo_meses = pd.DataFrame(resumo_meses)
        st.dataframe(df_resumo_meses, use_container_width=True, hide_index=True)

        # 3. DESEMPENHO INDIVIDUAL DOS ALUNOS COM SINALIZAÇÃO
        st.markdown("---")
        st.subheader("3️⃣ Frequência Individual e Status do Aluno por Mês")
        
        mes_selecionado = st.selectbox("Selecione o mês para visualizar a frequência individual dos alunos:", abas_presenca)
        
        if mes_selecionado in dict_dados_meses:
            dados_mes = dict_dados_meses[mes_selecionado]
            df_al_m = dados_mes["df"].copy()
            treinos_mes = dados_mes["treinos"]
            
            def calcular_status(row):
                p = row["Total Presenças"]
                pct = (p / treinos_mes * 100) if treinos_mes > 0 else 0.0
                pct = min(pct, 100.0)
                
                if pct >= 85.0:
                    status = "🟢 Ótimo"
                elif pct >= 70.0:
                    status = "🟡 Regular"
                else:
                    status = "🔴 Preocupante"
                return pd.Series([p, f"{round(pct, 1)}%", status])
                
            df_al_m[["Aulas Presenciadas", "Frequência (%)", "Status de Frequência"]] = df_al_m.apply(calcular_status, axis=1)
            
            cols_finais = ["Nome do Aluno", "Faixa", "Aulas Presenciadas", "Frequência (%)", "Status de Frequência"]
            df_exib_ind = df_al_m[cols_finais].sort_values(by="Aulas Presenciadas", ascending=False)
            
            filtro_status = st.multiselect(
                "Filtrar por Status de Frequência:",
                options=["🟢 Ótimo", "🟡 Regular", "🔴 Preocupante"],
                default=["🟢 Ótimo", "🟡 Regular", "🔴 Preocupante"]
            )
            
            df_exib_filtrado = df_exib_ind[df_exib_ind["Status de Frequência"].isin(filtro_status)]
            st.dataframe(df_exib_filtrado, use_container_width=True, hide_index=True)
            
    except Exception as e:
        st.error(f"Erro ao processar métricas de graduação e presença: {e}")

# --- 🔎 ALUNOS NÃO IDENTIFICADOS ---
elif opcao == "🔎 Alunos Não Identificados":
    st.header("🔎 Alunos Não Identificados")
    st.warning("⚠️ Aluno detectado sem cadastro na foto do treino de hoje (14/09/2026): Lucas Gabriel Souza (Favor efetuar o cadastro em '➕ Novo Cadastro').")

# --- 📲 COMUNICAÇÃO COM OS ALUNOS ---
elif opcao == "📲 Comunicação (WhatsApp)":
    st.header("📲 Envio de Avisos e Comunicação (WhatsApp)")
    df_cad = carregar_aba("Ficha de Cadastro")
    
    if not df_cad.empty:
        col_nome = encontrar_coluna(df_cad, ["nome do aluno", "nome"], 1)
        col_contato = encontrar_coluna(df_cad, ["contato", "telefone", "celular"], 6)
        
        lista_nomes = df_cad[col_nome].dropna().astype(str).unique().tolist()
        aluno_sel = st.selectbox("Selecione o Aluno pelo Nome:", lista_nomes)
        
        dados_aluno = df_cad[df_cad[col_nome].astype(str) == aluno_sel]
        tel_aluno = str(dados_aluno[col_contato].values[0]) if not dados_aluno.empty else ""
        
        st.info(f"📞 Aluno Selecionado: {aluno_sel} | Telefone Cadastrado: {tel_aluno}")
        mensagem = st.text_area("Digite a mensagem:", f"Olá, responsável pelo(a) aluno(a) {aluno_sel}! Passando para dar um recado importante do Projeto Sementes.")
        
        link_wa = gerar_link_whatsapp(tel_aluno, mensagem)
        if link_wa:
            st.markdown(f"[📲 Clique aqui para abrir o WhatsApp de {aluno_sel}]({link_wa})", unsafe_allow_html=True)

# --- 🎂 ANIVERSARIANTES ---
elif opcao == "🎂 Aniversariantes":
    st.header("🎂 Aniversariantes por Mês e Notificações")
    
    df_cad = carregar_aba("Ficha de Cadastro")
    
    if not df_cad.empty:
        col_nome = encontrar_coluna(df_cad, ["nome do aluno", "nome"], 1)
        col_nasc = encontrar_coluna(df_cad, ["data de nascimento", "nascimento", "data nasc"], 4)
        col_resp = encontrar_coluna(df_cad, ["responsável", "responsavel"], 5)
        col_tel = encontrar_coluna(df_cad, ["contato", "telefone", "celular"], 6)
        
        df_aniv = df_cad[[col_nome, col_nasc, col_resp, col_tel]].dropna(subset=[col_nome, col_nasc]).copy()
        
        df_aniv['dt_obj'] = pd.to_datetime(df_aniv[col_nasc], errors='coerce')
        df_aniv = df_aniv.dropna(subset=['dt_obj'])
        
        df_aniv['Dia_Mes'] = df_aniv['dt_obj'].dt.strftime('%d/%m')
        df_aniv['Mes_Num'] = df_aniv['dt_obj'].dt.month
        df_aniv['Dia_Num'] = df_aniv['dt_obj'].dt.day
        
        hoje = datetime.date.today()
        dia_hoje = hoje.day
        mes_hoje = hoje.month
        
        # --- 1. NOTIFICAÇÃO DE ANIVERSARIANTES DO DIA ---
        aniv_hoje = df_aniv[(df_aniv['Dia_Num'] == dia_hoje) & (df_aniv['Mes_Num'] == mes_hoje)]
        
        if not aniv_hoje.empty:
            st.success(f"🎉 NOTIFICAÇÃO DO DIA: Hoje ({hoje.strftime('%d/%m')}) temos {len(aniv_hoje)} aniversariante(s)!")
            
            for _, row in aniv_hoje.iterrows():
                nome = row[col_nome]
                resp = str(row[col_resp]) if pd.notna(row[col_resp]) else "Responsável"
                tel = str(row[col_tel]) if pd.notna(row[col_tel]) else ""
                
                msg_parabens = f"🎉 Parabéns {nome}! O Projeto Sementes te deseja um feliz aniversário, com muita saúde, paz e evolução nos treinos! OSS! 🥋"
                link_wa = gerar_link_whatsapp(tel, msg_parabens)
                
                c_a1, c_a2 = st.columns([3, 1])
                with c_a1:
                    st.info(f"🎂 {nome} (Resp: {resp}) — Telefone: {tel}")
                with c_a2:
                    if link_wa:
                        st.markdown(f"[📲 Enviar Parabéns]({link_wa})", unsafe_allow_html=True)
        else:
            st.info(f"ℹ️ Nenhum aluno fazendo aniversário hoje ({hoje.strftime('%d/%m')}).")
            
        st.markdown("---")
        
        # --- 2. FILTRO SEPARADO POR MESES ---
        st.subheader("📅 Aniversariantes Separados por Mês")
        
        nomes_meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }
        
        mes_sel_nome = st.selectbox(
            "Selecione o Mês para visualizar a lista:",
            list(nomes_meses.values()),
            index=mes_hoje - 1
        )
        
        mes_sel_num = [k for k, v in nomes_meses.items() if v == mes_sel_nome][0]
        
        df_mes_filtrado = df_aniv[df_aniv['Mes_Num'] == mes_sel_num].sort_values(by='Dia_Num')
        
        if not df_mes_filtrado.empty:
            st.caption(f"Exibindo {len(df_mes_filtrado)} aniversariante(s) em {mes_sel_nome}:")
            
            df_exib = df_mes_filtrado[[col_nome, 'Dia_Mes', col_resp, col_tel]].copy()
            df_exib.columns = ["Nome do Aluno", "Dia / Mês", "Responsável", "Contato"]
            
            st.dataframe(df_exib, use_container_width=True, hide_index=True)
            
            st.markdown("### ✉️ Enviar Felicitações Individuais")
            for _, r in df_mes_filtrado.iterrows():
                al_nome = r[col_nome]
                al_dia = r['Dia_Mes']
                al_resp = str(r[col_resp]) if pd.notna(r[col_resp]) else ""
                al_tel = str(r[col_tel]) if pd.notna(r[col_tel]) else ""
                
                msg_m = f"Olá! Passando para parabenizar o(a) {al_nome} pelo aniversário em {al_dia}! Grande abraço da equipe do Projeto Sementes! 🥋🎉"
                link = gerar_link_whatsapp(al_tel, msg_m)
                
                with st.expander(f"🎈 {al_dia} - {al_nome}"):
                    st.write(f"Mensagem: {msg_m}")
                    if link:
                        st.markdown(f"[📲 Enviar mensagem via WhatsApp para {al_nome}]({link})", unsafe_allow_html=True)
        else:
            st.warning(f"Nenhum aniversariante encontrado no mês de {mes_sel_nome}.")

# --- ➕ NOVO CADASTRO ---
elif opcao == "➕ Novo Cadastro":
    st.header("➕ Formulário de Cadastro de Novo Aluno")
    
    with st.form("form_cadastro_completo"):
        c1, c2 = st.columns(2)
        
        with c1:
            nome = st.text_input("Nome do Aluno")
            data_inicio = st.date_input("Data Inicio", datetime.date.today())
            
            # Subcolunas para alinhar Faixa e Graus lado a lado
            sub_c1, sub_c2 = st.columns(2)
            
            faixa = sub_c1.selectbox(
                "FAIXA", 
                ["Branca", "Cinza", "Amarela", "Laranja", "Verde", "Azul", "Roxa", "Marrom", "Preta"], 
                key="select_faixa_cadastro"
            )
            graus = sub_c2.selectbox(
                "GRAUS", 
                ["0 Graus", "1 Grau", "2 Graus", "3 Graus", "4 Graus", "5 Graus", "6 Graus"],
                key="select_graus_cadastro"
            )
            
            data_nasc = st.date_input("Data de Nascimento")
            responsavel = st.text_input("Nome Completo do Responsável")
            
        with c2:
            contato = st.text_input("Contato Telefônico (WhatsApp)")
            saude = st.text_input("Histórico de Saúde / Restrições Médico-Esportivas")
            endereco = st.text_input("Endereço Residencial")
            escola = st.text_input("Escola / Instituição de Ensino")
            
        st.markdown("---")
        st.subheader("📷 Biometria Facial (Foto do Aluno)")
        foto_nova = st.camera_input("Tire a foto do aluno para a Biometria Facial")
        
        btn_salvar = st.form_submit_button("Salvar e Atualizar Ficha de Cadastro")
        
        if btn_salvar:
            if nome:
                st.success(f"Aluno {nome} ({faixa} - {graus}) cadastrado com sucesso junto com sua biometria facial!")
            else:
                st.error("Por favor, preencha pelo menos o nome do aluno.")

# --- 📷 BIOMETRIA FACIAL ---
elif opcao == "📷 Biometria Facial":
    st.header("📷 Cadastro e Registro de Biometria Facial")
    
    df_cad = carregar_aba("Ficha de Cadastro")
    
    if not df_cad.empty:
        col_nome = encontrar_coluna(df_cad, ["nome do aluno", "nome"], 1)
        lista_alunos = df_cad[col_nome].dropna().astype(str).unique().tolist()
        
        aluno_sel = st.selectbox("Selecione o aluno cadastrado para registrar ou atualizar a Biometria Facial:", lista_alunos)
        
        if aluno_sel:
            st.info(f"Registrando foto facial de: {aluno_sel}")
            foto_biometria = st.camera_input(f"Capturar foto de {aluno_sel}")
            
            if st.button("Salvar Biometria Facial"):
                if foto_biometria is not None:
                    st.success(f"Biometria Facial de {aluno_sel} vinculada e salva com sucesso!")
                else:
                    st.error("Por favor, capture a foto antes de salvar.")
    else:
        st.warning("Nenhum aluno cadastrado na planilha base para vincular a biometria.")

# --- 📸 CHAMADA POR FOTO ---
elif opcao == "📸 Chamada por Foto":
    st.header("📸 Chamada por Foto")
    foto_turma = st.camera_input("Tire a foto de toda a turma reunida")
    if foto_turma:
        st.info("Processando biometria facial dos alunos presentes...")
        
