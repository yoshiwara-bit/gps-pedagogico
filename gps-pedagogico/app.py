import streamlit as st
import pandas as pd
from docx import Document
from pypdf import PdfReader
import smtplib
from email.message import EmailMessage
from datetime import datetime
from io import BytesIO
import urllib.parse
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="GPS — Guia Pedagógico Semanal", page_icon="🧭", layout="wide")

# ==============================================================================
# CONFIGURAÇÕES DA SUA URE
# ==============================================================================
ID_PLANILHA_GOOGLE = "1v0SdvumxgFhoAH7W1gTJJ_TH50kKB2qa9QaB5ZQMHsU"
EMAIL_DISPARADOR = "seu.email.sistema@gmail.com"
SENHA_APP_GMAIL = "COLE_AQUI_AS_16_LETRAS_DO_GMAIL"
EMAIL_DIRIGENTE = "dirigente@educacao.sp.gov.br"
EMAIL_CEC = "marcos.yoshiwara@educacao.sp.gov.br"

def normalizar_texto(txt):
    if pd.isna(txt):
        return ""
    s = unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode("ASCII")
    return s.strip().upper()

def mapear_pasta_chave(pasta_bruta):
    p = normalizar_texto(pasta_bruta)
    if "CONVIVA" in p:
        return "CONVIVA"
    if "MULTIPLICA" in p:
        return "MULTIPLICA"
    if "QUALIDADE" in p:
        return "QUALIDADE DA AULA"
    if "DESENVOLVIMENTO CURRICULAR" in p:
        return "DESENVOLVIMENTO CURRICULAR"
    if "RECOMPOSICAO" in p:
        return "RECOMPOSICAO"
    if "EDUCACAO ESPECIAL" in p:
        return "EDUCACAO ESPECIAL"
    if p in ["CEC", "AT"]:
        return p
    return p

# --- LEITURA AUTOMÁTICA DAS ABAS DO GOOGLE SHEETS ---
@st.cache_data(ttl=60)
def carregar_aba(nome_aba):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{ID_PLANILHA_GOOGLE}/gviz/tq?tqx=out:csv&sheet={nome_aba}"
        return pd.read_csv(url, dtype=str)
    except Exception:
        return pd.DataFrame()

df_usuarios = carregar_aba("Usuarios")
df_pastas_raw = carregar_aba("Config_Pastas")
if df_pastas_raw.empty:
    df_pastas_raw = carregar_aba("Confing_Pastas")
df_eventos = carregar_aba("Calendario_Eventos")
df_bloqueios = carregar_aba("Bloqueios_Agenda")

# --- CONSOLIDAÇÃO AUTOMÁTICA DA ABA DE PASTAS (MESMO COM INCISOS EM VÁRIAS LINHAS) ---
def consolidar_pastas(df_p):
    if df_p.empty:
        return pd.DataFrame([
            {"Pasta_Chave": "CONVIVA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 5° - Convivência e Proteção Escolar (Conviva SP)."},
            {"Pasta_Chave": "QUALIDADE DA AULA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 1° - Qualidade da Aula e formação CGP/CGPG."},
            {"Pasta_Chave": "DESENVOLVIMENTO CURRICULAR", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 3° - Desenvolvimento Curricular, materiais e avaliações."},
            {"Pasta_Chave": "MULTIPLICA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 2° - Programa Multiplica SP (Resolução SEDUC Nº 90/2025)."},
            {"Pasta_Chave": "EDUCACAO ESPECIAL", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 4° - Educação Especial (EC, PAEE, PEI)."},
            {"Pasta_Chave": "RECOMPOSICAO", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 6° - Recomposição das Aprendizagens."},
        ])
    df_c = df_p.copy()
    col_res = "Resolução" if "Resolução" in df_c.columns else ("Resolucao" if "Resolucao" in df_c.columns else df_c.columns[1])
    df_c["Pasta"] = df_c["Pasta"].ffill()
    df_c[col_res] = df_c[col_res].ffill()
    df_c = df_c.dropna(subset=["Foco_Pedagogico"])
    agrupado = df_c.groupby(["Pasta", col_res], as_index=False).agg({"Foco_Pedagogico": lambda x: "\n".join(x.astype(str))})
    agrupado["Pasta_Chave"] = agrupado["Pasta"].apply(mapear_pasta_chave)
    agrupado = agrupado.rename(columns={col_res: "Resolucao"})
    return agrupado

df_pastas = consolidar_pastas(df_pastas_raw)

# --- TRADUTOR INTELIGENTE DE PLANILHAS BRUTAS DO DRIVE ---
def padronizar_planilha_bruta(df_raw):
    mapa = {}
    for col in df_raw.columns:
        c_low = str(col).strip().lower()
        if any(k in c_low for k in ["escola", "unidade", "ue"]):
            mapa[col] = "Escola"
        elif "etapa" in c_low or "segmento" in c_low:
            mapa[col] = "Etapa"
        elif "turma" in c_low or "série" in c_low or "serie" in c_low or "ano" in c_low:
            mapa[col] = "Turma"
        elif any(k in c_low for k in ["estudante", "aluno", "discente"]):
            mapa[col] = "Estudante"
        elif any(k in c_low for k in ["professor", "docente", "cursista", "multiplicador"]):
            mapa[col] = "Professor"
        elif "componente" in c_low or "disciplina" in c_low:
            mapa[col] = "Componente"
        elif any(k in c_low for k in ["nível", "nivel", "proficiência", "proficiencia", "desempenho"]):
            mapa[col] = "Nivel"
        elif "freq" in c_low:
            mapa[col] = "Frequencia"
        elif "habilidade" in c_low:
            mapa[col] = "Habilidade_Critica"
        elif "plataforma" in c_low or "tarefa" in c_low or "leia" in c_low or "khan" in c_low:
            mapa[col] = "Plataforma_Critica"
    df_pad = df_raw.rename(columns=mapa).copy()
    for col_req, default_val in [
        ("Escola", "EE Armindo Ramos"), ("Etapa", "Todas"), ("Turma", "Geral"),
        ("Estudante", "Verificar Lista"), ("Componente", "Todos"), ("Nivel", "Abaixo do Básico"),
        ("Frequencia", 75), ("Habilidade_Critica", "Prioritária da Semana"), ("Plataforma_Critica", "Abaixo da Meta")
    ]:
        if col_req not in df_pad.columns:
            df_pad[col_req] = default_val
    return df_pad

# --- MEMÓRIA DE SESSÃO ---
if "senhas_atualizadas" not in st.session_state:
    st.session_state["senhas_atualizadas"] = {}
if "indicadores_ativos" not in st.session_state:
    st.session_state["indicadores_ativos"] = pd.DataFrame([
        {"Escola": "EE Armando Gonçalves Professor", "Etapa": "Ensino Médio", "Turma": "1ª Série A", "Estudante": "Ana Clara", "Componente": "Matematica", "Nivel": "Abaixo do Básico", "Frequencia": 68, "Habilidade_Critica": "EM13MAT301", "Plataforma_Critica": "Khan Academy (38%)"},
        {"Escola": "EE Elvira Silva Professora", "Etapa": "Anos Iniciais", "Turma": "5º Ano B", "Estudante": "Lucas Oliveira", "Componente": "Anos Iniciais", "Nivel": "Abaixo do Básico", "Frequencia": 71, "Habilidade_Critica": "EF05LP01", "Plataforma_Critica": "Elefante Letrado (29%)"},
        {"Escola": "EE Raposo Tavares", "Etapa": "Ensino Médio", "Turma": "2ª Série C", "Estudante": "Pedro Henrique", "Componente": "Língua Portuguesa", "Nivel": "Abaixo do Básico", "Frequencia": 64, "Habilidade_Critica": "EM13LP01", "Plataforma_Critica": "Leia SP (32%)"},
        {"Escola": "EE Jeremias Junior Coronel", "Etapa": "Anos Finais", "Turma": "9º Ano A", "Estudante": "Mariana Costa", "Componente": "Artes", "Nivel": "Adequado", "Frequencia": 94, "Habilidade_Critica": "Nenhuma", "Plataforma_Critica": "Adequado (85%)"},
    ])
if "cursistas_ativos" not in st.session_state:
    st.session_state["cursistas_ativos"] = pd.DataFrame([
        {"Escola": "EE Raposo Tavares", "Professor": "Prof. Roberto (Cursista)", "Componente": "Artes", "Etapa": "Ensino Médio", "Status": "Ativo - Baixa entrega"},
    ])
if "boletins_pdf_texto" not in st.session_state:
    st.session_state["boletins_pdf_texto"] = ""

for chave in ["visitas", "encaminhamentos", "agenda_cec"]:
    if chave not in st.session_state:
        st.session_state[chave] = []

def gerar_link_google_agenda(titulo, data_obj, hora_inicio, hora_fim, local, descricao):
    dt_ini = f"{data_obj.strftime('%Y%m%d')}T{hora_inicio.replace(':', '')}00"
    dt_fim = f"{data_obj.strftime('%Y%m%d')}T{hora_fim.replace(':', '')}00"
    params = {"action": "TEMPLATE", "text": titulo, "dates": f"{dt_ini}/{dt_fim}", "location": local, "details": descricao}
    return "https://calendar.google.com/calendar/render?" + urllib.parse.urlencode(params)

# --- TELA DE LOGIN COM TROCA OBRIGATÓRIA DE SENHA PROVISÓRIA (123) ---
if "user" not in st.session_state:
    st.title("🧭 GPS — Guia Pedagógico Semanal")
    st.subheader("Unidade Regional de Ensino | Planejamento, Intencionalidade e Auditoria")

    if "usuario_troca_senha" in st.session_state:
        u_temp = st.session_state["usuario_troca_senha"]
        st.warning(f"🔐 **Primeiro Acesso Detectado ({u_temp['Nome']}):** A senha `123` é provisória. Crie sua senha pessoal definitiva para acessar o GPS.")
        nova_s1 = st.text_input("Nova Senha Pessoal:", type="password")
        nova_s2 = st.text_input("Confirme a Nova Senha:", type="password")
        c_ok, c_canc = st.columns([1, 3])
        with c_ok:
            if st.button("Salvar Nova Senha e Entrar", type="primary"):
                if len(nova_s1.strip()) < 4 or nova_s1.strip() == "123":
                    st.error("Escolha uma senha com pelo menos 4 caracteres e diferente de 123.")
                elif nova_s1 != nova_s2:
                    st.error("As senhas digitadas não coincidem.")
                else:
                    st.session_state["senhas_atualizadas"][u_temp["Email"]] = nova_s1.strip()
                    st.session_state["user"] = u_temp
                    del st.session_state["usuario_troca_senha"]
                    st.rerun()
        with c_canc:
            if st.button("Cancelar"):
                del st.session_state["usuario_troca_senha"]
                st.rerun()
        st.stop()

    c1, c2 = st.columns([1, 1])
    with c1:
        email_in = st.text_input("E-mail Institucional", placeholder="ex: marcos.yoshiwara@educacao.sp.gov.br")
        senha_in = st.text_input("Senha", type="password", placeholder="Primeiro acesso: 123")
        if st.button("Entrar no GPS", type="primary"):
            email_clean = email_in.strip().lower()
            senha_clean = senha_in.strip()
            df_usuarios["Email_Clean"] = df_usuarios["Email"].astype(str).str.strip().str.lower()
            df_usuarios["Senha_Clean"] = (
                df_usuarios["Senha"].fillna("123").astype(str).str.strip()
                .str.replace(r"\.0$", "", regex=True)
                .replace({"": "123", "nan": "123", "None": "123"})
            )
            match = df_usuarios[df_usuarios["Email_Clean"] == email_clean]

            if not match.empty:
                u_dict = match.iloc[0].to_dict()
                senha_real = st.session_state["senhas_atualizadas"].get(email_clean, u_dict["Senha_Clean"])
                if senha_clean == senha_real:
                    if senha_real == "123":
                        u_dict["Email"] = email_clean
                        st.session_state["usuario_troca_senha"] = u_dict
                        st.rerun()
                    else:
                        st.session_state["user"] = u_dict
                        st.rerun()
                else:
                    st.error("Senha incorreta.")
            else:
                st.error("E-mail não cadastrado na aba Usuarios.")
    st.stop()

user = st.session_state["user"]
pasta_completa = str(user.get("Pasta", "CEC"))
pasta_chave = mapear_pasta_chave(pasta_completa)

match_pasta = df_pastas[df_pastas["Pasta_Chave"] == pasta_chave]
if not match_pasta.empty:
    info_pasta = match_pasta.iloc[0].to_dict()
else:
    info_pasta = {"Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": f"Coordenação e Acompanhamento Pedagógico ({pasta_completa})"}

df_indicadores = st.session_state["indicadores_ativos"]
df_cursistas = st.session_state["cursistas_ativos"]

# --- FILTRO INTELIGENTE DAS ESCOLAS DO USUÁRIO ---
col_escola_user = "Escola" if "Escola" in user else ("Escolas" if "Escolas" in user else None)
escolas_user_raw = str(user.get(col_escola_user, "Todas")).strip() if col_escola_user else "Todas"

escolas_base = sorted(df_indicadores["Escola"].unique().tolist())
if escolas_user_raw.lower() not in ["todas", "nan", "none", ""]:
    partes = [p.strip() for p in escolas_user_raw.replace(" e ", ";").split(";") if p.strip()]
    escolas_disponiveis = sorted(list(set(partes)))
else:
    escolas_disponiveis = escolas_base

# --- MENU LATERAL ---
st.sidebar.title("🧭 GPS Pedagógico")
st.sidebar.markdown(f"**Usuário:** {user['Nome']}")
st.sidebar.info(f"**Pasta:** {pasta_chave}\n\n**Frentes:** {pasta_completa}\n\n**Componente:** {user.get('Componente', 'Todos')} | **Etapa:** {user.get('Etapa', 'Todas')}")
if st.sidebar.button("Sair"):
    del st.session_state["user"]
    st.rerun()

menu = st.sidebar.radio("Navegação", [
    "1. GPS Semanal (Intencionalidade)",
    "2. Plano de Trabalho, E-mail & Agenda",
    "3. Roteiro de Visita, Protocolo 179 & Repasse",
    "4. Painel CEC (Drive, Agendamentos & 1:1)"
])

# ==============================================================================
# ABA 1: GPS SEMANAL
# ==============================================================================
if menu == "1. GPS Semanal (Intencionalidade)":
    st.header(f"🧭 GPS Semanal — {pasta_chave}")
    st.caption(f"**Base Legal:** {info_pasta['Resolucao']}")
    with st.expander("📜 Ver Atribuições Legais Completas desta Pasta (Resolução SEDUC nº 1/2026)", expanded=False):
        st.text(info_pasta["Foco_Pedagogico"])

    if st.session_state["boletins_pdf_texto"]:
        with st.expander("📰 Destaques Extraídos do Boletim Semanal / Circular (Drive)", expanded=False):
            st.write(st.session_state["boletins_pdf_texto"][:1800] + "...")

    if not df_eventos.empty and "Pasta" in df_eventos.columns:
        ev_pasta = df_eventos[(df_eventos["Pasta"].apply(mapear_pasta_chave) == pasta_chave) | (df_eventos["Pasta"].astype(str).str.lower() == "todas")]
        for _, ev in ev_pasta.iterrows():
            st.info(f"📅 **Alerta de Calendário Obrigatório ({ev.get('Data_Evento', '')}):** **{ev.get('Nome_Evento', '')}** — {ev.get('Acao_Obrigatoria', '')}")

    meus_comp = [c for c in st.session_state["agenda_cec"] if c["Destino"] in ["Toda a Equipe", pasta_chave, user["Nome"]]]
    if meus_comp:
        st.warning("📌 **Alinhamentos / Convocações Agendados pelo CEC na Sua Semana:**")
        for c in meus_comp:
            st.write(f"- **{c['Data']} ({c['Horario']})** — [{c['Tipo']}] {c['Assunto']}")

    escola_gps = st.selectbox("Selecione a Escola para visualizar a Intencionalidade da Semana:", escolas_disponiveis)

    if pasta_chave == "MULTIPLICA":
        curs_esc = df_cursistas[df_cursistas["Escola"].str.lower() == escola_gps.lower()]
        if not curs_esc.empty:
            st.success(f"🎯 **INTENCIONALIDADE: ACOMPANHAMENTO FORMATIVO** — Escola com cursistas/multiplicadores ativos em **{user.get('Componente', '')}**.")
            st.dataframe(curs_esc, use_container_width=True)
        else:
            st.warning(f"📢 **INTENCIONALIDADE: MOBILIZAÇÃO, DIVULGAÇÃO E APOIO OPERACIONAL** — Escola sem cursistas ativos mapeados em **{user.get('Componente', '')}**. Foco: alinhar divulgação com CGP/Direção e destravar adesão.")

    dados_esc = df_indicadores[df_indicadores["Escola"].str.lower() == escola_gps.lower()]
    st.subheader("🔍 Raio-X Cirúrgico: Turmas, Habilidades, Plataformas e Estudantes")
    if dados_esc.empty:
        st.info("Aguardando carregamento da planilha de indicadores desta unidade no Painel CEC (Aba 4).")
    else:
        st.dataframe(dados_esc, use_container_width=True)

# ==============================================================================
# ABA 2: PLANO DE TRABALHO, E-MAIL E GOOGLE AGENDA
# ==============================================================================
elif menu == "2. Plano de Trabalho, E-mail & Agenda":
    st.header("📄 Plano de Trabalho Oficial, E-mail e Google Agenda")

    c1, c2, c3 = st.columns(3)
    with c1:
        natureza_dia = st.selectbox("Natureza da Atividade:", [
            "Visita Presencial à Escola",
            "Alinhamento Semanal com a Dirigente (URE)",
            "Alinhamento Equipe / Individual com CEC (URE)",
            "Plantão Interno na URE (Análise GPS)",
            "Convocação / Formação SEDUC ou EFAPE",
            "Feriado / Ponto Facultativo"
        ])
        data_plano = st.date_input("Data:", datetime.today())
    with c2:
        hora_ini = st.text_input("Hora Início (HH:MM):", "08:00")
        hora_fim = st.text_input("Hora Término (HH:MM):", "11:30")
    with c3:
        if natureza_dia == "Visita Presencial à Escola":
            local_plano = st.selectbox("Escola:", escolas_disponiveis)
            publico_plano = st.text_input("Turma / Professor / CGP:", "CGP e Docentes do Componente")
        else:
            local_plano = st.text_input("Local / Referência:", "Sede da URE / EFAPE")
            publico_plano = st.text_input("Participantes / Boletim:", "Equipe Pedagógica / Convocação")

    data_str = data_plano.strftime("%d/%m/%Y")
    if not df_bloqueios.empty and "Data" in df_bloqueios.columns:
        bloq = df_bloqueios[(df_bloqueios["Data"].astype(str) == data_str) & ((df_bloqueios["Pasta_Afetada"].astype(str).str.lower() == "todas") | (df_bloqueios["Pasta_Afetada"].apply(mapear_pasta_chave) == pasta_chave))]
        if not bloq.empty and natureza_dia == "Visita Presencial à Escola":
            st.error(f"🚫 **Atenção — Data Bloqueada ({bloq.iloc[0]['Tipo']}):** {bloq.iloc[0]['Descricao']}. Altere a Natureza da Atividade acima.")

    if natureza_dia == "Visita Presencial à Escola":
        dados_esc = df_indicadores[df_indicadores["Escola"].str.lower() == local_plano.lower()]
        alunos_crit = ", ".join(dados_esc[dados_esc["Nivel"] == "Abaixo do Básico"]["Estudante"].astype(str).tolist()) if not dados_esc.empty else "Estudantes prioritários da unidade"
        habs_crit = ", ".join(dados_esc["Habilidade_Critica"].astype(str).unique().tolist()) if not dados_esc.empty else "Habilidades essenciais da semana"
        plats_crit = ", ".join(dados_esc["Plataforma_Critica"].astype(str).unique().tolist()) if not dados_esc.empty else "Plataformas educacionais SEDUC"
        primeira_linha_legal = info_pasta["Foco_Pedagogico"].split("\n")[0]

        if pasta_chave == "MULTIPLICA":
            curs_esc = df_cursistas[df_cursistas["Escola"].str.lower() == local_plano.lower()]
            if not curs_esc.empty:
                texto_auto = f"Intencionalidade Formativa ({info_pasta['Resolucao']} - {primeira_linha_legal}): Acompanhamento presencial dos cursistas/multiplicadores de {user.get('Componente', '')} e desdobramento em ATPC."
            else:
                texto_auto = f"Intencionalidade de Mobilização e Apoio Operacional ({info_pasta['Resolucao']} - {primeira_linha_legal}): Divulgação do Programa Multiplica SP ({user.get('Componente', '')}) junto à equipe gestora e mapeamento de docentes elegíveis."
        else:
            texto_auto = (
                f"Fundamentação: {info_pasta['Resolucao']} — {primeira_linha_legal} "
                f"Foco da semana na unidade ({pasta_completa}): Habilidades críticas ({habs_crit}); Plataformas ({plats_crit}); "
                f"Estudantes prioritários: {alunos_crit}."
            )
    else:
        texto_auto = f"Atividade institucional ({natureza_dia}) conforme cronograma da Unidade Regional de Ensino / Agenda SEDUC."

    texto_final = st.text_area("Intencionalidade / Justificativa Oficial (editável):", value=texto_auto, height=130)

    doc = Document()
    doc.add_heading("PLANO DE TRABALHO SEMANAL — GPS PEDAGÓGICO", 0)
    doc.add_paragraph(f"Especialista (PEC): {user['Nome']} | Pasta: {pasta_chave} ({user.get('Componente', '')} - {user.get('Etapa', '')})")
    doc.add_paragraph(f"Base Legal: {info_pasta['Resolucao']}")
    tab = doc.add_table(rows=2, cols=4)
    tab.style = "Table Grid"
    h = tab.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Data / Horário", "Atividade / Local", "Público / Referência", "Intencionalidade Pedagógica"
    r = tab.rows[1].cells
    r[0].text = f"{data_str}\n{hora_ini} às {hora_fim}"
    r[1].text = f"{natureza_dia}\n{local_plano}"
    r[2].text = publico_plano
    r[3].text = texto_final

    buf = BytesIO()
    doc.save(buf)
    bytes_doc = buf.getvalue()

    link_gcal = gerar_link_google_agenda(f"GPS [{pasta_chave}] — {local_plano}", data_plano, hora_ini, hora_fim, local_plano, texto_final)

    b1, b2, b3 = st.columns(3)
    with b1:
        st.download_button("📥 1. Baixar Word (.docx)", data=bytes_doc, file_name=f"Plano_GPS_{user['Nome']}.docx")
    with b2:
        st.link_button("📅 2. Salvar na Google Agenda", link_gcal)
    with b3:
        if st.button("📧 3. Enviar E-mail (PEC e Dirigente)", type="primary"):
            if "COLE_AQUI" in SENHA_APP_GMAIL:
                st.warning("⚠️ Cole as 16 letras da Senha de App do Gmail no topo do código (`SENHA_APP_GMAIL`) para ativar o envio automático.")
            else:
                try:
                    msg = EmailMessage()
                    msg["Subject"] = f"📋 Plano de Trabalho GPS — {user['Nome']} ({pasta_chave}) | {data_str}"
                    msg["From"] = EMAIL_DISPARADOR
                    msg["To"] = f"{user['Email']}, {EMAIL_DIRIGENTE}"
                    msg["Cc"] = EMAIL_CEC
                    msg.set_content(f"Segue anexo o Plano de Trabalho oficial.\n\nAtividade: {natureza_dia} - {local_plano}\nData: {data_str} ({hora_ini} às {hora_fim})\nIntencionalidade: {texto_final}")
                    msg.add_attachment(bytes_doc, maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=f"Plano_{user['Nome']}.docx")
                    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                        smtp.login(EMAIL_DISPARADOR, SENHA_APP_GMAIL)
                        smtp.send_message(msg)
                    st.success(f"✅ Enviado com sucesso para {user['Email']} e Dirigente ({EMAIL_DIRIGENTE})!")
                except Exception as e:
                    st.error(f"Erro no envio: {e}")

# ==============================================================================
# ABA 3: ROTEIRO DE VISITA, AUDITORIA PROTOCOLO 179 E REPASSE ENTRE PECS
# ==============================================================================
elif menu == "3. Roteiro de Visita, Protocolo 179 & Repasse":
    st.header("✅ Roteiro Presencial, Auditoria Protocolo 179 e Repasse entre PECs")
    esc_vis = st.selectbox("Escola Visitada:", escolas_disponiveis)

    pendentes = [e for e in st.session_state["encaminhamentos"] if e["Escola"] == esc_vis and e["Status"] == "Pendente"]
    if pendentes:
        st.error("🔄 **Encaminhamento(s) Pendente(s) nesta Escola (Verificação Obrigatória):**")
        for idx, p in enumerate(pendentes):
            st.write(f"- **{p['Origem']} ({p['Pasta_Origem']})**: {p['Descricao']}")
            if st.checkbox("Confirmar que este encaminhamento foi verificado/cumprido hoje", key=f"chk_p_{idx}"):
                p["Status"] = "Cumprido"

    st.subheader("📋 Roteiro de Acompanhamento Presencial")
    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown("**1. Observação de Aula**")
        a1 = st.checkbox("Possui plano de aula alinhado à habilidade da semana")
        a2 = st.checkbox("Utiliza Material Digital / Escopo-Sequência")
        a3 = st.checkbox("Atende estudantes Abaixo do Básico de forma focalizada (agrupamentos/mediação)")
    with r2:
        st.markdown("**2. Acompanhamento de ATPC**")
        t1 = st.checkbox("Possui Pauta formativa estruturada")
        t2 = st.checkbox("Possui Lista de Presença e Ata registradas")
        t3 = st.checkbox("Promove socialização de boas práticas")
    with r3:
        st.markdown("**3. Alinhamento com Coordenador (CGP/CGPG)**")
        c_1 = st.checkbox("Verificação de pautas, atas e registros")
        c_2 = st.checkbox("Análise de indicadores por turma e nominalmente por estudante")
        c_3 = st.checkbox("Apoio presencial em sala e registro de feedbacks")

    termo_pdf = st.file_uploader("Upload do Termo de Visita Assinado (PDF) para Auditoria Automática", type=["pdf"])
    if termo_pdf is not None:
        leitor = PdfReader(termo_pdf)
        texto_pdf = " ".join([pag.extract_text() or "" for pag in leitor.pages]).lower()
        palavras_risco = ["violência", "violencia", "bullying", "agressão", "agressao", "ameaça", "automutilação"]
        palavras_fluxo = ["protocolo 179", "placon", "conviva", "família", "familia", "conselho tutelar", "rede de proteção", "acolhimento"]
        if any(w in texto_pdf for w in palavras_risco) and not any(w in texto_pdf for w in palavras_fluxo):
            st.error("🚨 **Alerta de Auditoria (Protocolo 179):** O Termo cita ocorrência de conflito/violência, mas não menciona registro na Plataforma Conviva (Placon), acionamento da família ou fluxo do Protocolo 179!")
        else:
            st.success("✅ **Auditoria de Termo:** Documento lido e validado sem inconsistências de protocolo.")

    st.subheader("📤 Encaminhamento / Passar Bastão para Outra Pasta")
    desc_enc = st.text_input("Descreva o encaminhamento acordado com a escola ou repassado a outro especialista:")
    lista_pecs = [f"{r['Nome']} ({mapear_pasta_chave(r['Pasta'])} - {r['Componente']})" for _, r in df_usuarios.iterrows() if mapear_pasta_chave(r["Pasta"]) not in ["CEC", "AT"]]
    dest_enc = st.selectbox("Responsável pelo acompanhamento no próximo ciclo:", ["Manter comigo (Próxima visita)"] + lista_pecs)

    if st.button("Salvar Registro Completo da Visita", type="primary"):
        prioridade_ok = a3 or c_2
        if not prioridade_ok:
            st.warning("⚠️ Alerta do GPS: Itens de atendimento focalizado aos estudantes críticos (Aula / CGP) não foram marcados. Isso será sinalizado no acompanhamento.")
        if desc_enc:
            st.session_state["encaminhamentos"].append({
                "Escola": esc_vis, "Origem": user["Nome"], "Pasta_Origem": pasta_chave,
                "Destino": dest_enc, "Descricao": desc_enc, "Status": "Pendente"
            })
        st.session_state["visitas"].append({
            "Data": datetime.today().strftime("%d/%m/%Y"), "PEC": user["Nome"],
            "Pasta": pasta_chave, "Escola": esc_vis, "Prioridade_OK": prioridade_ok
        })
        st.success("✅ Registro da visita salvo com sucesso!")

# ==============================================================================
# ABA 4: PAINEL DO CEC (CARREGAR ARQUIVOS DO DRIVE, AGENDAR E GESTÃO DE SENHAS)
# ==============================================================================
elif menu == "4. Painel CEC (Drive, Agendamentos & 1:1)":
    st.header("🎛️ Painel do CEC — Atualização Semanal, Agendamentos e Pauta 1:1")

    st.subheader("📂 1. Atualizar Dados da Semana (Arquivos Originais do Google Drive)")
    st.write("Não precisa editar colunas nem digitar nada no Sheets. Selecione os arquivos originais baixados do Drive (Excel, CSV ou Boletim em PDF):")
    arquivos_semana = st.file_uploader("Planilhas de Indicadores, Cursistas Multiplica ou Boletins (PDF):", type=["xlsx", "csv", "pdf"], accept_multiple_files=True)
    if arquivos_semana:
        dfs_ind, dfs_curs, textos_pdf = [], [], []
        for arq in arquivos_semana:
            nome_low = arq.name.lower()
            if nome_low.endswith(".pdf"):
                leitor = PdfReader(arq)
                textos_pdf.append(f"**[{arq.name}]**\n" + "\n".join([p.extract_text() or "" for p in leitor.pages]))
            else:
                df_lido = pd.read_csv(arq) if nome_low.endswith(".csv") else pd.read_excel(arq)
                df_pad = padronizar_planilha_bruta(df_lido)
                if "multiplica" in nome_low or "cursista" in nome_low:
                    dfs_curs.append(df_pad)
                else:
                    dfs_ind.append(df_pad)
        if dfs_ind:
            st.session_state["indicadores_ativos"] = pd.concat(dfs_ind, ignore_index=True)
        if dfs_curs:
            st.session_state["cursistas_ativos"] = pd.concat(dfs_curs, ignore_index=True)
        if textos_pdf:
            st.session_state["boletins_pdf_texto"] = "\n\n".join(textos_pdf)
        st.success(f"✅ {len(arquivos_semana)} arquivo(s) processado(s) e padronizado(s) automaticamente para todas as 6 pastas!")

    st.divider()
    st.subheader("📅 2. Agendar Alinhamento (Equipe / Individual) ou Convocação SEDUC")
    ca1, ca2, ca3 = st.columns(3)
    with ca1:
        tipo_ag = st.selectbox("Tipo de Compromisso:", [
            "Alinhamento Geral com Toda a Equipe",
            "Alinhamento por Pasta",
            "Alinhamento Individual (1:1)",
            "Convocação Agenda SEDUC / EFAPE"
        ])
        data_ag = st.date_input("Data do Compromisso:", datetime.today())
    with ca2:
        hor_ag = st.text_input("Horário (ex: 09:00 às 11:00):", "09:00 às 11:00")
        if tipo_ag == "Alinhamento Individual (1:1)":
            dest_ag = st.selectbox("Selecione o PEC:", df_usuarios[~df_usuarios["Pasta"].isin(["CEC", "AT"])]["Nome"].tolist())
        elif tipo_ag == "Alinhamento por Pasta":
            dest_ag = st.selectbox("Selecione a Pasta:", ["CONVIVA", "QUALIDADE DA AULA", "DESENVOLVIMENTO CURRICULAR", "MULTIPLICA", "EDUCACAO ESPECIAL", "RECOMPOSICAO"])
        else:
            dest_ag = "Toda a Equipe"
    with ca3:
        assunto_ag = st.text_input("Assunto / Pauta:", "Alinhamento Semanal de Indicadores GPS")
        disparar_email_ag = st.checkbox("Disparar convite por e-mail agora", value=True)

    if st.button("Salvar e Distribuir Compromisso", type="primary"):
        st.session_state["agenda_cec"].append({
            "Data": data_ag.strftime("%d/%m/%Y"), "Horario": hor_ag,
            "Tipo": tipo_ag, "Destino": dest_ag, "Assunto": assunto_ag
        })
        st.success(f"✅ Compromisso distribuído para **{dest_ag}** em {data_ag.strftime('%d/%m/%Y')}!")

        if disparar_email_ag and "COLE_AQUI" not in SENHA_APP_GMAIL:
            if dest_ag == "Toda a Equipe":
                emails_dest = df_usuarios["Email"].dropna().tolist()
            elif dest_ag in ["CONVIVA", "QUALIDADE DA AULA", "DESENVOLVIMENTO CURRICULAR", "MULTIPLICA", "EDUCACAO ESPECIAL", "RECOMPOSICAO"]:
                emails_dest = df_usuarios[df_usuarios["Pasta"].apply(mapear_pasta_chave) == dest_ag]["Email"].dropna().tolist()
            else:
                emails_dest = df_usuarios[df_usuarios["Nome"] == dest_ag]["Email"].dropna().tolist()
            try:
                msg = EmailMessage()
                msg["Subject"] = f"📅 Convocação / Alinhamento CEC — {assunto_ag} ({data_ag.strftime('%d/%m/%Y')})"
                msg["From"] = EMAIL_DISPARADOR
                msg["To"] = ", ".join(emails_dest)
                msg.set_content(f"Olá,\n\nAgendamento CEC:\nTipo: {tipo_ag}\nData: {data_ag.strftime('%d/%m/%Y')}\nHorário: {hor_ag}\nPauta: {assunto_ag}\n\nEste horário já foi reservado no seu GPS Pedagógico.")
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login(EMAIL_DISPARADOR, SENHA_APP_GMAIL)
                    smtp.send_message(msg)
                st.info(f"📧 Convite enviado para: {', '.join(emails_dest)}")
            except Exception as e:
                st.error(f"Erro ao enviar e-mail: {e}")

    st.divider()
    st.subheader("🔑 3. Gestão de Senhas (Resetar para 123) & Pauta 1:1")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        pec_reset = st.selectbox("Resetar senha de um PEC para provisória (123):", df_usuarios["Email"].dropna().tolist())
        if st.button("Resetar Senha para 123"):
            st.session_state["senhas_atualizadas"][pec_reset.strip().lower()] = "123"
            st.success(f"Senha de {pec_reset} redefinida para `123`.")
    with col_s2:
        df_e = pd.DataFrame(st.session_state["encaminhamentos"])
        if not df_e.empty:
            st.write("**Encaminhamentos e Repasses Ativos:**")
            st.dataframe(df_e, use_container_width=True)
