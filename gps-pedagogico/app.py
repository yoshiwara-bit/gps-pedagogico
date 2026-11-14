import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import smtplib
from email.message import EmailMessage
from datetime import datetime
import calendar
from io import BytesIO
import urllib.parse
import urllib.request
import unicodedata
import os
import holidays
from icalendar import Calendar as ICalCalendar

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="GPS — Guia Pedagógico Semanal", page_icon="🧭", layout="wide")

# ==============================================================================
# CONFIGURAÇÕES DA URE DE MIRACATU (PREENCHA AS CREDENCIAIS ABAIXO)
# ==============================================================================
ID_PLANILHA_GOOGLE = "1v0SdvumxgFhoAH7W1gTJJ_TH50kKB2qa9QaB5ZQMHsU"
ID_AGENDA_SEDUC = "1wUt_4qcLgtQhlD6dZM4Ndae8aE0RLHYS-if84-F1zY8"
GID_AGENDA_SEDUC_EEC = "1592431035"
URL_PASTA_DRIVE_GPS = "https://drive.google.com/drive/folders/1FlLFdnhzUZJeR6ujwN4LP8jldIsMSg8m?usp=sharing"

URL_GOOGLE_AGENDA_CEC = "COLE_AQUI_SEU_ENDERECO_SECRETO_ICAL.ics"
EMAIL_DISPARADOR = "seu.email.sistema@gmail.com"
SENHA_APP_GMAIL = "COLE_AQUI_AS_16_LETRAS_DO_GMAIL"
EMAIL_DIRIGENTE = "mir@educacao.sp.gov.br"
EMAIL_CEC = "mir.eec@educacao.sp.gov.br"

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

# --- LEITURA AUTOMÁTICA DAS ABAS DO GOOGLE SHEETS E AGENDAS ---
@st.cache_data(ttl=60)
def carregar_aba(nome_aba):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{ID_PLANILHA_GOOGLE}/gviz/tq?tqx=out:csv&sheet={nome_aba}"
        return pd.read_csv(url, dtype=str)
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300)
def carregar_agenda_seduc_eec():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{ID_AGENDA_SEDUC}/export?format=csv&gid={GID_AGENDA_SEDUC_EEC}"
        return pd.read_csv(url, dtype=str)
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300)
def carregar_eventos_gcal(url_ics):
    if not url_ics or "COLE_AQUI" in url_ics:
        return []
    try:
        req = urllib.request.urlopen(url_ics, timeout=10)
        cal = ICalCalendar.from_ical(req.read())
        lista_ev = []
        for comp in cal.walk():
            if comp.name == "VEVENT":
                dtstart = comp.get("dtstart").dt
                dtend = comp.get("dtend").dt if comp.get("dtend") else dtstart
                resumo = str(comp.get("summary", "Compromisso URE / Indicadores"))
                local = str(comp.get("location", "Sede URE Miracatu"))
                desc = str(comp.get("description", ""))
                if hasattr(dtstart, "hour"):
                    hora_str = f"{dtstart.strftime('%H:%M')} às {dtend.strftime('%H:%M')}"
                else:
                    hora_str = "Dia Integral"
                lista_ev.append({
                    "ano": dtstart.year,
                    "mes": dtstart.month,
                    "dia": dtstart.day,
                    "data_str": dtstart.strftime("%d/%m/%Y"),
                    "horario": hora_str,
                    "resumo": resumo,
                    "local": local,
                    "descricao": desc
                })
        return lista_ev
    except Exception:
        return []

df_usuarios = carregar_aba("Usuarios")
df_pastas_raw = carregar_aba("Config_Pastas")
if df_pastas_raw.empty:
    df_pastas_raw = carregar_aba("Confing_Pastas")
df_eventos = carregar_aba("Calendario_Eventos")
df_bloqueios = carregar_aba("Bloqueios_Agenda")
df_agenda_seduc = carregar_agenda_seduc_eec()

MESES_INV = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}

def interpretar_data_seduc(data_extenso):
    if pd.isna(data_extenso):
        return None
    s = normalizar_texto(data_extenso).lower()
    if "," in s:
        s = s.split(",", 1)[1].strip()
    partes = [p.strip() for p in s.split(" de ") if p.strip()]
    if len(partes) == 3:
        try:
            dia = int(partes[0])
            mes = MESES_INV.get(partes[1], 0)
            ano = int(partes[2])
            if mes > 0:
                return datetime(ano, mes, dia)
        except Exception:
            return None
    return None

def evento_pertence_ao_pec(publico_alvo, pasta_pec, componente_pec=""):
    pub = normalizar_texto(publico_alvo)
    if not pub or "TODA A REDE" in pub:
        return True
    if pasta_pec in ["CEC", "AT"] and ("CEC" in pub or "EEC" in pub):
        return True
    if pasta_pec == "QUALIDADE DA AULA" and "QUALIDADE DA AULA" in pub:
        return True
    if pasta_pec == "CONVIVA" and "CONVIVA" in pub:
        return True
    if pasta_pec == "MULTIPLICA" and "MULTIPLICA" in pub:
        return True
    if pasta_pec == "EDUCACAO ESPECIAL" and "EDUCACAO ESPECIAL" in pub:
        return True
    if pasta_pec == "RECOMPOSICAO" and "RECOMPOSICAO" in pub:
        return True
    if pasta_pec == "DESENVOLVIMENTO CURRICULAR" and "DESENVOLVIMENTO CURRICULAR" in pub:
        comp_norm = normalizar_texto(componente_pec)
        if comp_norm and comp_norm != "TODOS":
            if comp_norm in pub or "DESENVOLVIMENTO CURRICULAR -" not in pub:
                return True
            return False
        return True
    return False

# --- CONSOLIDAÇÃO DA ABA DE PASTAS ---
def consolidar_pastas(df_p):
    padrao = pd.DataFrame([
        {"Pasta_Chave": "CONVIVA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 5° - Convivência e Proteção Escolar (Conviva SP)."},
        {"Pasta_Chave": "QUALIDADE DA AULA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 1° - Qualidade da Aula e formação CGP/CGPG."},
        {"Pasta_Chave": "DESENVOLVIMENTO CURRICULAR", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 3° - Desenvolvimento Curricular, materiais e avaliações."},
        {"Pasta_Chave": "MULTIPLICA", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 2° - Programa Multiplica SP (Resolução SEDUC Nº 90/2025)."},
        {"Pasta_Chave": "EDUCACAO ESPECIAL", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 4° - Educação Especial (EC, PAEE, PEI)."},
        {"Pasta_Chave": "RECOMPOSICAO", "Resolucao": "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026", "Foco_Pedagogico": "§ 6° - Recomposição das Aprendizagens."},
    ])
    if df_p.empty or len(df_p.columns) < 3:
        return padrao
    try:
        df_c = df_p.iloc[:, :3].copy()
        df_c.columns = ["Pasta", "Resolucao", "Foco_Pedagogico"]
        df_c["Pasta"] = df_c["Pasta"].ffill()
        df_c["Resolucao"] = df_c["Resolucao"].ffill()
        df_c = df_c.dropna(subset=["Pasta", "Foco_Pedagogico"])
        if df_c.empty:
            return padrao
        agrupado = df_c.groupby(["Pasta", "Resolucao"], as_index=False).agg({"Foco_Pedagogico": lambda x: "\n".join(x.astype(str))})
        agrupado["Pasta_Chave"] = agrupado["Pasta"].apply(mapear_pasta_chave)
        return agrupado
    except Exception:
        return padrao

df_pastas = consolidar_pastas(df_pastas_raw)

# --- TAREFAS E METAS MENSAIS PADRÃO POR PASTA (RESOLUÇÃO SEDUC Nº 1/2026) ---
TAREFAS_POR_PASTA = {
    "QUALIDADE DA AULA": [
        "Realizar visita presencial às escolas atribuídas (mínimo 1x a cada duas semanas)",
        "Realizar formação em serviço de CGP e CGPG com base no roteiro SEDUC",
        "Assistir a aulas previamente agendadas e registrar devolutivas formativas",
        "Analisar indicadores de frequência e metas das unidades escolares junto à Supervisão"
    ],
    "DESENVOLVIMENTO CURRICULAR": [
        "Articular currículo, materiais digitais/impressos e plataformas com SARESP, Prova Paulista, Provão Paulista, Diagnóstica, SAEB e ENEM",
        "Priorizar turmas, docentes e estudantes com proficiência Abaixo do Básico no componente",
        "Realizar encontros virtuais ou visitas presenciais nas escolas com menor desempenho nas avaliações",
        "Cascatear orientações técnicas e formações da SEDUC para CGP/CGPG e docentes"
    ],
    "CONVIVA": [
        "Realizar no mínimo 4 visitas presenciais às unidades escolares por semana",
        "Monitorar frequência crítica (< 75%), busca ativa e registros dos POC na Plataforma Conviva (Placon)",
        "Acompanhar atendimentos dos Psicólogos na Escola e fiscalizar prazos",
        "Revisar e acompanhar a implementação dos Planos de Melhoria de Convivência"
    ],
    "MULTIPLICA": [
        "Realizar visita presencial de no mínimo 2h/semana a Professor Multiplicador ou cursista",
        "Acompanhar painel de gestão do Programa Multiplica SP e intervir quando necessário",
        "Participar dos encontros formativos EFAPE e reuniões periódicas com a CAFF",
        "Mobilizar escolas sem cursistas ativos para adesão e divulgação do programa"
    ],
    "EDUCACAO ESPECIAL": [
        "Verificar Estudo de Caso (EC), PAEE, PEI e Relatórios Bimestrais nas escolas",
        "Priorizar visitas às unidades com maior demanda de apoio em inclusão e acessibilidade",
        "Apoiar CGP, CGPG, Professores do AEE e Ensino Colaborativo",
        "Reportar ao CEC relatório circunstanciado das ações formativas e atendimentos"
    ],
    "RECOMPOSICAO": [
        "Apoiar prioritariamente as escolas com maior porcentagem de estudantes em defasagem",
        "Monitorar evolução de proficiência, frequência e participação em Língua Portuguesa e Matemática",
        "Realizar visitas presenciais regulares e encontros virtuais nas escolas prioritárias",
        "Acionar PEC Desenvolvimento Curricular diante de dificuldades conceituais/materiais"
    ],
    "CEC": [
        "Conduzir alinhamento semanal com a Dirigente Regional e Equipe de Especialistas",
        "Consolidar Plano de Trabalho mensal da equipe e monitorar visitas/encaminhamentos",
        "Realizar reuniões individuais (1:1) e alinhamentos por pasta"
    ],
    "AT": [
        "Apoiar a coordenação dos alinhamentos regionais e consolidar registros da equipe",
        "Acompanhar cronograma de formações, convocações e indicadores da URE"
    ]
}

# --- TRADUTOR INTELIGENTE DE PLANILHAS BRUTAS DO DRIVE ---
def padronizar_planilha_bruta(df_raw):
    mapa = {}
    for col in df_raw.columns:
        c_low = normalizar_texto(col).lower()
        if any(k in c_low for k in ["escola", "unidade", "ue"]):
            mapa[col] = "Escola"
        elif any(k in c_low for k in ["avaliacao", "exame", "prova", "saresp", "saeb", "enem", "provao", "diagnostica"]):
            mapa[col] = "Avaliacao"
        elif "etapa" in c_low or "segmento" in c_low:
            mapa[col] = "Etapa"
        elif "turma" in c_low or "serie" in c_low or "ano" in c_low:
            mapa[col] = "Turma"
        elif any(k in c_low for k in ["professor", "docente", "regente", "cursista", "multiplicador"]):
            mapa[col] = "Professor"
        elif any(k in c_low for k in ["estudante", "aluno", "discente"]):
            mapa[col] = "Estudante"
        elif "componente" in c_low or "disciplina" in c_low or "area" in c_low:
            mapa[col] = "Componente"
        elif any(k in c_low for k in ["nivel", "proficiencia", "desempenho", "padrao"]):
            mapa[col] = "Nivel"
        elif any(k in c_low for k in ["nota", "acerto", "pontuacao", "media", "escore"]):
            mapa[col] = "Proficiencia_Nota"
        elif "freq" in c_low:
            mapa[col] = "Frequencia"
        elif "habilidade" in c_low or "descritor" in c_low:
            mapa[col] = "Habilidade_Critica"
        elif "plataforma" in c_low or "tarefa" in c_low or "leia" in c_low or "khan" in c_low:
            mapa[col] = "Plataforma_Critica"

    df_pad = df_raw.rename(columns=mapa).copy()
    for col_req, default_val in [
        ("Escola", "EE Armindo Ramos"),
        ("Avaliacao", "Prova Paulista / Diagnóstica"),
        ("Etapa", "Todas"),
        ("Turma", "Geral"),
        ("Professor", "Docente da Turma"),
        ("Estudante", "Verificar Lista"),
        ("Componente", "Todos"),
        ("Nivel", "Abaixo do Básico"),
        ("Proficiencia_Nota", "-"),
        ("Frequencia", 75),
        ("Habilidade_Critica", "Prioritária da Semana"),
        ("Plataforma_Critica", "Abaixo da Meta")
    ]:
        if col_req not in df_pad.columns:
            df_pad[col_req] = default_val
    return df_pad

# --- MEMÓRIA DE SESSÃO ---
if "senhas_atualizadas" not in st.session_state:
    st.session_state["senhas_atualizadas"] = {}
if "indicadores_ativos" not in st.session_state:
    st.session_state["indicadores_ativos"] = pd.DataFrame([
        {"Escola": "EE Armando Gonçalves Professor", "Avaliacao": "Prova Paulista", "Etapa": "Ensino Médio", "Turma": "1ª Série A", "Professor": "Prof. Carlos Mendes", "Estudante": "Ana Clara", "Componente": "Matematica", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "3,5", "Frequencia": 68, "Habilidade_Critica": "EM13MAT301", "Plataforma_Critica": "Khan Academy (38%)"},
        {"Escola": "EE Elvira Silva Professora", "Avaliacao": "Avaliação Diagnóstica", "Etapa": "Anos Iniciais", "Turma": "5º Ano B", "Professor": "Profa. Sandra Lima", "Estudante": "Lucas Oliveira", "Componente": "Anos Iniciais", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "4,0", "Frequencia": 71, "Habilidade_Critica": "EF05LP01", "Plataforma_Critica": "Elefante Letrado (29%)"},
        {"Escola": "EE Raposo Tavares", "Avaliacao": "SARESP / Provão Paulista", "Etapa": "Ensino Médio", "Turma": "3ª Série C", "Professor": "Profa. Helena Souza", "Estudante": "Pedro Henrique", "Componente": "Língua Portuguesa", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "198 pts", "Frequencia": 64, "Habilidade_Critica": "EM13LP01", "Plataforma_Critica": "Leia SP (32%)"},
        {"Escola": "EE Jeremias Junior Coronel", "Avaliacao": "SAEB / Prova Paulista", "Etapa": "Anos Finais", "Turma": "9º Ano A", "Professor": "Prof. Marcos Silva", "Estudante": "Mariana Costa", "Componente": "Artes", "Nivel": "Adequado", "Proficiencia_Nota": "8,5", "Frequencia": 94, "Habilidade_Critica": "Nenhuma", "Plataforma_Critica": "Adequado (85%)"},
    ])
if "cursistas_ativos" not in st.session_state:
    st.session_state["cursistas_ativos"] = pd.DataFrame([
        {"Escola": "EE Raposo Tavares", "Professor": "Prof. Roberto (Cursista)", "Componente": "Artes", "Etapa": "Ensino Médio", "Status": "Ativo - Baixa entrega"},
    ])
if "boletins_pdf_texto" not in st.session_state:
    st.session_state["boletins_pdf_texto"] = ""

for chave in ["visitas", "encaminhamentos", "agenda_cec", "compromissos_pec", "tarefas_extras_pec"]:
    if chave not in st.session_state:
        st.session_state[chave] = []

def gerar_link_google_agenda(titulo, data_obj, hora_inicio, hora_fim, local, descricao):
    dt_ini = f"{data_obj.strftime('%Y%m%d')}T{hora_inicio.replace(':', '')}00"
    dt_fim = f"{data_obj.strftime('%Y%m%d')}T{hora_fim.replace(':', '')}00"
    params = {"action": "TEMPLATE", "text": titulo, "dates": f"{dt_ini}/{dt_fim}", "location": local, "details": descricao}
    return "https://calendar.google.com/calendar/render?" + urllib.parse.urlencode(params)

def montar_docx_oficial(titulo_plano, nome_pec, pasta_str, data_ref_str, resolucao_str, linhas_tabela):
    doc = Document()
    secao = doc.sections[0]
    cabecalho = secao.header
    p_cab = cabecalho.paragraphs[0]
    p_cab.alignment = WD_ALIGN_PARAGRAPH.CENTER

    caminho_img = os.path.join(os.path.dirname(__file__), "cabecalho_ure.png")
    if os.path.exists(caminho_img):
        p_cab.add_run().add_picture(caminho_img, width=Inches(6.5))
    else:
        p_cab.add_run(
            "GOVERNO DO ESTADO DE SÃO PAULO\n"
            "UNIDADE REGIONAL DE ENSINO DE MIRACATU\n"
            "Av. Dona Evarista de Castro Ferreira, s/nº – Centro, Miracatu/SP, CEP: 11850000\n"
            "Telefone: (13) 3847-7050 / E-mail: mir@educacao.sp.gov.br"
        )

    titulo_doc = doc.add_heading(titulo_plano, 0)
    titulo_doc.alignment = WD_ALIGN_PARAGRAPH.CENTER

    tab_cab = doc.add_table(rows=2, cols=3)
    tab_cab.style = "Table Grid"
    hc = tab_cab.rows[0].cells
    hc[0].text, hc[1].text, hc[2].text = "Nome do PEC", "Pasta", "Data"
    rc = tab_cab.rows[1].cells
    rc[0].text = nome_pec
    rc[1].text = pasta_str
    rc[2].text = data_ref_str

    doc.add_paragraph(f"Base Legal: {resolucao_str}")

    tab = doc.add_table(rows=1 + len(linhas_tabela), cols=4)
    tab.style = "Table Grid"
    h = tab.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Data / Horário", "Atividade / Local", "Público / Referência", "Intencionalidade Pedagógica"
    for i, item in enumerate(linhas_tabela, start=1):
        r = tab.rows[i].cells
        r[0].text = item["data_hora"]
        r[1].text = item["atividade_local"]
        r[2].text = item["publico"]
        r[3].text = item["intencionalidade"]

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

# --- TELA DE LOGIN COM TROCA OBRIGATÓRIA DE SENHA PROVISÓRIA (123) ---
if "user" not in st.session_state:
    st.title("🧭 GPS — Guia Pedagógico Semanal")
    st.subheader("Unidade Regional de Ensino de Miracatu | Planejamento, Intencionalidade e Auditoria")

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
            if df_usuarios.empty:
                st.error("Não foi possível ler a aba Usuarios. Verifique se a planilha do Google Sheets está compartilhada como 'Qualquer pessoa com o link'.")
            else:
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

meses_pt = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# --- MENU LATERAL ---
st.sidebar.title("🧭 GPS Pedagógico")
st.sidebar.markdown(f"**Usuário:** {user['Nome']}")
st.sidebar.info(f"**Pasta:** {pasta_chave}\n\n**Frentes:** {pasta_completa}\n\n**Componente:** {user.get('Componente', 'Todos')} | **Etapa:** {user.get('Etapa', 'Todas')}")
st.sidebar.link_button("📁 Pasta GPS no Google Drive", URL_PASTA_DRIVE_GPS)

if st.sidebar.button("Sair"):
    del st.session_state["user"]
    st.rerun()

menu = st.sidebar.radio("Navegação", [
    "1. GPS Semanal (Intencionalidade)",
    "2. Planner Mensal & Tarefas da Pasta",
    "3. Plano de Trabalho, E-mail & Agenda",
    "4. Roteiro de Visita, Protocolo 179 & Repasse",
    "5. Painel CEC (Drive, Agendamentos & 1:1)"
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

    dados_esc = df_indicadores[df_indicadores["Escola"].str.lower() == escola_gps.lower()].copy()

    # Filtro automático por Componente quando for Desenvolvimento Curricular
    comp_user = normalizar_texto(user.get("Componente", "TODOS"))
    if pasta_chave == "DESENVOLVIMENTO CURRICULAR" and comp_user not in ["", "TODOS", "NAN", "NONE"]:
        dados_comp = dados_esc[dados_esc["Componente"].apply(normalizar_texto).str.contains(comp_user, na=False)]
        if not dados_comp.empty:
            dados_esc = dados_comp

    st.subheader("🔍 Raio-X Cirúrgico: Avaliações, Turmas, Professores e Estudantes")
    if dados_esc.empty:
        st.info("Aguardando carregamento das planilhas de indicadores/avaliações desta unidade no Painel CEC (Aba 5).")
    else:
        avaliacoes_disp = ["Todas as Avaliações"] + sorted(dados_esc["Avaliacao"].astype(str).unique().tolist())
        av_sel = st.selectbox("Filtrar por Avaliação (SARESP, Diagnóstica, Prova Paulista, Provão Paulista, SAEB, ENEM):", avaliacoes_disp)
        if av_sel != "Todas as Avaliações":
            dados_esc = dados_esc[dados_esc["Avaliacao"].astype(str) == av_sel]

        abaixo_basico = dados_esc[dados_esc["Nivel"].apply(normalizar_texto).str.contains("ABAIXO|INSUFICIENTE|DEFASAGEM", na=False)]
        freq_num = pd.to_numeric(dados_esc["Frequencia"], errors="coerce")
        freq_media = f"{freq_num.mean():.1f}%" if not freq_num.dropna().empty else "N/D"
        baixa_freq = len(dados_esc[freq_num < 75]) if not freq_num.dropna().empty else 0

        # Separação de foco: Desenvolvimento Curricular (Proficiência) vs Conviva/Qualidade (Frequência + Proficiência)
        if pasta_chave == "DESENVOLVIMENTO CURRICULAR":
            m1, m2, m3 = st.columns(3)
            m1.metric("Estudantes Abaixo do Básico", len(abaixo_basico))
            m2.metric("Turmas Críticas no Componente", dados_esc["Turma"].nunique())
            m3.metric("Docentes Mapeados", dados_esc["Professor"].nunique())

            st.markdown("#### 🎯 Priorização Curricular: Turmas, Professores e Proficiência")
            dados_esc["_critico"] = dados_esc["Nivel"].apply(lambda x: 1 if any(k in normalizar_texto(x) for k in ["ABAIXO", "INSUFICIENTE", "DEFASAGEM"]) else 0)
            resumo_turmas = (
                dados_esc.groupby(["Avaliacao", "Componente", "Turma", "Professor"], as_index=False)
                .agg(
                    Total_Estudantes=("Estudante", "count"),
                    Qtd_Abaixo_Basico=("_critico", "sum")
                )
                .sort_values(by=["Qtd_Abaixo_Basico"], ascending=False)
            )
            st.dataframe(resumo_turmas, use_container_width=True)

            st.markdown("#### 📋 Detalhamento Nominal: Estudantes, Proficiência e Plataformas")
            cols_exibir = [c for c in ["Avaliacao", "Componente", "Turma", "Professor", "Estudante", "Nivel", "Proficiencia_Nota", "Habilidade_Critica", "Plataforma_Critica"] if c in dados_esc.columns]
            st.dataframe(dados_esc[cols_exibir], use_container_width=True)
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Estudantes Abaixo do Básico", len(abaixo_basico))
            m2.metric("Turmas Mapeadas", dados_esc["Turma"].nunique())
            m3.metric("Frequência Média", freq_media)
            m4.metric("Alertas de Frequência (< 75%)", baixa_freq)

            st.markdown("#### 🎯 Priorização de Turmas e Professores (Frequência e Proficiência)")
            dados_esc["_critico"] = dados_esc["Nivel"].apply(lambda x: 1 if any(k in normalizar_texto(x) for k in ["ABAIXO", "INSUFICIENTE", "DEFASAGEM"]) else 0)
            dados_esc["_freq_num"] = freq_num
            resumo_turmas = (
                dados_esc.groupby(["Avaliacao", "Componente", "Turma", "Professor"], as_index=False)
                .agg(
                    Total_Estudantes=("Estudante", "count"),
                    Qtd_Abaixo_Basico=("_critico", "sum"),
                    Frequencia_Media=("_freq_num", "mean")
                )
                .sort_values(by=["Frequencia_Media", "Qtd_Abaixo_Basico"], ascending=[True, False])
            )
            resumo_turmas["Frequencia_Media"] = resumo_turmas["Frequencia_Media"].round(1).astype(str) + "%"
            st.dataframe(resumo_turmas, use_container_width=True)

            st.markdown("#### 📋 Detalhamento Nominal de Estudantes, Frequência e Habilidades")
            cols_exibir = [c for c in ["Avaliacao", "Componente", "Turma", "Professor", "Estudante", "Nivel", "Proficiencia_Nota", "Frequencia", "Habilidade_Critica", "Plataforma_Critica"] if c in dados_esc.columns]
            st.dataframe(dados_esc[cols_exibir], use_container_width=True)

# ==============================================================================
# ABA 2: PLANNER MENSAL VISUAL & TAREFAS DA PASTA
# ==============================================================================
elif menu == "2. Planner Mensal & Tarefas da Pasta":
    st.header(f"🗓️ Planner Mensal — {user['Nome']} ({pasta_chave})")

    cp1, cp2 = st.columns([1, 1])
    with cp1:
        mes_sel = st.selectbox("Mês:", list(meses_pt.keys()), index=datetime.today().month - 1, format_func=lambda m: meses_pt[m])
    with cp2:
        ano_sel = st.number_input("Ano:", min_value=2025, max_value=2030, value=datetime.today().year)

    mes_ano_planner = f"{meses_pt[mes_sel]}/{ano_sel}"
    st.caption("Legenda: 🔴 Feriado/Bloqueio | 🔵 Formação/Live SEDUC | 🟣 Alinhamento URE/CEC | 🟢 Visita/Compromisso PEC")

    eventos_por_dia = {d: [] for d in range(1, 32)}
    linhas_consolidadas_mes = []

    # 0. Feriados Nacionais e Estaduais (SP) Automáticos
    feriados_sp = holidays.Brazil(subdiv="SP", years=[ano_sel])
    for dt_f, nome_f in feriados_sp.items():
        if dt_f.month == mes_sel:
            d_txt = dt_f.strftime("%d/%m/%Y")
            eventos_por_dia[dt_f.day].append(f"🔴 [Feriado] {nome_f}")
            linhas_consolidadas_mes.append({
                "data_hora": f"{d_txt}\nDia Integral",
                "atividade_local": f"Feriado — {nome_f}",
                "publico": "Calendário Oficial Nacional/SP",
                "intencionalidade": "Suspensão de atividades e visitas presenciais conforme calendário oficial."
            })

    # 1. Bloqueios e Feriados Municipais (Google Sheets)
    if not df_bloqueios.empty and "Data" in df_bloqueios.columns:
        for _, b in df_bloqueios.iterrows():
            d_txt = str(b.get("Data", "")).strip()
            p_af = str(b.get("Pasta_Afetada", "Todas")).strip()
            if p_af.lower() == "todas" or mapear_pasta_chave(p_af) == pasta_chave:
                try:
                    dt_b = datetime.strptime(d_txt, "%d/%m/%Y")
                    if dt_b.month == mes_sel and dt_b.year == ano_sel:
                        rotulo = f"🔴 [{b.get('Tipo', 'Bloqueio')}] {b.get('Descricao', '')}"
                        eventos_por_dia[dt_b.day].append(rotulo)
                        linhas_consolidadas_mes.append({
                            "data_hora": f"{d_txt}\nDia Integral",
                            "atividade_local": f"{b.get('Tipo', 'Feriado')} — {b.get('Descricao', '')}",
                            "publico": "Calendário Regional/Municipal",
                            "intencionalidade": "Suspensão de visitas presenciais conforme calendário."
                        })
                except Exception:
                    pass

    # 2A. Formações e Eventos Manuais (Google Sheets)
    if not df_eventos.empty and "Data_Evento" in df_eventos.columns:
        for _, ev in df_eventos.iterrows():
            d_txt = str(ev.get("Data_Evento", "")).strip()
            p_ev = str(ev.get("Pasta", "Todas")).strip()
            if p_ev.lower() == "todas" or mapear_pasta_chave(p_ev) == pasta_chave:
                try:
                    dt_e = datetime.strptime(d_txt, "%d/%m/%Y")
                    if dt_e.month == mes_sel and dt_e.year == ano_sel:
                        rotulo = f"🔵 {ev.get('Nome_Evento', '')}"
                        eventos_por_dia[dt_e.day].append(rotulo)
                        linhas_consolidadas_mes.append({
                            "data_hora": f"{d_txt}\nConforme Pauta",
                            "atividade_local": ev.get("Nome_Evento", "Formação / Live"),
                            "publico": f"Pasta {pasta_chave}",
                            "intencionalidade": ev.get("Acao_Obrigatoria", "Participação e desdobramento formativo.")
                        })
                except Exception:
                    pass

    # 2B. Agenda Oficial SEDUC 2026 (Aba EEC - PEC e CEC)
    if not df_agenda_seduc.empty and "Data" in df_agenda_seduc.columns:
        for _, row_s in df_agenda_seduc.iterrows():
            status_ev = normalizar_texto(row_s.get("Status", ""))
            if "CANCELADA" in status_ev:
                continue
            dt_s = interpretar_data_seduc(row_s.get("Data", ""))
            if dt_s and dt_s.month == mes_sel and dt_s.year == ano_sel:
                pub_alvo = row_s.get("Público-alvo", "")
                if evento_pertence_ao_pec(pub_alvo, pasta_chave, user.get("Componente", "")):
                    h_ini_s = str(row_s.get("Início", "")).strip()
                    h_fim_s = str(row_s.get("Fim", "")).strip()
                    ativ_s = str(row_s.get("Atividade", "")).strip()
                    local_s = str(row_s.get("Local", "CMSP/Teams")).strip()
                    desc_s = str(row_s.get("Descrição", "")).strip()
                    eventos_por_dia[dt_s.day].append(f"🔵 [{h_ini_s}-{h_fim_s}] {ativ_s} ({local_s})")
                    linhas_consolidadas_mes.append({
                        "data_hora": f"{dt_s.strftime('%d/%m/%Y')}\n{h_ini_s} às {h_fim_s}",
                        "atividade_local": f"[Agenda SEDUC] {ativ_s} ({local_s})",
                        "publico": pub_alvo,
                        "intencionalidade": desc_s if desc_s and desc_s != "nan" else f"Participação na formação/orientação SEDUC ({ativ_s})."
                    })

    # 3A. Alinhamentos agendados pelo CEC no painel
    for c in st.session_state["agenda_cec"]:
        if c["Destino"] in ["Toda a Equipe", pasta_chave, user["Nome"]]:
            try:
                dt_c = datetime.strptime(c["Data"], "%d/%m/%Y")
                if dt_c.month == mes_sel and dt_c.year == ano_sel:
                    eventos_por_dia[dt_c.day].append(f"🟣 [{c['Horario']}] {c['Assunto']}")
                    linhas_consolidadas_mes.append({
                        "data_hora": f"{c['Data']}\n{c['Horario']}",
                        "atividade_local": f"{c['Tipo']} — Sede URE",
                        "publico": c["Destino"],
                        "intencionalidade": c["Assunto"]
                    })
            except Exception:
                pass

    # 3B. Eventos vindos direto da Google Agenda do CEC
    for ev_g in carregar_eventos_gcal(URL_GOOGLE_AGENDA_CEC):
        if ev_g["mes"] == mes_sel and ev_g["ano"] == ano_sel:
            eventos_por_dia[ev_g["dia"]].append(f"🟣 [{ev_g['horario']}] {ev_g['resumo']}")
            linhas_consolidadas_mes.append({
                "data_hora": f"{ev_g['data_str']}\n{ev_g['horario']}",
                "atividade_local": f"[Agenda URE/CEC] {ev_g['resumo']} ({ev_g['local']})",
                "publico": "Equipe EEC / URE Miracatu",
                "intencionalidade": ev_g["descricao"] if ev_g["descricao"] else ev_g["resumo"]
            })

    # 4. Visitas e compromissos próprios do PEC
    for comp in st.session_state["compromissos_pec"]:
        if comp["PEC"] == user["Nome"]:
            try:
                dt_p = datetime.strptime(comp["Data"], "%d/%m/%Y")
                if dt_p.month == mes_sel and dt_p.year == ano_sel:
                    eventos_por_dia[dt_p.day].append(f"🟢 [{comp['Horario']}] {comp['Local']}")
                    linhas_consolidadas_mes.append({
                        "data_hora": f"{comp['Data']}\n{comp['Horario']}",
                        "atividade_local": f"{comp['Natureza']} — {comp['Local']}",
                        "publico": comp["Publico"],
                        "intencionalidade": comp["Intencionalidade"]
                    })
            except Exception:
                pass

    # --- GRADE VISUAL DO CALENDÁRIO MENSAL (SEGUNDA A SÁBADO) ---
    cal = calendar.Calendar(firstweekday=0)
    semanas = cal.monthdayscalendar(ano_sel, mes_sel)
    dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"]

    cols_h = st.columns(6)
    for idx_d, nome_d in enumerate(dias_semana):
        cols_h[idx_d].markdown(f"**{nome_d}**")

    for sem in semanas:
        cols_s = st.columns(6)
        for idx_d in range(6):
            dia_num = sem[idx_d]
            with cols_s[idx_d]:
                if dia_num == 0:
                    st.markdown("<div style='min-height:95px; background-color:#f8f9fa; border-radius:6px; padding:6px;'></div>", unsafe_allow_html=True)
                else:
                    itens_dia = eventos_por_dia.get(dia_num, [])
                    conteudo_html = f"<strong>Dia {dia_num:02d}</strong><br>"
                    if itens_dia:
                        conteudo_html += "<br>".join([f"<span style='font-size:12px;'>{it}</span>" for it in itens_dia])
                    else:
                        conteudo_html += "<span style='font-size:11px; color:#888;'>Livre para visita/planejamento</span>"
                    borda = "#2b6cb0" if itens_dia else "#e2e8f0"
                    fundo = "#ebf8ff" if itens_dia else "#ffffff"
                    st.markdown(
                        f"<div style='min-height:105px; border:1px solid {borda}; background-color:{fundo}; border-radius:6px; padding:6px; margin-bottom:6px;'>{conteudo_html}</div>",
                        unsafe_allow_html=True
                    )

    st.divider()
    col_t1, col_t2 = st.columns([1.2, 1])

    with col_t1:
        st.subheader(f"✅ Tarefas e Atribuições Obrigatórias do Mês ({pasta_chave})")
        tarefas_base = TAREFAS_POR_PASTA.get(pasta_chave, TAREFAS_POR_PASTA["QUALIDADE DA AULA"])
        for idx_t, t_txt in enumerate(tarefas_base):
            st.checkbox(t_txt, key=f"tarefa_padrao_{pasta_chave}_{mes_sel}_{idx_t}")

        for idx_ex, t_ex in enumerate(st.session_state["tarefas_extras_pec"]):
            if t_ex["PEC"] == user["Nome"] and t_ex["Mes"] == mes_ano_planner:
                st.checkbox(f"📌 {t_ex['Descricao']}", key=f"tarefa_extra_{idx_ex}")

        nova_tarefa = st.text_input("Adicionar tarefa ou demanda específica deste mês:")
        if st.button("Adicionar Tarefa ao Planner"):
            if nova_tarefa.strip():
                st.session_state["tarefas_extras_pec"].append({
                    "PEC": user["Nome"], "Mes": mes_ano_planner, "Descricao": nova_tarefa.strip()
                })
                st.rerun()

    with col_t2:
        st.subheader("➕ Inserir Compromisso Rápido no Planner")
        dia_add = st.date_input("Data do Compromisso:", datetime(ano_sel, mes_sel, min(datetime.today().day, 28)), key="data_quick_planner")
        nat_add = st.selectbox("Tipo:", ["Visita Presencial à Escola", "Formação / Live SEDUC", "Alinhamento URE", "Plantão Interno GPS"], key="nat_quick")
        loc_add = st.selectbox("Escola / Local:", escolas_disponiveis + ["Sede da URE Miracatu", "Transmissão Online (EFAPE/CMSP)"])
        hor_add = st.text_input("Horário:", "08:00 às 12:00", key="hor_quick")
        int_add = st.text_input("Objetivo / Foco:", "Acompanhamento pedagógico conforme GPS")
        if st.button("Salvar no Calendário do Mês", type="primary"):
            st.session_state["compromissos_pec"].append({
                "PEC": user["Nome"], "Data": dia_add.strftime("%d/%m/%Y"),
                "Horario": hor_add, "Natureza": nat_add, "Local": loc_add,
                "Publico": "Equipe Escolar / URE", "Intencionalidade": int_add
            })
            st.success("Compromisso adicionado ao Planner!")
            st.rerun()

    st.divider()
    st.subheader(f"📥 Exportar Plano de Trabalho Consolidado de {mes_ano_planner}")
    if not linhas_consolidadas_mes:
        st.info("Adicione compromissos acima ou preencha eventos na Aba 3 para consolidar o documento completo do mês.")
    else:
        linhas_ordenadas = sorted(linhas_consolidadas_mes, key=lambda x: x["data_hora"])
        bytes_mes = montar_docx_oficial(
            f"Plano de Trabalho do Mês de {mes_ano_planner}",
            str(user["Nome"]),
            f"{pasta_chave} ({user.get('Componente', 'Todos')})",
            mes_ano_planner,
            info_pasta["Resolucao"],
            linhas_ordenadas
        )
        nome_arq_mes = f"Plano_Trabalho_{meses_pt[mes_sel]}_{ano_sel}_{user['Nome'].replace(' ', '_')}.docx"
        cm1, cm2 = st.columns(2)
        with cm1:
            st.download_button("📥 Baixar Plano do Mês Completo (.docx com Cabeçalho URE)", data=bytes_mes, file_name=nome_arq_mes)
        with cm2:
            if st.button("📧 Enviar Plano do Mês por E-mail (PEC, Dirigente e CEC)", type="primary"):
                if "COLE_AQUI" in SENHA_APP_GMAIL:
                    st.warning("⚠️ Configure `SENHA_APP_GMAIL` no topo do código para disparar e-mails.")
                else:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = f"📋 Plano de Trabalho do Mês de {mes_ano_planner} — {user['Nome']} ({pasta_chave})"
                        msg["From"] = EMAIL_DISPARADOR
                        msg["To"] = f"{user['Email']}, {EMAIL_DIRIGENTE}"
                        msg["Cc"] = EMAIL_CEC
                        msg.set_content(f"Prezados(as),\n\nSegue em anexo o Plano de Trabalho consolidado referente ao mês de {mes_ano_planner}.\n\nEspecialista: {user['Nome']}\nPasta: {pasta_chave}\n\nAtenciosamente,\nSistema GPS — URE Miracatu")
                        msg.add_attachment(bytes_mes, maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=nome_arq_mes)
                        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                            smtp.login(EMAIL_DISPARADOR, SENHA_APP_GMAIL)
                            smtp.send_message(msg)
                        st.success(f"✅ Plano mensal enviado para {user['Email']}, Dirigente ({EMAIL_DIRIGENTE}) e CEC ({EMAIL_CEC})!")
                    except Exception as e:
                        st.error(f"Erro no envio: {e}")

# ==============================================================================
# ABA 3: PLANO DE TRABALHO INDIVIDUAL, E-MAIL E GOOGLE AGENDA
# ==============================================================================
elif menu == "3. Plano de Trabalho, E-mail & Agenda":
    st.header("📄 Gerador de Plano de Trabalho (Por Atividade / Visita), E-mail e Agenda")

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
            local_plano = st.text_input("Local / Referência:", "Sede da URE Miracatu / EFAPE")
            publico_plano = st.text_input("Participantes / Boletim:", "Equipe Pedagógica / Convocação")

    data_str = data_plano.strftime("%d/%m/%Y")
    mes_ano_str = f"{meses_pt[data_plano.month]}/{data_plano.year}"

    if not df_bloqueios.empty and "Data" in df_bloqueios.columns:
        bloq = df_bloqueios[(df_bloqueios["Data"].astype(str) == data_str) & ((df_bloqueios["Pasta_Afetada"].astype(str).str.lower() == "todas") | (df_bloqueios["Pasta_Afetada"].apply(mapear_pasta_chave) == pasta_chave))]
        if not bloq.empty and natureza_dia == "Visita Presencial à Escola":
            st.error(f"🚫 **Atenção — Data Bloqueada ({bloq.iloc[0]['Tipo']}):** {bloq.iloc[0]['Descricao']}. Altere a Natureza da Atividade acima.")

    if natureza_dia == "Visita Presencial à Escola":
        dados_esc = df_indicadores[df_indicadores["Escola"].str.lower() == local_plano.lower()]
        crit_esc = dados_esc[dados_esc["Nivel"].apply(normalizar_texto).str.contains("ABAIXO|INSUFICIENTE|DEFASAGEM", na=False)] if not dados_esc.empty else pd.DataFrame()
        alunos_crit = ", ".join(crit_esc["Estudante"].astype(str).tolist()) if not crit_esc.empty else "Estudantes prioritários da unidade"
        turmas_crit = ", ".join(crit_esc["Turma"].astype(str).unique().tolist()) if not crit_esc.empty else "Turmas prioritárias"
        profs_crit = ", ".join(crit_esc["Professor"].astype(str).unique().tolist()) if not crit_esc.empty else "Docentes do componente"
        avs_crit = ", ".join(dados_esc["Avaliacao"].astype(str).unique().tolist()) if not dados_esc.empty else "Avaliações internas e externas"
        habs_crit = ", ".join(dados_esc["Habilidade_Critica"].astype(str).unique().tolist()) if not dados_esc.empty else "Habilidades essenciais da semana"
        plats_crit = ", ".join(dados_esc["Plataforma_Critica"].astype(str).unique().tolist()) if not dados_esc.empty else "Plataformas educacionais SEDUC"
        primeira_linha_legal = info_pasta["Foco_Pedagogico"].split("\n")[0]

        if pasta_chave == "MULTIPLICA":
            curs_esc = df_cursistas[df_cursistas["Escola"].str.lower() == local_plano.lower()]
            if not curs_esc.empty:
                texto_auto = f"Intencionalidade Formativa ({info_pasta['Resolucao']} - {primeira_linha_legal}): Acompanhamento presencial dos cursistas/multiplicadores de {user.get('Componente', '')} e desdobramento em ATPC."
            else:
                texto_auto = f"Intencionalidade de Mobilização e Apoio Operacional ({info_pasta['Resolucao']} - {primeira_linha_legal}): Divulgação do Programa Multiplica SP ({user.get('Componente', '')}) junto à equipe gestora e mapeamento de docentes elegíveis."
        elif pasta_chave == "DESENVOLVIMENTO CURRICULAR":
            texto_auto = (
                f"Fundamentação: {info_pasta['Resolucao']} — {primeira_linha_legal} "
                f"Análise de proficiência ({avs_crit}) na unidade ({pasta_completa}): "
                f"Priorização das turmas ({turmas_crit}) e alinhamento formativo com os docentes ({profs_crit}); "
                f"Habilidades críticas ({habs_crit}); Plataformas ({plats_crit}); "
                f"Estudantes em foco: {alunos_crit}."
            )
        else:
            texto_auto = (
                f"Fundamentação: {info_pasta['Resolucao']} — {primeira_linha_legal} "
                f"Acompanhamento de frequência, proteção escolar e indicadores na unidade ({pasta_completa}): "
                f"Turmas prioritárias ({turmas_crit}); Habilidades/Foco ({habs_crit}); Estudantes em acompanhamento: {alunos_crit}."
            )
    else:
        texto_auto = f"Atividade institucional ({natureza_dia}) conforme cronograma da Unidade Regional de Ensino de Miracatu / Agenda SEDUC."

    texto_final = st.text_area("Intencionalidade / Justificativa Oficial (editável):", value=texto_auto, height=130)

    if st.button("➕ Registrar no Planner Mensal"):
        st.session_state["compromissos_pec"].append({
            "PEC": user["Nome"], "Data": data_str,
            "Horario": f"{hora_ini} às {hora_fim}", "Natureza": natureza_dia,
            "Local": local_plano, "Publico": publico_plano, "Intencionalidade": texto_final
        })
        st.success(f"✅ Atividade de {data_str} adicionada ao seu Planner Mensal!")

    bytes_doc = montar_docx_oficial(
        f"Plano de Trabalho do Mês de {mes_ano_str}",
        str(user["Nome"]),
        f"{pasta_chave} ({user.get('Componente', 'Todos')})",
        data_str,
        info_pasta["Resolucao"],
        [{"data_hora": f"{data_str}\n{hora_ini} às {hora_fim}", "atividade_local": f"{natureza_dia}\n{local_plano}", "publico": publico_plano, "intencionalidade": texto_final}]
    )

    link_gcal = gerar_link_google_agenda(f"GPS [{pasta_chave}] — {local_plano}", data_plano, hora_ini, hora_fim, local_plano, texto_final)
    nome_arq_final = f"Plano_Trabalho_{meses_pt[data_plano.month]}_{user['Nome'].replace(' ', '_')}.docx"

    b1, b2, b3 = st.columns(3)
    with b1:
        st.download_button("📥 1. Baixar Word (.docx)", data=bytes_doc, file_name=nome_arq_final)
    with b2:
        st.link_button("📅 2. Salvar na Google Agenda", link_gcal)
    with b3:
        if st.button("📧 3. Enviar E-mail (PEC, Dirigente e CEC)", type="primary"):
            if "COLE_AQUI" in SENHA_APP_GMAIL:
                st.warning("⚠️ Cole as 16 letras da Senha de App do Gmail no topo do código (`SENHA_APP_GMAIL`) para ativar o envio automático.")
            else:
                try:
                    msg = EmailMessage()
                    msg["Subject"] = f"📋 Plano de Trabalho ({mes_ano_str}) | {local_plano} — {user['Nome']} ({pasta_chave})"
                    msg["From"] = EMAIL_DISPARADOR
                    msg["To"] = f"{user['Email']}, {EMAIL_DIRIGENTE}"
                    msg["Cc"] = EMAIL_CEC
                    msg.set_content(
                        f"Prezados(as),\n\n"
                        f"Segue em anexo o Plano de Trabalho do Mês de {mes_ano_str}.\n\n"
                        f"• Nome do PEC: {user['Nome']} ({user['Email']})\n"
                        f"• Pasta: {pasta_chave} | {user.get('Componente', 'Todos')}\n"
                        f"• Data: {data_str} ({hora_ini} às {hora_fim})\n"
                        f"• Atividade / Escola: {natureza_dia} — {local_plano}\n\n"
                        f"Intencionalidade Pedagógica:\n{texto_final}\n\n"
                        f"Atenciosamente,\n"
                        f"Sistema GPS — Guia Pedagógico Semanal (URE Miracatu)"
                    )
                    msg.add_attachment(bytes_doc, maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=nome_arq_final)
                    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                        smtp.login(EMAIL_DISPARADOR, SENHA_APP_GMAIL)
                        smtp.send_message(msg)
                    st.success(f"✅ Enviado com sucesso para {user['Email']}, Dirigente ({EMAIL_DIRIGENTE}) e cópia para CEC ({EMAIL_CEC})!")
                except Exception as e:
                    st.error(f"Erro no envio: {e}")

# ==============================================================================
# ABA 4: ROTEIRO DE VISITA, AUDITORIA PROTOCOLO 179 E REPASSE ENTRE PECS
# ==============================================================================
elif menu == "4. Roteiro de Visita, Protocolo 179 & Repasse":
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
# ABA 5: PAINEL DO CEC (CARREGAR ARQUIVOS DO DRIVE, AGENDAR E GESTÃO DE SENHAS)
# ==============================================================================
elif menu == "5. Painel CEC (Drive, Agendamentos & 1:1)":
    st.header("🎛️ Painel do CEC — Atualização Semanal, Agendamentos e Pauta 1:1")

    st.subheader("📂 1. Atualizar Dados da Semana (Arquivos Originais do Google Drive)")
    st.write("Selecione os arquivos baixados da pasta GPS do Drive (Excel, CSV ou Boletim em PDF):")
    arquivos_semana = st.file_uploader("Planilhas de Indicadores/Avaliações, Cursistas Multiplica ou Boletins (PDF):", type=["xlsx", "csv", "pdf"], accept_multiple_files=True)
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
        st.success(f"✅ {len(arquivos_semana)} arquivo(s) processado(s) e distribuído(s) conforme a atribuição de cada pasta!")

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
        st.success(f"✅ Compromisso distribuído para **{dest_ag}** em {data_ag.strftime('%d/%m/%Y')} (já visível no Planner Mensal)!")

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
                msg.set_content(f"Olá,\n\nAgendamento CEC:\nTipo: {tipo_ag}\nData: {data_ag.strftime('%d/%m/%Y')}\nHorário: {hor_ag}\nPauta: {assunto_ag}\n\nEste horário já foi inserido automaticamente no seu Planner Mensal do GPS Pedagógico.")
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
