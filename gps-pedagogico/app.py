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
import json
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
ID_PASTA_DRIVE_GPS = "1FlLFdnhzUZJeR6ujwN4LP8jldIsMSg8m"
URL_PASTA_DRIVE_GPS = f"https://drive.google.com/drive/folders/{ID_PASTA_DRIVE_GPS}?usp=sharing"

GOOGLE_DRIVE_API_KEY = st.secrets.get("GOOGLE_DRIVE_API_KEY", "COLE_AQUI_SUA_CHAVE_AIza...")
URL_GOOGLE_AGENDA_CEC = "COLE_AQUI_SEU_ENDERECO_SECRETO_ICAL.ics"
EMAIL_DISPARADOR = "seu.email.sistema@gmail.com"
SENHA_APP_GMAIL = "COLE_AQUI_AS_16_LETRAS_DO_GMAIL"
EMAIL_DIRIGENTE = "mir@educacao.sp.gov.br"
EMAIL_CEC = "mir.eec@educacao.sp.gov.br"

ESCOLAS_URE_MIRACATU = [
    "EE Alice Rodrigues Motta Professora",
    "EE Armando Goncalves Professor",
    "EE Bairro Do Engano",
    "EE Bairro Jaire",
    "EE Bairro Nova Itariri",
    "EE Bairro Pe Da Serra",
    "EE Bigua",
    "EE Clay Presgrave Do Amaral Engenheiro",
    "EE Clodonil Cardoso",
    "EE Dinora Rocha Professora",
    "EE Domingos Bauer Leite Poeta",
    "EE Elvira Silva Professora",
    "EE Itariri",
    "EE Jeremias Junior Coronel",
    "EE Joao Adorno Vassao",
    "EE Jofre Manoel Prefeito",
    "EE Jose Ermirio De Moraes Senador",
    "EE Jose Muniz Teixeira",
    "EE Judith Sant Ana Diegues Professora",
    "EE Maria Jose Morais De Carvalho Professora",
    "EE Oswaldo Florencio Professor",
    "EE Otaviano Soares Albuquerque",
    "EE Pedro Barros",
    "EE Raposo Tavares",
    "EE Santa Rita De Cassia",
    "EE Sebastiana Muniz Paiva",
    "EE Sylas Baltazar De Araujo Professor",
    "EE Tres Barras",
    "EE Ushisuke Miadaira",
    "EE Veiga Junior Professor"
]

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

# --- TRADUTOR INTELIGENTE DE PLANILHAS BRUTAS E MATRIZES DO BI SEDUC ---
MAPA_SIGLAS_BI = {
    "MAT": "Matematica",
    "PORT": "Lingua Portuguesa",
    "ING": "Lingua Inglesa",
    "HIST": "Historia",
    "GEO": "Geografia",
    "CIE": "Ciencias",
    "FILO": "Filosofia",
    "SOC": "Sociologia",
    "BIO": "Biologia",
    "FIS": "Fisica",
    "QUI": "Quimica",
    "FIN": "Educacao Financeira",
    "TEC": "Tecnologia"
}

def limpar_nome_escola_bi(nome_bruto):
    s = str(nome_bruto).strip()
    if " - " in s and s.split(" - ")[-1].strip().isdigit():
        s = " - ".join(s.split(" - ")[:-1]).strip()
    s_title = s.title()
    return s_title if s_title.upper().startswith("EE ") else f"EE {s_title}"

def padronizar_planilha_bruta(df_raw, nome_arquivo=""):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    cols_norm = {normalizar_texto(c): c for c in df_raw.columns}
    siglas_presentes = [c_orig for c_n, c_orig in cols_norm.items() if c_n in MAPA_SIGLAS_BI]

    if siglas_presentes and ("EXPORT TURMA" in cols_norm or "EXPORT ESCOLAS" in cols_norm):
        escola_detectada = "Geral URE Miracatu"
        for val in df_raw.iloc[:, 0].dropna().astype(str):
            if "NM_ESCOLA" in val:
                trecho = val.split("NM_ESCOLA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho:
                        escola_detectada = limpar_nome_escola_bi(trecho.split(sep)[1].split("CD_DIRETORIA")[0].strip())
                        break

        avaliacao_nome = "Prova Paulista — 3º Bimestre" if "PP3" in nome_arquivo.upper() else f"Avaliação BI ({nome_arquivo})"
        col_chave = cols_norm.get("EXPORT TURMA") or cols_norm.get("EXPORT ESCOLAS")
        eh_por_escola = "EXPORT ESCOLAS" in cols_norm
        col_part = cols_norm.get("(%) PARTICIPACAO")
        col_alunos = cols_norm.get("TOTAL ALUNOS")

        registros = []
        for _, row in df_raw.iterrows():
            item_chave = str(row.get(col_chave, "")).strip()
            if not item_chave or item_chave.lower() in ["nan", "total"] or "Filtros aplicados" in item_chave:
                continue

            escola_row = limpar_nome_escola_bi(item_chave) if eh_por_escola else escola_detectada
            turma_row = "Consolidado da Escola" if eh_por_escola else item_chave
            part_txt = str(row.get(col_part, "0")).replace("%", "").replace(",", ".").strip()
            try:
                part_num = float(part_txt)
            except Exception:
                part_num = 95.0

            total_al = str(row.get(col_alunos, "-")).strip()

            for col_sigla in siglas_presentes:
                val_acerto = str(row.get(col_sigla, "")).replace("%", "").replace(",", ".").strip()
                if not val_acerto or val_acerto.lower() in ["nan", "none", ""]:
                    continue
                try:
                    acerto_num = float(val_acerto)
                except Exception:
                    continue

                comp_nome = MAPA_SIGLAS_BI[normalizar_texto(col_sigla)]
                if acerto_num < 45.0:
                    nivel_str = "Abaixo do Básico"
                elif acerto_num < 60.0:
                    nivel_str = "Básico"
                else:
                    nivel_str = "Adequado"

                registros.append({
                    "Escola": escola_row,
                    "Avaliacao": avaliacao_nome,
                    "Etapa": "Ensino Médio" if "SERIE" in turma_row.upper() else ("Anos Finais" if "ANO" in turma_row.upper() else "Todas"),
                    "Turma": turma_row,
                    "Professor": "Consultar Atribuição / CGP",
                    "Estudante": f"Turma ({total_al} alunos)",
                    "Componente": comp_nome,
                    "Nivel": nivel_str,
                    "Proficiencia_Nota": f"{acerto_num:.1f}% acertos",
                    "Frequencia": round(part_num, 1),
                    "Habilidade_Critica": f"Acertos abaixo da meta ({acerto_num:.1f}%)" if acerto_num < 50 else "Dentro do esperado",
                    "Plataforma_Critica": f"Participação Prova: {part_num:.1f}%"
                })
        return pd.DataFrame(registros)

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
        ("Escola", "EE Armando Goncalves Professor"),
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

@st.cache_data(ttl=300)
def sincronizar_pasta_drive_automatico(folder_id, api_key):
    if not api_key or "COLE_AQUI" in api_key:
        return pd.DataFrame(), pd.DataFrame(), "", 0
    dfs_ind, dfs_curs, textos_pdf = [], [], []
    qtd_arquivos = 0

    def varrer_pasta(fid):
        nonlocal qtd_arquivos
        q = urllib.parse.quote(f"'{fid}' in parents and trashed = false")
        url_api = f"https://www.googleapis.com/drive/v3/files?q={q}&fields=files(id,name,mimeType)&key={api_key}"
        try:
            resp = urllib.request.urlopen(url_api, timeout=15)
            dados = json.loads(resp.read().decode("utf-8"))
        except Exception:
            return

        for f in dados.get("files", []):
            f_id, f_name, f_mime = f["id"], f["name"], f["mimeType"]
            if f_mime == "application/vnd.google-apps.folder":
                varrer_pasta(f_id)
            else:
                try:
                    if f_mime == "application/vnd.google-apps.spreadsheet":
                        dl_url = f"https://docs.google.com/spreadsheets/d/{f_id}/export?format=xlsx"
                    elif f_mime == "application/vnd.google-apps.document":
                        dl_url = f"https://docs.google.com/document/d/{f_id}/export?format=pdf"
                    else:
                        dl_url = f"https://www.googleapis.com/drive/v3/files/{f_id}?alt=media&key={api_key}"

                    arq_bytes = BytesIO(urllib.request.urlopen(dl_url, timeout=20).read())
                    n_low = f_name.lower()

                    if n_low.endswith(".pdf") or f_mime == "application/pdf":
                        leitor = PdfReader(arq_bytes)
                        textos_pdf.append(f"**[{f_name}]**\n" + "\n".join([p.extract_text() or "" for p in leitor.pages]))
                        qtd_arquivos += 1
                    elif n_low.endswith(".csv"):
                        df_l = pd.read_csv(arq_bytes)
                        df_p = padronizar_planilha_bruta(df_l, f_name)
                        if not df_p.empty:
                            (dfs_curs if "multiplica" in n_low or "cursista" in n_low else dfs_ind).append(df_p)
                            qtd_arquivos += 1
                    elif n_low.endswith(".xlsx") or "spreadsheet" in f_mime:
                        df_l = pd.read_excel(arq_bytes)
                        df_p = padronizar_planilha_bruta(df_l, f_name)
                        if not df_p.empty:
                            (dfs_curs if "multiplica" in n_low or "cursista" in n_low else dfs_ind).append(df_p)
                            qtd_arquivos += 1
                except Exception:
                    pass

    varrer_pasta(folder_id)
    df_ind_final = pd.concat(dfs_ind, ignore_index=True) if dfs_ind else pd.DataFrame()
    df_curs_final = pd.concat(dfs_curs, ignore_index=True) if dfs_curs else pd.DataFrame()
    return df_ind_final, df_curs_final, "\n\n".join(textos_pdf), qtd_arquivos

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

# ==============================================================================
# FUNÇÕES DE EXPORTAÇÃO (.DOCX) E DISPARO DE E-MAIL
# ==============================================================================
def gerar_docx_plano(nome_pec, email_pec, pasta_pec, comp_pec, escola_sel, semana_str, cronograma_dias, obs_texto):
    doc = Document()
    titulo = doc.add_heading("GPS — GUIA PEDAGÓGICO SEMANAL", level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("Unidade Regional de Ensino de Miracatu — Equipe de Especialistas em Currículo (EEC)")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_info = doc.add_paragraph()
    p_info.add_run("Especialista / Responsável: ").bold = True
    p_info.add_run(f"{nome_pec} ({email_pec})\n")
    p_info.add_run("Pasta / Eixo de Atuação: ").bold = True
    p_info.add_run(f"{pasta_pec} — Componente: {comp_pec}\n")
    p_info.add_run("Unidade Escolar Prioritária: ").bold = True
    p_info.add_run(f"{escola_sel}\n")
    p_info.add_run("Período / Semana: ").bold = True
    p_info.add_run(f"{semana_str}\n")

    doc.add_heading("1. Cronograma Semanal de Ações e Visitas", level=2)
    tabela = doc.add_table(rows=1, cols=4)
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Dia / Turno"
    hdr[1].text = "Local / Escola"
    hdr[2].text = "Ação Pedagógica (Resolução SEDUC 01/2026)"
    hdr[3].text = "Público-Alvo / Encaminhamento"

    for item in cronograma_dias:
        row_cells = tabela.add_row().cells
        row_cells[0].text = item.get("dia", "")
        row_cells[1].text = item.get("local", "")
        row_cells[2].text = item.get("acao", "")
        row_cells[3].text = item.get("publico", "")

    if obs_texto.strip():
        doc.add_heading("2. Pautas Formativas, Diagnóstico e Devolutivas", level=2)
        doc.add_paragraph(obs_texto.strip())

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def enviar_email_gps(destinatarios, assunto, corpo_texto, anexo_bytes, nome_anexo):
    if "COLE_AQUI" in SENHA_APP_GMAIL or "seu.email" in EMAIL_DISPARADOR:
        return False, "Configure EMAIL_DISPARADOR e SENHA_APP_GMAIL no código ou Secrets para habilitar o envio automático."
    try:
        msg = EmailMessage()
        msg["Subject"] = assunto
        msg["From"] = EMAIL_DISPARADOR
        msg["To"] = ", ".join(destinatarios)
        msg.set_content(corpo_texto)
        msg.add_attachment(
            anexo_bytes.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=nome_anexo
        )
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_DISPARADOR, SENHA_APP_GMAIL)
            smtp.send_message(msg)
        return True, "Plano enviado com sucesso!"
    except Exception as e:
        return False, f"Falha no envio do e-mail: {e}"

# ==============================================================================
# AUTENTICAÇÃO E SESSÃO
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario_nome"] = ""
    st.session_state["usuario_email"] = ""
    st.session_state["usuario_pasta"] = "QUALIDADE DA AULA"
    st.session_state["usuario_componente"] = "Todos"
    st.session_state["cronograma_semana"] = []

if not st.session_state["autenticado"]:
    st.title("🧭 GPS — Guia Pedagógico Semanal")
    st.subheader("Unidade Regional de Ensino de Miracatu — SEDUC-SP")
    st.markdown("---")

    col_esq, col_centro, col_dir = st.columns([1, 2, 1])
    with col_centro:
        with st.form("form_login_gps"):
            st.markdown("### Acesso da Equipe de Especialistas (EEC / CEC / AT)")
            email_in = st.text_input("E-mail Institucional (@educacao.sp.gov.br ou @prof.educacao.sp.gov.br)")
            senha_in = st.text_input("Senha / Código de Acesso", type="password")
            pasta_fallback = st.selectbox(
                "Pasta / Função (usado se não estiver cadastrado na aba Usuarios):",
                list(TAREFAS_POR_PASTA.keys())
            )
            entrar = st.form_submit_button("Entrar no Sistema GPS", use_container_width=True)

            if entrar:
                if not email_in.strip():
                    st.error("Por favor, informe seu e-mail institucional.")
                else:
                    nome_encontrado = email_in.split("@")[0].replace(".", " ").title()
                    pasta_encontrada = pasta_fallback
                    comp_encontrado = "Todos"
                    acesso_ok = True

                    if not df_usuarios.empty:
                        cols_u = {normalizar_texto(c): c for c in df_usuarios.columns}
                        col_email = next((v for k, v in cols_u.items() if "EMAIL" in k or "E-MAIL" in k), None)
                        col_nome = next((v for k, v in cols_u.items() if "NOME" in k), None)
                        col_pasta = next((v for k, v in cols_u.items() if "PASTA" in k or "FUNCAO" in k), None)
                        col_comp = next((v for k, v in cols_u.items() if "COMPONENTE" in k), None)
                        col_senha = next((v for k, v in cols_u.items() if "SENHA" in k or "PIN" in k), None)

                        if col_email:
                            match = df_usuarios[df_usuarios[col_email].astype(str).str.strip().str.lower() == email_in.strip().lower()]
                            if not match.empty:
                                row_u = match.iloc[0]
                                if col_senha and pd.notna(row_u[col_senha]) and str(row_u[col_senha]).strip() != "":
                                    if senha_in.strip() != str(row_u[col_senha]).strip():
                                        acesso_ok = False
                                        st.error("Senha incorreta para este usuário.")
                                if acesso_ok:
                                    if col_nome and pd.notna(row_u[col_nome]):
                                        nome_encontrado = str(row_u[col_nome]).strip()
                                    if col_pasta and pd.notna(row_u[col_pasta]):
                                        pasta_encontrada = mapear_pasta_chave(row_u[col_pasta]) or pasta_fallback
                                    if col_comp and pd.notna(row_u[col_comp]):
                                        comp_encontrado = str(row_u[col_comp]).strip()

                    if acesso_ok:
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_nome"] = nome_encontrado
                        st.session_state["usuario_email"] = email_in.strip()
                        st.session_state["usuario_pasta"] = pasta_encontrada if pasta_encontrada in TAREFAS_POR_PASTA else pasta_fallback
                        st.session_state["usuario_componente"] = comp_encontrado
                        st.rerun()
    st.stop()

# ==============================================================================
# PAINEL PRINCIPAL DO SISTEMA GPS
# ==============================================================================
with st.sidebar:
    st.markdown("## 🧭 GPS — URE Miracatu")
    st.write(f"**Especialista:** {st.session_state['usuario_nome']}")
    st.write(f"**Pasta:** {st.session_state['usuario_pasta']}")
    st.write(f"**Componente:** {st.session_state['usuario_componente']}")
    st.markdown("---")
    if st.button("🔄 Atualizar Dados das Planilhas", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    if st.button("🚪 Sair / Trocar Usuário", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

st.title("🧭 GPS — Guia Pedagógico Semanal")
st.caption("Planejamento, Acompanhamento Escolar e Alinhamento Regional — URE de Miracatu")

aba_plano, aba_agenda, aba_indicadores, aba_normativas = st.tabs([
    "📋 Plano Semanal & Exportação",
    "📅 Agenda SEDUC, CEC & Feriados",
    "📊 Indicadores & Google Drive",
    "📚 Diretrizes das Pastas (Res. SEDUC 01/2026)"
])

# --- ABA 1: PLANO SEMANAL ---
with aba_plano:
    st.subheader("Elaboração do Guia Pedagógico Semanal")
    col1, col2 = st.columns(2)
    with col1:
        escola_prioritaria = st.selectbox("Unidade Escolar Prioritária da Semana:", ESCOLAS_URE_MIRACATU)
    with col2:
        semana_ref = st.text_input("Período / Semana de Referência:", value=datetime.now().strftime("Semana de %d/%m/%Y"))

    pasta_atual = st.session_state["usuario_pasta"]
    tarefas_pasta = TAREFAS_POR_PASTA.get(pasta_atual, [])

    st.markdown("#### Adicionar Atividade ao Cronograma da Semana")
    with st.form("form_add_atividade", clear_on_submit=True):
        c_dia, c_loc = st.columns(2)
        with c_dia:
            dia_turno = st.selectbox("Dia da Semana / Turno:", [
                "Segunda-feira (Manhã)", "Segunda-feira (Tarde)",
                "Terça-feira (Manhã)", "Terça-feira (Tarde)",
                "Quarta-feira (Manhã)", "Quarta-feira (Tarde)",
                "Quinta-feira (Manhã)", "Quinta-feira (Tarde)",
                "Sexta-feira (Manhã)", "Sexta-feira (Tarde)"
            ])
        with c_loc:
            local_acao = st.selectbox("Local / Escola:", ["Sede URE Miracatu / Trabalho Interno", "Reunião Virtual / EFAPE / SEDUC"] + ESCOLAS_URE_MIRACATU)

        acao_escolhida = st.selectbox("Ação Pedagógica da Pasta:", tarefas_pasta + ["Outra ação específica (descrever abaixo)"])
        publico_enc = st.text_input("Público-Alvo / Detalhamento da Ação:", placeholder="Ex.: CGP/CGPG, Professores de Matemática, Estudantes Abaixo do Básico...")
        add_btn = st.form_submit_button("➕ Inserir no Cronograma")

        if add_btn:
            st.session_state["cronograma_semana"].append({
                "dia": dia_turno,
                "local": local_acao,
                "acao": acao_escolhida,
                "publico": publico_enc if publico_enc.strip() else "Equipe Gestora / Docentes"
            })
            st.success("Atividade adicionada ao cronograma!")

    if st.session_state["cronograma_semana"]:
        df_cron = pd.DataFrame(st.session_state["cronograma_semana"])
        st.dataframe(df_cron, use_container_width=True)
        if st.button("🗑️ Limpar Cronograma da Semana"):
            st.session_state["cronograma_semana"] = []
            st.rerun()
    else:
        st.info("Nenhuma atividade inserida ainda. Use o formulário acima para montar os dias da semana.")

    obs_texto = st.text_area("Pautas Formativas, Diagnóstico de Indicadores e Encaminhamentos:", height=120)

    st.markdown("---")
    docx_buf = gerar_docx_plano(
        st.session_state["usuario_nome"],
        st.session_state["usuario_email"],
        pasta_atual,
        st.session_state["usuario_componente"],
        escola_prioritaria,
        semana_ref,
        st.session_state["cronograma_semana"],
        obs_texto
    )
    nome_arq_docx = f"GPS_{pasta_atual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx"

    col_dl, col_em = st.columns(2)
    with col_dl:
        st.download_button(
            label="⬇️ Baixar Guia Pedagógico Semanal (.docx)",
            data=docx_buf,
            file_name=nome_arq_docx,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col_em:
        if st.button("📧 Enviar Plano por E-mail para CEC / Dirigente", use_container_width=True):
            ok, msg_envio = enviar_email_gps(
                [EMAIL_CEC, EMAIL_DIRIGENTE],
                f"[GPS URE Miracatu] {pasta_atual} — {st.session_state['usuario_nome']} ({semana_ref})",
                f"Segue em anexo o Guia Pedagógico Semanal de {st.session_state['usuario_nome']} ({pasta_atual}).",
                docx_buf,
                nome_arq_docx
            )
            if ok:
                st.success(msg_envio)
            else:
                st.warning(msg_envio)

# --- ABA 2: AGENDA SEDUC, CEC & FERIADOS ---
with aba_agenda:
    st.subheader("Agenda Integrada — SEDUC/EEC, Google Agenda e Bloqueios")

    col_m, col_a = st.columns(2)
    with col_m:
        mes_sel = st.selectbox("Mês:", list(range(1, 13)), index=datetime.now().month - 1)
    with col_a:
        ano_sel = st.number_input("Ano:", min_value=2025, max_value=2030, value=datetime.now().year)

    feriados_sp = holidays.country_holidays("BR", subdiv="SP", years=[ano_sel])
    feriados_mes = [{"Data": d.strftime("%d/%m/%Y"), "Feriado / Ponto Facultativo": nome} for d, nome in sorted(feriados_sp.items()) if d.month == mes_sel]
    if feriados_mes:
        st.markdown("**Feriados Oficiais (SP) no Mês:**")
        st.dataframe(pd.DataFrame(feriados_mes), use_container_width=True)

    if not df_bloqueios.empty:
        st.markdown("**Bloqueios de Agenda / Convocações URE:**")
        st.dataframe(df_bloqueios, use_container_width=True)

    eventos_gcal = carregar_eventos_gcal(URL_GOOGLE_AGENDA_CEC)
    eventos_gcal_mes = [e for e in eventos_gcal if e["ano"] == ano_sel and e["mes"] == mes_sel]
    if eventos_gcal_mes:
        st.markdown("**Compromissos Google Agenda CEC:**")
        st.dataframe(pd.DataFrame(eventos_gcal_mes)[["data_str", "horario", "resumo", "local"]], use_container_width=True)

    st.markdown("**Agenda SEDUC / EEC (Filtrada para sua Pasta):**")
    if df_agenda_seduc.empty:
        st.info("Nenhum registro carregado da planilha da Agenda SEDUC.")
    else:
        st.dataframe(df_agenda_seduc, use_container_width=True)

# --- ABA 3: INDICADORES & GOOGLE DRIVE ---
with aba_indicadores:
    st.subheader("Análise de Indicadores e Matrizes BI SEDUC (Google Drive)")
    st.markdown(f"📂 [Abrir Pasta Oficial do Drive GPS]({URL_PASTA_DRIVE_GPS})")

    col_drv, col_up = st.columns(2)
    with col_drv:
        st.markdown("#### 1. Sincronização Automática (Google Drive)")
        if "COLE_AQUI" in GOOGLE_DRIVE_API_KEY:
            st.warning("Chave `GOOGLE_DRIVE_API_KEY` ainda não preenchida. Use o upload manual ao lado ou configure a chave no código.")
        else:
            df_ind, df_curs, textos_pdf, qtd = sincronizar_pasta_drive_automatico(ID_PASTA_DRIVE_GPS, GOOGLE_DRIVE_API_KEY)
            st.success(f"{qtd} arquivo(s) lido(s) automaticamente da pasta do Drive.")
            if not df_ind.empty:
                escola_filtro = st.selectbox("Filtrar Indicadores por Escola:", ["Todas"] + sorted([str(x) for x in df_ind["Escola"].dropna().unique() if str(x).strip() != ""]))
                df_exibir = df_ind if escola_filtro == "Todas" else df_ind[df_ind["Escola"].astype(str) == escola_filtro]
                st.dataframe(df_exibir, use_container_width=True)

    with col_up:
        st.markdown("#### 2. Leitura Rápida de Planilha do BI (.xlsx / .csv)")
        arquivo_bi = st.file_uploader("Envie uma planilha exportada do BI SEDUC (Prova Paulista, Frequência, etc.):", type=["xlsx", "csv"])
        if arquivo_bi is not None:
            df_raw_up = pd.read_csv(arquivo_bi) if arquivo_bi.name.endswith(".csv") else pd.read_excel(arquivo_bi)
            df_pad_up = padronizar_planilha_bruta(df_raw_up, arquivo_bi.name)
            if not df_pad_up.empty:
                st.success(f"Planilha padronizada com {len(df_pad_up)} registros!")
                st.dataframe(df_pad_up, use_container_width=True)
            else:
                st.dataframe(df_raw_up, use_container_width=True)

# --- ABA 4: DIRETRIZES DAS PASTAS ---
with aba_normativas:
    st.subheader("Base Legal e Foco Pedagógico por Pasta (Resolução SEDUC Nº 01/2026)")
    st.dataframe(df_pastas, use_container_width=True)
