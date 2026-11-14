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

GOOGLE_DRIVE_API_KEY = "COLE_AQUI_SUA_CHAVE_AIza..."
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

    # Formato matriz do BI da SEDUC (GERAL PP3.xlsx e planilhas por escola)
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

    # Formato planilha nominal padrão (estudante por linha)
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

# --- LEITURA AUTOMÁTICA DAS ABAS DO GOOGLE SHEETS, AGENDAS E DRIVE ---
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

# --- MEMÓRIA DE SESSÃO E SINCRONIZAÇÃO AUTOMÁTICA DO DRIVE ---
if "senhas_atualizadas" not in st.session_state:
    st.session_state["senhas_atualizadas"] = {}

df_drive_ind, df_drive_curs, txt_drive_pdf, qtd_drive_sync = sincronizar_pasta_drive_automatico(ID_PASTA_DRIVE_GPS, GOOGLE_DRIVE_API_KEY)

if "indicadores_ativos" not in st.session_state:
    if not df_drive_ind.empty:
        st.session_state["indicadores_ativos"] = df_drive_ind
    else:
        st.session_state["indicadores_ativos"] = pd.DataFrame([
            {"Escola": "EE Armando Goncalves Professor", "Avaliacao": "Prova Paulista — 3º Bimestre", "Etapa": "Ensino Médio", "Turma": "1ª SERIE B INTEGRAL", "Professor": "Consultar Atribuição / CGP", "Estudante": "Turma (19 alunos)", "Componente": "Matematica", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "24.4% acertos", "Frequencia": 94.7, "Habilidade_Critica": "Acertos abaixo da meta (24.4%)", "Plataforma_Critica": "Participação Prova: 94.7%"},
            {"Escola": "EE Bairro Do Engano", "Avaliacao": "Prova Paulista — 3º Bimestre", "Etapa": "Ensino Médio", "Turma": "1ª SERIE A MANHA", "Professor": "Consultar Atribuição / CGP", "Estudante": "Turma (14 alunos)", "Componente": "Fisica", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "17.9% acertos", "Frequencia": 100.0, "Habilidade_Critica": "Acertos abaixo da meta (17.9%)", "Plataforma_Critica": "Participação Prova: 100.0%"},
            {"Escola": "EE Bairro Nova Itariri", "Avaliacao": "Prova Paulista — 3º Bimestre", "Etapa": "Ensino Médio", "Turma": "1ª SERIE A MANHA", "Professor": "Consultar Atribuição / CGP", "Estudante": "Turma (21 alunos)", "Componente": "Quimica", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "13.8% acertos", "Frequencia": 92.9, "Habilidade_Critica": "Acertos abaixo da meta (13.8%)", "Plataforma_Critica": "Participação Prova: 92.9%"},
            {"Escola": "EE Alice Rodrigues Motta Professora", "Avaliacao": "Prova Paulista — 3º Bimestre", "Etapa": "Anos Finais", "Turma": "7º ANO B INTEGRAL", "Professor": "Consultar Atribuição / CGP", "Estudante": "Turma (22 alunos)", "Componente": "Lingua Inglesa", "Nivel": "Abaixo do Básico", "Proficiencia_Nota": "32.7% acertos", "Frequencia": 100.0, "Habilidade_Critica": "Acertos abaixo da meta (32.7%)", "Plataforma_Critica": "Participação Prova: 100.0%"},
        ])

if "cursistas_ativos" not in st.session_state:
    if not df_drive_curs.empty:
        st.session_state["cursistas_ativos"] = df_drive_curs
    else:
        st.session_state["cursistas_ativos"] = pd.DataFrame([
            {"Escola": "EE Raposo Tavares", "Professor": "Prof. Roberto (Cursista)", "Componente": "Artes", "Etapa": "Ensino Médio", "Status": "Ativo - Baixa entrega"},
        ])

if "boletins_pdf_texto" not in st.session_state:
    st.session_state["boletins_pdf_texto"] = txt_drive_pdf

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

def exibir_norteador_planejamento(pasta_pec, componente_pec, escola_escolhida, df_ind, df_curs, escolas_lista):
    with st.expander(f"💡 Sugestão Pedagógica Baseada nos Indicadores — {pasta_pec} (Opcional)", expanded=True):
        st.caption("Sugestão orientativa para apoiar sua decisão. Você tem autonomia para priorizar outras unidades/turmas conforme necessidade da rede, o que será dialogado posteriormente no alinhamento 1:1.")
        comp_norm = normalizar_texto(componente_pec)

        if pasta_pec == "DESENVOLVIMENTO CURRICULAR" and not df_ind.empty:
            df_c = df_ind[df_ind["Escola"].apply(normalizar_texto).isin([normalizar_texto(e) for e in escolas_lista])].copy()
            if comp_norm not in ["", "TODOS", "NAN", "NONE"]:
                df_f = df_c[df_c["Componente"].apply(normalizar_texto).str.contains(comp_norm, na=False)]
                if not df_f.empty:
                    df_c = df_f
            df_c["_critico"] = df_c["Nivel"].apply(lambda x: 1 if any(k in normalizar_texto(x) for k in ["ABAIXO", "INSUFICIENTE", "DEFASAGEM"]) else 0)
            rank_dc = (
                df_c.groupby(["Escola", "Avaliacao", "Turma", "Professor", "Proficiencia_Nota"], as_index=False)
                .agg(Alertas_Criticos=("_critico", "sum"))
                .sort_values(by=["Alertas_Criticos", "Proficiencia_Nota"], ascending=[False, True])
            )
            st.dataframe(rank_dc.head(10), use_container_width=True)
            return

        if normalizar_texto(escola_escolhida) not in [normalizar_texto(e) for e in escolas_lista]:
            return

        if pasta_pec == "MULTIPLICA":
            curs_esc = df_curs[df_curs["Escola"].apply(normalizar_texto) == normalizar_texto(escola_escolhida)]
            if not curs_esc.empty:
                st.write(f"**Sugestão de Acompanhamento Formativo em {escola_escolhida}:**")
                st.dataframe(curs_esc, use_container_width=True)
            else:
                st.write(f"**Sugestão de Mobilização em {escola_escolhida}:** Unidade sem cursistas ativos mapeados neste componente.")
            return

        dados_u = df_ind[df_ind["Escola"].apply(normalizar_texto) == normalizar_texto(escola_escolhida)].copy()
        if dados_u.empty:
            st.info("Sem indicadores detalhados por turma carregados para esta unidade ainda (consta apenas o consolidado geral ou aguarda upload).")
            return

        dados_u["_critico"] = dados_u["Nivel"].apply(lambda x: 1 if any(k in normalizar_texto(x) for k in ["ABAIXO", "INSUFICIENTE", "DEFASAGEM"]) else 0)
        dados_u["_freq_num"] = pd.to_numeric(dados_u["Frequencia"], errors="coerce")

        if pasta_pec == "QUALIDADE DA AULA":
            rank_qa = (
                dados_u.groupby(["Turma", "Componente", "Professor", "Proficiencia_Nota"], as_index=False)
                .agg(Abaixo_Basico=("_critico", "sum"), Freq_Participacao=("_freq_num", "mean"))
                .sort_values(by=["Abaixo_Basico", "Freq_Participacao"], ascending=[False, True])
            )
            rank_qa["Freq_Participacao"] = rank_qa["Freq_Participacao"].round(1).astype(str) + "%"
            st.dataframe(rank_qa, use_container_width=True)
        else:
            rank_geral = (
                dados_u.groupby(["Turma", "Componente", "Professor", "Proficiencia_Nota"], as_index=False)
                .agg(Abaixo_Basico=("_critico", "sum"), Freq_Participacao=("_freq_num", "mean"))
                .sort_values(by=["Freq_Participacao", "Abaixo_Basico"], ascending=[True, False])
            )
            rank_geral["Freq_Participacao"] = rank_geral["Freq_Participacao"].round(1).astype(str) + "%"
            st.dataframe(rank_geral, use_container_width=True)

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

    col1, _ = st.columns([1, 2])
    with col1:
        email_in = st.text_input("E-mail Institucional:")
        senha_in = st.text_input("Senha:", type="password")
        if st.button("Entrar no Sistema", type="primary"):
            if df_usuarios.empty:
                st.error("Erro ao ler a aba 'Usuarios'. Verifique a conexão.")
            else:
                em_clean = email_in.strip().lower()
                sn_clean = senha_in.strip()
                user_match = df_usuarios[df_usuarios["Email"].str.strip().str.lower() == em_clean]
                if not user_match.empty:
                    u_row = user_match.iloc[0].to_dict()
                    em_key = str(u_row.get("Email", "")).strip()
                    senha_planilha = str(u_row.get("Senha", "")).strip()
                    senha_efetiva = st.session_state["senhas_atualizadas"].get(em_key, senha_planilha)

                    if sn_clean == senha_efetiva:
                        if sn_clean == "123":
                            st.session_state["usuario_troca_senha"] = u_row
                            st.rerun()
                        else:
                            st.session_state["user"] = u_row
                            st.rerun()
                    else:
                        st.error("Senha incorreta.")
                else:
                    st.error("E-mail não cadastrado na aba 'Usuarios'.")
    st.stop()

# --- PAINEL PRINCIPAL ---
user = st.session_state["user"]
nome_user = str(user.get("Nome", "Especialista"))
cargo_user = str(user.get("Cargo", "PEC")).strip().upper()
pasta_bruta_user = str(user.get("Pasta", "GERAL"))
pasta_user = mapear_pasta_chave(pasta_bruta_user)
comp_user = str(user.get("Componente", "Todos")).strip()
escolas_raw = str(user.get("Escolas_Atribuidas", "TODAS")).strip()

if normalizar_texto(escolas_raw) == "TODAS" or cargo_user in ["CEC", "AT"] or pasta_user == "DESENVOLVIMENTO CURRICULAR":
    escolas_pec = ESCOLAS_URE_MIRACATU
else:
    lista_bruta = [e.strip() for e in escolas_raw.split(",") if e.strip()]
    escolas_pec = []
    for e_b in lista_bruta:
        encontrada = next((of for of in ESCOLAS_URE_MIRACATU if normalizar_texto(e_b) in normalizar_texto(of)), e_b)
        escolas_pec.append(encontrada)
    if not escolas_pec:
        escolas_pec = ESCOLAS_URE_MIRACATU

info_pasta = df_pastas[df_pastas["Pasta_Chave"] == pasta_user]
if not info_pasta.empty:
    resolucao_atual = info_pasta.iloc[0]["Resolucao"]
    foco_atual = info_pasta.iloc[0]["Foco_Pedagogico"]
else:
    resolucao_atual = "RESOLUÇÃO SEDUC Nº 1, DE 6 DE JANEIRO DE 2026"
    foco_atual = "Atuação Técnica e Pedagógica conforme diretrizes da URE Miracatu."

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown(f"### 👤 {nome_user}")
    st.markdown(f"**Perfil:** `{cargo_user}` | **Pasta:** `{pasta_user}`")
    if comp_user and normalizar_texto(comp_user) not in ["", "TODOS", "NAN", "NONE"]:
        st.markdown(f"**Componente:** `{comp_user}`")
    st.caption(f"📜 {resolucao_atual}")
    st.divider()

    st.markdown("### 📂 Fontes de Dados (Drive URE)")
    if qtd_drive_sync > 0:
        st.success(f"🔄 Drive Sincronizado ({qtd_drive_sync} arquivo(s))")
    else:
        st.info("📂 Pasta GPS conectada (aguardando API Key no código ou novos arquivos)")
    st.link_button("📂 Abrir Pasta GPS no Drive", URL_PASTA_DRIVE_GPS, use_container_width=True)

    with st.expander("➕ Upload Manual Avulso (Opcional)", expanded=False):
        st.caption("Use apenas se quiser testar um arquivo local sem subir no Drive.")
        arquivos_up = st.file_uploader("Subir PDF, XLSX ou CSV:", type=["pdf", "xlsx", "csv"], accept_multiple_files=True)
        if arquivos_up and st.button("Processar Arquivos Locais"):
            novos_ind, novos_curs, textos_pdf = [], [], []
            for arq in arquivos_up:
                n_low = arq.name.lower()
                if n_low.endswith(".pdf"):
                    leitor = PdfReader(arq)
                    textos_pdf.append(f"**[{arq.name}]**\n" + "\n".join([p.extract_text() or "" for p in leitor.pages]))
                elif n_low.endswith(".csv"):
                    df_p = padronizar_planilha_bruta(pd.read_csv(arq), arq.name)
                    (novos_curs if "multiplica" in n_low or "cursista" in n_low else novos_ind).append(df_p)
                elif n_low.endswith(".xlsx"):
                    df_p = padronizar_planilha_bruta(pd.read_excel(arq), arq.name)
                    (novos_curs if "multiplica" in n_low or "cursista" in n_low else novos_ind).append(df_p)
            if novos_ind:
                st.session_state["indicadores_ativos"] = pd.concat([st.session_state["indicadores_ativos"]] + novos_ind, ignore_index=True)
            if novos_curs:
                st.session_state["cursistas_ativos"] = pd.concat(novos_curs, ignore_index=True)
            if textos_pdf:
                st.session_state["boletins_pdf_texto"] += "\n\n" + "\n\n".join(textos_pdf)
            st.success("Arquivos locais integrados!")
            st.rerun()

    st.divider()
    if st.button("Sair"):
        del st.session_state["user"]
        st.rerun()

st.title("🧭 GPS — Guia Pedagógico Semanal | URE Miracatu")

# --- ABAS DO SISTEMA ---
abas_nomes = [
    "📅 1. Calendário & Compromissos",
    "📋 2. Plano de Trabalho (Mensal)",
    "🚌 3. Plano de Visitas (Semanal)",
    "📊 4. Indicadores & Boletins",
    "🔄 5. Encaminhamentos"
]
if
