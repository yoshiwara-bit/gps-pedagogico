import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import smtplib
from email.message import EmailMessage
from datetime import datetime, date
import calendar
from io import BytesIO
import urllib.parse
import urllib.request
import unicodedata
import os
import re
import json
import holidays
from icalendar import Calendar as ICalCalendar

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="GPS — Guia Pedagógico Semanal", page_icon="🧭", layout="wide")

# ==============================================================================
# CONFIGURAÇÕES DA URE DE MIRACATU (MANTIDAS DO SISTEMA ORIGINAL)
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

# ==============================================================================
# MAPA DE ATUAÇÃO POR PASTA / ÁREA (COMPONENTES + PLATAFORMAS DIGITAIS)
# ==============================================================================
MAPA_ATUACAO_COMPONENTES_PLATAFORMAS = {
    "CIENCIAS DA NATUREZA": {
        "componentes": ["CIENCIAS", "BIOLOGIA", "FISICA", "QUIMICA", "INVESTIGACAO CIENTIFICA", "PRATICAS EXPERIMENTAIS"],
        "plataformas": ["ALURA", "PLATAFORMA DE CIENCIAS"]
    },
    "CIENCIAS HUMANAS": {
        "componentes": ["HISTORIA", "GEOGRAFIA", "SOCIOLOGIA", "FILOSOFIA", "ESTUDOS ORIENTADOS", "GEOPOLITICA"],
        "plataformas": ["PLATAFORMA HUMANAS"]
    },
    "MATEMATICA": {
        "componentes": ["MATEMATICA", "FUNDAMENTOS DE MATEMATICA", "EDUCACAO FINANCEIRA", "PENSAMENTO COMPUTACIONAL", "MATEMATICA BASICA"],
        "plataformas": ["KHAN ACADEMY", "MATIFIC", "ALURA"]
    },
    "LINGUA PORTUGUESA": {
        "componentes": ["LINGUA PORTUGUESA", "REDACAO E LEITURA", "FUNDAMENTOS DE LINGUA PORTUGUESA", "PRODUCAO DE TEXTO", "ORATORIA", "LEITURA"],
        "plataformas": ["LEIA SP", "REDACAO PAULISTA"]
    },
    "LINGUAGENS": {
        "componentes": ["LINGUA INGLESA", "INGLES", "EDUCACAO FISICA", "ARTE", "LINGUA ESTRANGEIRA"],
        "plataformas": ["SPEAK", "OPEN ENGLISH", "LEIA SP"]
    },
    "RECOMPOSICAO": {
        "componentes": ["LINGUA PORTUGUESA", "MATEMATICA", "REDACAO E LEITURA", "FUNDAMENTOS DE LINGUA PORTUGUESA", "FUNDAMENTOS DE MATEMATICA"],
        "plataformas": ["KHAN ACADEMY", "MATIFIC", "LEIA SP", "REDACAO PAULISTA"]
    }
}

def obter_regras_atuacao_pec(pasta_pec, componente_pec):
    p_norm = normalizar_texto(pasta_pec)
    c_norm = normalizar_texto(componente_pec)

    if p_norm in ["CEC", "AT"] or c_norm in ["", "TODOS"]:
        if p_norm == "RECOMPOSICAO":
            return MAPA_ATUACAO_COMPONENTES_PLATAFORMAS["RECOMPOSICAO"]
        return None  # Sem filtro restritivo (CEC/AT enxergam tudo)

    for chave_area, regras in MAPA_ATUACAO_COMPONENTES_PLATAFORMAS.items():
        if chave_area in c_norm or chave_area in p_norm:
            return regras
        if any(comp in c_norm for comp in regras["componentes"]):
            return regras

    return {"componentes": [c_norm], "plataformas": ["SPEAK", "OPEN ENGLISH", "KHAN ACADEMY", "MATIFIC", "LEIA SP", "REDACAO PAULISTA", "ALURA"]}

def filtrar_indicadores_por_pec(df_ind, pasta_pec, componente_pec, escolas_pec=None):
    if df_ind is None or df_ind.empty:
        return pd.DataFrame()
    df_f = df_ind.copy()

    # Filtra escolas vinculadas ao PEC (se houver lista de escolas atribuídas ao PEC na aba Usuarios)
    if escolas_pec and "Escola" in df_f.columns:
        esc_norm = [normalizar_texto(e) for e in escolas_pec if str(e).strip()]
        if esc_norm:
            df_f = df_f[df_f["Escola"].apply(lambda x: any(e in normalizar_texto(x) or normalizar_texto(x) in e for e in esc_norm))]

    regras = obter_regras_atuacao_pec(pasta_pec, componente_pec)
    if regras is None:
        return df_f

    comps_alvo = regras["componentes"]
    plats_alvo = regras["plataformas"]

    def linha_pertence(row):
        tipo_reg = normalizar_texto(row.get("Tipo_Indicador", "AVALIACAO"))
        if tipo_reg == "FREQUENCIA":
            return True
        if tipo_reg == "PLATAFORMA":
            plat = normalizar_texto(row.get("Plataforma_Critica", "")) + " " + normalizar_texto(row.get("Componente", ""))
            return any(p in plat for p in plats_alvo)
        comp = normalizar_texto(row.get("Componente", ""))
        if comp in ["", "TODOS", "GERAL"]:
            return True
        return any(c in comp or comp in c for c in comps_alvo)

    return df_f[df_f.apply(linha_pertence, axis=1)].copy()

# ==============================================================================
# CONTROLE DE PRAZO: 2º DIA ÚTIL DO MÊS (PLANO DE TRABALHO)
# ==============================================================================
def obter_segundo_dia_util(ano, mes):
    feriados_sp = holidays.country_holidays("BR", subdiv="SP", years=[ano])
    dias_uteis = 0
    dia = 1
    while True:
        dt = date(ano, mes, dia)
        if dt.weekday() < 5 and dt not in feriados_sp:
            dias_uteis += 1
            if dias_uteis == 2:
                return dt
        dia += 1

def verificar_status_prazo_data(dt_envio, ano, mes):
    limite = obter_segundo_dia_util(ano, mes)
    if dt_envio is None:
        hoje = date.today()
        if hoje > limite:
            return f"ATRASADO (Prazo 2º dia útil: {limite.strftime('%d/%m/%Y')})"
        return f"AGUARDANDO ENVIO (Prazo até {limite.strftime('%d/%m/%Y')})"
    if dt_envio <= limite:
        return f"NO PRAZO (Enviado em {dt_envio.strftime('%d/%m/%Y')} | Limite: {limite.strftime('%d/%m/%Y')})"
    return f"FORA DO PRAZO (Enviado em {dt_envio.strftime('%d/%m/%Y')} | Limite: {limite.strftime('%d/%m/%Y')})"

# --- TRADUTOR INTELIGENTE DE PLANILHAS DO DRIVE (AVALIAÇÕES, FREQUÊNCIA, PLATAFORMAS) ---
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
    "RED": "Redacao e Leitura",
    "EDF": "Educacao Fisica",
    "ART": "Arte",
    "TEC": "Tecnologia"
}

def limpar_nome_escola_bi(nome_bruto):
    s = str(nome_bruto).strip()
    if " - " in s and s.split(" - ")[-1].strip().isdigit():
        s = " - ".join(s.split(" - ")[:-1]).strip()
    s_title = s.title()
    return s_title if s_title.upper().startswith("EE ") else f"EE {s_title}"

def padronizar_planilha_bruta(df_raw, nome_arquivo="", pasta_origem=""):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    nome_up = normalizar_texto(nome_arquivo)
    pasta_up = normalizar_texto(pasta_origem)
    cols_norm = {normalizar_texto(c): c for c in df_raw.columns}
    siglas_presentes = [c_orig for c_n, c_orig in cols_norm.items() if c_n in MAPA_SIGLAS_BI]

    # Caso 1: Matriz BI SEDUC (SARESP livre / Prova Paulista / Diagnóstica)
    if siglas_presentes and ("EXPORT TURMA" in cols_norm or "EXPORT ESCOLAS" in cols_norm):
        escola_detectada = "Geral URE Miracatu"
        for val in df_raw.iloc[:, 0].dropna().astype(str):
            if "NM_ESCOLA" in val:
                trecho = val.split("NM_ESCOLA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho:
                        escola_detectada = limpar_nome_escola_bi(trecho.split(sep)[1].split("CD_DIRETORIA")[0].strip())
                        break

        avaliacao_nome = "SARESP (Livre)" if "SARESP" in nome_up else ("Prova Paulista" if "PP" in nome_up or "PAULISTA" in nome_up else f"Avaliação ({nome_arquivo})")
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
                if acerto_num <= 40.0:
                    nivel_str = "Vermelho (Crítico / Abaixo do Básico)"
                elif acerto_num <= 60.0:
                    nivel_str = "Amarelo (Atenção / Básico)"
                elif acerto_num <= 80.0:
                    nivel_str = "Verde (Adequado)"
                else:
                    nivel_str = "Azul (Avançado)"

                registros.append({
                    "Tipo_Indicador": "AVALIACAO",
                    "Escola": escola_row,
                    "Avaliacao": avaliacao_nome,
                    "Etapa": "Ensino Médio" if "SERIE" in turma_row.upper() else ("Anos Finais" if "ANO" in turma_row.upper() else "Todas"),
                    "Turma": turma_row,
                    "Professor": "Consultar Atribuição / CGP",
                    "Estudante": f"Turma ({total_al} alunos)",
                    "Componente": comp_nome,
                    "Nivel": nivel_str,
                    "Proficiencia_Nota": f"{acerto_num:.1f}%",
                    "Valor_Numerico": acerto_num,
                    "Frequencia": round(part_num, 1),
                    "Habilidade_Critica": f"Desempenho abaixo da meta ({acerto_num:.1f}%)" if acerto_num <= 60 else "Dentro do esperado",
                    "Plataforma_Critica": "-"
                })
        return pd.DataFrame(registros)

    # Caso 2: Planilhas tabulares (Avaliações, Frequência Semanal ou Plataformas Digitais)
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
        elif any(k in c_low for k in ["nota", "acerto", "pontuacao", "media", "escore", "uso", "engajamento", "taxa"]):
            mapa[col] = "Proficiencia_Nota"
        elif "freq" in c_low:
            mapa[col] = "Frequencia"
        elif "habilidade" in c_low or "descritor" in c_low:
            mapa[col] = "Habilidade_Critica"
        elif any(k in c_low for k in ["plataforma", "speak", "open english", "khan", "matific", "leia", "redacao", "alura"]):
            mapa[col] = "Plataforma_Critica"

    df_pad = df_raw.rename(columns=mapa).copy()

    # Identifica o tipo de indicador pela subpasta do Drive ou nome do arquivo
    if "FREQUENCIA" in pasta_up or "FREQUENCIA" in nome_up:
        tipo_ind = "FREQUENCIA"
    elif "PLATAFORMA" in pasta_up or any(p in nome_up for p in ["PLATAFORMA", "SPEAK", "OPEN", "KHAN", "MATIFIC", "LEIA"]):
        tipo_ind = "PLATAFORMA"
    else:
        tipo_ind = "AVALIACAO"

    df_pad["Tipo_Indicador"] = tipo_ind

    for col_req, default_val in [
        ("Escola", "Geral URE Miracatu"),
        ("Avaliacao", "SARESP / Prova Paulista" if tipo_ind == "AVALIACAO" else ("Frequência Semanal" if tipo_ind == "FREQUENCIA" else "Plataforma Digital")),
        ("Etapa", "Todas"),
        ("Turma", "Geral"),
        ("Professor", "Docente da Turma"),
        ("Estudante", "Consolidado"),
        ("Componente", "Todos"),
        ("Nivel", "Atenção / Monitorar"),
        ("Proficiencia_Nota", "-"),
        ("Frequencia", 85.0),
        ("Habilidade_Critica", "Acompanhamento Semanal"),
        ("Plataforma_Critica", nome_arquivo if tipo_ind == "PLATAFORMA" else "-")
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
                hora_str = f"{dtstart.strftime('%H:%M')} às {dtend.strftime('%H:%M')}" if hasattr(dtstart, "hour") else "Dia Integral"
                lista_ev.append({
                    "ano": dtstart.year, "mes": dtstart.month, "dia": dtstart.day,
                    "data_str": dtstart.strftime("%d/%m/%Y"), "horario": hora_str,
                    "resumo": resumo, "local": local, "descricao": desc
                })
        return lista_ev
    except Exception:
        return []

@st.cache_data(ttl=180)
def sincronizar_pasta_drive_completo(folder_id, api_key):
    """
    Lê todas as subpastas da pasta raiz do Google Drive:
      - Avaliações (SARESP livre, Prova Paulista, Diagnósticas)
      - Frequência Semanal das Escolas
      - Plataformas (Speak, Open English, Khan Academy, Matific, Leia SP, Redação, Alura)
      - Planos de Trabalho (.docx) enviados pelos PECs (verifica data vs 2º dia útil)
      - Termos de Visita (.pdf) enviados pelos PECs por escola/dia
    """
    if not api_key or "COLE_AQUI" in api_key:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0

    dfs_ind, dfs_curs = [], []
    registros_planos, registros_termos = [], []
    qtd_arquivos = 0

    def varrer_pasta(fid, caminho_pasta=""):
        nonlocal qtd_arquivos
        q = urllib.parse.quote(f"'{fid}' in parents and trashed = false")
        url_api = f"https://www.googleapis.com/drive/v3/files?q={q}&fields=files(id,name,mimeType,modifiedTime)&key={api_key}"
        try:
            resp = urllib.request.urlopen(url_api, timeout=15)
            dados = json.loads(resp.read().decode("utf-8"))
        except Exception:
            return

        for f in dados.get("files", []):
            f_id, f_name, f_mime = f["id"], f["name"], f["mimeType"]
            mod_time_str = f.get("modifiedTime", "")
            dt_mod = None
            if mod_time_str:
                try:
                    dt_mod = datetime.fromisoformat(mod_time_str.replace("Z", "+00:00")).date()
                except Exception:
                    pass

            if f_mime == "application/vnd.google-apps.folder":
                varrer_pasta(f_id, f"{caminho_pasta}/{f_name}")
            else:
                try:
                    if f_mime == "application/vnd.google-apps.spreadsheet":
                        dl_url = f"https://docs.google.com/spreadsheets/d/{f_id}/export?format=xlsx"
                    elif f_mime == "application/vnd.google-apps.document":
                        dl_url = f"https://docs.google.com/document/d/{f_id}/export?format=docx"
                    else:
                        dl_url = f"https://www.googleapis.com/drive/v3/files/{f_id}?alt=media&key={api_key}"

                    arq_bytes = BytesIO(urllib.request.urlopen(dl_url, timeout=20).read())
                    n_low = f_name.lower()
                    pasta_low = caminho_pasta.lower()

                    # 1. Termos de Visita em PDF
                    if n_low.endswith(".pdf") or f_mime == "application/pdf":
                        leitor = PdfReader(arq_bytes)
                        texto_pdf = "
".join([p.extract_text() or "" for p in leitor.pages])
                        pec_nome = os.path.basename(caminho_pasta).replace("Pasta_PEC_", "").strip() or "PEC"
                        escola_enc = next((e for e in ESCOLAS_URE_MIRACATU if normalizar_texto(e) in normalizar_texto(texto_pdf) or normalizar_texto(e) in normalizar_texto(f_name)), "Escola não identificada no PDF")
                        m_data = re.search(r"(\d{2}/\d{2}/\d{4})", texto_pdf)
                        data_vis = m_data.group(1) if m_data else (dt_mod.strftime("%d/%m/%Y") if dt_mod else "-")
                        registros_termos.append({
                            "PEC": pec_nome,
                            "Escola": escola_enc,
                            "Data_Visita": data_vis,
                            "Arquivo_Termo_PDF": f_name
                        })
                        qtd_arquivos += 1

                    # 2. Planos de Trabalho em Word (.docx)
                    elif (n_low.endswith(".docx") or f_mime == "application/vnd.google-apps.document") and not n_low.startswith("00_modelo"):
                        doc_p = Document(arq_bytes)
                        texto_doc = "
".join([p.text for p in doc_p.paragraphs])
                        for t in doc_p.tables:
                            for r in t.rows:
                                texto_doc += "
" + " | ".join([c.text.strip() for c in r.cells])

                        pec_nome = os.path.basename(caminho_pasta).replace("Pasta_PEC_", "").strip()
                        if not pec_nome or "plano" in pec_nome.lower():
                            pec_nome = f_name.replace(".docx", "").replace("GPS_", "").strip()

                        escolas_citadas = [e for e in ESCOLAS_URE_MIRACATU if normalizar_texto(e) in normalizar_texto(texto_doc)]
                        if not escolas_citadas:
                            escolas_citadas = ["Plano Geral / Sem escola específica"]

                        hoje = date.today()
                        status_pz = verificar_status_prazo_data(dt_mod, hoje.year, hoje.month)
                        for esc_c in escolas_citadas:
                            registros_planos.append({
                                "PEC": pec_nome,
                                "Escola": esc_c,
                                "Data_Envio": dt_mod.strftime("%d/%m/%Y") if dt_mod else "-",
                                "Status_Prazo_2DiaUtil": status_pz,
                                "Arquivo_Plano": f_name
                            })
                        qtd_arquivos += 1

                    # 3. Planilhas de Avaliações, Frequência Semanal e Plataformas
                    elif n_low.endswith(".csv"):
                        df_l = pd.read_csv(arq_bytes)
                        df_p = padronizar_planilha_bruta(df_l, f_name, caminho_pasta)
                        if not df_p.empty:
                            (dfs_curs if "multiplica" in n_low or "cursista" in n_low else dfs_ind).append(df_p)
                            qtd_arquivos += 1
                    elif n_low.endswith(".xlsx") or "spreadsheet" in f_mime:
                        df_l = pd.read_excel(arq_bytes)
                        df_p = padronizar_planilha_bruta(df_l, f_name, caminho_pasta)
                        if not df_p.empty:
                            (dfs_curs if "multiplica" in n_low or "cursista" in n_low else dfs_ind).append(df_p)
                            qtd_arquivos += 1
                except Exception:
                    pass

    varrer_pasta(folder_id, "")
    df_ind_final = pd.concat(dfs_ind, ignore_index=True) if dfs_ind else pd.DataFrame()
    df_curs_final = pd.concat(dfs_curs, ignore_index=True) if dfs_curs else pd.DataFrame()
    df_planos_final = pd.DataFrame(registros_planos).drop_duplicates() if registros_planos else pd.DataFrame()
    df_termos_final = pd.DataFrame(registros_termos).drop_duplicates() if registros_termos else pd.DataFrame()
    return df_ind_final, df_curs_final, df_planos_final, df_termos_final, qtd_arquivos

df_usuarios = carregar_aba("Usuarios")
df_pastas_raw = carregar_aba("Config_Pastas")
if df_pastas_raw.empty:
    df_pastas_raw = carregar_aba("Confing_Pastas")
df_eventos = carregar_aba("Calendario_Eventos")
df_bloqueios = carregar_aba("Bloqueios_Agenda")
df_agenda_seduc = carregar_agenda_seduc_eec()

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
        agrupado = df_c.groupby(["Pasta", "Resolucao"], as_index=False).agg({"Foco_Pedagogico": lambda x: "
".join(x.astype(str))})
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
        "Articular currículo, materiais digitais/impressos e plataformas (Speak, Open English, Khan Academy, Matific, Leia SP, Redação) com SARESP, Prova Paulista, Diagnóstica, SAEB e ENEM",
        "Priorizar turmas, docentes e estudantes com desempenho crítico dentro da sua pasta de atuação",
        "Realizar encontros virtuais ou visitas presenciais nas escolas com menor desempenho nas avaliações e plataformas",
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
        "Monitorar evolução de proficiência, frequência e plataformas (Khan Academy, Matific, Leia SP) em Língua Portuguesa e Matemática",
        "Realizar visitas presenciais regulares e encontros virtuais nas escolas prioritárias",
        "Acionar PEC Desenvolvimento Curricular diante de dificuldades conceituais/materiais"
    ],
    "CEC": [
        "Conduzir alinhamento semanal com a Dirigente Regional e Equipe de Especialistas",
        "Consolidar Plano de Trabalho mensal da equipe (prazo 2º dia útil) e monitorar Termos de Visita",
        "Realizar reuniões individuais (1:1) e auditar justificativas de fechamento mensal"
    ],
    "AT": [
        "Apoiar a coordenação dos alinhamentos regionais e consolidar registros da equipe",
        "Acompanhar cronograma de formações, convocações e indicadores da URE"
    ]
}

# ==============================================================================
# GERADOR DE PLANO DE TRABALHO EM WORD COM CABEÇALHO OFICIAL DA URE MIRACATU
# ==============================================================================
def gerar_docx_plano(nome_pec, email_pec, pasta_pec, comp_pec, escola_sel, semana_str, cronograma_dias, obs_texto, df_prioridades_pec=None):
    doc = Document()

    # Cabeçalho oficial URE Miracatu
    header = doc.sections[0].header
    p_cab = header.paragraphs[0]
    p_cab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_cab.add_run("SECRETARIA DA EDUCAÇÃO DO ESTADO DE SÃO PAULO
")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p_cab.add_run("UNIDADE REGIONAL DE ENSINO DE MIRACATU — EQUIPE DE ESPECIALISTAS EM CURRÍCULO (EEC)
")
    r2.bold = True
    r2.font.size = Pt(9.5)
    r3 = p_cab.add_run("PLANO DE TRABALHO / GUIA PEDAGÓGICO — ENTREGA ATÉ O 2º DIA ÚTIL DO MÊS")
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(90, 90, 90)

    titulo = doc.add_heading("PLANO DE TRABALHO — ACOMPANHAMENTO PEDAGÓGICO (GPS)", level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_info = doc.add_paragraph()
    p_info.add_run("Especialista / PEC Responsável: ").bold = True
    p_info.add_run(f"{nome_pec} ({email_pec})
")
    p_info.add_run("Pasta / Área de Atuação: ").bold = True
    p_info.add_run(f"{pasta_pec} — Componente: {comp_pec}
")
    p_info.add_run("Unidade Escolar Prioritária: ").bold = True
    p_info.add_run(f"{escola_sel}
")
    p_info.add_run("Período / Mês de Referência: ").bold = True
    p_info.add_run(f"{semana_str}
")

    doc.add_heading("1. Cronograma de Visitas e Ações Planejadas", level=2)
    tabela = doc.add_table(rows=1, cols=4)
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Data / Dia / Turno"
    hdr[1].text = "Escola / Local"
    hdr[2].text = "Ação Pedagógica (Resolução SEDUC 01/2026)"
    hdr[3].text = "Público-Alvo / Encaminhamento"

    for item in cronograma_dias:
        row_cells = tabela.add_row().cells
        row_cells[0].text = item.get("dia", "")
        row_cells[1].text = item.get("local", "")
        row_cells[2].text = item.get("acao", "")
        row_cells[3].text = item.get("publico", "")

    if df_prioridades_pec is not None and not df_prioridades_pec.empty:
        doc.add_heading("2. Indicadores Prioritários da Sua Pasta de Atuação (Avaliações, Frequência e Plataformas)", level=2)
        t_ind = doc.add_table(rows=1, cols=5)
        t_ind.style = "Table Grid"
        for idx, col_n in enumerate(["Escola", "Tipo / Avaliação", "Turma / Componente", "Resultado / Nível", "Plataforma / Observação"]):
            t_ind.rows[0].cells[idx].text = col_n
        for _, r_i in df_prioridades_pec.head(15).iterrows():
            rc = t_ind.add_row().cells
            rc[0].text = str(r_i.get("Escola", "-"))
            rc[1].text = str(r_i.get("Avaliacao", "-"))
            rc[2].text = f"{r_i.get('Turma', '-')} | {r_i.get('Componente', '-')}"
            rc[3].text = f"{r_i.get('Proficiencia_Nota', '-')} ({r_i.get('Nivel', '-')})"
            rc[4].text = str(r_i.get("Plataforma_Critica", "-"))

    if obs_texto.strip():
        doc.add_heading("3. Pautas Formativas, Diagnóstico e Devolutivas", level=2)
        doc.add_paragraph(obs_texto.strip())

    doc.add_paragraph("
")
    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass.add_run("_________________________________________          _________________________________________
")
    p_ass.add_run(f"PEC: {nome_pec}                                    CEC — Coordenador de Equipe Curricular")

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
# AUTENTICAÇÃO E SESSÃO (LÊ DADOS DE CADA PASTA E PEC DO GOOGLE SHEETS)
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario_nome"] = ""
    st.session_state["usuario_email"] = ""
    st.session_state["usuario_pasta"] = "QUALIDADE DA AULA"
    st.session_state["usuario_componente"] = "Todos"
    st.session_state["usuario_escolas"] = []
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
                    escolas_encontradas = []
                    acesso_ok = True

                    if not df_usuarios.empty:
                        cols_u = {normalizar_texto(c): c for c in df_usuarios.columns}
                        col_email = next((v for k, v in cols_u.items() if "EMAIL" in k or "E-MAIL" in k), None)
                        col_nome = next((v for k, v in cols_u.items() if "NOME" in k), None)
                        col_pasta = next((v for k, v in cols_u.items() if "PASTA" in k or "FUNCAO" in k), None)
                        col_comp = next((v for k, v in cols_u.items() if "COMPONENTE" in k or "AREA" in k), None)
                        col_esc = next((v for k, v in cols_u.items() if "ESCOLA" in k), None)
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
                                    if col_esc:
                                        escolas_encontradas = [
                                            e.strip() for e in match[col_esc].dropna().astype(str).tolist()
                                            for e in e.split(",") if e.strip()
                                        ]

                    if acesso_ok:
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_nome"] = nome_encontrado
                        st.session_state["usuario_email"] = email_in.strip()
                        st.session_state["usuario_pasta"] = pasta_encontrada if pasta_encontrada in TAREFAS_POR_PASTA else pasta_fallback
                        st.session_state["usuario_componente"] = comp_encontrado
                        st.session_state["usuario_escolas"] = escolas_encontradas
                        st.rerun()
    st.stop()

# ==============================================================================
# CARREGAMENTO GLOBAL DOS DADOS DO DRIVE (PARA TODAS AS ABAS)
# ==============================================================================
pasta_atual = st.session_state["usuario_pasta"]
comp_atual = st.session_state["usuario_componente"]
escolas_usuario = st.session_state["usuario_escolas"]
eh_cec_ou_at = pasta_atual in ["CEC", "AT"]

df_ind_global, df_curs_global, df_planos_drive, df_termos_drive, qtd_arqs_drive = (
    sincronizar_pasta_drive_completo(ID_PASTA_DRIVE_GPS, GOOGLE_DRIVE_API_KEY)
    if "COLE_AQUI" not in GOOGLE_DRIVE_API_KEY else (pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0)
)

df_ind_usuario = (
    df_ind_global if eh_cec_ou_at
    else filtrar_indicadores_por_pec(df_ind_global, pasta_atual, comp_atual, escolas_usuario)
)

# ==============================================================================
# BARRA LATERAL E PAINEL PRINCIPAL
# ==============================================================================
with st.sidebar:
    st.markdown("## 🧭 GPS — URE Miracatu")
    st.write(f"**Especialista:** {st.session_state['usuario_nome']}")
    st.write(f"**Pasta:** {pasta_atual}")
    st.write(f"**Área / Componente:** {comp_atual}")

    hoje = date.today()
    limite_2du = obter_segundo_dia_util(hoje.year, hoje.month)
    st.info(f"📅 **Prazo Plano de Trabalho:** até o 2º dia útil ({limite_2du.strftime('%d/%m/%Y')})")

    st.markdown("---")
    if st.button("🔄 Sincronizar Planilhas & Drive", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    if st.button("🚪 Sair / Trocar Usuário", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

st.title("🧭 GPS — Guia Pedagógico Semanal")
st.caption("Planejamento, Indicadores por Pasta de Atuação e Acompanhamento de Visitas — URE de Miracatu")

lista_abas = [
    "📋 Plano de Trabalho & Exportação",
    "📊 Indicadores, Plataformas & Frequência",
    "📅 Agenda SEDUC, CEC & Feriados",
    "📚 Diretrizes das Pastas (Res. SEDUC 01/2026)"
]
if eh_cec_ou_at:
    lista_abas.insert(2, "🛡️ Painel Exclusivo CEC (Confronto & Fechamento)")

abas = st.tabs(lista_abas)

# --- ABA 1: PLANO DE TRABALHO & EXPORTAÇÃO ---
with abas[0]:
    st.subheader("Elaboração do Plano de Trabalho (Modelo Oficial URE Miracatu)")
    st.caption("Submeta o arquivo gerado na pasta do Google Drive até o 2º dia útil do mês. Após cada visita, deposite o Termo de Visita em PDF na sua pasta do Drive.")

    col1, col2 = st.columns(2)
    with col1:
        lista_esc_select = escolas_usuario if escolas_usuario else ESCOLAS_URE_MIRACATU
        escola_prioritaria = st.selectbox("Unidade Escolar Prioritária:", lista_esc_select)
    with col2:
        semana_ref = st.text_input("Mês / Semana de Referência:", value=datetime.now().strftime("%m/%Y — Semana de %d/%m/%Y"))

    tarefas_pasta = TAREFAS_POR_PASTA.get(pasta_atual, [])

    st.markdown("#### Adicionar Visita / Atividade ao Plano")
    with st.form("form_add_atividade", clear_on_submit=True):
        c_dia, c_loc = st.columns(2)
        with c_dia:
            dia_turno = st.text_input("Data / Dia da Semana / Turno:", value=datetime.now().strftime("%d/%m/%Y (Manhã)"))
        with c_loc:
            local_acao = st.selectbox("Local / Escola:", ["Sede URE Miracatu / Trabalho Interno", "Reunião Virtual / EFAPE / SEDUC"] + lista_esc_select)

        acao_escolhida = st.selectbox("Ação Pedagógica da Pasta:", tarefas_pasta + ["Outra ação específica (descrever abaixo)"])
        publico_enc = st.text_input("Público-Alvo / Detalhamento da Ação:", placeholder="Ex.: CGP/CGPG, Professores da área, Plataformas (Speak, Khan, Matific)...")
        add_btn = st.form_submit_button("➕ Inserir no Plano")

        if add_btn:
            st.session_state["cronograma_semana"].append({
                "dia": dia_turno,
                "local": local_acao,
                "acao": acao_escolhida,
                "publico": publico_enc if publico_enc.strip() else "Equipe Gestora / Docentes"
            })
            st.success("Atividade adicionada ao Plano de Trabalho!")

    if st.session_state["cronograma_semana"]:
        df_cron = pd.DataFrame(st.session_state["cronograma_semana"])
        st.dataframe(df_cron, use_container_width=True)
        if st.button("🗑️ Limpar Atividades"):
            st.session_state["cronograma_semana"] = []
            st.rerun()
    else:
        st.info("Nenhuma atividade inserida ainda. Use o formulário acima para montar seu cronograma.")

    obs_texto = st.text_area("Pautas Formativas, Diagnóstico de Indicadores e Encaminhamentos:", height=110)

    # Prioridades filtradas da área do PEC para constar automaticamente no Word
    df_prio_word = pd.DataFrame()
    if not df_ind_usuario.empty:
        df_prio_word = df_ind_usuario[df_ind_usuario["Nivel"].astype(str).str.contains("Vermelho|Amarelo|Abaixo|Atenção", case=False, na=False)]

    docx_buf = gerar_docx_plano(
        st.session_state["usuario_nome"],
        st.session_state["usuario_email"],
        pasta_atual,
        comp_atual,
        escola_prioritaria,
        semana_ref,
        st.session_state["cronograma_semana"],
        obs_texto,
        df_prio_word
    )
    nome_arq_docx = f"Plano_Trabalho_URE_{st.session_state['usuario_nome'].replace(' ', '_')}_{datetime.now().strftime('%Y%m')}.docx"

    col_dl, col_em, col_drv_link = st.columns(3)
    with col_dl:
        st.download_button(
            label="⬇️ Baixar Plano de Trabalho (.docx com Cabeçalho URE)",
            data=docx_buf,
            file_name=nome_arq_docx,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col_em:
        if st.button("📧 Enviar Plano por E-mail ao CEC / Dirigente", use_container_width=True):
            ok, msg_envio = enviar_email_gps(
                [EMAIL_CEC, EMAIL_DIRIGENTE],
                f"[GPS URE Miracatu] Plano de Trabalho — {st.session_state['usuario_nome']} ({pasta_atual})",
                f"Segue em anexo o Plano de Trabalho de {st.session_state['usuario_nome']} ({pasta_atual} - {comp_atual}).",
                docx_buf,
                nome_arq_docx
            )
            if ok:
                st.success(msg_envio)
            else:
                st.warning(msg_envio)
    with col_drv_link:
        st.link_button("📂 Abrir Pasta do Drive para Depositar Plano / Termos", URL_PASTA_DRIVE_GPS, use_container_width=True)

    # Acompanhamento Individual do próprio PEC (Plano x Termos de Visita no Drive)
    if not eh_cec_ou_at and (not df_planos_drive.empty or not df_termos_drive.empty):
        st.markdown("---")
        st.markdown("#### 🔔 Meu Acompanhamento Mensal (Plano de Trabalho x Termos de Visita no Drive)")
        nome_u_norm = normalizar_texto(st.session_state["usuario_nome"])
        meus_planos = df_planos_drive[df_planos_drive["PEC"].apply(lambda x: nome_u_norm in normalizar_texto(x) or normalizar_texto(x) in nome_u_norm)] if not df_planos_drive.empty else pd.DataFrame()
        meus_termos = df_termos_drive[df_termos_drive["PEC"].apply(lambda x: nome_u_norm in normalizar_texto(x) or normalizar_texto(x) in nome_u_norm)] if not df_termos_drive.empty else pd.DataFrame()

        if not meus_planos.empty:
            st.dataframe(meus_planos, use_container_width=True)
            escolas_plan = set(meus_planos["Escola"].tolist())
            escolas_vis = set(meus_termos["Escola"].tolist()) if not meus_termos.empty else set()
            pendentes = [e for e in escolas_plan - escolas_vis if "Geral" not in e]
            if pendentes:
                st.warning(f"⚠️ **Atenção:** Você planejou visita e ainda não depositou o Termo de Visita (PDF) no Drive para: **{', '.join(pendentes)}**. Caso não ocorra até o fim do mês, registre a justificativa.")
            else:
                st.success("✅ Todas as escolas previstas no seu Plano de Trabalho possuem Termo de Visita registrado no Drive.")

# --- ABA 2: INDICADORES, PLATAFORMAS & FREQUÊNCIA ---
with abas[1]:
    st.subheader(f"Indicadores Pedagógicos — {pasta_atual} ({comp_atual})")
    if not eh_cec_ou_at:
        regras_info = obter_regras_atuacao_pec(pasta_atual, comp_atual)
        if regras_info:
            st.info(f"🎯 **Filtro automático da sua pasta de atuação:** Componentes `{', '.join(regras_info['componentes'])}` | Plataformas `{', '.join(regras_info['plataformas'])}` + Frequência Semanal.")

    col_drv, col_up = st.columns([2, 1])
    with col_drv:
        if "COLE_AQUI" in GOOGLE_DRIVE_API_KEY:
            st.warning("Configure `GOOGLE_DRIVE_API_KEY` nos Secrets do Streamlit para leitura automática das pastas do Google Drive.")
        else:
            st.success(f"📂 {qtd_arqs_drive} arquivo(s) sincronizado(s) do Google Drive (Avaliações SARESP livre/PP, Frequência Semanal e Plataformas).")

    with col_up:
        arquivo_bi = st.file_uploader("Testar/Visualizar planilha avulsa (.xlsx / .csv):", type=["xlsx", "csv"])
        if arquivo_bi is not None:
            df_raw_up = pd.read_csv(arquivo_bi) if arquivo_bi.name.endswith(".csv") else pd.read_excel(arquivo_bi)
            df_pad_up = padronizar_planilha_bruta(df_raw_up, arquivo_bi.name)
            df_ind_usuario = pd.concat([df_ind_usuario, df_pad_up], ignore_index=True)

    if not df_ind_usuario.empty:
        sub_aval, sub_freq, sub_plat = st.tabs(["📝 Avaliações (SARESP / PP / Diagnóstica)", "🏫 Frequência Semanal", "💻 Plataformas Digitais"])

        with sub_aval:
            df_av = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "AVALIACAO"]
            if not df_av.empty:
                esc_sel = st.selectbox("Filtrar Escola (Avaliações):", ["Todas"] + sorted(df_av["Escola"].dropna().unique().tolist()), key="f_esc_av")
                df_av_show = df_av if esc_sel == "Todas" else df_av[df_av["Escola"] == esc_sel]
                st.dataframe(df_av_show.drop(columns=["Tipo_Indicador", "Valor_Numerico"], errors="ignore"), use_container_width=True)
            else:
                st.info("Nenhum dado de avaliação encontrado para esta seleção.")

        with sub_freq:
            df_fr = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "FREQUENCIA"]
            if not df_fr.empty:
                st.dataframe(df_fr.drop(columns=["Tipo_Indicador", "Valor_Numerico"], errors="ignore"), use_container_width=True)
            else:
                st.info("Nenhuma planilha de frequência semanal carregada ainda.")

        with sub_plat:
            df_pl = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "PLATAFORMA"]
            if not df_pl.empty:
                st.dataframe(df_pl.drop(columns=["Tipo_Indicador", "Valor_Numerico"], errors="ignore"), use_container_width=True)
            else:
                st.info("Nenhuma planilha de plataformas digitais carregada ainda.")
    else:
        st.info("Deposite as planilhas nas pastas do Google Drive para visualizar os indicadores filtrados automaticamente.")

# --- ABA EXCLUSIVA DO CEC / AT: CONFRONTO & FECHAMENTO MENSAL ---
idx_proxima_aba = 2
if eh_cec_ou_at:
    with abas[2]:
        st.subheader("🛡️ Painel Exclusivo CEC — Controle de Prazos (2º Dia Útil), Confronto & Justificativas")
        st.caption("Visão global de todos os PECs, confrontando Planos de Trabalho enviados até o 2º dia útil x Termos de Visita depositados no Google Drive.")

        c_cec1, c_cec2 = st.columns(2)
        with c_cec1:
            st.markdown("#### 1. Planos de Trabalho Enviados (Auditoria do 2º Dia Útil)")
            if not df_planos_drive.empty:
                st.dataframe(df_planos_drive, use_container_width=True)
            else:
                st.info("Nenhum Plano de Trabalho (.docx) localizado no Google Drive até o momento.")

        with c_cec2:
            st.markdown("#### 2. Termos de Visita Depositados no Drive (PDF)")
            if not df_termos_drive.empty:
                st.dataframe(df_termos_drive, use_container_width=True)
            else:
                st.info("Nenhum Termo de Visita (.pdf) localizado no Google Drive até o momento.")

        st.markdown("#### 3. Relatório de Fechamento Mensal (Planejado x Visitado x Justificativa)")
        if not df_planos_drive.empty:
            confronto = df_planos_drive.merge(
                df_termos_drive[["PEC", "Escola", "Data_Visita", "Arquivo_Termo_PDF"]] if not df_termos_drive.empty else pd.DataFrame(columns=["PEC", "Escola", "Data_Visita", "Arquivo_Termo_PDF"]),
                on=["PEC", "Escola"],
                how="outer",
                indicator=True
            )
            confronto["Status_Execucao"] = confronto["_merge"].map({
                "both": "✅ CUMPRIDA (Planejada e Visitada)",
                "left_only": "⚠️ NÃO CUMPRIDA (Exige Justificativa no Fechamento)",
                "right_only": "➕ VISITA EXTRA (Fora do Plano Inicial)"
            })
            confronto.drop(columns=["_merge"], inplace=True)
            confronto["Justificativa_PEC_Fim_Mes"] = ""

            pec_filtro_cec = st.selectbox("Filtrar Visão Individual por PEC:", ["Todos os PECs (Visão Geral)"] + sorted(confronto["PEC"].dropna().unique().tolist()))
            confronto_exibir = confronto if pec_filtro_cec == "Todos os PECs (Visão Geral)" else confronto[confronto["PEC"] == pec_filtro_cec]

            confronto_editado = st.data_editor(confronto_exibir, use_container_width=True, num_rows="dynamic")

            buf_cec = BytesIO()
            with pd.ExcelWriter(buf_cec, engine="openpyxl") as writer:
                confronto_editado.to_excel(writer, sheet_name="Confronto_Fechamento_CEC", index=False)
                if not df_ind_global.empty:
                    df_ind_global.to_excel(writer, sheet_name="Indicadores_Globais_URE", index=False)
            buf_cec.seek(0)

            st.download_button(
                label="⬇️ Baixar Relatório Gerencial Consolidado CEC (.xlsx)",
                data=buf_cec,
                file_name=f"Relatorio_Gerencial_CEC_URE_Miracatu_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    idx_proxima_aba = 3

# --- ABA AGENDA SEDUC, CEC & FERIADOS ---
with abas[idx_proxima_aba]:
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

    st.markdown("**Agenda SEDUC / EEC:**")
    if df_agenda_seduc.empty:
        st.info("Nenhum registro carregado da planilha da Agenda SEDUC.")
    else:
        st.dataframe(df_agenda_seduc, use_container_width=True)

# --- ABA DIRETRIZES DAS PASTAS ---
with abas[idx_proxima_aba + 1]:
    st.subheader("Base Legal e Foco Pedagógico por Pasta (Resolução SEDUC Nº 01/2026)")
    st.dataframe(df_pastas, use_container_width=True)
