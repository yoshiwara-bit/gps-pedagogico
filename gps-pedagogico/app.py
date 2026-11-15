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
# CONFIGURAÇÕES DA URE DE MIRACATU
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

DIAS_SEMANA_PT = {
    0: "SEGUNDA-FEIRA",
    1: "TERCA-FEIRA",
    2: "QUARTA-FEIRA",
    3: "QUINTA-FEIRA",
    4: "SEXTA-FEIRA",
    5: "SABADO",
    6: "DOMINGO"
}

DIAS_SEMANA_LABEL = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo"
}

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
        return None

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
            plat = normalizar_texto(row.get("Subcategoria", "")) + " " + normalizar_texto(row.get("Plataforma_Critica", ""))
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

# ==============================================================================
# CONVERSÃO DE PERCENTUAL + DETECÇÃO DE PROVA / PERÍODO / NÍVEL
# ==============================================================================
def converter_para_percentual(val):
    if pd.isna(val):
        return None
    s = str(val).replace("%", "").replace(",", ".").strip()
    if not s or s.lower() in ["nan", "none", "-", ""]:
        return None
    try:
        num = float(s)
        if 0.0 <= num <= 1.0001:
            return round(num * 100.0, 1)
        return round(num, 1)
    except Exception:
        return None

def formatar_pct_str(num):
    if num is None or pd.isna(num):
        return "-"
    return f"{num:.1f}%"

def classificar_farol_nivel(pct):
    if pct is None or pd.isna(pct):
        return "Sem Dado"
    if pct <= 40.0:
        return "🔴 Vermelho (Crítico / Abaixo do Básico)"
    elif pct <= 60.0:
        return "🟡 Amarelo (Atenção / Básico)"
    elif pct <= 80.0:
        return "🟢 Verde (Adequado)"
    else:
        return "🔵 Azul (Avançado)"

def detectar_classificacao_arquivo(nome_arquivo, caminho_pasta=""):
    ref = normalizar_texto(f"{caminho_pasta} {nome_arquivo}")

    if "FREQUENCIA" in ref:
        tipo = "FREQUENCIA"
    elif any(k in ref for k in ["PLATAFORMA", "KHAN", "MATIFIC", "SPEAK", "OPEN ENGLISH", "LEIA SP", "REDACAO", "ALURA"]):
        tipo = "PLATAFORMA"
    else:
        tipo = "AVALIACAO"

    if tipo == "PLATAFORMA":
        if "KHAN" in ref:
            sub = "Khan Academy"
        elif "MATIFIC" in ref:
            sub = "Matific"
        elif "SPEAK" in ref:
            sub = "Speak"
        elif "OPEN" in ref:
            sub = "Open English"
        elif "LEIA" in ref:
            sub = "Leia SP"
        elif "REDACAO" in ref:
            sub = "Redação Paulista"
        elif "ALURA" in ref:
            sub = "Alura"
        else:
            sub = "Outras Plataformas"
    elif tipo == "FREQUENCIA":
        sub = "Frequência Escolar"
    else:
        if "PAULISTA" in ref or re.search(r"\bPP\b", ref):
            sub = "Prova Paulista"
        elif "SARESP" in ref:
            sub = "SARESP"
        elif "SAEB" in ref:
            sub = "SAEB"
        elif "ENEM" in ref:
            sub = "ENEM"
        elif "DIAGNOSTICA" in ref or "ENTRADA" in ref:
            sub = "Avaliação Diagnóstica"
        else:
            sub = "Outras Avaliações"

    m_bim = re.search(r"([1-4])\s*[º°O]?\s*(?:BIM|B\b)", ref)
    m_ano = re.search(r"\b(202[3-9])\b", ref)
    m_sem = re.search(r"SEM(?:ANA)?\s*([0-9]{1,2})", ref)
    if m_bim:
        periodo = f"{m_bim.group(1)}º Bimestre" + (f" ({m_ano.group(1)})" if m_ano else "")
    elif m_ano:
        periodo = f"Ano {m_ano.group(1)}"
    elif m_sem:
        periodo = f"Semana {m_sem.group(1)}"
    else:
        periodo = "Geral / Período Atual"

    escola_pasta = None
    for esc in ESCOLAS_URE_MIRACATU:
        if normalizar_texto(esc) in ref or normalizar_texto(esc).replace("EE ", "") in ref:
            escola_pasta = esc
            break

    return tipo, sub, periodo, escola_pasta

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

def corresponder_escola_oficial(nome_bruto):
    norm_b = normalizar_texto(limpar_nome_escola_bi(nome_bruto)).replace("EE ", "")
    for esc in ESCOLAS_URE_MIRACATU:
        norm_o = normalizar_texto(esc).replace("EE ", "")
        if norm_b == norm_o or norm_b in norm_o or norm_o in norm_b:
            return esc
    return limpar_nome_escola_bi(nome_bruto)

# ==============================================================================
# TRADUTOR DE PLANILHA DE ATPC POR ESCOLA (DIA DA SEMANA + HORÁRIO + ÁREA)
# ==============================================================================
def formatar_valor_hora_atpc(val):
    if val is None or pd.isna(val):
        return ""
    if hasattr(val, "strftime"):
        return val.strftime("%H:%M")
    s = str(val).strip()
    s = re.sub(r"(\d{1,2}:\d{2}):\d{2}", r"\1", s)
    return s

def detectar_dias_semana_texto(texto_norm):
    dias = []
    if "SEG" in texto_norm or "2A" in texto_norm:
        dias.append(0)
    if "TER" in texto_norm or "3A" in texto_norm:
        dias.append(1)
    if "QUA" in texto_norm or "4A" in texto_norm:
        dias.append(2)
    if "QUI" in texto_norm or "5A" in texto_norm:
        dias.append(3)
    if "SEX" in texto_norm or "6A" in texto_norm:
        dias.append(4)
    if "SAB" in texto_norm:
        dias.append(5)
    return dias

def padronizar_planilha_atpc(df_raw):
    """
    Lê planilhas de ATPC do Google Drive ou Google Sheets contendo Escola, Dia da Semana, Horário e Área/Componente.
    Suporta tanto formato em linhas (Escola | Dia | Horário | Área) quanto formato matriz (Escola | Segunda | Terça | ...).
    Retorna DataFrame padronizado com: Escola, Dia_Semana_Num (0=Seg..6=Dom), Dia_Semana_Str, Horario, Area_Componente.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    cols_norm = {normalizar_texto(c): c for c in df_raw.columns}
    # Evita confundir planilhas de proficiência BI ou frequência com planilhas de ATPC
    if any(k in MAPA_SIGLAS_BI for k in cols_norm):
        return pd.DataFrame()
    if any(any(x in k for x in ["PROFICIENCIA", "ACERTO", "FREQUENCIA", "ENGAJAMENTO"]) for k in cols_norm):
        return pd.DataFrame()

    col_esc = next((v for k, v in cols_norm.items() if any(x in k for x in ["ESCOLA", "UNIDADE", "UE"])), None)
    if not col_esc:
        return pd.DataFrame()

    col_dia = next((v for k, v in cols_norm.items() if any(x in k for x in ["DIA", "SEMANA", "ATPC"]) and not detectar_dias_semana_texto(k)), None)
    col_hor = next((v for k, v in cols_norm.items() if any(x in k for x in ["HORA", "HORARIO", "TURNO", "PERIODO"])), None)
    col_ini = next((v for k, v in cols_norm.items() if "INICIO" in k), None)
    col_fim = next((v for k, v in cols_norm.items() if "TERMINO" in k or "FIM" in k), None)
    col_area = next((v for k, v in cols_norm.items() if any(x in k for x in ["AREA", "COMPONENTE", "DISCIPLINA", "SEGMENTO"])), None)

    # Verifica se existe formato matriz onde os cabeçalhos das colunas são os dias da semana (Segunda, Terça, Quarta, Quinta, Sexta)
    cols_dias_matriz = {}
    for k, v in cols_norm.items():
        if v == col_esc:
            continue
        dias_col = detectar_dias_semana_texto(k)
        if len(dias_col) == 1:
            cols_dias_matriz[v] = dias_col[0]

    registros = []

    if col_dia:
        for _, row in df_raw.iterrows():
            esc_raw = str(row[col_esc]).strip() if pd.notna(row[col_esc]) else ""
            dia_raw = normalizar_texto(row[col_dia]) if pd.notna(row[col_dia]) else ""
            if not esc_raw or not dia_raw or esc_raw.lower() in ["nan", "none", "total"]:
                continue

            escola_of = corresponder_escola_oficial(esc_raw)

            if col_ini and col_fim and pd.notna(row[col_ini]) and pd.notna(row[col_fim]):
                h_i = formatar_valor_hora_atpc(row[col_ini])
                h_f = formatar_valor_hora_atpc(row[col_fim])
                horario_str = f"{h_i} às {h_f}" if h_i and h_f else "Horário ATPC"
            elif col_hor and pd.notna(row[col_hor]):
                horario_str = formatar_valor_hora_atpc(row[col_hor]) or "Horário ATPC"
            else:
                horario_str = "Horário ATPC"

            area_str = str(row[col_area]).strip() if col_area and pd.notna(row[col_area]) and str(row[col_area]).strip().lower() not in ["nan", "none", ""] else "Geral / Todas as Áreas"

            for d_num in detectar_dias_semana_texto(dia_raw):
                registros.append({
                    "Escola": escola_of,
                    "Dia_Semana_Num": d_num,
                    "Dia_Semana_Str": DIAS_SEMANA_LABEL[d_num],
                    "Horario": horario_str,
                    "Area_Componente": area_str
                })
    elif cols_dias_matriz:
        for _, row in df_raw.iterrows():
            esc_raw = str(row[col_esc]).strip() if pd.notna(row[col_esc]) else ""
            if not esc_raw or esc_raw.lower() in ["nan", "none", "total"]:
                continue
            escola_of = corresponder_escola_oficial(esc_raw)
            area_base = str(row[col_area]).strip() if col_area and pd.notna(row[col_area]) and str(row[col_area]).strip().lower() not in ["nan", "none", ""] else "Geral / Todas as Áreas"
            for col_orig, d_num in cols_dias_matriz.items():
                val_cel = formatar_valor_hora_atpc(row[col_orig])
                if not val_cel or val_cel.lower() in ["nan", "none", "-", "nao", "não"]:
                    continue
                registros.append({
                    "Escola": escola_of,
                    "Dia_Semana_Num": d_num,
                    "Dia_Semana_Str": DIAS_SEMANA_LABEL[d_num],
                    "Horario": val_cel,
                    "Area_Componente": area_base
                })

    return pd.DataFrame(registros).drop_duplicates() if registros else pd.DataFrame()

def obter_atpcs_escola(df_atpc, escola_alvo, area_pec="Todos"):
    if df_atpc is None or df_atpc.empty or not escola_alvo:
        return pd.DataFrame()
    esc_norm = normalizar_texto(escola_alvo)
    sub = df_atpc[df_atpc["Escola"].apply(lambda x: esc_norm in normalizar_texto(x) or normalizar_texto(x) in esc_norm)].copy()
    if sub.empty:
        return sub
    # Se houver filtro de área e existirem linhas específicas, prioriza as da área ou Gerais
    if area_pec and normalizar_texto(area_pec) not in ["", "TODOS", "CEC", "AT"]:
        a_norm = normalizar_texto(area_pec)
        match_area = sub[sub["Area_Componente"].apply(lambda x: a_norm in normalizar_texto(x) or normalizar_texto(x) in a_norm or "GERAL" in normalizar_texto(x) or "TODAS" in normalizar_texto(x))]
        if not match_area.empty:
            return match_area
    return sub

def sugerir_datas_atpc_no_mes(ano, mes, df_atpc_escola, dias_bloqueados_set=None):
    """Retorna lista de dicionários com datas do mês em que ocorre ATPC na escola selecionada (excluindo feriados/bloqueios)."""
    if df_atpc_escola is None or df_atpc_escola.empty:
        return []
    feriados_sp = holidays.country_holidays("BR", subdiv="SP", years=[ano])
    dias_bloqueados_set = dias_bloqueados_set or set()
    _, num_dias = calendar.monthrange(ano, mes)

    sugestoes = []
    for dia in range(1, num_dias + 1):
        dt = date(ano, mes, dia)
        if dt in feriados_sp or dt in dias_bloqueados_set:
            continue
        matches = df_atpc_escola[df_atpc_escola["Dia_Semana_Num"] == dt.weekday()]
        for _, m in matches.iterrows():
            sugestoes.append({
                "data": dt,
                "data_str": dt.strftime("%d/%m/%Y"),
                "dia_semana": m["Dia_Semana_Str"],
                "horario": m["Horario"],
                "area": m["Area_Componente"]
            })
    return sugestoes

# ==============================================================================
# TRADUTOR INTELIGENTE DE PLANILHAS DE INDICADORES (ESCOLA -> TURMA -> ESTUDANTE)
# ==============================================================================
def padronizar_planilha_bruta(df_raw, nome_arquivo="", pasta_origem=""):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    tipo_ind, subcategoria, periodo_ref, escola_detectada_path = detectar_classificacao_arquivo(nome_arquivo, pasta_origem)
    cols_norm = {normalizar_texto(c): c for c in df_raw.columns}
    siglas_presentes = [c_orig for c_n, c_orig in cols_norm.items() if c_n in MAPA_SIGLAS_BI]

    escola_rodape = None
    turma_rodape = None
    for val in df_raw.iloc[:, 0].dropna().astype(str):
        if "FILTROS APLICADOS" in normalizar_texto(val):
            if "NM_ESCOLA" in val:
                trecho = val.split("NM_ESCOLA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho:
                        escola_rodape = corresponder_escola_oficial(trecho.split(sep)[1].split("CD_DIRETORIA")[0].split("NM_TURMA")[0].strip())
                        break
            if "NM_TURMA" in val:
                trecho_t = val.split("NM_TURMA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho_t:
                        turma_rodape = trecho_t.split(sep)[1].split("NM_ESCOLA")[0].split("CD_DIRETORIA")[0].strip()
                        break

    escola_padrao = escola_rodape or escola_detectada_path or "Geral URE Miracatu"
    col_prof = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["PROFESSOR", "DOCENTE"])), None)
    col_recomp = next((c_orig for c_n, c_orig in cols_norm.items() if "RECOMPOSICAO" in c_n), None)
    col_per_row = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["BIMESTRE", "PERIODO", "EDICAO"])), None)
    col_sub_row = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["AVALIACAO", "PROVA", "PLATAFORMA", "INDICADOR"])), None)

    if siglas_presentes:
        col_escola = next((c_orig for c_n, c_orig in cols_norm.items() if "EXPORT ESCOLA" in c_n or c_n == "ESCOLA" or c_n == "NM_ESCOLA"), None)
        col_turma = next((c_orig for c_n, c_orig in cols_norm.items() if "EXPORT TURMA" in c_n or c_n == "TURMA" or c_n == "NM_TURMA"), None)
        col_aluno = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ALUNO", "ESTUDANTE", "NOME", "RA"]) and "TOTAL" not in c_n and "PROFESSOR" not in c_n), None)
        col_part = next((c_orig for c_n, c_orig in cols_norm.items() if "PARTICIPACAO" in c_n or "FREQ" in c_n), None)
        col_total_al = next((c_orig for c_n, c_orig in cols_norm.items() if "TOTAL ALUNO" in c_n or "ESTUDANTE" == c_n), None)

        registros = []
        for _, row in df_raw.iterrows():
            valores_linha_str = " ".join([str(v) for v in row.values if pd.notna(v)])
            if not valores_linha_str.strip() or "Filtros aplicados" in valores_linha_str:
                continue

            val_esc = str(row[col_escola]).strip() if col_escola and pd.notna(row[col_escola]) else ""
            val_tur = str(row[col_turma]).strip() if col_turma and pd.notna(row[col_turma]) else ""
            val_alu = str(row[col_aluno]).strip() if col_aluno and pd.notna(row[col_aluno]) else ""
            val_prof = str(row[col_prof]).strip() if col_prof and pd.notna(row[col_prof]) else ""
            val_rec = normalizar_texto(row[col_recomp]) in ["SIM", "TRUE", "1", "S", "X"] if col_recomp else False

            if any(x.lower() in ["total", "none", "nan"] for x in [val_esc, val_tur, val_alu] if x):
                continue

            if col_aluno and val_alu:
                nivel_gran = "ESTUDANTE"
                escola_row = corresponder_escola_oficial(val_esc) if val_esc else escola_padrao
                turma_row = val_tur if val_tur else (turma_rodape or "Turma Geral")
                estudante_row = val_alu
            elif col_turma and val_tur:
                nivel_gran = "TURMA"
                escola_row = corresponder_escola_oficial(val_esc) if val_esc else escola_padrao
                turma_row = val_tur
                qtd_al = str(row[col_total_al]).strip() if col_total_al and pd.notna(row[col_total_al]) else "-"
                estudante_row = f"Consolidado Turma ({qtd_al} alunos)"
            elif col_escola and val_esc:
                nivel_gran = "ESCOLA"
                escola_row = corresponder_escola_oficial(val_esc)
                turma_row = "Consolidado da Escola"
                estudante_row = "Consolidado Escola"
            else:
                continue

            part_pct = converter_para_percentual(row[col_part]) if col_part else None

            for col_sigla in siglas_presentes:
                acerto_pct = converter_para_percentual(row[col_sigla])
                if acerto_pct is None:
                    continue
                comp_nome = MAPA_SIGLAS_BI[normalizar_texto(col_sigla)]
                registros.append({
                    "Tipo_Indicador": tipo_ind,
                    "Subcategoria": (str(row[col_sub_row]).strip() if col_sub_row and pd.notna(row[col_sub_row]) and str(row[col_sub_row]).strip() else subcategoria),
                    "Periodo": (str(row[col_per_row]).strip() if col_per_row and pd.notna(row[col_per_row]) and str(row[col_per_row]).strip() else periodo_ref),
                    "Nivel_Granularidade": nivel_gran,
                    "Escola": escola_row,
                    "Turma": turma_row,
                    "Estudante": estudante_row,
                    "Professor": val_prof,
                    "Recomposicao": val_rec,
                    "Componente": comp_nome,
                    "Proficiencia_Pct": acerto_pct,
                    "Participacao_Pct": part_pct,
                    "Nivel": classificar_farol_nivel(acerto_pct)
                })
        return pd.DataFrame(registros)

    col_esc = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ESCOLA", "UNIDADE", "UE"])), None)
    col_tur = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["TURMA", "SERIE", "ANO"])), None)
    col_alu = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ALUNO", "ESTUDANTE", "DISCENTE", "NOME"]) and "TOTAL" not in c_n and "PROFESSOR" not in c_n), None)
    col_comp = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["COMPONENTE", "DISCIPLINA", "AREA"])), None)
    col_val = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["FREQ", "USO", "ENGAJAMENTO", "TAXA", "ACERTO", "MEDIA", "NOTA", "PERCENTUAL", "PROFICIENCIA", "CONCLUSAO", "ATIVIDADE"])), None)

    registros_b = []
    for _, row in df_raw.iterrows():
        valores_linha_str = " ".join([str(v) for v in row.values if pd.notna(v)])
        if not valores_linha_str.strip() or "Filtros aplicados" in valores_linha_str:
            continue

        val_esc = str(row[col_esc]).strip() if col_esc and pd.notna(row[col_esc]) else ""
        val_tur = str(row[col_tur]).strip() if col_tur and pd.notna(row[col_tur]) else ""
        val_alu = str(row[col_alu]).strip() if col_alu and pd.notna(row[col_alu]) else ""
        val_prof = str(row[col_prof]).strip() if col_prof and pd.notna(row[col_prof]) else ""
        val_rec = normalizar_texto(row[col_recomp]) in ["SIM", "TRUE", "1", "S", "X"] if col_recomp else False

        if any(x.lower() in ["total", "none", "nan"] for x in [val_esc, val_tur, val_alu] if x):
            continue

        if col_alu and val_alu:
            nivel_gran = "ESTUDANTE"
            escola_row = corresponder_escola_oficial(val_esc) if val_esc else escola_padrao
            turma_row = val_tur if val_tur else (turma_rodape or "Turma Geral")
            estudante_row = val_alu
        elif col_tur and val_tur:
            nivel_gran = "TURMA"
            escola_row = corresponder_escola_oficial(val_esc) if val_esc else escola_padrao
            turma_row = val_tur
            estudante_row = "Consolidado Turma"
        elif col_esc and val_esc:
            nivel_gran = "ESCOLA"
            escola_row = corresponder_escola_oficial(val_esc)
            turma_row = "Consolidado da Escola"
            estudante_row = "Consolidado Escola"
        else:
            continue

        val_pct = converter_para_percentual(row[col_val]) if col_val else None
        if val_pct is None:
            continue

        comp_row = str(row[col_comp]).strip() if col_comp and pd.notna(row[col_comp]) else ("Frequência" if tipo_ind == "FREQUENCIA" else subcategoria)
        registros_b.append({
            "Tipo_Indicador": tipo_ind,
            "Subcategoria": (str(row[col_sub_row]).strip() if col_sub_row and pd.notna(row[col_sub_row]) and str(row[col_sub_row]).strip() else subcategoria),
            "Periodo": (str(row[col_per_row]).strip() if col_per_row and pd.notna(row[col_per_row]) and str(row[col_per_row]).strip() else periodo_ref),
            "Nivel_Granularidade": nivel_gran,
            "Escola": escola_row,
            "Turma": turma_row,
            "Estudante": estudante_row,
            "Professor": val_prof,
            "Recomposicao": val_rec,
            "Componente": comp_row,
            "Proficiencia_Pct": val_pct,
            "Participacao_Pct": val_pct if tipo_ind == "FREQUENCIA" else None,
            "Nivel": classificar_farol_nivel(val_pct)
        })

    return pd.DataFrame(registros_b)

# ==============================================================================
# MOTOR DE PRIORIZAÇÃO PÓS-FILTRO POR INDICADOR (PROVA PAULISTA, SARESP, FREQUÊNCIA, PLATAFORMA, TODOS)
# ==============================================================================
OPCOES_CRITERIO_PRIORIDADE = {
    "TODOS": "Todos os Indicadores (Análise Multi-Indicador Consolidada)",
    "PP_3BIM": "Prova Paulista — 3º Bimestre",
    "PP_1BIM": "Prova Paulista — 1º Bimestre",
    "PP_2BIM": "Prova Paulista — 2º Bimestre",
    "PP_4BIM": "Prova Paulista — 4º Bimestre",
    "PP_TODAS": "Prova Paulista (Todos os Bimestres)",
    "SARESP": "SARESP",
    "FREQUENCIA": "Frequência Escolar",
    "PLATAFORMA": "Plataformas Digitais (Todas)",
    "DIAGNOSTICA": "Avaliação Diagnóstica",
    "SAEB": "SAEB",
    "ENEM": "ENEM"
}

def filtrar_base_por_criterio_prioridade(df_ind, criterio_chave):
    if df_ind is None or df_ind.empty:
        return pd.DataFrame()
    df_b = df_ind.copy()

    if criterio_chave == "TODOS":
        return df_b
    elif criterio_chave == "PP_3BIM":
        return df_b[
            (df_b["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))) &
            (df_b["Periodo"].apply(lambda x: "3" in normalizar_texto(x)))
        ]
    elif criterio_chave == "PP_1BIM":
        return df_b[
            (df_b["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))) &
            (df_b["Periodo"].apply(lambda x: "1" in normalizar_texto(x)))
        ]
    elif criterio_chave == "PP_2BIM":
        return df_b[
            (df_b["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))) &
            (df_b["Periodo"].apply(lambda x: "2" in normalizar_texto(x)))
        ]
    elif criterio_chave == "PP_4BIM":
        return df_b[
            (df_b["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))) &
            (df_b["Periodo"].apply(lambda x: "4" in normalizar_texto(x)))
        ]
    elif criterio_chave == "PP_TODAS":
        return df_b[df_b["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))]
    elif criterio_chave == "SARESP":
        return df_b[df_b["Subcategoria"].apply(lambda x: "SARESP" in normalizar_texto(x))]
    elif criterio_chave == "FREQUENCIA":
        return df_b[df_b["Tipo_Indicador"] == "FREQUENCIA"]
    elif criterio_chave == "PLATAFORMA":
        return df_b[df_b["Tipo_Indicador"] == "PLATAFORMA"]
    elif criterio_chave == "DIAGNOSTICA":
        return df_b[df_b["Subcategoria"].apply(lambda x: "DIAGNOSTICA" in normalizar_texto(x))]
    elif criterio_chave == "SAEB":
        return df_b[df_b["Subcategoria"].apply(lambda x: "SAEB" in normalizar_texto(x))]
    elif criterio_chave == "ENEM":
        return df_b[df_b["Subcategoria"].apply(lambda x: "ENEM" in normalizar_texto(x))]
    elif criterio_chave.startswith("PLAT::"):
        nome_plat = normalizar_texto(criterio_chave.split("::", 1)[1])
        return df_b[(df_b["Tipo_Indicador"] == "PLATAFORMA") & (df_b["Subcategoria"].apply(lambda x: nome_plat in normalizar_texto(x)))]
    return df_b

def calcular_metricas_por_escola_pos_filtro(df_filtrado, criterio_chave):
    """
    Calcula a média e o resumo detalhado por escola após aplicar o filtro de indicador.
    Se criterio_chave == 'TODOS', analisa equilibradamente Avaliações + Plataformas + Frequência.
    """
    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame(columns=["Escola", "Indicador_Medio_Pct", "Indice_Criticidade_Pct", "Detalhe_Indicador", "Farol"])

    df_esc = df_filtrado[df_filtrado["Nivel_Granularidade"] == "ESCOLA"]
    if df_esc.empty:
        df_esc = df_filtrado

    if criterio_chave == "TODOS":
        linhas_multi = []
        for esc, grp_e in df_esc.groupby("Escola"):
            partes_resumo = []
            medias_dim = []
            for tipo_d, rot_d in [("AVALIACAO", "Avaliações"), ("PLATAFORMA", "Plataformas"), ("FREQUENCIA", "Frequência")]:
                sub_d = grp_e[grp_e["Tipo_Indicador"] == tipo_d]
                if not sub_d.empty:
                    m_d = round(sub_d["Proficiencia_Pct"].mean(), 1)
                    medias_dim.append(m_d)
                    partes_resumo.append(f"{rot_d}: {m_d:.1f}%")
            if medias_dim:
                media_composta = round(float(np.mean(medias_dim)), 1)
                linhas_multi.append({
                    "Escola": esc,
                    "Indicador_Medio_Pct": media_composta,
                    "Indice_Criticidade_Pct": round(100.0 - media_composta, 1),
                    "Detalhe_Indicador": " | ".join(partes_resumo),
                    "Farol": classificar_farol_nivel(media_composta)
                })
        return pd.DataFrame(linhas_multi)

    grp = df_esc.groupby("Escola", as_index=False).agg(
        Indicador_Medio_Pct=("Proficiencia_Pct", "mean")
    )
    grp["Indicador_Medio_Pct"] = grp["Indicador_Medio_Pct"].round(1)
    grp["Indice_Criticidade_Pct"] = (100.0 - grp["Indicador_Medio_Pct"]).round(1)
    nome_crit = OPCOES_CRITERIO_PRIORIDADE.get(criterio_chave, criterio_chave.replace("PLAT::", "Plataforma: "))
    grp["Detalhe_Indicador"] = grp["Indicador_Medio_Pct"].apply(lambda v: f"{nome_crit}: {v:.1f}%")
    grp["Farol"] = grp["Indicador_Medio_Pct"].apply(classificar_farol_nivel)
    return grp

def obter_ranking_escolas_prioritarias(df_ind, pasta_pec, escolas_delimitadas, criterio_chave="TODOS", df_recomp_sheets=None):
    # Recomposição com lista pré-estabelecida quando o filtro geral "TODOS" está ativo
    if pasta_pec == "RECOMPOSICAO" and criterio_chave == "TODOS" and df_recomp_sheets is not None and not df_recomp_sheets.empty:
        cols_r = {normalizar_texto(c): c for c in df_recomp_sheets.columns}
        col_esc_r = next((v for k, v in cols_r.items() if "ESCOLA" in k or "UNIDADE" in k), None)
        if col_esc_r:
            df_r_clean = df_recomp_sheets.dropna(subset=[col_esc_r]).copy()
            df_r_clean["Escola_Oficial"] = df_r_clean[col_esc_r].apply(corresponder_escola_oficial)
            contagem = df_r_clean.groupby("Escola_Oficial").size().reset_index(name="Qtd_Estudantes_Recomposicao")
            contagem = contagem.sort_values(by="Qtd_Estudantes_Recomposicao", ascending=False).head(30).reset_index(drop=True)
            contagem.insert(0, "Prioridade", [f"{i+1}º" for i in range(len(contagem))])
            contagem["Detalhe_Indicador"] = contagem["Qtd_Estudantes_Recomposicao"].apply(lambda q: f"{q} estudantes em Recomposição")
            contagem["Farol"] = "Lista Prioritária Recomposição"
            return contagem.rename(columns={"Escola_Oficial": "Escola"})

    df_filtrado = filtrar_base_por_criterio_prioridade(df_ind, criterio_chave)
    df_metricas = calcular_metricas_por_escola_pos_filtro(df_filtrado, criterio_chave)
    nome_crit_label = OPCOES_CRITERIO_PRIORIDADE.get(criterio_chave, criterio_chave.replace("PLAT::", "Plataforma: "))

    # PEC Qualidade da Aula: ordena sua carteira de escolas atribuídas com base no indicador filtrado
    if pasta_pec == "QUALIDADE DA AULA" and escolas_delimitadas:
        esc_oficiais = [corresponder_escola_oficial(e) for e in escolas_delimitadas if str(e).strip()]
        rows_q = []
        for esc in esc_oficiais:
            media_val = np.nan
            det_val = f"Sem dado em {nome_crit_label}"
            farol_val = "Sem Dado no Filtro"
            if not df_metricas.empty:
                sub_e = df_metricas[df_metricas["Escola"].apply(lambda x: normalizar_texto(esc) in normalizar_texto(x) or normalizar_texto(x) in normalizar_texto(esc))]
                if not sub_e.empty:
                    media_val = sub_e.iloc[0]["Indicador_Medio_Pct"]
                    det_val = sub_e.iloc[0]["Detalhe_Indicador"]
                    farol_val = sub_e.iloc[0]["Farol"]
            rows_q.append({
                "Escola": esc,
                "Indicador_Medio_Pct": media_val,
                "Indice_Criticidade_Pct": (100.0 - media_val) if pd.notna(media_val) else -1.0,
                "Detalhe_Indicador": det_val,
                "Farol": farol_val
            })
        df_q = pd.DataFrame(rows_q)
        df_q = df_q.sort_values(by=["Indice_Criticidade_Pct", "Escola"], ascending=[False, True]).reset_index(drop=True)
        df_q.insert(0, "Prioridade", [f"{i+1}º" for i in range(len(df_q))])
        return df_q

    # Demais pastas (ou CEC/AT): ranqueia as 30 escolas da URE pós-filtro
    if not df_metricas.empty:
        grp = df_metricas.copy()
        escolas_presentes_norm = {normalizar_texto(e) for e in grp["Escola"]}
        faltantes = [
            {
                "Escola": esc,
                "Indicador_Medio_Pct": np.nan,
                "Indice_Criticidade_Pct": -1.0,
                "Detalhe_Indicador": f"Sem dado em {nome_crit_label}",
                "Farol": "Sem Dado no Filtro"
            }
            for esc in ESCOLAS_URE_MIRACATU
            if normalizar_texto(esc) not in escolas_presentes_norm
        ]
        if faltantes:
            grp = pd.concat([grp, pd.DataFrame(faltantes)], ignore_index=True)

        grp = grp.sort_values(by=["Indice_Criticidade_Pct", "Escola"], ascending=[False, True]).head(30).reset_index(drop=True)
        grp.insert(0, "Prioridade", [f"{i+1}º" for i in range(len(grp))])
        return grp

    return pd.DataFrame({
        "Prioridade": [f"{i+1}º" for i in range(len(ESCOLAS_URE_MIRACATU[:30]))],
        "Escola": ESCOLAS_URE_MIRACATU[:30],
        "Indicador_Medio_Pct": [np.nan] * len(ESCOLAS_URE_MIRACATU[:30]),
        "Indice_Criticidade_Pct": [-1.0] * len(ESCOLAS_URE_MIRACATU[:30]),
        "Detalhe_Indicador": [f"Aguardando planilhas de {nome_crit_label} no Drive"] * len(ESCOLAS_URE_MIRACATU[:30]),
        "Farol": ["Sem Dado no Filtro"] * len(ESCOLAS_URE_MIRACATU[:30])
    })

# ==============================================================================
# CONSTRUTOR DO PLANNER MENSAL VISUAL (ESTILO GOOGLE AGENDA - MÊS INTEIRO)
# ==============================================================================
def parsear_data_flexivel(texto_data, ano_padrao=None, mes_padrao=None):
    if not texto_data or pd.isna(texto_data):
        return None
    s = str(texto_data).strip()
    m1 = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m1:
        try:
            return date(int(m1.group(3)), int(m1.group(2)), int(m1.group(1)))
        except Exception:
            return None
    m2 = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m2:
        try:
            return date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except Exception:
            return None
    m3 = re.search(r"(\d{1,2})/(\d{1,2})", s)
    if m3 and ano_padrao:
        try:
            return date(int(ano_padrao), int(m3.group(2)), int(m3.group(1)))
        except Exception:
            return None
    return None

def coletar_eventos_mes_planner(ano, mes, df_bloqueios, df_agenda_seduc, eventos_gcal, cronograma_pec, pasta_pec=""):
    """
    Reúne automaticamente para o mês inteiro:
    - Feriados oficiais SP (vermelho)
    - Convocações / Bloqueios URE (laranja/amarelo)
    - Lives e Agenda SEDUC / EEC (roxo/azul)
    - Compromissos Google Agenda CEC (amarelo/dourado)
    - Visitas / Ações agendadas pelo PEC no formulário (verde)
    """
    eventos_por_dia = {}

    def add_ev(dt, tipo, horario, titulo, detalhe=""):
        if not dt or dt.year != ano or dt.month != mes:
            return
        eventos_por_dia.setdefault(dt.day, []).append({
            "tipo": tipo,
            "horario": horario,
            "titulo": titulo,
            "detalhe": detalhe
        })

    # 1. Feriados Oficiais SP (texto sucinto no Planner)
    feriados_sp = holidays.country_holidays("BR", subdiv="SP", years=[ano])
    for d_fer, nome_fer in feriados_sp.items():
        if d_fer.year == ano and d_fer.month == mes:
            nome_curto = str(nome_fer).replace("Dia Nacional de Zumbi e da ", "").replace("Dia de ", "").strip()
            add_ev(d_fer, "FERIADO", "", f"FERIADO {nome_curto}")

    # 2. Bloqueios / Convocações URE (aba Bloqueios_Agenda)
    if df_bloqueios is not None and not df_bloqueios.empty:
        cols_b = {normalizar_texto(c): c for c in df_bloqueios.columns}
        c_data = next((v for k, v in cols_b.items() if "DATA" in k or "DIA" in k), None)
        c_hora = next((v for k, v in cols_b.items() if "HORA" in k or "TURNO" in k), None)
        c_tit = next((v for k, v in cols_b.items() if any(x in k for x in ["EVENTO", "CONVOCACAO", "MOTIVO", "DESCRICAO", "TITULO", "PAUTA"])), None)
        if c_data:
            for _, row in df_bloqueios.iterrows():
                dt_b = parsear_data_flexivel(row[c_data], ano, mes)
                if dt_b:
                    h_b = str(row[c_hora]).strip() if c_hora and pd.notna(row[c_hora]) else "Integral / URE"
                    t_b = str(row[c_tit]).strip() if c_tit and pd.notna(row[c_tit]) else "Convocação / Bloqueio URE"
                    add_ev(dt_b, "CONVOCACAO", h_b, f"📢 {t_b}")

    # 3. Agenda SEDUC / EEC e Lives (planilha SEDUC)
    if df_agenda_seduc is not None and not df_agenda_seduc.empty:
        cols_s = {normalizar_texto(c): c for c in df_agenda_seduc.columns}
        c_data_s = next((v for k, v in cols_s.items() if "DATA" in k or "DIA" in k), None)
        c_hora_s = next((v for k, v in cols_s.items() if "HORA" in k or "HORARIO" in k), None)
        c_tit_s = next((v for k, v in cols_s.items() if any(x in k for x in ["PAUTA", "ASSUNTO", "EVENTO", "LIVE", "ACAO", "DESCRICAO", "TEMA"])), None)
        c_pub_s = next((v for k, v in cols_s.items() if any(x in k for x in ["PASTA", "PUBLICO", "AREA"])), None)
        if c_data_s:
            for _, row in df_agenda_seduc.iterrows():
                dt_s = parsear_data_flexivel(row[c_data_s], ano, mes)
                if dt_s:
                    h_s = str(row[c_hora_s]).strip() if c_hora_s and pd.notna(row[c_hora_s]) else "Horário SEDUC"
                    t_s = str(row[c_tit_s]).strip() if c_tit_s and pd.notna(row[c_tit_s]) else "Agenda SEDUC / Live"
                    pub_s = str(row[c_pub_s]).strip() if c_pub_s and pd.notna(row[c_pub_s]) else ""
                    icone = "📺" if "LIVE" in normalizar_texto(t_s) or "CMSP" in normalizar_texto(t_s) else "🏛️"
                    add_ev(dt_s, "SEDUC_LIVE", h_s, f"{icone} {t_s}", pub_s)

    # 4. Google Agenda CEC (iCal)
    for ev in (eventos_gcal or []):
        if ev.get("ano") == ano and ev.get("mes") == mes:
            dt_g = date(ano, mes, int(ev["dia"]))
            res_g = str(ev.get("resumo", "Compromisso URE")).strip()
            if "FERIADO" in normalizar_texto(res_g) or "CONSCIENCIA NEGRA" in normalizar_texto(res_g) or "PROCLAMACAO DA REPUBLICA" in normalizar_texto(res_g):
                nome_f_g = re.sub(r"(?i)^feriado\s*:\s*", "", res_g).replace("Dia Nacional de Zumbi e da ", "").replace("Dia de ", "").strip()
                add_ev(dt_g, "FERIADO", "", f"FERIADO {nome_f_g}")
            else:
                add_ev(dt_g, "GCAL", ev.get("horario", ""), f"📅 {res_g}", ev.get("local", ""))

    # 5. Visitas e Atividades lançadas pelo PEC (atualização imediata no planner!)
    for idx_item, item in enumerate(cronograma_pec or []):
        dt_p = parsear_data_flexivel(item.get("data_iso") or item.get("dia", ""), ano, mes)
        if dt_p:
            loc_p = item.get("local", "Escola")
            if normalizar_texto(loc_p) == "URE MIRACATU":
                add_ev(dt_p, "URE_INTERNO", "", "🏢 URE Miracatu", "")
            else:
                add_ev(dt_p, "VISITA_PEC", "", f"🏫 {loc_p}", "")

    return eventos_por_dia

def renderizar_html_planner_mensal(ano, mes, eventos_por_dia):
    """Renderiza uma grade mensal completa (Domingo a Sábado) inspirada no Google Agenda."""
    cal = calendar.Calendar(firstweekday=6)  # Começa no Domingo (DOM, SEG, TER, QUA, QUI, SEX, SÁB)
    semanas = cal.monthdayscalendar(ano, mes)
    nomes_meses = [
        "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    estilos_badge = {
        "FERIADO": "background:#FFEBEE; color:#B71C1C; border-left:3px solid #D32F2F;",
        "CONVOCACAO": "background:#FFF3E0; color:#E65100; border-left:3px solid #F57C00;",
        "SEDUC_LIVE": "background:#EDE7F6; color:#4527A0; border-left:3px solid #5E35B1;",
        "GCAL": "background:#FFF8E1; color:#F57F17; border-left:3px solid #FBC02D;",
        "VISITA_PEC": "background:#E8F5E9; color:#1B5E20; border-left:3px solid #2E7D32;",
        "URE_INTERNO": "background:#F9A825; color:#212121; font-weight:600; border-left:3px solid #E65100;"
    }

    html = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; border: 1px solid #dadce0; border-radius: 8px; overflow: hidden; background: #fff;">
        <div style="background: #f8f9fa; padding: 10px 14px; border-bottom: 1px solid #dadce0; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 16px; font-weight: 600; color: #202124;">🗓️ {nomes_meses[mes]} de {ano}</span>
            <span style="font-size: 11px; color: #5f6368;">
                <span style="background:#FFEBEE;color:#B71C1C;padding:2px 5px;border-radius:3px;">🚩 Feriado</span>
                <span style="background:#FFF3E0;color:#E65100;padding:2px 5px;border-radius:3px;">📢 Convocação</span>
                <span style="background:#EDE7F6;color:#4527A0;padding:2px 5px;border-radius:3px;">📺 Live/SEDUC</span>
                <span style="background:#E8F5E9;color:#1B5E20;padding:2px 5px;border-radius:3px;">🏫 Visita PEC</span>
                <span style="background:#F9A825;color:#212121;font-weight:600;padding:2px 5px;border-radius:3px;">🏢 URE Miracatu</span>
            </span>
        </div>
        <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
            <thead>
                <tr style="background: #f1f3f4; color: #5f6368; font-size: 11px; text-align: center;">
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">DOM.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">SEG.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">TER.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">QUA.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">QUI.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">SEX.</th>
                    <th style="padding: 6px; border: 1px solid #e0e0e0;">SÁB.</th>
                </tr>
            </thead>
            <tbody>
    """

    hoje = date.today()
    for sem in semanas:
        html += "<tr>"
        for idx_col, dia_num in enumerate(sem):
            if dia_num == 0:
                html += '<td style="height: 98px; background: #fafafa; border: 1px solid #e0e0e0;"></td>'
            else:
                eh_hoje = (hoje.year == ano and hoje.month == mes and hoje.day == dia_num)
                bg_cel = "#f1f3f4" if idx_col in [0, 6] else "#ffffff"
                badge_dia = (
                    f'<span style="background:#1a73e8;color:#fff;border-radius:50%;padding:1px 6px;font-weight:bold;">{dia_num}</span>'
                    if eh_hoje else f'<span style="font-weight:600;color:#3c4043;">{dia_num}</span>'
                )
                evs = eventos_por_dia.get(dia_num, [])
                evs_html = ""
                for ev in evs:
                    st_b = estilos_badge.get(ev["tipo"], "background:#E3F2FD;color:#0D47A1;")
                    hora_txt = f"<b>{ev['horario']}</b> — " if ev["horario"] else ""
                    tit_txt = ev["titulo"]
                    det_txt = f" ({ev['detalhe']})" if ev.get("detalhe") else ""
                    evs_html += f'<div style="{st_b} font-size: 10px; line-height: 1.25; padding: 2px 4px; margin-top: 2px; border-radius: 3px; word-wrap: break-word;">{hora_txt}{tit_txt}{det_txt}</div>'

                html += f'<td style="height: 98px; vertical-align: top; padding: 4px; background: {bg_cel}; border: 1px solid #e0e0e0;">{badge_dia}{evs_html}</td>'
        html += "</tr>"

    html += "</tbody></table></div>"
    return html

# ==============================================================================
# ESTILIZAÇÃO VISUAL DO FAROL (CORES POR PROFICIÊNCIA / FREQUÊNCIA / USO)
# ==============================================================================
def estilo_celula_farol(val):
    if pd.isna(val):
        return ""
    s = str(val).replace("%", "").replace(",", ".").strip()
    try:
        num = float(s)
        if num <= 40.0:
            return "background-color: #FFCDD2; color: #B71C1C; font-weight: bold;"
        elif num <= 60.0:
            return "background-color: #FFF9C4; color: #F57F17; font-weight: bold;"
        elif num <= 80.0:
            return "background-color: #C8E6C9; color: #1B5E20; font-weight: bold;"
        else:
            return "background-color: #BBDEFB; color: #0D47A1; font-weight: bold;"
    except Exception:
        if "Vermelho" in s or "Crítico" in s:
            return "background-color: #FFCDD2; color: #B71C1C; font-weight: bold;"
        elif "Amarelo" in s or "Atenção" in s:
            return "background-color: #FFF9C4; color: #F57F17; font-weight: bold;"
        elif "Verde" in s or "Adequado" in s:
            return "background-color: #C8E6C9; color: #1B5E20; font-weight: bold;"
        elif "Azul" in s or "Avançado" in s:
            return "background-color: #BBDEFB; color: #0D47A1; font-weight: bold;"
        return ""

def renderizar_tabela_hierarquica(df_sub, titulo_secao, rotulo_metrica="Média (%)"):
    if df_sub.empty:
        st.info(f"Nenhum dado carregado para **{titulo_secao}**. Deposite as planilhas na pasta correspondente no Google Drive.")
        return

    st.caption("🎨 **Legenda do Farol:** 🔴 **0% a 40%** (Crítico) | 🟡 **40,1% a 60%** (Atenção) | 🟢 **60,1% a 80%** (Adequado) | 🔵 **80,1% a 100%** (Avançado)")

    st.markdown(f"#### 🏫 Visão Geral por Escola — {titulo_secao}")
    df_esc_base = df_sub[df_sub["Nivel_Granularidade"] == "ESCOLA"]
    if df_esc_base.empty:
        df_esc_base = df_sub

    piv_esc = df_esc_base.pivot_table(
        index="Escola",
        columns="Componente",
        values="Proficiencia_Pct",
        aggfunc="mean"
    ).round(1)
    piv_esc[rotulo_metrica] = piv_esc.mean(axis=1).round(1)
    cols_ord = [rotulo_metrica] + [c for c in piv_esc.columns if c != rotulo_metrica]
    piv_esc = piv_esc[cols_ord].sort_values(by=rotulo_metrica, ascending=True).reset_index()

    cols_num_esc = [c for c in piv_esc.columns if c != "Escola"]
    styler_esc = piv_esc.style.map(estilo_celula_farol, subset=cols_num_esc).format({c: "{:.1f}%" for c in cols_num_esc}, na_rep="-")
    st.dataframe(styler_esc, use_container_width=True, hide_index=True)

    escolas_disp = piv_esc["Escola"].dropna().tolist()
    escola_escolhida = st.selectbox(f"🔍 Entrar na Escola (ver Turmas e Estudantes):", ["Selecione uma escola..."] + escolas_disp, key=f"sel_esc_{titulo_secao}")

    if escola_escolhida != "Selecione uma escola...":
        df_da_escola = df_sub[df_sub["Escola"] == escola_escolhida]
        st.markdown(f"#### 📋 Geral por Turma — {escola_escolhida}")

        df_turma_base = df_da_escola[df_da_escola["Nivel_Granularidade"].isin(["TURMA", "ESTUDANTE"])]
        if not df_turma_base.empty:
            piv_tur = df_turma_base.pivot_table(
                index="Turma",
                columns="Componente",
                values="Proficiencia_Pct",
                aggfunc="mean"
            ).round(1)
            piv_tur[rotulo_metrica] = piv_tur.mean(axis=1).round(1)
            cols_ord_t = [rotulo_metrica] + [c for c in piv_tur.columns if c != rotulo_metrica]
            piv_tur = piv_tur[cols_ord_t].sort_values(by=rotulo_metrica, ascending=True).reset_index()

            cols_num_tur = [c for c in piv_tur.columns if c != "Turma"]
            styler_tur = piv_tur.style.map(estilo_celula_farol, subset=cols_num_tur).format({c: "{:.1f}%" for c in cols_num_tur}, na_rep="-")
            st.dataframe(styler_tur, use_container_width=True, hide_index=True)

            turmas_disp = [t for t in piv_tur["Turma"].dropna().tolist() if "Consolidado" not in t]
            if turmas_disp:
                turma_escolhida = st.selectbox(f"🎓 Entrar na Turma (Visão Estudante por Estudante):", ["Selecione uma turma..."] + turmas_disp, key=f"sel_tur_{titulo_secao}_{escola_escolhida}")
                if turma_escolhida != "Selecione uma turma...":
                    df_alunos = df_da_escola[(df_da_escola["Turma"] == turma_escolhida) & (df_da_escola["Nivel_Granularidade"] == "ESTUDANTE")]
                    st.markdown(f"#### 🧑‍🎓 Visão Estudante por Estudante — {turma_escolhida}")
                    if not df_alunos.empty:
                        piv_alu = df_alunos.pivot_table(
                            index="Estudante",
                            columns="Componente",
                            values="Proficiencia_Pct",
                            aggfunc="mean"
                        ).round(1)
                        piv_alu[rotulo_metrica] = piv_alu.mean(axis=1).round(1)
                        cols_ord_a = [rotulo_metrica] + [c for c in piv_alu.columns if c != rotulo_metrica]
                        piv_alu = piv_alu[cols_ord_a].sort_values(by=rotulo_metrica, ascending=True).reset_index()

                        cols_num_alu = [c for c in piv_alu.columns if c != "Estudante"]
                        styler_alu = piv_alu.style.map(estilo_celula_farol, subset=cols_num_alu).format({c: "{:.1f}%" for c in cols_num_alu}, na_rep="-")
                        st.dataframe(styler_alu, use_container_width=True, hide_index=True)
                    else:
                        st.info("Para esta turma só foi carregada a planilha consolidada por turma. Suba a planilha nominal de estudantes da turma no Drive para abrir a visão estudante por estudante.")
        else:
            st.info("Para esta escola só foi carregada a planilha consolidada geral da escola. Suba as planilhas por turma/estudante na pasta desta escola no Drive.")

# ==============================================================================
# CARREGAMENTO DE PLANILHAS GOOGLE E DRIVE
# ==============================================================================
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
    if not api_key or "COLE_AQUI" in api_key:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0

    dfs_ind, dfs_curs, dfs_atpc_drv = [], [], []
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

                    if n_low.endswith(".pdf") or f_mime == "application/pdf":
                        leitor = PdfReader(arq_bytes)
                        texto_pdf = "\n".join([p.extract_text() or "" for p in leitor.pages])
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

                    elif (n_low.endswith(".docx") or f_mime == "application/vnd.google-apps.document") and not n_low.startswith("00_modelo"):
                        doc_p = Document(arq_bytes)
                        texto_doc = "\n".join([p.text for p in doc_p.paragraphs])
                        for t in doc_p.tables:
                            for r in t.rows:
                                texto_doc += "\n" + " | ".join([c.text.strip() for c in r.cells])

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

                    elif n_low.endswith(".csv"):
                        df_l = pd.read_csv(arq_bytes)
                        df_a = padronizar_planilha_atpc(df_l)
                        if not df_a.empty:
                            dfs_atpc_drv.append(df_a)
                            qtd_arquivos += 1
                        else:
                            df_p = padronizar_planilha_bruta(df_l, f_name, caminho_pasta)
                            if not df_p.empty:
                                (dfs_curs if "multiplica" in n_low or "cursista" in n_low else dfs_ind).append(df_p)
                                qtd_arquivos += 1
                    elif n_low.endswith(".xlsx") or "spreadsheet" in f_mime:
                        abas_excel = pd.read_excel(arq_bytes, sheet_name=None)
                        contou_arq = False
                        for nome_aba_xls, df_l in abas_excel.items():
                            if df_l is None or df_l.empty:
                                continue
                            df_a = padronizar_planilha_atpc(df_l)
                            if not df_a.empty:
                                dfs_atpc_drv.append(df_a)
                                contou_arq = True
                            else:
                                df_p = padronizar_planilha_bruta(df_l, f"{f_name} - {nome_aba_xls}", caminho_pasta)
                                if not df_p.empty:
                                    (dfs_curs if "multiplica" in n_low or "cursista" in n_low or "multiplica" in nome_aba_xls.lower() else dfs_ind).append(df_p)
                                    contou_arq = True
                        if contou_arq:
                            qtd_arquivos += 1
                except Exception:
                    pass

    varrer_pasta(folder_id, "")
    df_ind_final = pd.concat(dfs_ind, ignore_index=True) if dfs_ind else pd.DataFrame()
    df_curs_final = pd.concat(dfs_curs, ignore_index=True) if dfs_curs else pd.DataFrame()
    df_atpc_final = pd.concat(dfs_atpc_drv, ignore_index=True).drop_duplicates() if dfs_atpc_drv else pd.DataFrame()
    df_planos_final = pd.DataFrame(registros_planos).drop_duplicates() if registros_planos else pd.DataFrame()
    df_termos_final = pd.DataFrame(registros_termos).drop_duplicates() if registros_termos else pd.DataFrame()
    return df_ind_final, df_curs_final, df_atpc_final, df_planos_final, df_termos_final, qtd_arquivos

df_usuarios = carregar_aba("Usuarios")
df_pastas_raw = carregar_aba("Config_Pastas")
if df_pastas_raw.empty:
    df_pastas_raw = carregar_aba("Confing_Pastas")
df_eventos = carregar_aba("Calendario_Eventos")
df_bloqueios = carregar_aba("Bloqueios_Agenda")
df_agenda_seduc = carregar_agenda_seduc_eec()

# Abas para ATPC, Atribuição Docente e Recomposição
df_atpc_sheets_raw = carregar_aba("ATPC_Escolas")
if df_atpc_sheets_raw.empty:
    df_atpc_sheets_raw = carregar_aba("Horarios_ATPC")
if df_atpc_sheets_raw.empty:
    df_atpc_sheets_raw = carregar_aba("ATPC")
df_atpc_sheets = padronizar_planilha_atpc(df_atpc_sheets_raw)

df_atribuicao_docentes = carregar_aba("Atribuicao_Docentes")
if df_atribuicao_docentes.empty:
    df_atribuicao_docentes = carregar_aba("Professores_Componentes")

df_recomposicao_sheets = carregar_aba("Recomposicao_Alunos")
if df_recomposicao_sheets.empty:
    df_recomposicao_sheets = carregar_aba("Alunos_Recomposicao")

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

COMANDOS_CURTOS_ATIVIDADES = [
    "Apoio Presencial",
    "Alinhamento com o/a Coordenador(a)",
    "Alinhamento 1x1",
    "Alinhamento de Equipe",
    "Organização de Formação Presencial",
    "Formação em Serviço (CGP/CGPG)",
    "Observação de Aula",
    "Participação em ATPC",
    "Live / Webconferência",
    "Análise de Indicadores",
    "Busca Ativa / Frequência",
    "Acompanhamento de Plataformas",
    "Devolutiva Formativa",
    "Reunião com Gestão Escolar",
    "Atendimento Educação Especial (AEE/PEI)",
    "Acompanhamento Multiplica SP",
    "Sistematização / Relatório"
]

TAREFAS_POR_PASTA = {
    "QUALIDADE DA AULA": [
        "Apoio Presencial",
        "Alinhamento com o/a Coordenador(a)",
        "Formação em Serviço (CGP/CGPG)",
        "Observação de Aula",
        "Participação em ATPC",
        "Devolutiva Formativa",
        "Análise de Indicadores"
    ],
    "DESENVOLVIMENTO CURRICULAR": [
        "Apoio Presencial",
        "Alinhamento com o/a Coordenador(a)",
        "Participação em ATPC",
        "Acompanhamento de Plataformas",
        "Análise de Indicadores",
        "Organização de Formação Presencial",
        "Live / Webconferência"
    ],
    "CONVIVA": [
        "Apoio Presencial",
        "Busca Ativa / Frequência",
        "Reunião com Gestão Escolar",
        "Alinhamento com o/a Coordenador(a)",
        "Acompanhamento Placon / Conviva",
        "Análise de Indicadores"
    ],
    "MULTIPLICA": [
        "Apoio Presencial",
        "Acompanhamento Multiplica SP",
        "Alinhamento com o/a Coordenador(a)",
        "Live / Webconferência",
        "Organização de Formação Presencial",
        "Análise de Indicadores"
    ],
    "EDUCACAO ESPECIAL": [
        "Apoio Presencial",
        "Atendimento Educação Especial (AEE/PEI)",
        "Alinhamento com o/a Coordenador(a)",
        "Participação em ATPC",
        "Reunião com Gestão Escolar",
        "Sistematização / Relatório"
    ],
    "RECOMPOSICAO": [
        "Apoio Presencial",
        "Análise de Indicadores",
        "Acompanhamento de Plataformas",
        "Alinhamento com o/a Coordenador(a)",
        "Participação em ATPC",
        "Devolutiva Formativa"
    ],
    "CEC": [
        "Alinhamento de Equipe",
        "Alinhamento 1x1",
        "Alinhamento com o/a Coordenador(a)",
        "Análise de Indicadores",
        "Organização de Formação Presencial",
        "Live / Webconferência",
        "Sistematização / Relatório"
    ],
    "AT": [
        "Alinhamento de Equipe",
        "Análise de Indicadores",
        "Organização de Formação Presencial",
        "Live / Webconferência",
        "Sistematização / Relatório"
    ]
}

# ==============================================================================
# GERADOR DE PLANO DE TRABALHO EM WORD COM CABEÇALHO OFICIAL DA URE MIRACATU
# ==============================================================================
def renderizar_preview_plano_html(nome_pec, email_pec, pasta_pec, comp_pec, escola_sel, semana_str, cronograma_dias, obs_texto, df_prioridades_pec=None):
    cron_ordenado = sorted(cronograma_dias, key=lambda x: x.get("data_iso", "9999-99-99"))
    linhas_cron = ""
    if cron_ordenado:
        for item in cron_ordenado:
            linhas_cron += f"""
            <tr>
                <td style="border:1px solid #ccc;padding:6px;">{item.get('dia', '')}</td>
                <td style="border:1px solid #ccc;padding:6px;"><b>{item.get('local', '')}</b></td>
                <td style="border:1px solid #ccc;padding:6px;">{item.get('acao', '')}</td>
                <td style="border:1px solid #ccc;padding:6px;">{item.get('publico', '')}</td>
            </tr>
            """
    else:
        linhas_cron = '<tr><td colspan="4" style="border:1px solid #ccc;padding:8px;text-align:center;color:#666;">Nenhuma visita ou ação adicionada ainda.</td></tr>'

    secao_ind = ""
    if df_prioridades_pec is not None and not df_prioridades_pec.empty:
        linhas_ind = ""
        for _, r_i in df_prioridades_pec.head(15).iterrows():
            linhas_ind += f"""
            <tr>
                <td style="border:1px solid #ccc;padding:5px;">{r_i.get('Escola', '-')}</td>
                <td style="border:1px solid #ccc;padding:5px;">{r_i.get('Subcategoria', '-')} ({r_i.get('Periodo', '-')})</td>
                <td style="border:1px solid #ccc;padding:5px;">{r_i.get('Turma', '-')} | {r_i.get('Estudante', '-')}</td>
                <td style="border:1px solid #ccc;padding:5px;">{r_i.get('Componente', '-')}</td>
                <td style="border:1px solid #ccc;padding:5px;">{formatar_pct_str(r_i.get('Proficiencia_Pct'))} - {r_i.get('Nivel', '-')}</td>
            </tr>
            """
        secao_ind = f"""
        <h4 style="color:#1a365d;margin-top:16px;margin-bottom:6px;">2. Indicadores Prioritários da Sua Pasta de Atuação</h4>
        <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead>
                <tr style="background:#f1f5f9;">
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Escola</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Subcategoria / Período</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Turma / Estudante</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Componente</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Resultado / Farol</th>
                </tr>
            </thead>
            <tbody>{linhas_ind}</tbody>
        </table>
        """

    secao_obs = ""
    if obs_texto and obs_texto.strip():
        obs_fmt = obs_texto.strip().replace("\n", "<br>")
        secao_obs = f"""
        <h4 style="color:#1a365d;margin-top:16px;margin-bottom:6px;">3. Pautas Formativas, Diagnóstico e Devolutivas</h4>
        <div style="border:1px solid #e2e8f0;background:#f8fafc;padding:10px;border-radius:4px;font-size:13px;">{obs_fmt}</div>
        """

    return f"""
    <div style="background:#ffffff;border:1px solid #cbd5e1;border-radius:8px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;">
        <div style="text-align:center;border-bottom:2px solid #1e3a8a;padding-bottom:10px;margin-bottom:16px;">
            <div style="font-size:13px;font-weight:bold;">SECRETARIA DA EDUCAÇÃO DO ESTADO DE SÃO PAULO</div>
            <div style="font-size:12px;font-weight:bold;">UNIDADE REGIONAL DE ENSINO DE MIRACATU — EQUIPE DE ESPECIALISTAS EM CURRÍCULO (EEC)</div>
            <div style="font-size:11px;color:#475569;">PLANO DE TRABALHO / GUIA PEDAGÓGICO — ENTREGA ATÉ O 2º DIA ÚTIL DO MÊS</div>
        </div>
        <h3 style="text-align:center;color:#0f172a;margin:10px 0 16px 0;">PLANO DE TRABALHO — ACOMPANHAMENTO PEDAGÓGICO (GPS)</h3>
        <div style="font-size:13px;line-height:1.6;margin-bottom:14px;">
            <b>Especialista / PEC Responsável:</b> {nome_pec} ({email_pec})<br>
            <b>Pasta / Área de Atuação:</b> {pasta_pec} — <b>Componente:</b> {comp_pec}<br>
            <b>Unidade Escolar Prioritária Principal:</b> {escola_sel}<br>
            <b>Período / Mês de Referência:</b> {semana_str}
        </div>
        <h4 style="color:#1a365d;margin-bottom:6px;">1. Cronograma Mensal de Visitas, ATPCs e Ações Planejadas</h4>
        <table style="width:100%;border-collapse:collapse;font-size:12.5px;">
            <thead>
                <tr style="background:#f1f5f9;">
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Data / Dia / Horário</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Escola / Local</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Ação Pedagógica (Resolução SEDUC 01/2026)</th>
                    <th style="border:1px solid #ccc;padding:6px;text-align:left;">Público-Alvo / ATPC / Encaminhamento</th>
                </tr>
            </thead>
            <tbody>{linhas_cron}</tbody>
        </table>
        {secao_ind}
        {secao_obs}
        <div style="margin-top:32px;display:flex;justify-content:space-around;text-align:center;font-size:12px;">
            <div>_________________________________________<br><b>PEC: {nome_pec}</b></div>
            <div>_________________________________________<br><b>CEC — Coordenador de Equipe Curricular</b></div>
        </div>
    </div>
    """

def gerar_docx_plano(nome_pec, email_pec, pasta_pec, comp_pec, escola_sel, semana_str, cronograma_dias, obs_texto, df_prioridades_pec=None):
    doc = Document()

    header = doc.sections[0].header
    p_cab = header.paragraphs[0]
    p_cab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_cab.add_run("SECRETARIA DA EDUCAÇÃO DO ESTADO DE SÃO PAULO\n")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p_cab.add_run("UNIDADE REGIONAL DE ENSINO DE MIRACATU — EQUIPE DE ESPECIALISTAS EM CURRÍCULO (EEC)\n")
    r2.bold = True
    r2.font.size = Pt(9.5)
    r3 = p_cab.add_run("PLANO DE TRABALHO / GUIA PEDAGÓGICO — ENTREGA ATÉ O 2º DIA ÚTIL DO MÊS")
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(90, 90, 90)

    titulo = doc.add_heading("PLANO DE TRABALHO — ACOMPANHAMENTO PEDAGÓGICO (GPS)", level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_info = doc.add_paragraph()
    p_info.add_run("Especialista / PEC Responsável: ").bold = True
    p_info.add_run(f"{nome_pec} ({email_pec})\n")
    p_info.add_run("Pasta / Área de Atuação: ").bold = True
    p_info.add_run(f"{pasta_pec} — Componente: {comp_pec}\n")
    p_info.add_run("Unidade Escolar Prioritária Principal: ").bold = True
    p_info.add_run(f"{escola_sel}\n")
    p_info.add_run("Período / Mês de Referência: ").bold = True
    p_info.add_run(f"{semana_str}\n")

    doc.add_heading("1. Cronograma Mensal de Visitas, ATPCs e Ações Planejadas", level=2)
    tabela = doc.add_table(rows=1, cols=4)
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Data / Dia / Horário"
    hdr[1].text = "Escola / Local"
    hdr[2].text = "Ação Pedagógica (Resolução SEDUC 01/2026)"
    hdr[3].text = "Público-Alvo / ATPC / Encaminhamento"

    cron_ordenado = sorted(cronograma_dias, key=lambda x: x.get("data_iso", "9999-99-99"))
    for item in cron_ordenado:
        row_cells = tabela.add_row().cells
        row_cells[0].text = item.get("dia", "")
        row_cells[1].text = item.get("local", "")
        row_cells[2].text = item.get("acao", "")
        row_cells[3].text = item.get("publico", "")

    if df_prioridades_pec is not None and not df_prioridades_pec.empty:
        doc.add_heading("2. Indicadores Prioritários da Sua Pasta de Atuação (Avaliações, Frequência e Plataformas)", level=2)
        t_ind = doc.add_table(rows=1, cols=5)
        t_ind.style = "Table Grid"
        for idx, col_n in enumerate(["Escola", "Subcategoria / Período", "Turma / Estudante", "Componente", "Resultado / Farol"]):
            t_ind.rows[0].cells[idx].text = col_n
        for _, r_i in df_prioridades_pec.head(15).iterrows():
            rc = t_ind.add_row().cells
            rc[0].text = str(r_i.get("Escola", "-"))
            rc[1].text = f"{r_i.get('Subcategoria', '-')} ({r_i.get('Periodo', '-')})"
            rc[2].text = f"{r_i.get('Turma', '-')} | {r_i.get('Estudante', '-')}"
            rc[3].text = str(r_i.get("Componente", "-"))
            rc[4].text = f"{formatar_pct_str(r_i.get('Proficiencia_Pct'))} - {r_i.get('Nivel', '-')}"

    if obs_texto.strip():
        doc.add_heading("3. Pautas Formativas, Diagnóstico e Devolutivas", level=2)
        doc.add_paragraph(obs_texto.strip())

    doc.add_paragraph("\n")
    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass.add_run("_________________________________________          _________________________________________\n")
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
# AUTENTICAÇÃO E SESSÃO
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
                                            corresponder_escola_oficial(e.strip())
                                            for e in match[col_esc].dropna().astype(str).tolist()
                                            for e in re.split(r"[,;/]", e) if e.strip()
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
# CARREGAMENTO GLOBAL DOS DADOS DO DRIVE
# ==============================================================================
pasta_atual = st.session_state["usuario_pasta"]
comp_atual = st.session_state["usuario_componente"]
escolas_usuario = st.session_state["usuario_escolas"]
eh_cec_ou_at = pasta_atual in ["CEC", "AT"]
eh_pec_qualidade = pasta_atual == "QUALIDADE DA AULA"
eh_recomposicao = pasta_atual == "RECOMPOSICAO"

df_ind_global, df_curs_global, df_atpc_drive, df_planos_drive, df_termos_drive, qtd_arqs_drive = (
    sincronizar_pasta_drive_completo(ID_PASTA_DRIVE_GPS, GOOGLE_DRIVE_API_KEY)
    if "COLE_AQUI" not in GOOGLE_DRIVE_API_KEY else (pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0)
)

# Consolida base de ATPC (Sheets + Drive + Upload na sessão)
dfs_atpc_lista = [d for d in [df_atpc_sheets, df_atpc_drive] if not d.empty]
df_atpc_global = pd.concat(dfs_atpc_lista, ignore_index=True).drop_duplicates() if dfs_atpc_lista else pd.DataFrame()

df_ind_usuario = (
    df_ind_global if eh_cec_ou_at
    else filtrar_indicadores_por_pec(df_ind_global, pasta_atual, comp_atual, escolas_usuario if eh_pec_qualidade else None)
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
st.caption("Planejamento Mensal Integrado, Indicadores Hierárquicos e Acompanhamento — URE de Miracatu")

lista_abas = [
    "📋 Plano de Trabalho & Planner Mensal",
    "📊 Indicadores (Avaliações, Plataformas & Frequência)",
    "📅 Agenda SEDUC, CEC & Feriados",
    "📚 Diretrizes das Pastas (Res. SEDUC 01/2026)"
]
if eh_cec_ou_at:
    lista_abas.insert(2, "🛡️ Painel Exclusivo CEC (Confronto & Fechamento)")

abas = st.tabs(lista_abas)

# --- ABA 1: PLANO DE TRABALHO & PLANNER MENSAL LADO A LADO ---
with abas[0]:
    st.subheader("Elaboração do Plano de Trabalho (Modelo Oficial URE Miracatu)")
    st.caption("Submeta o arquivo gerado na pasta do Google Drive até o 2º dia útil do mês. Após cada visita, deposite o Termo de Visita em PDF na sua pasta do Drive.")

    # Monta dinamicamente as opções de indicadores disponíveis (Prova Paulista 3º Bimestre, SARESP, Frequência, Plataformas e Todos)
    opcoes_criterio_dict = dict(OPCOES_CRITERIO_PRIORIDADE)
    if not df_ind_usuario.empty:
        plats_presentes = sorted(df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "PLATAFORMA"]["Subcategoria"].dropna().unique().tolist())
        for pl_nome in plats_presentes:
            if pl_nome != "Outras Plataformas":
                opcoes_criterio_dict[f"PLAT::{pl_nome}"] = f"Plataforma — {pl_nome}"

    # 1. Seletor do Mês e Ano do Planejamento (ATPCs são lidas automaticamente do Google Drive / Sheets)
    c_top1, c_top2 = st.columns(2)
    with c_top1:
        mes_plano = st.selectbox(
            "Mês do Plano:",
            list(range(1, 13)),
            index=datetime.now().month - 1,
            format_func=lambda m: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][m - 1],
            key="mes_planner_top"
        )
    with c_top2:
        ano_plano = st.number_input("Ano do Plano:", min_value=2025, max_value=2030, value=datetime.now().year, key="ano_planner_top")

    st.markdown("---")

    # ==========================================================================
    # TELA DIVIDIDA EM DOIS: ESQUERDA = PLANNER MENSAL | DIREITA = LANÇAMENTO
    # ==========================================================================
    col_planner, col_lancamento = st.columns([1.35, 1.0], gap="large")

    # --- LADO DIREITO: FILTRO DE PRIORIDADE + LANÇAMENTO + SUGESTÃO AUTOMÁTICA DE ATPC DO DRIVE ---
    with col_lancamento:
        st.markdown("### ✍️ Lançar Visita / Ação no Plano")

        criterio_prio_sel = st.selectbox(
            "🎯 Filtrar Escolas Prioritárias pelo Indicador:",
            options=list(opcoes_criterio_dict.keys()),
            format_func=lambda k: opcoes_criterio_dict[k],
            key="filtro_criterio_prioridade_live",
            help="Selecione o indicador (Prova Paulista 3º Bimestre, SARESP, Frequência, Plataforma ou Todos) para listar abaixo as escolas prioritárias pós-filtro."
        )

        # Calcula o ranking pós-filtro conforme o indicador escolhido
        df_ranking_prio = obter_ranking_escolas_prioritarias(
            df_ind_usuario,
            pasta_atual,
            escolas_usuario,
            criterio_chave=criterio_prio_sel,
            df_recomp_sheets=df_recomposicao_sheets
        )

        # Monta lista pós-filtro com rótulos ricos mostrando posição de prioridade e índice do indicador
        mapa_rotulo_para_escola = {}
        opcoes_locais_rotulos = []
        if not df_ranking_prio.empty:
            for _, r_p in df_ranking_prio.iterrows():
                esc_nome = str(r_p["Escola"]).strip()
                prio_txt = str(r_p.get("Prioridade", "")).strip()
                pct_val = r_p.get("Indicador_Medio_Pct", np.nan)
                farol_txt = str(r_p.get("Farol", ""))
                icone_f = "🔴" if "Vermelho" in farol_txt else ("🟡" if "Amarelo" in farol_txt else ("🟢" if "Verde" in farol_txt else ("🔵" if "Azul" in farol_txt else "⚪")))
                if pd.notna(pct_val):
                    rotulo = f"{prio_txt} Prioritária | {esc_nome} ({pct_val:.1f}% {icone_f})"
                elif "Qtd_Estudantes_Recomposicao" in r_p and pd.notna(r_p["Qtd_Estudantes_Recomposicao"]):
                    rotulo = f"{prio_txt} Prioritária | {esc_nome} ({int(r_p['Qtd_Estudantes_Recomposicao'])} est. Recomposição)"
                else:
                    rotulo = f"{prio_txt} | {esc_nome}"
                opcoes_locais_rotulos.append(rotulo)
                mapa_rotulo_para_escola[rotulo] = esc_nome
        else:
            for esc_nome in (escolas_usuario if escolas_usuario else ESCOLAS_URE_MIRACATU):
                opcoes_locais_rotulos.append(esc_nome)
                mapa_rotulo_para_escola[esc_nome] = esc_nome

        for local_extra in ["URE Miracatu", "Reunião Virtual / EFAPE / SEDUC"]:
            opcoes_locais_rotulos.append(local_extra)
            mapa_rotulo_para_escola[local_extra] = local_extra

        busca_inicio = st.text_input(
            "🔎 Identificar escola pelo início (opcional):",
            placeholder="Digite o início do nome para localizar na lista pós-filtro...",
            key="busca_inicio_escola"
        )

        indice_sugerido = 0
        if busca_inicio.strip():
            termo_norm = normalizar_texto(busca_inicio).replace("EE ", "")
            match_inicio = next(
                (i for i, rot in enumerate(opcoes_locais_rotulos)
                 if normalizar_texto(mapa_rotulo_para_escola[rot]).replace("EE ", "").startswith(termo_norm)
                 or normalizar_texto(mapa_rotulo_para_escola[rot]).startswith(normalizar_texto(busca_inicio))),
                None
            )
            if match_inicio is None:
                match_inicio = next(
                    (i for i, rot in enumerate(opcoes_locais_rotulos) if termo_norm in normalizar_texto(mapa_rotulo_para_escola[rot])),
                    0
                )
            indice_sugerido = match_inicio

        rotulo_escolhido = st.selectbox(
            "1. Escola Prioritária (Pós-Filtro) ou Local Selecionado:",
            opcoes_locais_rotulos,
            index=indice_sugerido,
            key="esc_lancamento_live"
        )
        escola_sel_form = mapa_rotulo_para_escola[rotulo_escolhido]

        # Exibe o diagnóstico pós-filtro da escola selecionada
        if not df_ranking_prio.empty and escola_sel_form in df_ranking_prio["Escola"].values:
            linha_diag = df_ranking_prio[df_ranking_prio["Escola"] == escola_sel_form].iloc[0]
            st.caption(f"📊 **Diagnóstico Pós-Filtro ({opcoes_criterio_dict[criterio_prio_sel]}):** {linha_diag.get('Prioridade', '')} lugar — {linha_diag.get('Detalhe_Indicador', '')}")

        # Identificação automática dos dias e horários de ATPC a partir do Google Drive / Sheets
        atpcs_da_escola = obter_atpcs_escola(df_atpc_global, escola_sel_form, comp_atual)
        sugestoes_atpc = []
        if not atpcs_da_escola.empty:
            resumo_atpc = " | ".join([f"**{r['Dia_Semana_Str']}** ({r['Horario']} — {r['Area_Componente']})" for _, r in atpcs_da_escola.iterrows()])
            st.info(f"🕒 **ATPC identificada no Drive para esta Escola:** {resumo_atpc}")
            sugestoes_atpc = sugerir_datas_atpc_no_mes(ano_plano, mes_plano, atpcs_da_escola)
        elif escola_sel_form in ESCOLAS_URE_MIRACATU:
            st.caption("ℹ️ Nenhum horário de ATPC localizado para esta escola nas planilhas do Google Drive.")

        # Sugestão automática imediata de data e horário da ATPC identificada no Drive
        data_sugerida_dt = None
        hora_ini_default = datetime.strptime("08:00", "%H:%M").time()
        hora_fim_default = datetime.strptime("12:00", "%H:%M").time()
        incluir_atpc_auto = False

        if sugestoes_atpc:
            opcoes_sug = [
                f"✨ Sugestão ATPC (Drive): {s['data_str']} ({s['dia_semana']}) — {s['horario']} [{s['area']}]"
                for s in sugestoes_atpc
            ] + ["Escolher outra data manualmente no calendário abaixo..."]
            escolha_sug = st.selectbox("💡 Data & Horário de ATPC Sugeridos Automaticamente:", opcoes_sug, index=0, key="sug_atpc_sel")
            if escolha_sug != "Escolher outra data manualmente no calendário abaixo...":
                idx_s = opcoes_sug.index(escolha_sug)
                data_sugerida_dt = sugestoes_atpc[idx_s]["data"]
                incluir_atpc_auto = True
                m_horas = re.findall(r"(\d{1,2}:\d{2})", str(sugestoes_atpc[idx_s]["horario"]))
                if len(m_horas) >= 2:
                    try:
                        hora_ini_default = datetime.strptime(m_horas[0], "%H:%M").time()
                        hora_fim_default = datetime.strptime(m_horas[1], "%H:%M").time()
                    except Exception:
                        pass

        # Calcula data padrão dentro do mês/ano selecionado
        _, max_dia_mes = calendar.monthrange(ano_plano, mes_plano)
        dia_default = min(datetime.now().day, max_dia_mes)
        data_default = data_sugerida_dt if data_sugerida_dt else date(ano_plano, mes_plano, dia_default)

        with st.form("form_add_atividade_split", clear_on_submit=False):
            data_visita = st.date_input(
                "2. Data (clique para abrir o calendário):",
                value=data_default,
                min_value=date(ano_plano, mes_plano, 1),
                max_value=date(ano_plano, mes_plano, max_dia_mes),
                format="DD/MM/YYYY"
            )

            c_h1, c_h2 = st.columns(2)
            with c_h1:
                hora_inicio = st.time_input("3. Horário Início (clicável):", value=hora_ini_default, step=900)
            with c_h2:
                hora_fim = st.time_input("4. Horário Término (clicável):", value=hora_fim_default, step=900)

            opcoes_atividades = list(dict.fromkeys(TAREFAS_POR_PASTA.get(pasta_atual, []) + COMANDOS_CURTOS_ATIVIDADES))
            default_atividades = ["Participação em ATPC", "Apoio Presencial"] if incluir_atpc_auto else [opcoes_atividades[0]]
            atividades_selecionadas = st.multiselect(
                "5. Atividades (clique para selecionar uma ou várias):",
                options=opcoes_atividades,
                default=default_atividades,
                placeholder="Selecione as atividades que irá realizar..."
            )

            info_adicional = st.text_input(
                "6. Informação Adicional (opcional — vai apenas para o Plano de Trabalho):",
                value="",
                placeholder="Digite alguma informação extra se precisar (ex.: pauta, turma, público)..."
            )

            add_btn = st.form_submit_button("➕ Adicionar ao Planner e ao Plano de Trabalho", use_container_width=True)

            if add_btn:
                if not atividades_selecionadas and not info_adicional.strip():
                    st.warning("Selecione pelo menos uma atividade ou preencha uma informação adicional.")
                else:
                    dia_sem_nome = DIAS_SEMANA_LABEL[data_visita.weekday()]
                    horario_visita = f"{hora_inicio.strftime('%H:%M')} às {hora_fim.strftime('%H:%M')}"
                    acao_combinada = " • ".join(atividades_selecionadas) if atividades_selecionadas else "Atividade Específica"
                    detalhe_plano = info_adicional.strip() if info_adicional.strip() else "Conforme diretrizes da pasta"
                    st.session_state["cronograma_semana"].append({
                        "data_iso": data_visita.isoformat(),
                        "dia": f"{data_visita.strftime('%d/%m/%Y')} ({dia_sem_nome}) — {horario_visita}",
                        "horario": horario_visita,
                        "local": escola_sel_form,
                        "acao": acao_combinada,
                        "publico": detalhe_plano,
                        "inclui_atpc": ("Participação em ATPC" in atividades_selecionadas)
                    })
                    st.rerun()

        # Lista compacta das visitas lançadas com botão individual de remoção
        if st.session_state["cronograma_semana"]:
            st.markdown("#### 📌 Visitas / Ações Agendadas no Mês")
            for idx_it, it in enumerate(st.session_state["cronograma_semana"]):
                c_it1, c_it2 = st.columns([5, 1])
                with c_it1:
                    st.markdown(f"**{it['dia']}** — `{it['local']}`<br><small>{it['acao']} | {it['publico']}</small>", unsafe_allow_html=True)
                with c_it2:
                    if st.button("🗑️", key=f"del_item_{idx_it}", help="Remover agendamento"):
                        st.session_state["cronograma_semana"].pop(idx_it)
                        st.rerun()
            if st.button("🧹 Limpar Todos os Agendamentos", use_container_width=True):
                st.session_state["cronograma_semana"] = []
                st.rerun()

        # Tabela diagnóstica completa pós-filtro recolhível logo abaixo ou consultável
        with st.expander(f"📊 Ver Ranking Completo Pós-Filtro ({opcoes_criterio_dict[criterio_prio_sel]})", expanded=False):
            cols_show = [c for c in df_ranking_prio.columns if c != "Indice_Criticidade_Pct"]
            styler_prio = df_ranking_prio[cols_show].style
            if "Indicador_Medio_Pct" in cols_show:
                styler_prio = styler_prio.map(estilo_celula_farol, subset=["Indicador_Medio_Pct"]).format({"Indicador_Medio_Pct": "{:.1f}%"}, na_rep="-")
            st.dataframe(styler_prio, use_container_width=True, hide_index=True)

    # --- LADO ESQUERDO: PLANNER MENSAL COMPLETO ATUALIZADO EM TEMPO REAL ---
    with col_planner:
        st.markdown("### 🗓️ Planner Mensal Integrado (Tempo Real)")
        eventos_gcal = carregar_eventos_gcal(URL_GOOGLE_AGENDA_CEC)
        eventos_mes = coletar_eventos_mes_planner(
            ano_plano,
            mes_plano,
            df_bloqueios,
            df_agenda_seduc,
            eventos_gcal,
            st.session_state["cronograma_semana"],
            pasta_atual
        )
        html_cal = renderizar_html_planner_mensal(ano_plano, mes_plano, eventos_mes)
        st.markdown(html_cal, unsafe_allow_html=True)

    st.markdown("---")
    obs_texto = st.text_area("Pautas Formativas, Diagnóstico de Indicadores e Encaminhamentos:", height=100)

    df_prio_word = pd.DataFrame()
    if not df_ind_usuario.empty:
        df_prio_word = df_ind_usuario[df_ind_usuario["Proficiencia_Pct"] <= 60.0]

    mes_ref_str = f"{mes_plano:02d}/{ano_plano}"
    escola_principal_doc = st.session_state["cronograma_semana"][0]["local"] if st.session_state["cronograma_semana"] else escola_sel_form

    docx_buf = gerar_docx_plano(
        st.session_state["usuario_nome"],
        st.session_state["usuario_email"],
        pasta_atual,
        comp_atual,
        escola_principal_doc,
        mes_ref_str,
        st.session_state["cronograma_semana"],
        obs_texto,
        df_prio_word
    )
    nome_arq_docx = f"Plano_Trabalho_URE_{st.session_state['usuario_nome'].replace(' ', '_')}_{ano_plano}{mes_plano:02d}.docx"

    with st.expander("👁️ Pré-Visualizar Plano de Trabalho Completo (Antes de Baixar ou Enviar)", expanded=True):
        html_preview = renderizar_preview_plano_html(
            st.session_state["usuario_nome"],
            st.session_state["usuario_email"],
            pasta_atual,
            comp_atual,
            escola_principal_doc,
            mes_ref_str,
            st.session_state["cronograma_semana"],
            obs_texto,
            df_prio_word
        )
        st.markdown(html_preview, unsafe_allow_html=True)

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
                f"Segue em anexo o Plano de Trabalho de {st.session_state['usuario_nome']} ({pasta_atual} - {comp_atual}) referente a {mes_ref_str}.",
                docx_buf,
                nome_arq_docx
            )
            if ok:
                st.success(msg_envio)
            else:
                st.warning(msg_envio)
    with col_drv_link:
        st.link_button("📂 Abrir Pasta do Drive para Depositar Plano / Termos", URL_PASTA_DRIVE_GPS, use_container_width=True)

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

# --- ABA 2: INDICADORES HIERÁRQUICOS (AVALIAÇÕES, PLATAFORMAS E FREQUÊNCIA) ---
with abas[1]:
    st.subheader(f"Indicadores Pedagógicos — {pasta_atual} ({comp_atual})")
    if not eh_cec_ou_at:
        regras_info = obter_regras_atuacao_pec(pasta_atual, comp_atual)
        if regras_info:
            st.info(f"🎯 **Filtro automático da sua pasta de atuação:** Componentes `{', '.join(regras_info['componentes'])}` | Plataformas `{', '.join(regras_info['plataformas'])}` + Frequência.")

    col_drv, col_up = st.columns([2, 1])
    with col_drv:
        if "COLE_AQUI" in GOOGLE_DRIVE_API_KEY:
            st.warning("Configure `GOOGLE_DRIVE_API_KEY` nos Secrets do Streamlit para leitura automática das pastas do Google Drive.")
        else:
            st.success(f"📂 {qtd_arqs_drive} arquivo(s) sincronizado(s) do Google Drive.")

    with col_up:
        arquivo_bi = st.file_uploader("Testar planilha avulsa (.xlsx / .csv):", type=["xlsx", "csv"])
        if arquivo_bi is not None:
            df_raw_up = pd.read_csv(arquivo_bi) if arquivo_bi.name.endswith(".csv") else pd.read_excel(arquivo_bi)
            df_pad_up = padronizar_planilha_bruta(df_raw_up, arquivo_bi.name)
            df_ind_usuario = pd.concat([df_ind_usuario, df_pad_up], ignore_index=True)

    sub_aval, sub_plat, sub_freq = st.tabs([
        "📝 Avaliações (Prova Paulista, SARESP, SAEB, ENEM, Diagnóstica)",
        "💻 Plataformas Digitais (Khan, Matific, Speak, Leia SP...)",
        "🏫 Frequência Escolar"
    ])

    with sub_aval:
        df_av = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "AVALIACAO"] if not df_ind_usuario.empty else pd.DataFrame()
        c_av1, c_av2 = st.columns(2)
        with c_av1:
            tipo_prova = st.selectbox(
                "Selecione a Avaliação:",
                ["Prova Paulista", "SARESP", "SAEB", "ENEM", "Avaliação Diagnóstica", "Todas as Avaliações"]
            )
        with c_av2:
            if tipo_prova == "Prova Paulista":
                periodo_filtro = st.selectbox("Bimestre (Prova Paulista):", ["Todos os Bimestres", "1º Bimestre", "2º Bimestre", "3º Bimestre", "4º Bimestre"])
            elif tipo_prova in ["SARESP", "SAEB", "ENEM", "Avaliação Diagnóstica"]:
                anos_disp = sorted(df_av["Periodo"].dropna().unique().tolist()) if not df_av.empty else []
                periodo_filtro = st.selectbox(f"Ano / Edição ({tipo_prova}):", ["Todos os Anos / Edições"] + anos_disp)
            else:
                periodo_filtro = "Todos"

        df_av_filt = df_av.copy()
        if not df_av_filt.empty:
            if tipo_prova != "Todas as Avaliações":
                df_av_filt = df_av_filt[df_av_filt["Subcategoria"] == tipo_prova]
            if not periodo_filtro.startswith("Todos"):
                df_av_filt = df_av_filt[df_av_filt["Periodo"].str.contains(periodo_filtro, case=False, na=False)]

        renderizar_tabela_hierarquica(df_av_filt, f"{tipo_prova} ({periodo_filtro})", "Proficiência Média (%)")

    with sub_plat:
        df_pl = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "PLATAFORMA"] if not df_ind_usuario.empty else pd.DataFrame()
        plats_lista = ["Todas as Plataformas", "Khan Academy", "Matific", "Speak", "Open English", "Leia SP", "Redação Paulista", "Alura"]
        plat_escolhida = st.selectbox("Selecione a Plataforma Digital:", plats_lista)

        df_pl_filt = df_pl.copy()
        if not df_pl_filt.empty and plat_escolhida != "Todas as Plataformas":
            df_pl_filt = df_pl_filt[df_pl_filt["Subcategoria"] == plat_escolhida]

        renderizar_tabela_hierarquica(df_pl_filt, f"Plataforma — {plat_escolhida}", "Engajamento / Uso (%)")

    with sub_freq:
        df_fr = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "FREQUENCIA"] if not df_ind_usuario.empty else pd.DataFrame()
        periodos_fr = sorted(df_fr["Periodo"].dropna().unique().tolist()) if not df_fr.empty else []
        per_fr_sel = st.selectbox("Período / Semana de Frequência:", ["Todos os Períodos"] + periodos_fr)

        df_fr_filt = df_fr.copy()
        if not df_fr_filt.empty and per_fr_sel != "Todos os Períodos":
            df_fr_filt = df_fr_filt[df_fr_filt["Periodo"] == per_fr_sel]

        renderizar_tabela_hierarquica(df_fr_filt, f"Frequência Escolar ({per_fr_sel})", "Frequência Média (%)")

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
