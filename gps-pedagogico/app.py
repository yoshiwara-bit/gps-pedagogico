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
import hashlib
import hmac
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

# Fontes documentais: IDs informados pela URE. Não alterar o Drive original.
PASTAS_FONTES_GPS = {
    "01 - Cadastro das Escolas": "1ACYYFNq7igDwWWm-wfG-OGb7qtTwxmWO",
    "02 - Avaliações": "10zb-l_xNYiCmVlgUjpNpGSHmvXfv_-dJ",
    "03 - Frequência Escolar": "1LbYPCLoFMTxqho5dgF0-c_TF-x_XowN9",
    "04 - Plataformas Digitais": "1SfLLugy9Ncxtit-oasqgiYUy7c6yX-h7",
}

@st.cache_data(ttl=300, show_spinner=False)
def inventariar_fonte_drive(folder_id, api_key):
    """Lista metadados recursivos, sem baixar PDFs nem classificar visitas.

    Uma chave de API só permite listar pastas acessíveis a ela; 403/404 são
    reportados explicitamente para não parecer que uma pasta está vazia.
    """
    if not api_key or "COLE_AQUI" in api_key:
        return [], "Chave de API não configurada."
    arquivos = []
    visitados = set()
    def visitar(fid, caminho, profundidade):
        if profundidade > 8 or fid in visitados:
            return
        visitados.add(fid)
        token = None
        while True:
            params = {"q": f"'{fid}' in parents and trashed = false",
                      "fields": "nextPageToken,files(id,name,mimeType,modifiedTime)",
                      "pageSize": "1000", "key": api_key}
            if token:
                params["pageToken"] = token
            url = "https://www.googleapis.com/drive/v3/files?" + urllib.parse.urlencode(params)
            try:
                with urllib.request.urlopen(url, timeout=15) as response:
                    dados = json.loads(response.read().decode("utf-8"))
            except Exception as exc:
                raise RuntimeError(f"Não foi possível listar {caminho or 'pasta principal'}: {exc}") from exc
            for item in dados.get("files", []):
                nome = item.get("name", "")
                if item.get("mimeType") == "application/vnd.google-apps.folder":
                    visitar(item["id"], f"{caminho}/{nome}".strip("/"), profundidade+1)
                else:
                    arquivos.append({"Pasta": caminho or "(raiz)", "Arquivo": nome,
                                     "Tipo": item.get("mimeType", ""),
                                     "Atualização": item.get("modifiedTime", "")})
            token = dados.get("nextPageToken")
            if not token:
                break
    try:
        visitar(folder_id, "", 0)
        return arquivos, None
    except Exception as exc:
        return arquivos, str(exc)


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

def extrair_regras_config_pastas_sheets(df_pastas_raw, pasta_pec, componente_pec):
    """Lê na aba Config_Pastas do Sheets eventuais componentes e plataformas vinculados à subpasta do PEC."""
    if df_pastas_raw is None or df_pastas_raw.empty:
        return [], []
    p_norm = normalizar_texto(pasta_pec)
    c_norm = normalizar_texto(componente_pec)
    tokens_pec = [t.strip() for t in re.split(r"[,;/|-]", c_norm) if t.strip() and t.strip() not in ["TODOS", "GERAL"]]

    comps_sheets, plats_sheets = [], []
    todas_plats_conhecidas = ["KHAN ACADEMY", "MATIFIC", "SPEAK", "OPEN ENGLISH", "LEIA SP", "REDACAO PAULISTA", "ALURA", "PLATAFORMA DE CIENCIAS", "PLATAFORMA HUMANAS"]
    todos_comps_conhecidos = [c for v in MAPA_ATUACAO_COMPONENTES_PLATAFORMAS.values() for c in v["componentes"]]

    for _, row in df_pastas_raw.iterrows():
        linha_txt = normalizar_texto(" ".join([str(v) for v in row.values if pd.notna(v)]))
        if not linha_txt:
            continue
        match_pasta = (p_norm in linha_txt) or any(tok in linha_txt for tok in tokens_pec)
        if match_pasta and (not tokens_pec or any(tok in linha_txt for tok in tokens_pec)):
            for c_k in todos_comps_conhecidos:
                if c_k in linha_txt and c_k not in comps_sheets:
                    comps_sheets.append(c_k)
            for pl_k in todas_plats_conhecidas:
                if pl_k in linha_txt and pl_k not in plats_sheets:
                    plats_sheets.append(pl_k)
    return comps_sheets, plats_sheets

def obter_regras_atuacao_pec(pasta_pec, componente_pec, df_pastas_raw=None):
    p_norm = normalizar_texto(pasta_pec)
    c_norm = normalizar_texto(componente_pec)

    # CEC e AT têm visão global de todos os componentes e plataformas
    if p_norm in ["CEC", "AT"]:
        return None

    # Qualidade da Aula (QA): acompanha integralmente suas escolas atribuídas no Sheets,
    # com acesso a TODAS as plataformas digitais e componentes dessas escolas.
    if p_norm == "QUALIDADE DA AULA":
        return {"componentes": None, "plataformas": None}

    if p_norm == "RECOMPOSICAO":
        return MAPA_ATUACAO_COMPONENTES_PLATAFORMAS["RECOMPOSICAO"]

    # Desenvolvimento Curricular (e demais pastas específicas): filtra conforme pasta/componente do Sheets
    if c_norm in ["", "TODOS"]:
        return None

    tokens_comp = [t.strip() for t in re.split(r"[,;/|-]", c_norm) if t.strip()]
    comps_encontrados = list(dict.fromkeys(tokens_comp))
    plats_encontradas = []

    for chave_area, regras in MAPA_ATUACAO_COMPONENTES_PLATAFORMAS.items():
        if chave_area in c_norm or chave_area in p_norm or any(tok in chave_area or chave_area in tok for tok in tokens_comp):
            comps_encontrados.extend(regras["componentes"])
            plats_encontradas.extend(regras["plataformas"])
        elif any(comp in tok or tok in comp for tok in tokens_comp for comp in regras["componentes"]):
            comps_encontrados.extend(regras["componentes"])
            plats_encontradas.extend(regras["plataformas"])

    # Complementa com vínculos explícitos da aba Config_Pastas do Google Sheets
    comps_sh, plats_sh = extrair_regras_config_pastas_sheets(df_pastas_raw, pasta_pec, componente_pec)
    comps_encontrados.extend(comps_sh)
    plats_encontradas.extend(plats_sh)

    comps_unicos = list(dict.fromkeys([c for c in comps_encontrados if c]))
    plats_unicas = list(dict.fromkeys([p for p in plats_encontradas if p]))

    return {
        "componentes": comps_unicos if comps_unicos else [c_norm],
        "plataformas": plats_unicas
    }

def filtrar_indicadores_por_pec(df_ind, pasta_pec, componente_pec, escolas_pec=None, df_pastas_raw=None):
    if df_ind is None or df_ind.empty:
        return pd.DataFrame()
    df_f = df_ind.copy()
    p_norm = normalizar_texto(pasta_pec)

    # Qualidade da Aula: restringe ESTRITAMENTE às escolas de acompanhamento atribuídas ao PEC no Sheets
    if p_norm == "QUALIDADE DA AULA":
        esc_norm = [normalizar_texto(e) for e in (escolas_pec or []) if str(e).strip()]
        if not esc_norm:
            return pd.DataFrame(columns=df_f.columns)
        df_f = df_f[df_f["Escola"].apply(lambda x: any(e in normalizar_texto(x) or normalizar_texto(x) in e for e in esc_norm))]
    elif escolas_pec and "Escola" in df_f.columns:
        esc_norm = [normalizar_texto(e) for e in escolas_pec if str(e).strip()]
        if esc_norm:
            df_f = df_f[df_f["Escola"].apply(lambda x: any(e in normalizar_texto(x) or normalizar_texto(x) in e for e in esc_norm))]

    regras = obter_regras_atuacao_pec(pasta_pec, componente_pec, df_pastas_raw=df_pastas_raw)
    if regras is None:
        return df_f

    comps_alvo = regras.get("componentes")
    plats_alvo = regras.get("plataformas")

    def linha_pertence(row):
        tipo_reg = normalizar_texto(row.get("Tipo_Indicador", "AVALIACAO"))
        if tipo_reg == "FREQUENCIA":
            return True
        if tipo_reg == "PLATAFORMA":
            # Se plats_alvo for None (caso de Qualidade da Aula), libera TODAS as plataformas
            if plats_alvo is None:
                return True
            if not plats_alvo:
                return False
            plat = normalizar_texto(row.get("Subcategoria", "")) + " " + normalizar_texto(row.get("Plataforma_Critica", ""))
            return any(p in plat for p in plats_alvo)
        # Avaliações
        if comps_alvo is None:
            return True
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

def identificar_nome_subpasta_especifica(caminho_pasta, tipo_categoria):
    """Extrai dinamicamente o nome da subpasta dentro de Avaliações, Plataformas ou Frequência no Drive."""
    if not caminho_pasta:
        return None
    partes = [p.strip() for p in caminho_pasta.split("/") if p.strip()]
    ignorar = {
        "INDICADORES", "DADOS", "PLANILHAS", "BASE", "URE MIRACATU", "GPS",
        "1 BIMESTRE", "2 BIMESTRE", "3 BIMESTRE", "1º BIMESTRE", "2º BIMESTRE", "3º BIMESTRE"
    }
    for idx, p in enumerate(partes):
        pn = normalizar_texto(p)
        if tipo_categoria == "AVALIACAO" and any(k in pn for k in ["AVALIAC", "PROVA"]):
            for sub_p in partes[idx + 1:]:
                sn = normalizar_texto(sub_p)
                if sn not in ignorar and not re.match(r"^[1-4]\s*[º°O]?\s*BIM", sn) and not any(normalizar_texto(e).replace("EE ", "") in sn for e in ESCOLAS_URE_MIRACATU):
                    return sub_p.strip()
        elif tipo_categoria == "PLATAFORMA" and "PLATAFORMA" in pn:
            for sub_p in partes[idx + 1:]:
                sn = normalizar_texto(sub_p)
                if sn not in ignorar and not re.match(r"^[1-4]\s*[º°O]?\s*BIM", sn) and not any(normalizar_texto(e).replace("EE ", "") in sn for e in ESCOLAS_URE_MIRACATU):
                    return sub_p.strip()
        elif tipo_categoria == "FREQUENCIA" and "FREQUENCIA" in pn:
            for sub_p in partes[idx + 1:]:
                sn = normalizar_texto(sub_p)
                if sn not in ignorar and not any(normalizar_texto(e).replace("EE ", "") in sn for e in ESCOLAS_URE_MIRACATU):
                    return sub_p.strip()
    return None

def detectar_classificacao_arquivo(nome_arquivo, caminho_pasta=""):
    ref = normalizar_texto(f"{caminho_pasta} {nome_arquivo}")
    ref_pasta = normalizar_texto(caminho_pasta)

    if "FREQUENCIA" in ref_pasta or "FREQUENCIA" in ref:
        tipo = "FREQUENCIA"
    elif "PLATAFORMA" in ref_pasta or any(k in ref for k in ["PLATAFORMA", "KHAN", "MATIFIC", "SPEAK", "OPEN ENGLISH", "LEIA SP", "REDACAO", "ALURA"]):
        tipo = "PLATAFORMA"
    else:
        tipo = "AVALIACAO"

    subpasta_dir = identificar_nome_subpasta_especifica(caminho_pasta, tipo)

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
        elif subpasta_dir:
            sub = subpasta_dir
        else:
            sub = "Outras Plataformas"
    elif tipo == "FREQUENCIA":
        sub = subpasta_dir if subpasta_dir else "Frequência Escolar"
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
        elif subpasta_dir:
            sub = subpasta_dir
        else:
            sub = "Outras Avaliações"

    # Prova Paulista possui apenas 1º, 2º e 3º Bimestres (não existe 4º Bimestre)
    max_bim = 3 if sub == "Prova Paulista" else 4
    m_bim = re.search(r"([1-4])\s*[º°O]?\s*(?:BIM|B\b)", ref)
    m_ano = re.search(r"\b(202[3-9])\b", ref)
    m_sem = re.search(r"SEM(?:ANA)?\s*([0-9]{1,2})", ref)
    if m_bim and int(m_bim.group(1)) <= max_bim:
        periodo = f"{m_bim.group(1)}º Bimestre" + (f" ({m_ano.group(1)})" if m_ano else "")
    elif m_ano:
        periodo = f"Ano {m_ano.group(1)}"
    elif m_sem:
        periodo = f"Semana {m_sem.group(1)}"
    elif subpasta_dir and tipo == "FREQUENCIA":
        periodo = subpasta_dir
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
        if dt in feriados_sp or dt in FERIADOS_EEC or dt in dias_bloqueados_set:
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
# DETECÇÃO AUTOMÁTICA DA PROVA PAULISTA MAIS ATUAL DO PERÍODO (1º A 3º BIMESTRE)
# ==============================================================================
def obter_max_bimestre_prova_paulista_por_mes(mes_ref):
    """Prova Paulista ocorre até o 3º Bimestre (não existe 4º Bimestre)."""
    if mes_ref is None:
        return 3
    m = int(mes_ref)
    if m <= 5:
        return 1
    elif m <= 8:
        return 2
    else:
        return 3

def obter_edicao_mais_atual_prova_paulista(df_ind, ano_ref=None, mes_ref=None):
    """
    Identifica automaticamente a edição mais atual da Prova Paulista disponível
    para o período em que o plano ou a pesquisa está sendo realizado (1º, 2º ou 3º Bimestre).
    Retorna (bimestre_num, rotulo_edicao).
    """
    max_bim_periodo = obter_max_bimestre_prova_paulista_por_mes(mes_ref)
    if df_ind is None or df_ind.empty:
        return max_bim_periodo, f"{max_bim_periodo}º Bimestre"

    df_pp = df_ind[df_ind["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))].copy()
    if df_pp.empty:
        return max_bim_periodo, f"{max_bim_periodo}º Bimestre"

    candidatos = []
    for per_val in df_pp["Periodo"].dropna().unique():
        p_norm = normalizar_texto(per_val)
        m_b = re.search(r"([1-3])\s*[º°O]?\s*BIM", p_norm)
        m_a = re.search(r"\b(202[3-9])\b", p_norm)
        b_num = int(m_b.group(1)) if m_b else 0
        a_num = int(m_a.group(1)) if m_a else (int(ano_ref) if ano_ref else date.today().year)
        if b_num in [1, 2, 3]:
            candidatos.append((a_num, b_num, str(per_val).strip()))

    if not candidatos:
        return max_bim_periodo, f"{max_bim_periodo}º Bimestre"

    if ano_ref:
        cands_ano = [c for c in candidatos if c[0] == int(ano_ref)]
        if cands_ano:
            candidatos = cands_ano

    cands_validos = [c for c in candidatos if c[1] <= max_bim_periodo]
    escolhido = max(cands_validos, key=lambda x: (x[0], x[1])) if cands_validos else max(candidatos, key=lambda x: (x[0], x[1]))
    return escolhido[1], escolhido[2]

def filtrar_prova_paulista_mais_atual(df_ind, ano_ref=None, mes_ref=None):
    if df_ind is None or df_ind.empty:
        return pd.DataFrame(), "3º Bimestre"
    df_pp = df_ind[df_ind["Subcategoria"].apply(lambda x: "PAULISTA" in normalizar_texto(x))].copy()
    if df_pp.empty:
        _, rot = obter_edicao_mais_atual_prova_paulista(df_ind, ano_ref, mes_ref)
        return df_pp, rot

    bim_alvo, rot_alvo = obter_edicao_mais_atual_prova_paulista(df_ind, ano_ref, mes_ref)
    df_pp_bim = df_pp[df_pp["Periodo"].apply(lambda x: bool(re.search(rf"\b{bim_alvo}\s*[º°O]?\s*BIM", normalizar_texto(x))))]
    if not df_pp_bim.empty:
        if ano_ref:
            df_pp_ano = df_pp_bim[df_pp_bim["Periodo"].astype(str).str.contains(str(ano_ref), na=False)]
            if not df_pp_ano.empty:
                return df_pp_ano, rot_alvo
        return df_pp_bim, rot_alvo
    return df_pp, rot_alvo

# ==============================================================================
# MOTOR DE PRIORIZAÇÃO PÓS-FILTRO POR INDICADOR
# ==============================================================================
def construir_opcoes_criterio_prioridade(df_ind, ano_ref=None, mes_ref=None):
    _, rot_pp_atual = obter_edicao_mais_atual_prova_paulista(df_ind, ano_ref=ano_ref, mes_ref=mes_ref)
    opcoes = {
        "TODOS": "Todos os Indicadores (Análise Multi-Indicador Consolidada)",
        "PP_ATUAL": f"Prova Paulista — Edição Atual ({rot_pp_atual})",
        "FREQUENCIA": "Frequência Escolar",
        "PLATAFORMA": "Plataformas Digitais (Todas)"
    }
    if df_ind is not None and not df_ind.empty:
        # Adiciona dinamicamente outras avaliações encontradas nas subpastas (SARESP, SAEB, ENEM, Diagnóstica, etc.)
        avals_presentes = sorted([
            s for s in df_ind[df_ind["Tipo_Indicador"] == "AVALIACAO"]["Subcategoria"].dropna().unique().tolist()
            if "PAULISTA" not in normalizar_texto(s) and s != "Outras Avaliações"
        ])
        for av_nome in avals_presentes:
            opcoes[f"AVAL::{av_nome}"] = f"Avaliação — {av_nome}"

        # Adiciona dinamicamente todas as plataformas encontradas nas subpastas
        plats_presentes = sorted([
            s for s in df_ind[df_ind["Tipo_Indicador"] == "PLATAFORMA"]["Subcategoria"].dropna().unique().tolist()
            if s != "Outras Plataformas"
        ])
        for pl_nome in plats_presentes:
            opcoes[f"PLAT::{pl_nome}"] = f"Plataforma — {pl_nome}"
    return opcoes

def filtrar_base_por_criterio_prioridade(df_ind, criterio_chave, ano_ref=None, mes_ref=None):
    if df_ind is None or df_ind.empty:
        return pd.DataFrame()
    df_b = df_ind.copy()

    if criterio_chave == "TODOS":
        return df_b
    elif criterio_chave == "PP_ATUAL":
        df_pp_atual, _ = filtrar_prova_paulista_mais_atual(df_b, ano_ref=ano_ref, mes_ref=mes_ref)
        return df_pp_atual
    elif criterio_chave == "FREQUENCIA":
        return df_b[df_b["Tipo_Indicador"] == "FREQUENCIA"]
    elif criterio_chave == "PLATAFORMA":
        return df_b[df_b["Tipo_Indicador"] == "PLATAFORMA"]
    elif criterio_chave.startswith("AVAL::"):
        nome_av = normalizar_texto(criterio_chave.split("::", 1)[1])
        return df_b[(df_b["Tipo_Indicador"] == "AVALIACAO") & (df_b["Subcategoria"].apply(lambda x: nome_av in normalizar_texto(x)))]
    elif criterio_chave.startswith("PLAT::"):
        nome_plat = normalizar_texto(criterio_chave.split("::", 1)[1])
        return df_b[(df_b["Tipo_Indicador"] == "PLATAFORMA") & (df_b["Subcategoria"].apply(lambda x: nome_plat in normalizar_texto(x)))]
    return df_b

def calcular_metricas_por_escola_pos_filtro(df_filtrado, criterio_chave, nome_crit_label=""):
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
    rot_uso = nome_crit_label or criterio_chave.replace("PLAT::", "Plataforma: ").replace("AVAL::", "Avaliação: ")
    grp["Detalhe_Indicador"] = grp["Indicador_Medio_Pct"].apply(lambda v: f"{rot_uso}: {v:.1f}%")
    grp["Farol"] = grp["Indicador_Medio_Pct"].apply(classificar_farol_nivel)
    return grp

def obter_ranking_escolas_prioritarias(
    df_ind,
    pasta_pec,
    escolas_delimitadas,
    criterio_chave="TODOS",
    df_recomp_sheets=None,
    ano_ref=None,
    mes_ref=None,
    nome_crit_label="",
    componente_filtro="Todos os Componentes"
):
    # Recomposição com lista pré-estabelecida quando o filtro geral "TODOS" está ativo
    if pasta_pec == "RECOMPOSICAO" and criterio_chave == "TODOS" and componente_filtro == "Todos os Componentes" and df_recomp_sheets is not None and not df_recomp_sheets.empty:
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

    df_filtrado = filtrar_base_por_criterio_prioridade(df_ind, criterio_chave, ano_ref=ano_ref, mes_ref=mes_ref)

    if componente_filtro and componente_filtro != "Todos os Componentes" and not df_filtrado.empty:
        comp_f_norm = normalizar_texto(componente_filtro)
        df_filtrado = df_filtrado[
            df_filtrado["Componente"].apply(
                lambda x: normalizar_texto(x) == comp_f_norm or comp_f_norm in normalizar_texto(x)
            )
        ]

    rot_base = nome_crit_label or criterio_chave
    rot_label = (
        f"{rot_base} — {componente_filtro}"
        if (componente_filtro and componente_filtro != "Todos os Componentes")
        else rot_base
    )

    df_metricas = calcular_metricas_por_escola_pos_filtro(
        df_filtrado, criterio_chave, nome_crit_label=rot_label
    )

    # PEC Qualidade da Aula: restringe ESTRITAMENTE às suas escolas de acompanhamento conforme o Sheets
    if pasta_pec == "QUALIDADE DA AULA":
        esc_oficiais = [corresponder_escola_oficial(e) for e in (escolas_delimitadas or []) if str(e).strip()]
        if not esc_oficiais:
            return pd.DataFrame(columns=["Prioridade", "Escola", "Indicador_Medio_Pct", "Indice_Criticidade_Pct", "Detalhe_Indicador", "Farol"])
        rows_q = []
        for esc in esc_oficiais:
            media_val = np.nan
            det_val = f"Sem dado em {rot_label}"
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
                "Detalhe_Indicador": f"Sem dado em {rot_label}",
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
        "Detalhe_Indicador": [f"Aguardando planilhas de {rot_label} no Drive"] * len(ESCOLAS_URE_MIRACATU[:30]),
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
    for d_fer, nome_fer in list(feriados_sp.items()) + list(FERIADOS_EEC.items()):
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
    return "\n".join(linha.lstrip() for linha in html.splitlines() if linha.strip())

# ==============================================================================
# ESTILIZAÇÃO VISUAL DO FAROL (CORES POR PROFICIÊNCIA / FREQUÊNCIA / USO)
# ==============================================================================
def cor_farol_percentual(valor, tipo="AVALIACAO"):
    """Critérios aprovados: avaliação 40/60/90; frequência 85/90.
    Usa o valor original, sem arredondamento prévio.
    """
    try:
        numero = float(str(valor).replace("%", "").replace(",", ".").strip())
        if not 0 <= numero <= 100:
            return ""
    except (TypeError, ValueError):
        return ""
    if tipo == "FREQUENCIA":
        return ("background-color: #F5C4B2; color: #8A2A0B; font-weight: bold;" if numero < 85 else
                "background-color: #FFF0A8; color: #735700; font-weight: bold;" if numero < 90 else
                "background-color: #BCE8C5; color: #075E23; font-weight: bold;")
    return ("background-color: #FFCDD2; color: #B71C1C; font-weight: bold;" if numero <= 40 else
            "background-color: #FFF9C4; color: #7B6100; font-weight: bold;" if numero <= 60 else
            "background-color: #C8E6C9; color: #1B5E20; font-weight: bold;" if numero < 90 else
            "background-color: #BBDEFB; color: #0D47A1; font-weight: bold;")

def estilo_celula_farol(val):
    return cor_farol_percentual(val, "AVALIACAO")

def renderizar_tabela_hierarquica(df_sub, titulo_secao, rotulo_metrica="Média (%)"):
    if df_sub.empty:
        st.info(f"Nenhum dado carregado para **{titulo_secao}**. Deposite as planilhas na pasta correspondente no Google Drive.")
        return

    frequencia = "Frequência" in titulo_secao or "Frequência" in rotulo_metrica
    if frequencia:
        st.caption("Farol da frequência: 🟠 <85% | 🟡 85% a <90% | 🟢 ≥90%")
    else:
        st.caption("Farol da Prova Paulista: 🔴 ≤40% | 🟡 >40% a 60% | 🟢 >60% a <90% | 🔵 ≥90%")
    ordenar = st.radio("Ordenação", ["Alfabética (A–Z)", "Farol ↓", "Farol ↑"], horizontal=True,
                      key=f"ord_gps_{titulo_secao}")
    estilo_atual = (lambda v: cor_farol_percentual(v, "FREQUENCIA" if frequencia else "AVALIACAO"))

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
    piv_esc = piv_esc[cols_ord].reset_index()
    piv_esc = (piv_esc.sort_values("Escola") if ordenar.startswith("Alfabética") else
               piv_esc.sort_values(rotulo_metrica, ascending=(ordenar == "Farol ↑"), na_position="last"))

    cols_num_esc = [c for c in piv_esc.columns if c != "Escola"]
    styler_esc = piv_esc.style.map(estilo_atual, subset=cols_num_esc).format({c: "{:.1f}%" for c in cols_num_esc}, na_rep="-")
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
            piv_tur = piv_tur[cols_ord_t].reset_index()
            piv_tur = (piv_tur.sort_values("Turma") if ordenar.startswith("Alfabética") else
                       piv_tur.sort_values(rotulo_metrica, ascending=(ordenar == "Farol ↑"), na_position="last"))

            cols_num_tur = [c for c in piv_tur.columns if c != "Turma"]
            styler_tur = piv_tur.style.map(estilo_atual, subset=cols_num_tur).format({c: "{:.1f}%" for c in cols_num_tur}, na_rep="-")
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
                        # Somente termos de visita são contabilizados como acompanhamento.
                        # Relatórios SARESP/SAEB, fichas e outros PDFs não são termos.
                        caminho_norm = normalizar_texto(caminho_pasta)
                        nome_norm = normalizar_texto(f_name)
                        texto_norm = normalizar_texto(texto_pdf[:5000])
                        pasta_termo = any(m in caminho_norm for m in ("TERMOS_DE_VISITA", "TERMOS DE VISITA", "PASTA_PEC_"))
                        documento_termo = any(m in nome_norm or m in texto_norm for m in ("TERMO DE VISITA", "TERMO DE ACOMPANHAMENTO", "TERMO_DE_VISITA", "ACOMPANHAMENTO PEDAGOGICO"))
                        if not (pasta_termo and documento_termo):
                            continue
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

    html_final = f"""
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
    return "\n".join(linha.lstrip() for linha in html_final.splitlines() if linha.strip())

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

# Acesso temporário exclusivo para desenvolvimento; remover após configurar OAuth.
# Credenciais ficam somente em Streamlit Secrets, nunca em planilhas públicas.
EMAIL_UNICO_GPS = str(st.secrets.get("GPS_ADMIN_EMAIL", "")).strip().lower()
HASH_UNICO_GPS = str(st.secrets.get("GPS_ADMIN_PASSWORD_HASH", "")).strip()


def validar_hash_gps(senha, registro):
    """Valida PBKDF2-SHA256 com comparação em tempo constante."""
    try:
        algoritmo, iteracoes, salt_hex, hash_hex = str(registro).strip().split("$")
        if algoritmo != "pbkdf2_sha256" or not 600000 <= int(iteracoes) <= 2000000:
            return False
        calculado = hashlib.pbkdf2_hmac(
            "sha256", senha.encode("utf-8"), bytes.fromhex(salt_hex), int(iteracoes)
        )
        return hmac.compare_digest(calculado, bytes.fromhex(hash_hex))
    except (ValueError, TypeError, OverflowError):
        return False


if not st.session_state["autenticado"]:
    st.title("🧭 GPS — Guia Pedagógico Semanal")
    st.subheader("Unidade Regional de Ensino de Miracatu — SEDUC-SP")
    st.markdown("---")
    col_esq, col_centro, col_dir = st.columns([1, 2, 1])
    with col_centro:
        with st.form("form_login_gps"):
            st.markdown("### Acesso temporário — Coordenação")
            email_in = st.text_input("E-mail institucional")
            senha_in = st.text_input("Senha provisória do GPS", type="password")
            entrar = st.form_submit_button("Entrar no Sistema GPS", use_container_width=True)

        if entrar:
            email_normalizado = email_in.strip().lower()
            if not EMAIL_UNICO_GPS or not HASH_UNICO_GPS:
                st.error("Autenticação não configurada. Verifique os Secrets do aplicativo.")
            elif (email_normalizado != EMAIL_UNICO_GPS or not senha_in
                  or not validar_hash_gps(senha_in, HASH_UNICO_GPS)):
                st.error("Acesso não autorizado ou credenciais inválidas.")
            else:
                # Perfil fixo de coordenação: não depende da aba Usuarios.
                st.session_state["autenticado"] = True
                st.session_state["usuario_nome"] = "Coordenação"
                st.session_state["usuario_email"] = EMAIL_UNICO_GPS
                st.session_state["usuario_pasta"] = "CEC"
                st.session_state["usuario_componente"] = "Todos"
                st.session_state["usuario_escolas"] = []
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

# Frequência local: disponibiliza os dados oficiais também nas telas de indicadores,
# independentemente da autenticação do Google Drive.
@st.cache_data(ttl=300, show_spinner=False)
def gps_frequencia_local():
    raiz = pathlib.Path(__file__).resolve().parent
    arq_freq = raiz / "GPS_frequencia_escolas_2026.csv"
    arq_cad = raiz / "GPS_cadastro_40_unidades.csv"
    if not arq_freq.exists() or not arq_cad.exists():
        return pd.DataFrame()
    try:
        freq = pd.read_csv(arq_freq, dtype={"codigo_seduc": str}, encoding="utf-8-sig")
        cad = pd.read_csv(arq_cad, dtype={"codigo_seduc": str}, encoding="utf-8-sig")
        freq["codigo_seduc"] = freq["codigo_seduc"].astype(str).str.strip()
        cad["codigo_seduc"] = cad["codigo_seduc"].astype(str).str.strip()
        dados = freq.merge(cad[["codigo_seduc", "escola"]].drop_duplicates("codigo_seduc"), on="codigo_seduc", how="inner")
        dados["frequencia_pct"] = pd.to_numeric(dados["frequencia_pct"], errors="coerce")
        dados = dados[dados["frequencia_pct"].between(0, 100)].drop_duplicates("codigo_seduc", keep="last")
        return pd.DataFrame({
            "Tipo_Indicador": "FREQUENCIA", "Subcategoria": "Frequência Escolar",
            "Periodo": "Consolidado 2026 (arquivo local)", "Nivel_Granularidade": "ESCOLA",
            "Escola": dados["escola"].values, "Turma": "Consolidado da Escola",
            "Estudante": "Consolidado Escola", "Professor": "", "Recomposicao": "",
            "Componente": "Frequência", "Proficiencia_Pct": dados["frequencia_pct"].values,
            "Participacao_Pct": float("nan"), "Nivel": "",
        })
    except (OSError, ValueError, KeyError, pd.errors.ParserError) as exc:
        st.warning(f"Não foi possível carregar a frequência local: {exc}")
        return pd.DataFrame()

freq_local_gps = gps_frequencia_local()
if not freq_local_gps.empty:
    if df_ind_global.empty:
        df_ind_global = freq_local_gps
    else:
        # O Drive tem precedência se trouxer frequência real; evita duplicidade.
        codigos_existentes = set()
        drive_freq = df_ind_global[df_ind_global["Tipo_Indicador"] == "FREQUENCIA"]
        nomes_drive = set(drive_freq["Escola"].astype(str).map(normalizar_texto))
        locais_novos = freq_local_gps[~freq_local_gps["Escola"].astype(str).map(normalizar_texto).isin(nomes_drive)]
        if not locais_novos.empty:
            df_ind_global = pd.concat([df_ind_global, locais_novos], ignore_index=True)

# Consolida base de ATPC (Sheets + Drive + Upload na sessão)
dfs_atpc_lista = [d for d in [df_atpc_sheets, df_atpc_drive] if not d.empty]
df_atpc_global = pd.concat(dfs_atpc_lista, ignore_index=True).drop_duplicates() if dfs_atpc_lista else pd.DataFrame()

df_ind_usuario = (
    df_ind_global if eh_cec_ou_at
    else filtrar_indicadores_por_pec(
        df_ind_global,
        pasta_atual,
        comp_atual,
        escolas_usuario if eh_pec_qualidade else None,
        df_pastas_raw=df_pastas_raw
    )
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
    st.markdown("### Navegação")
    grupos_gps = {
        "🗺️ Mapa Territorial": ["Visão Regional", "Cadastro das Unidades", "Rotas e Distâncias", "Cobertura Territorial"],
        "📋 Planejamento": ["Plano de Trabalho PEC", "Planner Mensal", "Agenda e Calendário"],
        "📝 Avaliações": ["SARESP", "Prova Paulista", "SAEB", "Avaliação Diagnóstica", "ENEM e Simulados"],
        "🏫 Frequência Escolar": ["Geral URE", "Escolas", "Turmas"],
        "💻 Plataformas Digitais": ["Tarefa SP", "Leia SP", "Redação Paulista", "Matific Anos Finais", "Khan Academy", "Speak", "Open English", "Elefante Letrado", "Matific Anos Iniciais", "Alura", "Robótica", "Programação"],
        "🛡️ Gestão CEC": ["Painel Exclusivo CEC", "Acompanhamentos", "Relatórios", "Administração — Fontes de Dados"],
        "📚 Diretrizes": ["Diretrizes das Pastas"],
    }
    if not eh_cec_ou_at:
        grupos_gps.pop("🛡️ Gestão CEC")
    if "gps_destino" not in st.session_state:
        st.session_state["gps_destino"] = "Visão Regional"
    for titulo_grupo, subitens in grupos_gps.items():
        with st.expander(titulo_grupo, expanded=st.session_state["gps_destino"] in subitens):
            for item in subitens:
                if st.button(item, key=f"gps_menu_{titulo_grupo}_{item}", use_container_width=True,
                             type="primary" if st.session_state["gps_destino"] == item else "secondary"):
                    st.session_state["gps_destino"] = item
    destino_gps = st.session_state["gps_destino"]
    pagina_gps = ("🗺️ Mapa Territorial" if destino_gps in grupos_gps["🗺️ Mapa Territorial"] else
                  "📁 Fontes de Dados" if destino_gps == "Administração — Fontes de Dados" else
                  "📋 Planejamento e Indicadores")
    st.markdown("---")
    if st.button("🔄 Sincronizar Planilhas & Drive", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    if st.button("🚪 Sair / Trocar Usuário", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# ==============================================================================
# MAPA TERRITORIAL — HOME (cadastro validado por código SEDUC)
# ==============================================================================
# Mapa: geometrias oficiais consultadas ao IBGE; seleção por clique e ficha da unidade.
@st.cache_data(ttl=86400, show_spinner=False)
def gps_limites_ibge():
    """Retorna GeoJSON somente dos seis municípios; falha explícita sem polígonos fictícios."""
    nomes = {"MIRACATU", "IGUAPE", "ITARIRI", "JUQUIA", "PEDRO DE TOLEDO", "ILHA COMPRIDA"}
    try:
        url = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/35/municipios"
        with urllib.request.urlopen(url, timeout=12) as resp:
            municipios_ibge = json.load(resp)
        codigos = {normalizar_texto(x["nome"]): str(x["id"]) for x in municipios_ibge}
        features = []
        for nome in sorted(nomes):
            codigo = codigos.get(nome)
            if not codigo:
                raise ValueError(f"Município não localizado no IBGE: {nome}")
            url_geo = f"https://servicodados.ibge.gov.br/api/v3/malhas/municipios/{codigo}?formato=application/vnd.geo+json&qualidade=minima"
            with urllib.request.urlopen(url_geo, timeout=15) as resp:
                geometria = json.load(resp)
            if geometria.get("type") == "FeatureCollection":
                geo = geometria["features"][0]["geometry"]
            elif geometria.get("type") == "Feature":
                geo = geometria["geometry"]
            else:
                geo = geometria
            if not geo or geo.get("type") not in ("Polygon", "MultiPolygon"):
                raise ValueError(f"Geometria inválida: {nome}")
            features.append({"type":"Feature", "geometry":geo, "properties":{"municipio":nome}})
        return {"type":"FeatureCollection", "features":features}, None
    except Exception as exc:
        return None, str(exc)


def gps_centroide_rotulo_municipio(geometria):
    """Centroide ponderado dos anéis externos, para rotulagem visual (não geodésica)."""
    poligonos = ([geometria["coordinates"]] if geometria["type"] == "Polygon"
                 else geometria["coordinates"])
    partes = []
    for poligono in poligonos:
        anel = poligono[0]
        area_dupla = cx = cy = 0.0
        for i in range(len(anel) - 1):
            x0, y0 = anel[i]
            x1, y1 = anel[i + 1]
            cruz = x0 * y1 - x1 * y0
            area_dupla += cruz
            cx += (x0 + x1) * cruz
            cy += (y0 + y1) * cruz
        if abs(area_dupla) > 1e-12:
            partes.append((abs(area_dupla), cx / (3 * area_dupla), cy / (3 * area_dupla)))
    if not partes:
        raise ValueError("Geometria sem área para posicionar o nome do município")
    total = sum(parte[0] for parte in partes)
    return (sum(a * x for a, x, y in partes) / total,
            sum(a * y for a, x, y in partes) / total)


def gps_identificar_escola(texto, cadastro):
    """Correspondência conservadora; códigos SEDUC prevalecem sobre nomes."""
    texto = str(texto or "")
    for _, escola in cadastro.iterrows():
        codigo = str(escola["codigo_seduc"])
        if re.search(r"(?<!\d)" + re.escape(codigo) + r"(?!\d)", texto):
            return codigo
    texto_norm = normalizar_texto(texto)
    # Ignore generic prefix and titles; compare distinctive school name tokens.
    texto_norm = re.sub(r"\b(ESCOLA|ESTADUAL|EE|PROFESSORA?|PREFEITO|ENGENHEIRO|DE|DA|DO|DOS|DAS)\b", " ", texto_norm)
    texto_norm = " ".join(texto_norm.split())
    achados = []
    for _, escola in cadastro.iterrows():
        nome = normalizar_texto(str(escola["escola"]))
        nome = re.sub(r"\b(ESCOLA|ESTADUAL|EE|PROFESSORA?|PREFEITO|ENGENHEIRO|DE|DA|DO|DOS|DAS)\b", " ", nome)
        nome = " ".join(nome.split())
        if len(nome) >= 6 and (nome in texto_norm or texto_norm in nome and len(texto_norm) >= 9):
            achados.append(str(escola["codigo_seduc"]))
    return achados[0] if len(achados) == 1 else None


# Ultimate Pro: distância rodoviária via OSRM; falha não é substituída por quilometragem fictícia.
@st.cache_data(ttl=86400, show_spinner=False)
def gps_distancia_rota(pontos):
    """Retorna km e minutos por percurso rodoviário; None se serviço indisponível."""
    if len(pontos) < 2:
        return 0.0, 0.0
    coords = ";".join(f"{float(lon):.6f},{float(lat):.6f}" for lat, lon in pontos)
    url = f"https://router.project-osrm.org/route/v1/driving/{coords}?overview=false&steps=false"
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent":"GPS-URE-Miracatu/1.0"}), timeout=8) as response:
            data = json.load(response)
        if data.get("code") == "Ok" and data.get("routes"):
            rota = data["routes"][0]
            return round(rota["distance"] / 1000, 1), round(rota["duration"] / 60)
    except Exception:
        pass
    return None, None


def gps_distancia_direta(p1, p2):
    from math import radians, sin, cos, asin, sqrt
    a, b = p1; c, d = p2
    dl = radians(c-a); dn = radians(d-b)
    h = sin(dl/2)**2 + cos(radians(a))*cos(radians(c))*sin(dn/2)**2
    return round(12742 * asin(sqrt(h)), 1)


def gps_resumo_quilometragem(agenda, cadastro, sede=(-24.285437,-47.458814)):
    """Cada PEC/dia: URE -> escolas na ordem do horário quando informado -> URE."""
    cols = ["PEC", "Data", "Escolas", "Visitas previstas", "Km rodoviários", "Minutos estimados", "Situação da rota"]
    if agenda.empty:
        return pd.DataFrame(columns=cols)
    coords = {str(r["codigo_seduc"]):(float(r["latitude"]),float(r["longitude"])) for _,r in cadastro.iterrows()}
    saida=[]
    for (pec, dia), grupo in agenda.groupby(["PEC", "Data"]):
        grupo = grupo.sort_values("Horário", kind="stable")
        codigos = list(dict.fromkeys(str(x) for x in grupo["codigo_seduc"] if str(x) in coords))
        if not codigos:
            continue
        pontos = [sede] + [coords[c] for c in codigos] + [sede]
        km, minutos = gps_distancia_rota(tuple(pontos))
        saida.append({"PEC":pec,"Data":dia,"Escolas":len(codigos),"Visitas previstas":len(grupo),
                      "Km rodoviários":km,"Minutos estimados":minutos,
                      "Situação da rota":"Calculada (estimativa)" if km is not None else "Rota indisponível"})
    return pd.DataFrame(saida,columns=cols)


def gps_carregar_agenda_consolidada(cadastro):
    """Agenda normalizada, auditável, opcional; nunca confunde plano de ação com visita datada."""
    from pathlib import Path
    arq = Path(__file__).resolve().parent / "GPS_agenda_pecs_novembro_2026_COMPLETA.csv"
    cols = ["codigo_seduc", "PEC", "Pasta", "Data", "Atividade", "Fonte", "Situação", "Objetivo", "Horário"]
    if not arq.exists():
        st.warning("Agenda de visitas não encontrada: envie GPS_agenda_pecs_novembro_2026_COMPLETA.csv para a mesma pasta do app.py no GitHub. Sem esse arquivo, o mapa não consegue mostrar visitas programadas que não estejam nos planos já importados.")
        return pd.DataFrame(columns=cols)
    try:
        base = pd.read_csv(arq, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        obrig = {"PEC", "Data", "Escola", "Atividade", "Objetivo", "Fonte"}
        if not obrig.issubset(base.columns):
            st.warning("Agenda consolidada incompleta: faltam colunas " + ", ".join(sorted(obrig-set(base.columns))))
            return pd.DataFrame(columns=cols)
        saida = []
        ignorados = []
        for indice, r in base.iterrows():
            data = pd.to_datetime(str(r["Data"]), format="%d/%m/%Y", errors="coerce")
            codigo = str(r.get("codigo_seduc", "")).strip()
            if codigo not in set(cadastro["codigo_seduc"].astype(str)):
                codigo = gps_identificar_escola(r["Escola"], cadastro)
            if pd.isna(data) or not codigo:
                ignorados.append({"linha_csv": indice + 2, "PEC": r.get("PEC", ""),
                                  "Escola": r.get("Escola", ""), "Data": r.get("Data", ""),
                                  "Motivo": "Data inválida" if pd.isna(data) else "Escola não identificada com segurança"})
                continue
            saida.append({"codigo_seduc": codigo, "PEC": r["PEC"], "Pasta": r.get("Pasta", "Não informada"),
                "Data": data.date(), "Atividade": r["Atividade"], "Fonte": r["Fonte"],
                "Situação": "Programada", "Objetivo": r["Objetivo"], "Horário": r.get("Horário", "Não informado")})
        if ignorados:
            st.warning(f"Agenda: {len(ignorados)} registro(s) não foram reconhecidos. Confira as linhas abaixo.")
            with st.expander("Ver registros rejeitados da agenda"):
                st.dataframe(pd.DataFrame(ignorados), hide_index=True, use_container_width=True)
        return pd.DataFrame(saida, columns=cols)
    except Exception as exc:
        st.warning(f"Não foi possível ler a agenda consolidada: {exc}")
        return pd.DataFrame(columns=cols)


def gps_agenda_planos(df_planos, cadastro):
    """Exibe somente informações comprovadas; resumo sem data não vira visita de hoje."""
    if df_planos is None or df_planos.empty:
        return pd.DataFrame(columns=["codigo_seduc", "PEC", "Pasta", "Data", "Atividade", "Fonte", "Situação", "Objetivo", "Horário"])
    registros = []
    for _, r in df_planos.iterrows():
        data_txt = str(r.get("Data_Visita", r.get("Data_Prevista", r.get("Data", ""))))
        data_dt = pd.to_datetime(data_txt, dayfirst=True, errors="coerce")
        if pd.isna(data_dt):
            continue  # Data_Envio não é data da visita.
        codigo = gps_identificar_escola(str(r.get("Escola", "")), cadastro)
        if not codigo:
            continue
        registros.append({"codigo_seduc":codigo, "PEC":str(r.get("PEC", "Não informado")),
                          "Pasta":str(r.get("Pasta", r.get("Área", "Não informada"))),
                          "Data":data_dt.date(),
                          "Atividade":str(r.get("Atividade", r.get("Objetivo", "Não detalhada no plano"))),
                          "Fonte":str(r.get("Arquivo_Plano", "Plano de Trabalho")),
                          "Situação":"Programada",
                          "Objetivo":str(r.get("Objetivo", r.get("Foco", "Não informado"))),
                          "Horário":str(r.get("Horário", r.get("Horario", "Não informado")))})
    return pd.DataFrame(registros, columns=["codigo_seduc", "PEC", "Pasta", "Data", "Atividade", "Fonte", "Situação", "Objetivo", "Horário"])


def gps_ficha_escola(registro, agenda, termos, data_referencia=None):
    """V11 Plus: painel da escola com menu lateral e cartões temáticos.

    Não cria indicadores não recebidos e não confunde plano com execução.
    """
    from html import escape as _esc
    escola = str(registro.get("escola", "Escola"))
    codigo = str(registro.get("codigo_seduc", ""))
    data_referencia = data_referencia or date.today()
    mes, ano = data_referencia.month, data_referencia.year
    st.markdown("""<style>
    .gps11-title{font-size:1.45rem;font-weight:800;color:#17313d;margin-bottom:4px}
    .gps11-card{padding:18px 17px;border-radius:17px;border:1px solid rgba(24,45,56,.08);min-height:145px}
    .gps11-card .gps11-label{font-size:.88rem;font-weight:700;letter-spacing:.01em}
    .gps11-card .gps11-number{font-size:2rem;font-weight:800;margin:12px 0 2px}
    .gps11-card .gps11-note{font-size:.8rem;opacity:.8}
    .gps11-blue{background:#e8f3fc;color:#174d78}
    .gps11-purple{background:#f0eafb;color:#57328a}
    .gps11-orange{background:#fff1e3;color:#8b4a0c}
    .gps11-green{background:#e6f4ee;color:#126344}
    .gps11-teal{background:#e5f5f4;color:#125f67}
    .gps11-red{background:#fcebed;color:#9d3441}
    </style>""", unsafe_allow_html=True)
    st.markdown(f'<div class="gps11-title">🏫 {_esc(escola)}</div>', unsafe_allow_html=True)
    st.caption(f'{registro.get("municipio", "")} · Código SEDUC {codigo} · {registro.get("tipo", "")}')
    if agenda is None or agenda.empty or "codigo_seduc" not in agenda:
        itens = pd.DataFrame()
    else:
        itens = agenda[agenda["codigo_seduc"].astype(str) == codigo].copy()
    if not itens.empty:
        datas = pd.to_datetime(itens["Data"], errors="coerce")
        itens_mes = itens[(datas.dt.month == mes) & (datas.dt.year == ano)].copy()
        itens_dia = itens[datas.dt.date == data_referencia].copy()
    else:
        itens_mes, itens_dia = pd.DataFrame(), pd.DataFrame()
    st.caption(f'Consulta de {mes:02d}/{ano} · {len(itens_mes)} visita(s) programada(s) nos dados carregados · {len(itens_dia)} na data selecionada')
    lateral, conteudo = st.columns([1, 3], gap="large")
    secoes = {
        "🏫 Identificação": ["Dados da unidade", "Equipe gestora", "Matrículas e turmas"],
        "👥 Frequência Escolar": ["Frequência geral", "Frequência por turma", "Turmas prioritárias"],
        "📊 Avaliações": ["Raio-X Pedagógico", "Prova Paulista", "SARESP", "SAEB", "Avaliação Diagnóstica", "ENEM e Simulados"],
        "💻 Plataformas Digitais": ["Visão geral", "Tarefa SP", "Leia SP", "Redação Paulista", "Matific", "Khan Academy", "Speak / Open English", "Alura e Tecnologia"],
        "📅 Acompanhamentos": ["Visitas do dia", "Visitas do mês", "Termos de Visita", "Justificativas", "Pendências documentais"],
    }
    with lateral:
        st.markdown("**Navegação da escola**")
        grupo = st.radio("Área", list(secoes), key="gps11_area", label_visibility="collapsed")
        subitem = st.radio("Subitem", secoes[grupo], key=f"gps11_sub_{grupo}", label_visibility="collapsed")
    with conteudo:
        st.subheader(subitem)
        def cartao(titulo, valor, detalhe, cor):
            st.markdown(f'<div class="gps11-card gps11-{cor}"><div class="gps11-label">{_esc(titulo)}</div>'
                        f'<div class="gps11-number">{_esc(str(valor))}</div><div class="gps11-note">{_esc(detalhe)}</div></div>',
                        unsafe_allow_html=True)
        if subitem == "Raio-X Pedagógico":
            a,b = st.columns(2)
            with a: cartao("👥 Frequência geral", "—", "Farol: sem dados validados", "blue")
            with b: cartao("📊 Prova Paulista", "—", "Farol: sem dados validados", "purple")
            a,b = st.columns(2)
            with a: cartao("📖 Abaixo do Básico · Português", "—", "Quantidade e percentual não disponíveis", "orange")
            with b: cartao("🔢 Abaixo do Básico · Matemática", "—", "Quantidade e percentual não disponíveis", "green")
            st.caption("Aguardando dados oficiais. Não é possível inferir níveis de proficiência a partir do percentual de acertos.")
        elif subitem in ("Visitas do dia", "Visitas do mês"):
            dados = itens_dia if subitem == "Visitas do dia" else itens_mes
            a,b = st.columns(2)
            with a: cartao("📅 Visitas previstas", len(dados), "Conforme os planos efetivamente carregados", "teal")
            with b: cartao("👤 PECs distintos", dados["PEC"].nunique() if not dados.empty and "PEC" in dados else 0,
                           "PECs identificados na programação", "blue")
            if dados.empty:
                st.info("Nenhuma visita localizada nos documentos carregados para este período. Isso não comprova ausência de planejamento ou visita.")
            else:
                campos = [c for c in ("Data", "PEC", "Horário", "Atividade", "Objetivo", "Situação", "Fonte") if c in dados]
                st.dataframe(dados[campos].sort_values("Data"), hide_index=True, use_container_width=True)
        elif subitem == "Dados da unidade":
            st.write(f'**Endereço:** {registro.get("logradouro", "Não informado")}, {registro.get("numero", "")}, {registro.get("municipio", "")}')
            lat, lon = registro.get("latitude"), registro.get("longitude")
            if pd.notna(lat) and pd.notna(lon):
                st.link_button("📍 Ver rota no Google Maps", f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}")
        elif subitem == "Termos de Visita":
            if termos is None or termos.empty or "Escola" not in termos:
                st.info("Nenhum Termo de Visita validado disponível nesta fonte.")
            else:
                nome = normalizar_texto(escola)
                correspondentes = termos[termos["Escola"].astype(str).map(lambda x: normalizar_texto(x) == nome)]
                if correspondentes.empty:
                    st.info("Nenhum termo associado com segurança a esta escola. Isso não comprova que a visita não ocorreu.")
                else:
                    st.dataframe(correspondentes, hide_index=True, use_container_width=True)
        elif subitem == "Pendências documentais":
            cartao("📄 Conferência documental", "—", "Aguardando cruzamento validado de Plano × Termo × Justificativa", "red")
            st.caption("Ausência de termo não comprova não realização; ausência de justificativa exige conferência.")
        elif subitem == "Justificativas":
            st.info("As justificativas serão exibidas após importação e vinculação segura por PEC, escola e data.")
        elif grupo == "👥 Frequência Escolar":
            cartao("👥 " + subitem, "—", "Sem dados oficiais sincronizados", "blue")
            st.info("A frequência por turma depende da importação validada da fonte institucional.")
        elif grupo == "📊 Avaliações":
            cartao("📊 " + subitem, "—", "Sem dados oficiais sincronizados", "purple")
        elif grupo == "💻 Plataformas Digitais":
            cartao("💻 " + subitem, "—", "Sem dados oficiais sincronizados", "teal")
        else:
            cartao("🏫 " + subitem, "—", "Informações aguardando cadastro validado", "blue")

FERIADOS_EEC = {date(2026, 11, 30): "Aniversário de Miracatu — feriado municipal da sede da EEC"}
PERIODOS_CEC = [(date(2026, 11, 16), date(2026, 11, 29)),
                (date(2026, 11, 30), date(2026, 12, 11))]

def gps_dia_util_eec(dia):
    return dia.weekday() < 5 and dia not in holidays.country_holidays("BR", subdiv="SP", years=[dia.year]) and dia not in FERIADOS_EEC

def gps_fim_quinzena(data_visita):
    for inicio, fim in PERIODOS_CEC:
        if inicio <= data_visita <= fim:
            return fim
    return None  # Não inventar quinzenas institucionais futuras.

def gps_data_segura(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    txt = str(valor or "").strip()
    if not txt or txt.lower() in ("nan", "nat", "none", "-"):
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:10], fmt).date()
        except ValueError:
            pass
    return None

def gps_chave_pec(valor):
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(valor))

def gps_chave_escola(valor):
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(valor).removeprefix("EE "))

def gps_conciliar_documentos(agenda, termos, justificativas=None, hoje=None):
    """Conciliação por PEC, escola e DATA; nunca confunde ausência de termo com falta à visita.

    Justificativas estruturadas: PEC, Escola, Data_Prevista, Situacao,
    Motivo, Nova_Data, Nova_Escola (as duas últimas opcionais).
    """
    hoje = hoje or date.today()
    justificativas = justificativas if justificativas is not None else pd.DataFrame()
    agenda = agenda if agenda is not None else pd.DataFrame()
    termos = termos if termos is not None else pd.DataFrame()
    colunas = ["PEC", "Escola", "Data prevista", "Situação", "Alerta", "Motivo / observação", "Termo", "Justificativa", "Fim do período"]
    if agenda.empty or not {"PEC", "Escola", "Data"}.issubset(agenda.columns):
        return pd.DataFrame(columns=colunas)
    indice_termos = {}
    for _, t in termos.iterrows():
        dia = gps_data_segura(t.get("Data_Visita"))
        if dia and t.get("PEC") and t.get("Escola"):
            chave = (gps_chave_pec(t["PEC"]), gps_chave_escola(t["Escola"]), dia)
            indice_termos.setdefault(chave, []).append(str(t.get("Arquivo_Termo_PDF", "Termo registrado")))
    indice_just = {}
    if not justificativas.empty:
        for _, j in justificativas.iterrows():
            dia = gps_data_segura(j.get("Data_Prevista"))
            if dia and j.get("PEC") and j.get("Escola"):
                chave = (gps_chave_pec(j["PEC"]), gps_chave_escola(j["Escola"]), dia)
                indice_just.setdefault(chave, []).append(j)
    saida, chaves_agenda, termos_usados = [], set(), set()
    for _, a in agenda.iterrows():
        dia = gps_data_segura(a.get("Data"))
        if not dia:
            continue
        pec, escola = str(a.get("PEC", "")), str(a.get("Escola", ""))
        if not pec.strip() or not escola.strip():
            continue
        chave = (gps_chave_pec(pec), gps_chave_escola(escola), dia)
        if chave in chaves_agenda:
            continue
        chaves_agenda.add(chave)
        fim = gps_fim_quinzena(dia)
        docs = indice_termos.get(chave, [])
        justs = indice_just.get(chave, [])
        justificativa = justs[0] if justs else None
        motivo = str(justificativa.get("Motivo", "")).strip() if justificativa is not None else ""
        situacao = normalizar_texto(justificativa.get("Situacao", "")) if justificativa is not None else ""
        termo_nome = "; ".join(docs)
        justificativa_txt = situacao if justificativa is not None else "Não localizada"
        if dia in FERIADOS_EEC:
            status, alerta, obs = "Programação em feriado institucional", "🔴 Verificar planejamento", "30/11: nenhum PEC em exercício; revisar agendamento"
        elif docs:
            termos_usados.add(chave)
            if "NAO REALIZ" in situacao or "CANCEL" in situacao:
                status, alerta, obs = "Documentos contraditórios", "🔴 Divergência", "Termo e justificativa de não realização para mesma visita"
            else:
                status, alerta, obs = "Termo localizado", "🟢 Regular", "Visita documentada por termo correspondente"
        elif "NAO REALIZ" in situacao or "CANCEL" in situacao:
            if motivo:
                status, alerta, obs = "Não realizada — justificada", "🟢 Justificativa registrada", motivo
            else:
                status, alerta, obs = "Não realizada — motivo ausente", "🔴 Justificativa incompleta", "Informar por que a visita não ocorreu"
        elif "ALTER" in situacao or "REMAR" in situacao or "REAGEND" in situacao:
            nova_data = gps_data_segura(justificativa.get("Nova_Data"))
            nova_escola = str(justificativa.get("Nova_Escola", "")).strip() or escola
            nova_chave = (gps_chave_pec(pec), gps_chave_escola(nova_escola), nova_data)
            docs_novos = indice_termos.get(nova_chave, []) if nova_data else []
            if not motivo or not nova_data:
                status, alerta, obs = "Alteração incompleta", "🔴 Justificativa incompleta", "Informar motivo e nova data"
            elif docs_novos:
                termos_usados.add(nova_chave)
                termo_nome = "; ".join(docs_novos)
                status, alerta, obs = "Alteração documentada", "🟢 Regular", motivo
            elif nova_data > hoje:
                status, alerta, obs = "Alteração justificada — futura", "⚪ Aguardando visita", motivo
            else:
                status, alerta, obs = "Alteração justificada — termo não localizado", "🟡 Verificar execução / termo", motivo
        elif dia > hoje:
            status, alerta, obs = "Visita programada", "⚪ Aguardando data", "Ainda não há evidência de realização"
        elif fim and hoje > fim:
            status, alerta, obs = "Execução não esclarecida", "🔴 Pendência documental / esclarecimento", "Sem termo e sem justificativa correspondente; não concluir que a visita não ocorreu"
        elif (hoje - dia).days >= 7:
            status, alerta, obs = "Execução não esclarecida", "🟡 Verificar termo / execução", "Sem termo e sem justificativa correspondente"
        else:
            status, alerta, obs = "Execução não esclarecida", "⚪ Aguardando conferência", "Sem termo; realização ainda não confirmada"
        saida.append({"PEC": pec, "Escola": escola, "Data prevista": dia.strftime("%d/%m/%Y"),
                      "Situação": status, "Alerta": alerta, "Motivo / observação": obs,
                      "Termo": termo_nome or "Não localizado", "Justificativa": justificativa_txt,
                      "Fim do período": fim.strftime("%d/%m/%Y") if fim else "Período não definido"})
    # Termos não associados a uma visita prevista exigem explicação de alteração do plano.
    for chave, docs in indice_termos.items():
        if chave in termos_usados or chave in chaves_agenda:
            continue
        pec, escola, dia = chave
        saida.append({"PEC": pec, "Escola": escola, "Data prevista": dia.strftime("%d/%m/%Y"),
                      "Situação": "Termo sem previsão correspondente", "Alerta": "🟡 Verificar alteração do plano",
                      "Motivo / observação": "Conferir justificativa de inclusão ou mudança da programação",
                      "Termo": "; ".join(docs), "Justificativa": "Verificar", "Fim do período": "—"})
    return pd.DataFrame(saida, columns=colunas)

def gps_alertas_termos(agenda, termos, hoje=None):
    """Compatibilidade com o painel V8; retorna todas as situações auditáveis."""
    return gps_conciliar_documentos(agenda, termos, hoje=hoje)


if pagina_gps == "🗺️ Mapa Territorial":
    import pathlib
    import pydeck as pdk
    from urllib.parse import quote as gps_url_quote
    st.title("🗺️ Mapa Territorial — URE Miracatu | Ultimate Pro")
    caminho_csv = pathlib.Path(__file__).resolve().parent / "GPS_cadastro_40_unidades.csv"
    if not caminho_csv.exists():
        st.error("Arquivo GPS_cadastro_40_unidades.csv não encontrado junto ao app.py.")
        st.stop()
    cadastro = pd.read_csv(caminho_csv, dtype={"codigo_seduc":str}).fillna("")
    for campo in ("latitude", "longitude"):
        cadastro[campo] = pd.to_numeric(cadastro[campo], errors="coerce")
    cadastro = cadastro.dropna(subset=["latitude", "longitude"]).drop_duplicates("codigo_seduc").copy()
    cores = {"MIRACATU":[15,107,125], "IGUAPE":[138,70,166], "ITARIRI":[208,106,44],
             "JUQUIA":[38,135,75], "PEDRO DE TOLEDO":[191,66,100], "ILHA COMPRIDA":[42,99,189]}
    cadastro["municipio_norm"] = cadastro["municipio"].map(lambda x: normalizar_texto(str(x)))
    cadastro["cor"] = cadastro["municipio_norm"].map(lambda x: cores.get(x,[90,90,90]) + [240])
    cadastro["endereco"] = cadastro["logradouro"].astype(str) + ", " + cadastro["numero"].astype(str)
    agenda_drive = gps_agenda_planos(df_planos_drive, cadastro)
    agenda_csv = gps_carregar_agenda_consolidada(cadastro)
    agenda = pd.concat([agenda_drive, agenda_csv], ignore_index=True).drop_duplicates(
        subset=["codigo_seduc", "PEC", "Data", "Atividade"], keep="last")

    st.caption("Localização **prevista**, não rastreamento físico. Histórico documentado por termos e alterações justificadas quando disponíveis. Atualize as fontes para recalcular.")
    hoje_mapa = date.today()
    # O calendário e a justificativa são fontes auditáveis; ausência de termo não significa ausência da visita.
    justificativas_mapa = pd.DataFrame()
    from pathlib import Path as _GPSPath
    arquivo_just_mapa = _GPSPath(__file__).resolve().parent / "GPS_justificativas_validadas.csv"
    if arquivo_just_mapa.exists():
        try:
            justificativas_mapa = pd.read_csv(arquivo_just_mapa, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            if not {"PEC","Escola","Data_Prevista","Situacao","Motivo"}.issubset(justificativas_mapa.columns):
                st.warning("Justificativas estruturadas: colunas obrigatórias ausentes. Arquivo ignorado.")
                justificativas_mapa = pd.DataFrame()
        except Exception as exc:
            st.warning(f"Não foi possível carregar justificativas validadas: {exc}")
    # Termos com data validada podem corrigir retrospectivamente a cobertura; nunca alterar o plano original.
    registros_termos = []
    if df_termos_drive is not None and not df_termos_drive.empty:
        for _, tr in df_termos_drive.iterrows():
            data_t = pd.to_datetime(str(tr.get("Data_Visita", "")), dayfirst=True, errors="coerce")
            cod_t = gps_identificar_escola(str(tr.get("Escola", "")), cadastro)
            if pd.notna(data_t) and cod_t and str(tr.get("PEC", "")).strip():
                registros_termos.append({"codigo_seduc":cod_t,"PEC":str(tr["PEC"]),"Data":data_t.date(),
                    "Atividade":"Acompanhamento documentado","Objetivo":"Consultar termo original",
                    "Horário":"Não informado","Pasta":"Não informada","Fonte":str(tr.get("Arquivo_Termo_PDF", "Termo")),
                    "Situação":"Documentada"})
    termos_mapa = pd.DataFrame(registros_termos)
    # Consolidação mensal: somente termos vinculados comprovam visitas executadas.
    # Justificativas válidas encerram pendências, mas não somam visitas não realizadas.
    def gps_consolidacao_mensal(ano, mes):
        previstos = agenda[agenda["Data"].apply(lambda d: isinstance(d, date) and d.year == ano and d.month == mes)].copy()
        documentos = termos_mapa[termos_mapa["Data"].apply(lambda d: isinstance(d, date) and d.year == ano and d.month == mes)].copy() if not termos_mapa.empty else termos_mapa.copy()
        if previstos.empty:
            return False, documentos.iloc[0:0].copy(), 0, 0, "Sem programação mensal validada para conciliar."
        previstos["_chave"] = previstos.apply(lambda r: (gps_chave_pec(r["PEC"]), str(r["codigo_seduc"]), r["Data"]), axis=1)
        chaves_previstas = set(previstos["_chave"])
        chaves_termos = set()
        if not documentos.empty:
            documentos["_chave"] = documentos.apply(lambda r: (gps_chave_pec(r["PEC"]), str(r["codigo_seduc"]), r["Data"]), axis=1)
            chaves_termos = set(documentos["_chave"])
        chaves_justificadas = set()
        if not justificativas_mapa.empty:
            for _, j in justificativas_mapa.iterrows():
                dia = gps_data_segura(j.get("Data_Prevista"))
                if not dia or dia.year != ano or dia.month != mes or not str(j.get("Motivo", "")).strip():
                    continue
                situacao = normalizar_texto(j.get("Situacao", ""))
                # Remanejamento só encerra a pendência se houver termo da nova visita.
                if any(x in situacao for x in ("NAO REALIZ", "CANCEL")):
                    codigo = gps_identificar_escola(str(j.get("Escola", "")), cadastro)
                    if codigo:
                        chaves_justificadas.add((gps_chave_pec(j.get("PEC", "")), str(codigo), dia))
                elif any(x in situacao for x in ("ALTER", "REMAR", "REAGEND")):
                    nova_data = gps_data_segura(j.get("Nova_Data"))
                    nova_escola = str(j.get("Nova_Escola", "")).strip() or str(j.get("Escola", ""))
                    codigo_novo = gps_identificar_escola(nova_escola, cadastro)
                    codigo_original = gps_identificar_escola(str(j.get("Escola", "")), cadastro)
                    if nova_data and codigo_novo and codigo_original and (gps_chave_pec(j.get("PEC", "")), str(codigo_novo), nova_data) in chaves_termos:
                        chaves_justificadas.add((gps_chave_pec(j.get("PEC", "")), str(codigo_original), dia))
        pendentes = chaves_previstas - chaves_termos - chaves_justificadas
        # Termos sem correspondência com a programação não são consolidados automaticamente.
        confirmados = documentos[documentos["_chave"].isin(chaves_previstas)].drop_duplicates("_chave") if not documentos.empty else documentos.copy()
        completo = not pendentes and not (chaves_termos - chaves_previstas)
        motivo = "Consolidação documental concluída." if completo else "Existem registros sem termo ou justificativa válida, ou termos sem previsão correspondente."
        return completo, confirmados, len(chaves_previstas), len(pendentes), motivo

    col_data, col_pec = st.columns([1, 2])
    data_ref = col_data.date_input("Data de referência", value=hoje_mapa, key="gps_mapa_data_ref")
    pecs_mapa = sorted(set(agenda.get("PEC", pd.Series(dtype=str)).dropna().astype(str).tolist()) |
                       set(termos_mapa.get("PEC", pd.Series(dtype=str)).dropna().astype(str).tolist()))
    pec_mapa = col_pec.selectbox("PEC", ["Todos"] + pecs_mapa, key="gps_mapa_pec")
    st.markdown("#### 📊 Consolidação mensal de visitas")
    mes_grafico = st.selectbox("Mês para o mapa de calor e os gráficos", options=list(range(1, 13)),
                              index=data_ref.month - 1, format_func=lambda m: f"{m:02d}/{data_ref.year}",
                              key="gps_mes_graficos")
    consolidado, visitas_confirmadas_mes, total_previstas_mes, pendencias_mes, motivo_mes = gps_consolidacao_mensal(data_ref.year, mes_grafico)
    # Atualização progressiva: cada termo vinculado ao planejamento já alimenta
    # os gráficos e o mapa de calor, mesmo com outras visitas pendentes.
    base_graficos = visitas_confirmadas_mes.copy()
    if not base_graficos.empty:
        st.success(f"{mes_grafico:02d}/{data_ref.year}: {len(base_graficos)} visita(s) documentada(s) até o momento.")
        if pendencias_mes:
            st.caption(f"Consolidação parcial: {pendencias_mes} de {total_previstas_mes} visitas programadas ainda aguardam termo ou justificativa conclusiva.")
    else:
        st.info(f"Ainda não há termos vinculados à programação de {mes_grafico:02d}/{data_ref.year}. Os gráficos e o mapa de calor serão preenchidos conforme os termos forem registrados.")
    if not base_graficos.empty:
        # Atualiza progressivamente, sem aguardar o fechamento do mês.
        if pec_mapa != "Todos":
            base_graficos = base_graficos[base_graficos["PEC"] == pec_mapa]
        if not base_graficos.empty:
            base_graficos["Escola"] = base_graficos["codigo_seduc"].map(cadastro.set_index("codigo_seduc")["escola"].to_dict())
            base_graficos["Município"] = base_graficos["codigo_seduc"].map(cadastro.set_index("codigo_seduc")["municipio"].to_dict())
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("**Visitas documentadas por PEC**")
                st.bar_chart(base_graficos.groupby("PEC").size().rename("Visitas"))
                st.markdown("**Visitas documentadas por município**")
                st.bar_chart(base_graficos.groupby("Município").size().rename("Visitas"))
            with g2:
                st.markdown("**Visitas documentadas por escola**")
                st.bar_chart(base_graficos.groupby("Escola").size().rename("Visitas"))
            st.caption("Somente termos de visita vinculados ao plano. Justificativas de não realização encerram pendências, mas não são contadas como visitas.")
        else:
            st.info("Nenhuma visita documentada corresponde ao PEC selecionado neste mês.")
    # Calendário único: data futura/atual = agenda; data passada = documentos
    # e programação original, sem presumir visita realizada por ausência de termo.
    modo_tempo = "Passado" if data_ref < hoje_mapa else ("Presente" if data_ref == hoje_mapa else "Futuro")
    agenda_dia = agenda[agenda["Data"] == data_ref].copy() if not agenda.empty else agenda.copy()
    termos_dia = termos_mapa[termos_mapa["Data"] == data_ref].copy() if not termos_mapa.empty else termos_mapa.copy()
    # Exibir toda a programação, inclusive datas passadas. O termo confirma
    # a documentação, não é requisito para que uma visita prevista apareça.
    visitas_periodo = agenda_dia.copy()
    if modo_tempo == "Passado":
        st.caption("Data passada: programação original preservada; a ausência de termo não comprova não realização. Confira a situação documental na ficha.")
        if not termos_dia.empty:
            chaves_previstas = set(zip(visitas_periodo.get("codigo_seduc", pd.Series(dtype=str)).astype(str),
                                       visitas_periodo.get("PEC", pd.Series(dtype=str)).astype(str)))
            adicionais = termos_dia[[ (str(r["codigo_seduc"]),str(r["PEC"])) not in chaves_previstas for _,r in termos_dia.iterrows() ]]
            if not adicionais.empty:
                visitas_periodo = pd.concat([visitas_periodo, adicionais], ignore_index=True)
    if data_ref in FERIADOS_EEC:
        if not agenda_dia.empty:
            st.warning("30/11/2026: feriado institucional para toda a EEC. Existem registros planejados para essa data; verifique o remanejamento.")
        visitas_periodo = visitas_periodo.iloc[0:0].copy()
    if pec_mapa != "Todos" and not visitas_periodo.empty:
        visitas_periodo = visitas_periodo[visitas_periodo["PEC"]==pec_mapa]
    resumo = {}
    if not visitas_periodo.empty:
        for cod, g in visitas_periodo.groupby("codigo_seduc"):
            resumo[str(cod)] = " | ".join(f"{r['PEC']} — {r['Data'].strftime('%d/%m/%Y')} — {r['Atividade']} — {r['Objetivo']} — {r['Horário']}" for _,r in g.iterrows())
    # Frequência oficial por código SEDUC: CSV exportado da planilha MIRACATU.
    # Atualizar o CSV no GitHub substitui os percentuais sem alterar o código.
    caminho_frequencia = pathlib.Path(__file__).resolve().parent / "GPS_frequencia_escolas_2026.csv"
    cadastro["frequencia_pct"] = float("nan")
    if caminho_frequencia.exists():
        try:
            freq = pd.read_csv(caminho_frequencia, dtype={"codigo_seduc": str})
            if {"codigo_seduc", "frequencia_pct"}.issubset(freq.columns):
                freq = freq.drop_duplicates("codigo_seduc", keep="last")
                valores = pd.to_numeric(freq.set_index("codigo_seduc")["frequencia_pct"], errors="coerce")
                cadastro["frequencia_pct"] = cadastro["codigo_seduc"].astype(str).map(valores)
            else:
                st.warning("CSV de frequência sem as colunas obrigatórias: codigo_seduc, frequencia_pct.")
        except Exception as exc:
            st.warning(f"Não foi possível carregar a frequência oficial: {exc}")
    cadastro["agenda_hoje"] = cadastro["codigo_seduc"].astype(str).map(lambda x: resumo.get(x,"Nenhum registro correspondente"))
    contagens = visitas_periodo.groupby("codigo_seduc").size().to_dict() if not visitas_periodo.empty else {}
    cadastro["intensidade"] = cadastro["codigo_seduc"].map(lambda c: int(contagens.get(c,0)))
    c1,c2,c3 = st.columns(3)
    c1.metric("Unidades georreferenciadas",len(cadastro))
    c2.metric("Escolas no período",sum(v>0 for v in contagens.values()))
    c3.metric("Visitas com termo" if modo_tempo=="Passado" else "Visitas programadas",len(visitas_periodo))
    if modo_tempo != "Passado" and agenda.empty:
        st.error("Nenhuma visita programada foi carregada. Confira o arquivo CSV da agenda e a importação dos Planos de Trabalho. Zero neste painel significa ausência de dados válidos, não ausência de visitas.")
    elif modo_tempo != "Passado" and visitas_periodo.empty:
        st.info("Há agenda carregada, mas nenhuma visita corresponde à data e aos filtros selecionados. Confira o dia de referência, o PEC e a abrangência dos dados importados.")
    f1,f2,f3 = st.columns([1,2,1])
    municipios = sorted(cadastro["municipio"].unique())
    municipio = f1.selectbox("Município",["Todos"]+municipios,key="gps_municipio_mapa")
    busca = f2.text_input("Buscar escola ou código SEDUC",key="gps_busca_mapa")
    modo = f3.selectbox("Visualização",["Mapa de calor", "Territorial", "Visitas no período"],key="gps_modo_mapa")
    filtrado = cadastro.copy()
    if municipio != "Todos":
        filtrado = filtrado[filtrado["municipio"] == municipio]
    if busca.strip():
        filtrado = filtrado[filtrado["escola"].str.contains(busca,case=False,regex=False) | filtrado["codigo_seduc"].str.contains(busca,regex=False)]
    if filtrado.empty:
        st.info("Nenhuma unidade encontrada.")
        st.stop()
    if modo == "Mapa de calor":
        # Mapa cumulativo dos termos já recebidos; visitas pendentes não são contadas.
        contagens_mes = base_graficos.groupby("codigo_seduc").size().to_dict() if not base_graficos.empty else {}
        filtrado["intensidade"] = filtrado["codigo_seduc"].map(lambda c: int(contagens_mes.get(c, 0)))
    if modo == "Visitas no período":
        filtrado["cor"] = filtrado["intensidade"].apply(lambda n: [22,163,74,245] if n else [150,150,150,180])
    if modo == "Mapa de calor":
        max_cont = max(1, int(filtrado["intensidade"].max()))
        filtrado["cor"] = filtrado["intensidade"].apply(lambda n: [190,40,45,240] if n>=max_cont and n>0 else ([240,130,35,225] if n>0 else [175,180,185,120]))
    # Marcadores tipo pino, conforme referência visual fornecida: cada escola
    # recebe a cor do seu município, sem alteração das coordenadas cadastrais.
    # SVG incorporado como data URI: não depende de arquivo extra no GitHub.
    def gps_icone_escola(cor_rgb):
        r, g, b = [int(v) for v in cor_rgb[:3]]
        preenchimento = f"#{r:02x}{g:02x}{b:02x}"
        svg = ("<svg xmlns='http://www.w3.org/2000/svg' width='64' height='84' viewBox='0 0 64 84'>"
               "<path d='M32 3C15 3 3 16 3 33c0 22 29 47 29 47s29-25 29-47C61 16 49 3 32 3Z' "
               f"fill='{preenchimento}' stroke='#14232d' stroke-width='3.5'/>"
               "<circle cx='32' cy='32' r='13' fill='#ffffff' stroke='#14232d' stroke-width='1.5'/>"
               "</svg>")
        return {"url": "data:image/svg+xml;charset=utf-8," + gps_url_quote(svg, safe=''),
                "width": 64, "height": 84, "anchorX": 32, "anchorY": 80}
    icones_municipios = {nome: gps_icone_escola(rgb) for nome, rgb in cores.items()}
    icone_padrao = gps_icone_escola([85, 85, 85])
    filtrado["icone"] = filtrado["municipio_norm"].map(
        lambda nome: icones_municipios.get(nome, icone_padrao))
    filtrado["quantidade_visitas"] = filtrado["intensidade"]
    geojson, erro_ibge = gps_limites_ibge()
    camadas = []
    if geojson:
        feats = []
        for f in geojson["features"]:
            m = f["properties"]["municipio"]
            f["properties"]["fill"] = cores.get(m,[100,100,100]) + [55]
            f["properties"]["line"] = cores.get(m,[100,100,100]) + [220]
            feats.append(f)
        camadas.append(pdk.Layer("GeoJsonLayer",data={"type":"FeatureCollection","features":feats},
                                 filled=True,stroked=True,get_fill_color="properties.fill",get_line_color="properties.line",
                                 line_width_min_pixels=2,pickable=False))
        # Os nomes ficam acima dos polígonos, com fonte grande e fundo branco.
        # O rótulo de Pedro de Toledo é criado a partir do próprio polígono do IBGE.
        rotulos_municipios = []
        for feicao in feats:
            try:
                lon, lat = gps_centroide_rotulo_municipio(feicao["geometry"])
                nome = feicao["properties"]["municipio"].replace("JUQUIA", "JUQUIÁ")
                rotulos_municipios.append({"longitude": lon, "latitude": lat, "nome": nome})
            except (ValueError, KeyError, TypeError, ZeroDivisionError):
                continue
        # O TextLayer será inserido depois dos pontos das escolas para garantir legibilidade.
    else:
        st.caption("Limites oficiais do IBGE temporariamente indisponíveis. Os pontos cadastrados continuam visíveis.")
        with st.expander("Detalhes da consulta ao IBGE"):
            st.code(str(erro_ibge))
    # A intensidade continua visível como halo atrás do pino, mas nunca substitui
    # a cor municipal do georreferenciamento da escola.
    if modo == "Mapa de calor":
        camadas.append(pdk.Layer("ScatterplotLayer", data=filtrado,
                                 get_position="[longitude,latitude]", get_fill_color="cor",
                                 get_radius="600 + intensidade * 180", radius_min_pixels=9,
                                 radius_max_pixels=29, stroked=False, pickable=False,
                                 id="intensidade_escolas"))
    if geojson and rotulos_municipios:
        # Ícones SVG transparentes garantem rótulos visíveis no deck.gl/Streamlit.
        # Mantidos abaixo dos marcadores, sem interceptar cliques.
        def _gps_icone_nome_municipio(nome):
            from html import escape as _escape_xml
            texto = _escape_xml(str(nome).upper())
            svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="360" height="48" viewBox="0 0 360 48">'
                   '<text x="180" y="31" text-anchor="middle" '
                   'font-family="Arial,Helvetica,sans-serif" font-size="24" '
                   'font-weight="900" fill="#000000">' + texto + '</text></svg>')
            return {"url": "data:image/svg+xml;charset=utf-8," + gps_url_quote(svg, safe=''),
                    "width": 360, "height": 48, "anchorX": 180, "anchorY": 24}
        for rotulo in rotulos_municipios:
            rotulo["icone_nome"] = _gps_icone_nome_municipio(rotulo["nome"])
        camadas.append(pdk.Layer(
            "IconLayer", data=rotulos_municipios,
            get_position="[longitude, latitude]", get_icon="icone_nome",
            get_size=240, size_scale=1, size_min_pixels=150,
            size_max_pixels=270, billboard=True, pickable=False,
            id="nomes_municipios"
        ))
    camadas.append(pdk.Layer("IconLayer", data=filtrado,
                             get_position="[longitude,latitude]", get_icon="icone",
                             get_size=18.5, size_scale=1, size_min_pixels=14,
                             size_max_pixels=24, billboard=True, pickable=True,
                             id="escolas"))
    # Sede: marcador distinto somente após coordenadas cadastradas em Secrets.
    sede_lat = st.secrets.get("GPS_SEDE_URE_LATITUDE", -24.285437)
    sede_lon = st.secrets.get("GPS_SEDE_URE_LONGITUDE", -47.458814)
    if sede_lat is not None and sede_lon is not None:
        try:
            sede = pd.DataFrame([{"latitude":float(sede_lat),"longitude":float(sede_lon),"escola":"SEDE URE MIRACATU", "municipio":"Sede regional", "codigo_seduc":"URE", "agenda_hoje":"Sede administrativa", "conteudo_popup": ("<b>UNIDADE REGIONAL DE ENSINO DE MIRACATU</b><br/>Município: Miracatu/SP<br/>Telefone: (13) 3847-7050<br/>E-mail: <a href=\"mailto:mir@educacao.sp.gov.br\" style=\"color:#ffffff;text-decoration:underline\">mir@educacao.sp.gov.br</a>")}])
            # A sede mantém o tamanho original dos pinos das escolas (37 px),
            # mas usa dourado exclusivo, não utilizado por nenhum município.
            sede["icone"] = [gps_icone_escola([245, 184, 39])]
            camadas.append(pdk.Layer("IconLayer", data=sede,
                                     get_position="[longitude,latitude]", get_icon="icone",
                                     get_size=37, size_scale=1, size_min_pixels=28,
                                     size_max_pixels=48, billboard=True, pickable=True,
                                     id="sede_ure"))
        except (TypeError,ValueError):
            st.warning("Coordenadas da sede da URE inválidas em Secrets.")
    # Conteúdo único por ponto: pydeck não substitui variáveis aninhadas em {raio_x}.
    # Mantém a cor e as dimensões originais do card.
    from html import escape as _gps_escape
    sede_coord = (-24.285437, -47.458814)
    def _gps_pct(valor):
        try:
            if pd.isna(valor):
                return "Sem dados"
            return f"{float(valor):.1f}%".replace(".", ",")
        except (TypeError, ValueError):
            return "Sem dados"
    def _gps_card(titulo, valor, fundo, cor):
        return (f"<div style='background:{fundo};color:{cor};padding:7px;border-radius:6px'>"
                f"<b>{titulo}</b><br/>{_gps_escape(str(valor))}</div>")
    def _gps_popup_escola(r):
        raio_x = ("<div style='display:grid;grid-template-columns:1fr 1fr;gap:5px;margin:8px 0'>"
            + _gps_card("Frequência geral", _gps_pct(r.get("frequencia_pct")), "#dceffb", "#174c73")
            + _gps_card("Prova Paulista", "Sem dados", "#ede5fa", "#60378f")
            + _gps_card("LP abaixo do básico", "Sem dados", "#fff0df", "#88480f")
            + _gps_card("MAT abaixo do básico", "Sem dados", "#dff3e9", "#126443") + "</div>")
        distancia = gps_distancia_direta(sede_coord, (float(r["latitude"]), float(r["longitude"])))
        agenda_html = _gps_escape(str(r["agenda_hoje"])).replace(" | ", "<br/>")
        return (f"<b>{_gps_escape(str(r['escola']))}</b><br/>"
                f"Município: {_gps_escape(str(r['municipio']))} · SEDUC: {_gps_escape(str(r['codigo_seduc']))}<br/>"
                f"<b>Visitas na data: {int(r['quantidade_visitas'])}</b><br/>"
                f"{agenda_html}{raio_x}<hr style='opacity:.25'/>"
                f"<small>Distância da URE: {distancia:.1f} km em linha reta (referência)</small>")
    filtrado["conteudo_popup"] = filtrado.apply(_gps_popup_escola, axis=1)
    # A sede tem o mesmo estilo de tooltip, mas conteúdo administrativo próprio.
    # Não reutiliza placeholders destinados às escolas.
    if sede_lat is not None and sede_lon is not None and 'sede' in locals():
        sede["conteudo_popup"] = ("<b>UNIDADE REGIONAL DE ENSINO DE MIRACATU</b><br/>"
            "Município: Miracatu/SP<br/>Telefone: (13) 3847-7050<br/>"
            "E-mail: <a href='mailto:mir@educacao.sp.gov.br' style='color:#ffffff;text-decoration:underline'>"
            "mir@educacao.sp.gov.br</a>")
    # IMPORTANTE: pydeck serializa os dados da camada ao criá-la.
    # Os cartões foram preenchidos depois da criação inicial dos IconLayers;
    # recriá-los agora evita que o tooltip exiba literalmente {conteudo_popup}.
    camadas = [camada for camada in camadas if getattr(camada, "id", None) not in ("escolas", "sede_ure")]
    camadas.append(pdk.Layer("IconLayer", data=filtrado,
                             get_position="[longitude,latitude]", get_icon="icone",
                             get_size=18.5, size_scale=1, size_min_pixels=14,
                             size_max_pixels=24, billboard=True, pickable=True,
                             id="escolas"))
    if 'sede' in locals() and not sede.empty:
        camadas.append(pdk.Layer("IconLayer", data=sede,
                                 get_position="[longitude,latitude]", get_icon="icone",
                                 get_size=37, size_scale=1, size_min_pixels=28,
                                 size_max_pixels=48, billboard=True, pickable=True,
                                 id="sede_ure"))
    tooltip = {"html":"{conteudo_popup}",
               "style":{"backgroundColor":"#17313d","color":"white","fontSize":"13px","maxWidth":"430px"}}
    vista = pdk.ViewState(latitude=float(filtrado.latitude.mean()),longitude=float(filtrado.longitude.mean()),zoom=9 if municipio=="Todos" else 11)
    evento = st.pydeck_chart(pdk.Deck(layers=camadas,initial_view_state=vista,tooltip=tooltip,map_style=None),
                             use_container_width=True,on_select="rerun",selection_mode="single-object",key="gps_mapa_selecao")
    selecionado = None
    try:
        selecoes = evento.selection.objects.get("escolas",[])
        if selecoes:
            selecionado = str(selecoes[0].get("codigo_seduc",""))
    except (AttributeError,TypeError,KeyError,IndexError):
        pass
    # Alternativa acessível ao clique: seleção explícita por escola, confiável
    # em navegadores que não propagam eventos de IconLayer ao Streamlit.
    opcoes_escolas = {str(r["codigo_seduc"]): f'{r["escola"]} — {r["municipio"]} (SEDUC {r["codigo_seduc"]})'
                      for _,r in filtrado.sort_values("escola").iterrows()}
    codigos_opcoes = [""] + list(opcoes_escolas)
    codigo_escolhido = st.selectbox(
        "🏫 Abrir ficha da escola (alternativa ao clique no pino)",
        options=codigos_opcoes,
        format_func=lambda c: opcoes_escolas.get(c, "Selecione uma escola"),
        key="gps_mapa_escola_ficha")
    if codigo_escolhido:
        selecionado = codigo_escolhido
    st.subheader(f"📅 Visitas e atividades — {data_ref.strftime('%d/%m/%Y')}")
    if visitas_periodo.empty:
        st.info("Nenhuma visita encontrada para esta data nos documentos atualmente carregados. Isso não comprova ausência de atividade.")
    else:
        linhas_visitas = visitas_periodo.copy()
        if municipio != "Todos":
            codigos_municipio = set(filtrado["codigo_seduc"].astype(str))
            linhas_visitas = linhas_visitas[linhas_visitas["codigo_seduc"].astype(str).isin(codigos_municipio)]
        if busca.strip():
            codigos_busca = set(filtrado["codigo_seduc"].astype(str))
            linhas_visitas = linhas_visitas[linhas_visitas["codigo_seduc"].astype(str).isin(codigos_busca)]
        nomes_escolas = cadastro.set_index("codigo_seduc")["escola"].to_dict()
        linhas_visitas.insert(0, "Escola", linhas_visitas["codigo_seduc"].astype(str).map(nomes_escolas))
        campos_visitas = [c for c in ["Escola","PEC","Data","Horário","Atividade","Objetivo","Situação","Fonte"] if c in linhas_visitas.columns]
        st.dataframe(linhas_visitas[campos_visitas], use_container_width=True, hide_index=True)
        st.caption("Fonte: programação disponível e termos identificados. Dados incompletos não devem ser interpretados como ausência de visita.")
    if selecionado and selecionado in set(cadastro["codigo_seduc"].astype(str)):
        st.session_state["gps_escola_selecionada"] = selecionado
    codigo_ficha = st.session_state.get("gps_escola_selecionada")
    if codigo_ficha and codigo_ficha in set(cadastro["codigo_seduc"].astype(str)):
        escola_ficha = cadastro[cadastro["codigo_seduc"].astype(str)==codigo_ficha].iloc[0]
        st.divider()
        gps_ficha_escola(escola_ficha, agenda, df_termos_drive, data_ref)
    with st.expander("📋 Conferência histórica: Plano × Termo × Justificativa", expanded=False):
        if agenda.empty:
            st.info("Nenhuma programação datada disponível para conferência.")
        else:
            agenda_aud = agenda.copy()
            agenda_aud["Escola"] = agenda_aud["codigo_seduc"].map(cadastro.set_index("codigo_seduc")["escola"].to_dict())
            conf = gps_conciliar_documentos(agenda_aud, df_termos_drive, justificativas_mapa)
            if pec_mapa != "Todos" and not conf.empty:
                conf = conf[conf["PEC"].map(gps_chave_pec)==gps_chave_pec(pec_mapa)]
            st.dataframe(conf, use_container_width=True, hide_index=True)
            st.caption("Termos e justificativas podem corrigir o status do passado sem apagar o plano original. Justificativas automáticas requerem GPS_justificativas_validadas.csv ou integração documental validada.")
    if destino_gps == "Cadastro das Unidades":
        st.dataframe(filtrado[["codigo_seduc","escola","municipio","tipo","endereco"]],hide_index=True,use_container_width=True)
    if selecionado and selecionado in set(cadastro["codigo_seduc"]):
        st.session_state["gps_escola_selecionada"] = selecionado
    if destino_gps != "Cadastro das Unidades":
        st.caption("Passe o mouse para ver os registros do período; clique no marcador para abrir a ficha. Mapa passado = termos documentados; presente/futuro = programação, sem rastreamento físico.")
    codigo_atual = st.session_state.get("gps_escola_selecionada")
    # Ficha da data selecionada já exibida imediatamente abaixo do mapa.
    if destino_gps in ("Rotas e Distâncias", "Cobertura Territorial"):
        st.subheader("🚗 Quilometragem planejada dos PECs")
        st.caption("Percurso estimado: URE → escolas programadas no dia → URE. Não representa rastreamento real nem comprovação da visita. A ordem das escolas segue os horários informados, quando disponíveis.")
        agenda_dist = agenda.copy()
        if pec_mapa != "Todos":
            agenda_dist = agenda_dist[agenda_dist["PEC"] == pec_mapa]
        if not agenda_dist.empty:
            datas = pd.to_datetime(agenda_dist["Data"])
            ano_dist = st.selectbox("Ano", sorted(datas.dt.year.unique(), reverse=True), key="gps_km_ano")
            agenda_dist = agenda_dist[agenda_dist["Data"].map(lambda d:d.year == ano_dist)]
            periodo = st.selectbox("Período", ["Dia", "Semana", "Mês", "Ano"], key="gps_km_periodo")
            if periodo == "Dia":
                agenda_dist = agenda_dist[agenda_dist["Data"] == data_ref]
            elif periodo == "Semana":
                inicio = data_ref - __import__("datetime").timedelta(days=data_ref.weekday())
                fim = inicio + __import__("datetime").timedelta(days=6)
                agenda_dist = agenda_dist[(agenda_dist["Data"] >= inicio)&(agenda_dist["Data"] <= fim)]
            elif periodo == "Mês":
                agenda_dist = agenda_dist[agenda_dist["Data"].map(lambda d:d.month == data_ref.month and d.year == data_ref.year)]
            if st.button("Calcular distâncias rodoviárias", key="gps_km_calcular"):
                with st.spinner("Consultando as rotas rodoviárias..."):
                    km_df = gps_resumo_quilometragem(agenda_dist, cadastro)
                st.session_state["gps_km_resultado"] = km_df
            km_df = st.session_state.get("gps_km_resultado")
            if isinstance(km_df, pd.DataFrame) and not km_df.empty:
                st.dataframe(km_df, hide_index=True, use_container_width=True)
                resumo_km = km_df.groupby("PEC",as_index=False).agg({"Km rodoviários":"sum", "Visitas previstas":"sum", "Data":"nunique"}).rename(columns={"Data":"Dias programados"})
                st.subheader("Total por PEC")
                st.dataframe(resumo_km, hide_index=True, use_container_width=True)
                st.download_button("Baixar quilometragem CSV", km_df.to_csv(index=False).encode("utf-8-sig"),"GPS_quilometragem_pecs.csv", "text/csv")
                if km_df["Km rodoviários"].isna().any():
                    st.warning("Algumas rotas não foram calculadas. Os totais são parciais e não incluem percursos indisponíveis.")
        else:
            st.info("Não há visitas datadas no período para calcular trajetos.")
    st.stop()

if pagina_gps == "📁 Fontes de Dados":
    st.title("📁 Fontes de Dados — GPS URE Miracatu")
    st.caption("Inventário de arquivos nas pastas do Drive. A listagem não importa nem consolida resultados automaticamente.")
    fonte = st.selectbox("Fonte", list(PASTAS_FONTES_GPS), key="gps_fonte_drive")
    pasta_id = PASTAS_FONTES_GPS[fonte]
    st.link_button("Abrir pasta no Google Drive", f"https://drive.google.com/drive/folders/{pasta_id}")
    with st.spinner("Verificando acesso e arquivos..."):
        itens, erro = inventariar_fonte_drive(pasta_id, GOOGLE_DRIVE_API_KEY)
    if erro:
        st.error("Acesso ou leitura não confirmado: " + erro)
        st.info("O acesso a outras pastas do Drive não garante acesso a esta. Não torne os arquivos públicos; uma autorização privada poderá ser necessária.")
    if itens:
        inventario = pd.DataFrame(itens)
        st.metric("Arquivos encontrados", len(inventario))
        st.dataframe(inventario, use_container_width=True, hide_index=True)
        st.download_button("Baixar inventário CSV", inventario.to_csv(index=False).encode("utf-8-sig"),
                           file_name="gps_inventario_drive.csv", mime="text/csv")
    elif not erro:
        st.info("Nenhum arquivo encontrado nesta pasta ou nas subpastas acessíveis.")
    st.info("Este inventário é diagnóstico administrativo, não representa resultados pedagógicos consolidados.")
    st.stop()

st.title("🧭 GPS — Guia Pedagógico Semanal")
st.caption("Planejamento Mensal Integrado, Indicadores Hierárquicos e Acompanhamento — URE de Miracatu")

lista_abas = [
    "📋 Plano de Trabalho & Planner Mensal",
    "📝 Avaliações",
    "💻 Plataformas Digitais",
    "🏫 Frequência Escolar",
    "📅 Agenda SEDUC, CEC & Feriados",
    "📚 Diretrizes das Pastas (Res. SEDUC 01/2026)"
]
if eh_cec_ou_at:
    lista_abas.insert(4, "🛡️ Painel Exclusivo CEC (Confronto & Fechamento)")

# A navegação lateral substitui as abas horizontais.
# A escolha de cada subitem mostra apenas o módulo correspondente.

# --- ABA 1: PLANO DE TRABALHO & PLANNER MENSAL LADO A LADO ---
if destino_gps in ("Plano de Trabalho PEC", "Planner Mensal"):
    st.subheader("Elaboração do Plano de Trabalho (Modelo Oficial URE Miracatu)")
    st.caption("Submeta o arquivo gerado na pasta do Google Drive até o 2º dia útil do mês. Após cada visita, deposite o Termo de Visita em PDF na sua pasta do Drive.")

    # 1. Seletor do Mês e Ano do Planejamento (ATPCs e Prova Paulista mais atual são identificadas automaticamente)
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

    # Monta dinamicamente as opções de indicadores já detectando a Prova Paulista mais atual do período
    opcoes_criterio_dict = construir_opcoes_criterio_prioridade(df_ind_usuario, ano_ref=ano_plano, mes_ref=mes_plano)

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

        # Base filtrada pelo indicador para listar dinamicamente os componentes disponíveis
        df_base_indicador = filtrar_base_por_criterio_prioridade(
            df_ind_usuario,
            criterio_prio_sel,
            ano_ref=ano_plano,
            mes_ref=mes_plano
        )

        componentes_disp = ["Todos os Componentes"]
        if not df_base_indicador.empty and "Componente" in df_base_indicador.columns:
            comps_unicos = sorted([
                str(c).strip()
                for c in df_base_indicador["Componente"].dropna().unique().tolist()
                if str(c).strip() and str(c).strip().lower() not in ["nan", "none", "frequência", "frequencia"]
            ])
            componentes_disp.extend(comps_unicos)

        comp_filtro_sel = st.selectbox(
            "📚 Filtrar por Componente Curricular (recalcula prioridade):",
            options=componentes_disp,
            key="filtro_componente_prioridade_live",
            help="Filtre por componente específico (ex.: Matemática, Língua Portuguesa) para recalcular o ranking das escolas e liberar o filtro por turma."
        )

        # Calcula o ranking pós-filtro conforme Indicador + Componente escolhidos
        df_ranking_prio = obter_ranking_escolas_prioritarias(
            df_ind_usuario,
            pasta_atual,
            escolas_usuario,
            criterio_chave=criterio_prio_sel,
            df_recomp_sheets=df_recomposicao_sheets,
            ano_ref=ano_plano,
            mes_ref=mes_plano,
            nome_crit_label=opcoes_criterio_dict.get(criterio_prio_sel, ""),
            componente_filtro=comp_filtro_sel
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
            lista_esc_fb = escolas_usuario if (eh_pec_qualidade or escolas_usuario) else ESCOLAS_URE_MIRACATU
            for esc_nome in lista_esc_fb:
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

        # Exibe o diagnóstico pós-filtro da escola selecionada (Indicador + Componente)
        if not df_ranking_prio.empty and escola_sel_form in df_ranking_prio["Escola"].values:
            linha_diag = df_ranking_prio[df_ranking_prio["Escola"] == escola_sel_form].iloc[0]
            rot_comp_diag = f" | {comp_filtro_sel}" if comp_filtro_sel != "Todos os Componentes" else ""
            st.caption(
                f"📊 **Diagnóstico Pós-Filtro ({opcoes_criterio_dict[criterio_prio_sel]}{rot_comp_diag}):** "
                f"{linha_diag.get('Prioridade', '')} lugar — {linha_diag.get('Detalhe_Indicador', '')}"
            )

        # Filtro em cascata por Turma da Escola selecionada (ranqueada por criticidade)
        turma_sel_form = "Todas as Turmas"
        df_base_comp_turma = df_base_indicador.copy()
        if comp_filtro_sel != "Todos os Componentes" and not df_base_comp_turma.empty:
            comp_f_norm = normalizar_texto(comp_filtro_sel)
            df_base_comp_turma = df_base_comp_turma[
                df_base_comp_turma["Componente"].apply(
                    lambda x: normalizar_texto(x) == comp_f_norm or comp_f_norm in normalizar_texto(x)
                )
            ]

        if not df_base_comp_turma.empty and escola_sel_form not in ["URE Miracatu", "Reunião Virtual / EFAPE / SEDUC"]:
            esc_norm_sel = normalizar_texto(escola_sel_form)
            df_esc_turmas = df_base_comp_turma[
                (df_base_comp_turma["Escola"].apply(lambda x: esc_norm_sel in normalizar_texto(x) or normalizar_texto(x) in esc_norm_sel))
                & (df_base_comp_turma["Nivel_Granularidade"].isin(["TURMA", "ESTUDANTE"]))
                & (~df_base_comp_turma["Turma"].astype(str).str.contains("Consolidado", case=False, na=False))
            ].copy()

            if not df_esc_turmas.empty:
                df_nivel_tur = df_esc_turmas[df_esc_turmas["Nivel_Granularidade"] == "TURMA"]
                base_agg_tur = df_nivel_tur if not df_nivel_tur.empty else df_esc_turmas

                df_rank_turmas = (
                    base_agg_tur.groupby("Turma", as_index=False)
                    .agg(Media_Pct=("Proficiencia_Pct", "mean"))
                    .sort_values(by=["Media_Pct", "Turma"], ascending=[True, True])
                    .reset_index(drop=True)
                )
                df_rank_turmas["Media_Pct"] = df_rank_turmas["Media_Pct"].round(1)

                mapa_rotulo_turma = {"Todas as Turmas (Atuação Geral na Escola)": "Todas as Turmas"}
                opcoes_turmas_rotulos = ["Todas as Turmas (Atuação Geral na Escola)"]

                for idx_t, r_t in df_rank_turmas.iterrows():
                    t_nome = str(r_t["Turma"]).strip()
                    t_pct = r_t["Media_Pct"]
                    t_farol = classificar_farol_nivel(t_pct)
                    t_ico = "🔴" if "Vermelho" in t_farol else ("🟡" if "Amarelo" in t_farol else ("🟢" if "Verde" in t_farol else "🔵"))
                    rot_t = f"{idx_t + 1}º Prioritária | {t_nome} ({t_pct:.1f}% {t_ico})"
                    opcoes_turmas_rotulos.append(rot_t)
                    mapa_rotulo_turma[rot_t] = t_nome

                rot_turma_escolhida = st.selectbox(
                    "🎓 Filtrar por Turma da Escola (Pós-Filtro):",
                    options=opcoes_turmas_rotulos,
                    key=f"turma_lancamento_live_{escola_sel_form}_{comp_filtro_sel}"
                )
                turma_sel_form = mapa_rotulo_turma[rot_turma_escolhida]

                if turma_sel_form != "Todas as Turmas":
                    df_det_turma = df_esc_turmas[df_esc_turmas["Turma"] == turma_sel_form]
                    media_t = round(df_det_turma["Proficiencia_Pct"].mean(), 1)
                    farol_t = classificar_farol_nivel(media_t)
                    profs_t = [
                        p for p in df_det_turma["Professor"].dropna().unique().tolist()
                        if str(p).strip() and str(p).strip().lower() not in ["nan", "none", "-"]
                    ] if "Professor" in df_det_turma.columns else []
                    alunos_crit = df_det_turma[
                        (df_det_turma["Nivel_Granularidade"] == "ESTUDANTE") & (df_det_turma["Proficiencia_Pct"] <= 40.0)
                    ]["Estudante"].nunique()

                    detalhes_t_txt = [f"📊 **Foco na Turma {turma_sel_form}:** {media_t:.1f}% ({farol_t})"]
                    if profs_t:
                        detalhes_t_txt.append(f"👩‍🏫 **Docente:** {', '.join(profs_t)}")
                    if alunos_crit > 0:
                        detalhes_t_txt.append(f"🚨 **{alunos_crit} estudante(s)** em nível crítico (≤ 40%)")
                    st.caption(" | ".join(detalhes_t_txt))
            elif escola_sel_form in ESCOLAS_URE_MIRACATU:
                st.caption("ℹ️ Nenhuma planilha detalhada por turma localizada no Drive para esta escola/componente.")

        # Identificação automática dos dias e horários de ATPC (priorizando o componente filtrado, se houver)
        area_busca_atpc = comp_filtro_sel if comp_filtro_sel != "Todos os Componentes" else comp_atual
        atpcs_da_escola = obter_atpcs_escola(df_atpc_global, escola_sel_form, area_busca_atpc)
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

                    partes_foco = []
                    if comp_filtro_sel != "Todos os Componentes":
                        partes_foco.append(f"Componente: {comp_filtro_sel}")
                    if turma_sel_form != "Todas as Turmas":
                        partes_foco.append(f"Turma: {turma_sel_form}")
                    if info_adicional.strip():
                        partes_foco.append(info_adicional.strip())

                    detalhe_plano = " | ".join(partes_foco) if partes_foco else "Conforme diretrizes da pasta"
                    st.session_state["cronograma_semana"].append({
                        "data_iso": data_visita.isoformat(),
                        "dia": f"{data_visita.strftime('%d/%m/%Y')} ({dia_sem_nome}) — {horario_visita}",
                        "horario": horario_visita,
                        "local": escola_sel_form,
                        "acao": acao_combinada,
                        "publico": detalhe_plano,
                        "componente": comp_filtro_sel,
                        "turma": turma_sel_form,
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
        titulo_exp_prio = f"{opcoes_criterio_dict[criterio_prio_sel]} | {comp_filtro_sel}" if comp_filtro_sel != "Todos os Componentes" else opcoes_criterio_dict[criterio_prio_sel]
        with st.expander(f"📊 Ver Ranking Completo Pós-Filtro ({titulo_exp_prio})", expanded=False):
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
        df_prio_word = df_ind_usuario[df_ind_usuario["Proficiencia_Pct"] <= 60.0].copy()
        if comp_filtro_sel != "Todos os Componentes":
            df_sub_c = df_prio_word[df_prio_word["Componente"].apply(lambda x: normalizar_texto(comp_filtro_sel) in normalizar_texto(x))]
            if not df_sub_c.empty:
                df_prio_word = df_sub_c
        if turma_sel_form != "Todas as Turmas":
            df_sub_t = df_prio_word[df_prio_word["Turma"] == turma_sel_form]
            if not df_sub_t.empty:
                df_prio_word = df_sub_t

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

def gps_filtrar_tipo_escola(df_indicadores, chave_filtro):
    """Aplica filtro cadastral às escolas dos indicadores, sem alterar os dados originais.

    A classificação vem exclusivamente do cadastro de unidades (tipo == EE).
    Registros sem correspondência segura permanecem em Todas, nunca em Regulares.
    """
    opcao = st.radio(
        "Tipo de escola:", ["Todas", "Regulares"], horizontal=True,
        key=chave_filtro,
    )
    if opcao == "Todas" or df_indicadores is None or df_indicadores.empty:
        return df_indicadores
    if "Escola" not in df_indicadores.columns:
        st.warning("Os indicadores não possuem a coluna Escola para aplicar o filtro.")
        return df_indicadores.iloc[0:0].copy()

    import pathlib
    caminho = pathlib.Path(__file__).resolve().parent / "GPS_cadastro_40_unidades.csv"
    if not caminho.exists():
        st.warning("Cadastro de unidades não encontrado: não é possível identificar as escolas regulares.")
        return df_indicadores.iloc[0:0].copy()
    try:
        cadastro_tipo = pd.read_csv(caminho, dtype=str, encoding="utf-8-sig").fillna("")
        if not {"escola", "tipo"}.issubset(cadastro_tipo.columns):
            raise ValueError("colunas escola e tipo ausentes")
        nomes_regulares = {
            normalizar_texto(nome) for nome in
            cadastro_tipo.loc[cadastro_tipo["tipo"].map(normalizar_texto) == "EE", "escola"]
        }
        # Aceita apenas equivalências exatas, com prefixos institucionais opcionais.
        # Não utiliza correspondência parcial, que poderia incluir CEL, CEEJA ou EEI.
        def eh_regular(nome):
            chave = normalizar_texto(nome)
            if chave in nomes_regulares:
                return True
            for prefixo in ("EE ", "E.E. ", "ESCOLA ESTADUAL "):
                if chave.startswith(prefixo) and chave[len(prefixo):] in nomes_regulares:
                    return True
            return False
        return df_indicadores.loc[df_indicadores["Escola"].map(eh_regular)].copy()
    except Exception as exc:
        st.warning(f"Não foi possível aplicar o filtro de escolas regulares: {exc}")
        return df_indicadores.iloc[0:0].copy()


def exibir_banner_perfil_indicadores():
    if eh_cec_ou_at:
        st.caption(f"📂 {qtd_arqs_drive} arquivo(s) sincronizado(s) do Google Drive — Visão Global da URE Miracatu." + (" Frequência escolar também disponível na base local." if not freq_local_gps.empty else ""))
    elif eh_pec_qualidade:
        if escolas_usuario:
            st.info(f"🏫 **Carteira Qualidade da Aula (Google Sheets):** {', '.join(escolas_usuario)} — Exibindo exclusivamente suas escolas de acompanhamento com **todas as plataformas digitais** disponíveis.")
        else:
            st.warning("⚠️ Nenhuma escola de acompanhamento foi localizada para o seu e-mail na aba `Usuarios` do Google Sheets. Verifique o cadastro da sua carteira de escolas.")
    else:
        regras_info = obter_regras_atuacao_pec(pasta_atual, comp_atual, df_pastas_raw=df_pastas_raw)
        if regras_info:
            comps_txt = ", ".join(regras_info["componentes"]) if regras_info.get("componentes") else "Todos"
            plats_txt = ", ".join(regras_info["plataformas"]) if regras_info.get("plataformas") else "Nenhuma específica"
            st.info(f"🎯 **Filtro específico da sua pasta ({pasta_atual} — {comp_atual}):** Componentes `{comps_txt}` | Plataformas `{plats_txt}`.")

# --- ABA 2: AVALIAÇÕES (COM SUBPASTAS DINÂMICAS E PROVA PAULISTA MAIS ATUAL) ---
if destino_gps in ("SARESP", "Prova Paulista", "SAEB", "Avaliação Diagnóstica", "ENEM e Simulados"):
    st.subheader(f"📝 Avaliações — {pasta_atual} ({comp_atual})")
    exibir_banner_perfil_indicadores()

    df_av = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "AVALIACAO"] if not df_ind_usuario.empty else pd.DataFrame()

    df_av = gps_filtrar_tipo_escola(df_av, "gps_tipo_escola_avaliacoes")

    # Descobre dinamicamente todas as subpastas/avaliações presentes no Drive (garantindo Prova Paulista primeiro)
    subs_av_drive = sorted([s for s in df_av["Subcategoria"].dropna().unique().tolist() if s != "Outras Avaliações"]) if not df_av.empty else []
    opcoes_avaliacoes = ["Prova Paulista"] + [s for s in subs_av_drive if s != "Prova Paulista"]
    for padrao_av in ["SARESP", "SAEB", "ENEM", "Avaliação Diagnóstica"]:
        if padrao_av not in opcoes_avaliacoes:
            opcoes_avaliacoes.append(padrao_av)
    if len(opcoes_avaliacoes) > 1:
        opcoes_avaliacoes.append("Todas as Avaliações")

    c_av1, c_av2 = st.columns(2)
    with c_av1:
        tipo_prova = st.selectbox(
            "Selecione a Avaliação / Subpasta:",
            opcoes_avaliacoes,
            key="sel_subpasta_avaliacao"
        )

    df_av_filt = df_av.copy()
    with c_av2:
        if tipo_prova == "Prova Paulista":
            _, rot_pp_pesquisa = obter_edicao_mais_atual_prova_paulista(df_av_filt)
            # Sem 4º Bimestre; padrão sempre na edição mais atual disponível
            bims_pp_disp = [f"Edição Mais Atual ({rot_pp_pesquisa})"]
            if not df_av_filt.empty:
                df_pp_only = df_av_filt[df_av_filt["Subcategoria"] == "Prova Paulista"]
                for b_n in [3, 2, 1]:
                    if any(re.search(rf"\b{b_n}\s*[º°O]?\s*BIM", normalizar_texto(p)) for p in df_pp_only["Periodo"].dropna().unique()):
                        bims_pp_disp.append(f"{b_n}º Bimestre")
            periodo_filtro = st.selectbox("Edição / Bimestre (Prova Paulista):", bims_pp_disp, key="sel_bim_pp_pesquisa")
        elif tipo_prova != "Todas as Avaliações":
            anos_disp = sorted(df_av_filt[df_av_filt["Subcategoria"] == tipo_prova]["Periodo"].dropna().unique().tolist()) if not df_av_filt.empty else []
            periodo_filtro = st.selectbox(f"Ano / Edição ({tipo_prova}):", ["Todos os Anos / Edições"] + anos_disp, key="sel_per_outra_av")
        else:
            periodo_filtro = "Todas as Edições"
            st.caption("Consolidando todas as avaliações disponíveis nas subpastas.")

    if not df_av_filt.empty:
        if tipo_prova == "Prova Paulista":
            if periodo_filtro.startswith("Edição Mais Atual"):
                df_av_filt, rot_pp_usada = filtrar_prova_paulista_mais_atual(df_av_filt)
                st.caption(f"✨ Analisando automaticamente a edição mais atual da Prova Paulista: **{rot_pp_usada}**.")
            else:
                df_av_filt = df_av_filt[
                    (df_av_filt["Subcategoria"] == "Prova Paulista") &
                    (df_av_filt["Periodo"].str.contains(periodo_filtro, case=False, na=False))
                ]
        elif tipo_prova != "Todas as Avaliações":
            df_av_filt = df_av_filt[df_av_filt["Subcategoria"] == tipo_prova]
            if not periodo_filtro.startswith("Todos"):
                df_av_filt = df_av_filt[df_av_filt["Periodo"].str.contains(periodo_filtro, case=False, na=False)]

    renderizar_tabela_hierarquica(df_av_filt, f"{tipo_prova} — {periodo_filtro}", "Proficiência Média (%)")

# --- ABA 3: PLATAFORMAS DIGITAIS (ESCOLA → TURMA) ---
if destino_gps in ("Tarefa SP", "Leia SP", "Redação Paulista", "Matific Anos Finais", "Khan Academy", "Speak", "Open English", "Elefante Letrado", "Matific Anos Iniciais", "Alura", "Robótica", "Programação"):
    st.subheader(f"💻 {destino_gps} — {pasta_atual} ({comp_atual})")
    exibir_banner_perfil_indicadores()

    df_pl = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "PLATAFORMA"].copy() if not df_ind_usuario.empty else pd.DataFrame()
    df_pl = gps_filtrar_tipo_escola(df_pl, "gps_tipo_escola_plataformas")

    # O menu lateral já define a plataforma. Não repetir essa escolha no painel.
    # Mantém compatibilidade com nomes de subpastas usados nas planilhas.
    equivalencias_plataforma = {
        "Matific Anos Finais": ["Matific Anos Finais", "Matific"],
        "Matific Anos Iniciais": ["Matific Anos Iniciais", "Matific"],
    }
    nomes_plataforma = equivalencias_plataforma.get(destino_gps, [destino_gps])
    if not df_pl.empty:
        df_pl = df_pl[df_pl["Subcategoria"].fillna("").map(normalizar_texto).isin(
            {normalizar_texto(nome) for nome in nomes_plataforma}
        )].copy()

    # Escolas com dados da plataforma são a fonte preferencial. Se a plataforma
    # ainda não foi sincronizada, mostrar o cadastro, sem inventar resultados.
    escolas_disponiveis = sorted(df_pl["Escola"].dropna().astype(str).unique().tolist()) if not df_pl.empty else []
    if not escolas_disponiveis:
        from pathlib import Path
        cadastro_path = Path(__file__).resolve().parent / "GPS_cadastro_40_unidades.csv"
        if cadastro_path.exists():
            try:
                cadastro_pl = pd.read_csv(cadastro_path, dtype=str, encoding="utf-8-sig").fillna("")
                if st.session_state.get("gps_tipo_escola_plataformas", "Todas") == "Regulares":
                    cadastro_pl = cadastro_pl[cadastro_pl["tipo"].map(normalizar_texto) == "EE"]
                escolas_disponiveis = sorted(cadastro_pl["escola"].loc[lambda x: x.str.strip() != ""].unique().tolist())
                # Para os perfis com carteira escolar, não expor escolas fora dela.
                if not eh_cec_ou_at and escolas_usuario:
                    carteira = {normalizar_texto(e) for e in escolas_usuario}
                    escolas_disponiveis = [e for e in escolas_disponiveis if normalizar_texto(e) in carteira]
            except Exception as exc:
                st.warning(f"Não foi possível carregar o cadastro de escolas: {exc}")

    escola_sel = st.selectbox(
        "Selecione a escola:", ["Selecione uma escola..."] + escolas_disponiveis,
        key=f"gps_pl_escola_{destino_gps}"
    )
    if escola_sel != "Selecione uma escola...":
        df_escola_pl = df_pl[df_pl["Escola"] == escola_sel].copy() if not df_pl.empty else pd.DataFrame()
        turmas_pl = sorted({str(t).strip() for t in df_escola_pl.get("Turma", pd.Series(dtype=str)).dropna()
                            if str(t).strip() and normalizar_texto(t) not in ("CONSOLIDADO", "GERAL", "NAN")})
        turma_sel = st.selectbox(
            "Selecione a turma:", ["Todas as turmas"] + turmas_pl,
            key=f"gps_pl_turma_{destino_gps}_{escola_sel}"
        )
        if not df_escola_pl.empty:
            if turma_sel != "Todas as turmas":
                df_escola_pl = df_escola_pl[df_escola_pl["Turma"].astype(str) == turma_sel]
            if not df_escola_pl.empty:
                st.markdown(f"#### Resultados — {destino_gps} | {escola_sel}" +
                            (f" | {turma_sel}" if turma_sel != "Todas as turmas" else ""))
                colunas_visiveis = [c for c in ["Escola", "Turma", "Componente", "Periodo", "Proficiencia_Pct", "Nivel_Granularidade"]
                                    if c in df_escola_pl.columns]
                st.dataframe(df_escola_pl[colunas_visiveis], use_container_width=True, hide_index=True)
            else:
                st.info("Não há resultados disponíveis para esta turma na plataforma selecionada.")
        else:
            st.info(f"Ainda não há resultados sincronizados de {destino_gps} para {escola_sel}.")
        if not turmas_pl:
            st.caption("As turmas aparecerão automaticamente quando houver dados detalhados desta escola e plataforma.")

# --- ABA 4: FREQUÊNCIA ESCOLAR (COM SUBPASTAS / PERÍODOS DINÂMICOS) ---
if destino_gps in ("Geral URE", "Escolas", "Turmas"):
    st.subheader(f"🏫 Frequência Escolar — {pasta_atual}")
    exibir_banner_perfil_indicadores()

    df_fr = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "FREQUENCIA"] if not df_ind_usuario.empty else pd.DataFrame()
    df_fr = gps_filtrar_tipo_escola(df_fr, "gps_tipo_escola_frequencia")
    subs_fr_drive = sorted(df_fr["Periodo"].dropna().unique().tolist()) if not df_fr.empty else []
    per_fr_sel = st.selectbox("Selecione o Período / Subpasta de Frequência:", ["Todos os Períodos"] + subs_fr_drive, key="sel_subpasta_frequencia")

    df_fr_filt = df_fr.copy()
    if not df_fr_filt.empty and per_fr_sel != "Todos os Períodos":
        df_fr_filt = df_fr_filt[df_fr_filt["Periodo"] == per_fr_sel]

    if destino_gps == "Turmas" and not df_fr_filt.empty and not df_fr_filt["Nivel_Granularidade"].isin(["TURMA", "ESTUDANTE"]).any():
        st.info("A base local contém frequência por escola. Os resultados por turma serão exibidos quando as planilhas detalhadas forem integradas.")
    renderizar_tabela_hierarquica(df_fr_filt, f"Frequência Escolar ({per_fr_sel})", "Frequência Média (%)")

# ==============================================================================
# ALERTAS CEC — TERMO DE VISITA: 7 DIAS (AMARELO), FIM DA QUINZENA (VERMELHO)
# ==============================================================================
# V9 — Calendário institucional e conciliação documental
# Configurações institucionais definidas antes da renderização do mapa.

# --- ABA EXCLUSIVA DO CEC / AT: CONFRONTO & FECHAMENTO MENSAL ---
idx_proxima_aba = 4
if eh_cec_ou_at:
    if destino_gps in ("Painel Exclusivo CEC", "Acompanhamentos", "Relatórios"):
        st.subheader("🛡️ Painel Exclusivo CEC — Controle de Prazos (2º Dia Útil), Confronto & Justificativas")
        st.caption("Visão global de todos os PECs, confrontando Planos de Trabalho enviados até o 2º dia útil x Termos de Visita depositados no Google Drive.")

        st.markdown("#### 🚦 Conciliação: Plano de Trabalho × Termo × Justificativa")
        st.caption("Ausência de termo não comprova ausência da visita. Justificativas explicam não realização e mudanças do plano. O cruzamento exige PEC, escola e data correspondentes.")
        st.info("Períodos institucionais configurados: 16–29/11/2026 e 30/11–11/12/2026. Em 30/11 não há expediente para nenhum PEC (feriado da sede em Miracatu).")
        st.markdown("**Justificativas estruturadas** — importe um CSV para conferir motivos e remanejamentos. O conteúdo deve ser conferido com o documento original.")
        arq_just = st.file_uploader("Justificativas mensais (CSV)", type=["csv"], key="gps_just_v9")
        justificativas_cec = pd.DataFrame()
        if arq_just is not None:
            try:
                justificativas_cec = pd.read_csv(arq_just, dtype=str, keep_default_na=False, sep=None, engine="python")
                col_obrig = {"PEC", "Escola", "Data_Prevista", "Situacao", "Motivo"}
                if not col_obrig.issubset(justificativas_cec.columns):
                    st.error("CSV de justificativas: faltam " + ", ".join(sorted(col_obrig - set(justificativas_cec.columns))))
                    justificativas_cec = pd.DataFrame()
            except Exception as exc:
                st.error(f"Não foi possível ler justificativas: {exc}")
        from pathlib import Path
        arquivo_agenda_alertas = Path(__file__).resolve().parent / "GPS_agenda_pecs_novembro_2026_COMPLETA.csv"
        agenda_cec = pd.DataFrame()
        if arquivo_agenda_alertas.exists():
            try:
                agenda_cec = pd.read_csv(arquivo_agenda_alertas, dtype=str, keep_default_na=False)
            except Exception as exc:
                st.warning(f"Erro ao ler agenda consolidada: {exc}")
        if agenda_cec.empty and not df_planos_drive.empty and "Data_Visita" in df_planos_drive.columns:
            agenda_cec = df_planos_drive.rename(columns={"Data_Visita": "Data"})
        if agenda_cec.empty:
            st.warning("Não há agenda datada para o cruzamento. Os resumos de planos sem data de visita não comprovam a programação diária.")
        else:
            confronto_v9 = gps_conciliar_documentos(agenda_cec, df_termos_drive, justificativas_cec)
            if confronto_v9.empty:
                st.info("Nenhum registro com PEC, escola e data válidos para conciliar.")
            else:
                filtro_pec_v9 = st.selectbox("PEC — conformidade documental", ["Todos"] + sorted(confronto_v9["PEC"].dropna().unique().tolist()), key="filtro_pec_v9")
                exibicao_v9 = confronto_v9 if filtro_pec_v9 == "Todos" else confronto_v9[confronto_v9["PEC"] == filtro_pec_v9]
                a, b, c = st.columns(3)
                a.metric("Registros confrontados", len(exibicao_v9))
                b.metric("Pendências / divergências", int(exibicao_v9["Alerta"].str.contains("🔴|🟡", regex=True).sum()))
                c.metric("Regulares / justificados", int(exibicao_v9["Alerta"].str.contains("🟢", regex=False).sum()))
                st.dataframe(exibicao_v9, hide_index=True, use_container_width=True)
                st.download_button("⬇️ Baixar conciliação V9 (CSV)", exibicao_v9.to_csv(index=False).encode("utf-8-sig"), "GPS_conciliacao_documental_v9.csv", "text/csv")
        st.caption("Limitação atual: a leitura automática de justificativas Word/PDF do Drive ainda não está validada; CSV estruturado é uma entrada de conferência, não substitui o documento assinado.")

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

        st.markdown("#### 3. Fechamento mensal")
        st.caption("Utilize a conciliação V9 acima. Não se classifica visita como não cumprida somente pela ausência do termo. Planos sem datas de visita devem ser detalhados antes do confronto.")
    idx_proxima_aba = 5

# --- ABA AGENDA SEDUC, CEC & FERIADOS ---
if destino_gps == "Agenda e Calendário":
    st.subheader("Agenda Integrada — SEDUC/EEC, Google Agenda e Bloqueios")

    col_m, col_a = st.columns(2)
    with col_m:
        mes_sel = st.selectbox("Mês:", list(range(1, 13)), index=datetime.now().month - 1)
    with col_a:
        ano_sel = st.number_input("Ano:", min_value=2025, max_value=2030, value=datetime.now().year)

    feriados_sp = holidays.country_holidays("BR", subdiv="SP", years=[ano_sel])
    feriados_mes = [{"Data": d.strftime("%d/%m/%Y"), "Feriado / Ponto Facultativo": nome} for d, nome in sorted(feriados_sp.items()) if d.month == mes_sel]
    feriados_mes.extend({"Data": d.strftime("%d/%m/%Y"), "Feriado / Ponto Facultativo": nome} for d, nome in FERIADOS_EEC.items() if d.year == ano_sel and d.month == mes_sel)
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
if destino_gps == "Diretrizes das Pastas":
    st.subheader("Base Legal e Foco Pedagógico por Pasta (Resolução SEDUC Nº 01/2026)")
    st.dataframe(df_pastas, use_container_width=True)
