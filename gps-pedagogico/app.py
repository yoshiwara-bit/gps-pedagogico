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
# CONVERSÃO DE PERCENTUAL (DECIMAL -> 0 a 100%) + DETECÇÃO DE PROVA / PERÍODO / NÍVEL
# ==============================================================================
def converter_para_percentual(val):
    if pd.isna(val):
        return None
    s = str(val).replace("%", "").replace(",", ".").strip()
    if not s or s.lower() in ["nan", "none", "-", ""]:
        return None
    try:
        num = float(s)
        # Se vier como decimal do BI (ex: 0.6033 ou 1.0), converte para 60.33 ou 100.0
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

    # 1. Tipo Principal
    if "FREQUENCIA" in ref:
        tipo = "FREQUENCIA"
    elif any(k in ref for k in ["PLATAFORMA", "KHAN", "MATIFIC", "SPEAK", "OPEN ENGLISH", "LEIA SP", "REDACAO", "ALURA"]):
        tipo = "PLATAFORMA"
    else:
        tipo = "AVALIACAO"

    # 2. Subcategoria (Prova ou Plataforma)
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

    # 3. Período (Bimestre, Ano ou Semana/Mês)
    m_bim = re.search(r"([1-4])\s*[º°O]?\s*BIM", ref)
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

    # 4. Escola pelo nome do arquivo ou pasta
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

# ==============================================================================
# TRADUTOR INTELIGENTE DE PLANILHAS (ESCOLA -> TURMA -> ESTUDANTE)
# ==============================================================================
def padronizar_planilha_bruta(df_raw, nome_arquivo="", pasta_origem=""):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    tipo_ind, subcategoria, periodo_ref, escola_detectada_path = detectar_classificacao_arquivo(nome_arquivo, pasta_origem)
    cols_norm = {normalizar_texto(c): c for c in df_raw.columns}
    siglas_presentes = [c_orig for c_n, c_orig in cols_norm.items() if c_n in MAPA_SIGLAS_BI]

    # Tenta capturar NM_ESCOLA ou NM_TURMA da linha de rodapé "Filtros aplicados:" do BI SEDUC
    escola_rodape = None
    turma_rodape = None
    for val in df_raw.iloc[:, 0].dropna().astype(str):
        if "FILTROS APLICADOS" in normalizar_texto(val):
            if "NM_ESCOLA" in val:
                trecho = val.split("NM_ESCOLA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho:
                        escola_rodape = limpar_nome_escola_bi(trecho.split(sep)[1].split("CD_DIRETORIA")[0].split("NM_TURMA")[0].strip())
                        break
            if "NM_TURMA" in val:
                trecho_t = val.split("NM_TURMA")[1]
                for sep in [" é ", " = "]:
                    if sep in trecho_t:
                        turma_rodape = trecho_t.split(sep)[1].split("NM_ESCOLA")[0].split("CD_DIRETORIA")[0].strip()
                        break

    escola_padrao = escola_rodape or escola_detectada_path or "Geral URE Miracatu"

    # CASO A: EXPORTAÇÃO PADRÃO BI SEDUC (COM SIGLAS MAT, PORT, ING, HIST, GEO...)
    if siglas_presentes:
        col_escola = next((c_orig for c_n, c_orig in cols_norm.items() if "EXPORT ESCOLA" in c_n or c_n == "ESCOLA" or c_n == "NM_ESCOLA"), None)
        col_turma = next((c_orig for c_n, c_orig in cols_norm.items() if "EXPORT TURMA" in c_n or c_n == "TURMA" or c_n == "NM_TURMA"), None)
        col_aluno = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ALUNO", "ESTUDANTE", "NOME", "RA"]) and "TOTAL" not in c_n), None)
        col_part = next((c_orig for c_n, c_orig in cols_norm.items() if "PARTICIPACAO" in c_n or "FREQ" in c_n), None)
        col_total_al = next((c_orig for c_n, c_orig in cols_norm.items() if "TOTAL ALUNO" in c_n or "ESTUDANTE" == c_n), None)

        registros = []
        for _, row in df_raw.iterrows():
            # Identifica chave principal da linha e descarta rodapés / linhas vazias
            valores_linha_str = " ".join([str(v) for v in row.values if pd.notna(v)])
            if not valores_linha_str.strip() or "Filtros aplicados" in valores_linha_str:
                continue

            val_esc = str(row[col_escola]).strip() if col_escola and pd.notna(row[col_escola]) else ""
            val_tur = str(row[col_turma]).strip() if col_turma and pd.notna(row[col_turma]) else ""
            val_alu = str(row[col_aluno]).strip() if col_aluno and pd.notna(row[col_aluno]) else ""

            if any(x.lower() in ["total", "none", "nan"] for x in [val_esc, val_tur, val_alu] if x):
                continue

            if col_aluno and val_alu:
                nivel_gran = "ESTUDANTE"
                escola_row = limpar_nome_escola_bi(val_esc) if val_esc else escola_padrao
                turma_row = val_tur if val_tur else (turma_rodape or "Turma Geral")
                estudante_row = val_alu
            elif col_turma and val_tur:
                nivel_gran = "TURMA"
                escola_row = limpar_nome_escola_bi(val_esc) if val_esc else escola_padrao
                turma_row = val_tur
                qtd_al = str(row[col_total_al]).strip() if col_total_al and pd.notna(row[col_total_al]) else "-"
                estudante_row = f"Consolidado Turma ({qtd_al} alunos)"
            elif col_escola and val_esc:
                nivel_gran = "ESCOLA"
                escola_row = limpar_nome_escola_bi(val_esc)
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
                    "Subcategoria": subcategoria,
                    "Periodo": periodo_ref,
                    "Nivel_Granularidade": nivel_gran,
                    "Escola": escola_row,
                    "Turma": turma_row,
                    "Estudante": estudante_row,
                    "Componente": comp_nome,
                    "Proficiencia_Pct": acerto_pct,
                    "Participacao_Pct": part_pct,
                    "Nivel": classificar_farol_nivel(acerto_pct)
                })
        return pd.DataFrame(registros)

    # CASO B: PLANILHAS DE FREQUÊNCIA, PLATAFORMAS (KHAN, MATIFIC, SPEAK, LEIA SP) OU OUTROS FORMATOS
    col_esc = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ESCOLA", "UNIDADE", "UE"])), None)
    col_tur = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["TURMA", "SERIE", "ANO"])), None)
    col_alu = next((c_orig for c_n, c_orig in cols_norm.items() if any(k in c_n for k in ["ALUNO", "ESTUDANTE", "DISCENTE", "NOME"]) and "TOTAL" not in c_n), None)
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

        if any(x.lower() in ["total", "none", "nan"] for x in [val_esc, val_tur, val_alu] if x):
            continue

        if col_alu and val_alu:
            nivel_gran = "ESTUDANTE"
            escola_row = limpar_nome_escola_bi(val_esc) if val_esc else escola_padrao
            turma_row = val_tur if val_tur else (turma_rodape or "Turma Geral")
            estudante_row = val_alu
        elif col_tur and val_tur:
            nivel_gran = "TURMA"
            escola_row = limpar_nome_escola_bi(val_esc) if val_esc else escola_padrao
            turma_row = val_tur
            estudante_row = "Consolidado Turma"
        elif col_esc and val_esc:
            nivel_gran = "ESCOLA"
            escola_row = limpar_nome_escola_bi(val_esc)
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
            "Subcategoria": subcategoria,
            "Periodo": periodo_ref,
            "Nivel_Granularidade": nivel_gran,
            "Escola": escola_row,
            "Turma": turma_row,
            "Estudante": estudante_row,
            "Componente": comp_row,
            "Proficiencia_Pct": val_pct,
            "Participacao_Pct": val_pct if tipo_ind == "FREQUENCIA" else None,
            "Nivel": classificar_farol_nivel(val_pct)
        })

    return pd.DataFrame(registros_b)

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
            return "background-color: #FFCDD2; color: #B71C1C; font-weight: bold;"  # Vermelho
        elif num <= 60.0:
            return "background-color: #FFF9C4; color: #F57F17; font-weight: bold;"  # Amarelo
        elif num <= 80.0:
            return "background-color: #C8E6C9; color: #1B5E20; font-weight: bold;"  # Verde
        else:
            return "background-color: #BBDEFB; color: #0D47A1; font-weight: bold;"  # Azul
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

    # 1. VISÃO GERAL POR ESCOLA (Sempre consolidada por Escola, nunca misturada por turma)
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

    cols_num_esc = [c for c in piv_esc.columns if c != "Escola"]
    styler_esc = piv_esc.style.map(estilo_celula_farol, subset=cols_num_esc).format({c: "{:.1f}%" for c in cols_num_esc}, na_rep="-")
    st.dataframe(styler_esc, use_container_width=True, hide_index=True)

    # 2. DRILL-DOWN: ENTRAR NA ESCOLA -> GERAL POR TURMA DA ESCOLA
    escolas_disp = sorted(df_sub["Escola"].dropna().unique().tolist())
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

            cols_num_tur = [c for c in piv_tur.columns if c != "Turma"]
            styler_tur = piv_tur.style.map(estilo_celula_farol, subset=cols_num_tur).format({c: "{:.1f}%" for c in cols_num_tur}, na_rep="-")
            st.dataframe(styler_tur, use_container_width=True, hide_index=True)

            # 3. DRILL-DOWN: ENTRAR NA TURMA -> VISÃO ESTUDANTE POR ESTUDANTE (GERAL DA TURMA)
            turmas_disp = sorted([t for t in df_turma_base["Turma"].dropna().unique().tolist() if "Consolidado" not in t])
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
                        piv_alu = piv_alu[cols_ord_a].reset_index()

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
    p_info.add_run("Unidade Escolar Prioritária: ").bold = True
    p_info.add_run(f"{escola_sel}\n")
    p_info.add_run("Período / Mês de Referência: ").bold = True
    p_info.add_run(f"{semana_str}\n")

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
# CARREGAMENTO GLOBAL DOS DADOS DO DRIVE
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
st.caption("Planejamento, Indicadores Hierárquicos (Escola → Turma → Estudante) e Acompanhamento — URE de Miracatu")

lista_abas = [
    "📋 Plano de Trabalho & Exportação",
    "📊 Indicadores (Avaliações, Plataformas & Frequência)",
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

    df_prio_word = pd.DataFrame()
    if not df_ind_usuario.empty:
        df_prio_word = df_ind_usuario[df_ind_usuario["Proficiencia_Pct"] <= 60.0]

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

    # 1. SUB-ABA AVALIAÇÕES
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

    # 2. SUB-ABA PLATAFORMAS
    with sub_plat:
        df_pl = df_ind_usuario[df_ind_usuario["Tipo_Indicador"] == "PLATAFORMA"] if not df_ind_usuario.empty else pd.DataFrame()
        plats_lista = ["Todas as Plataformas", "Khan Academy", "Matific", "Speak", "Open English", "Leia SP", "Redação Paulista", "Alura"]
        plat_escolhida = st.selectbox("Selecione a Plataforma Digital:", plats_lista)

        df_pl_filt = df_pl.copy()
        if not df_pl_filt.empty and plat_escolhida != "Todas as Plataformas":
            df_pl_filt = df_pl_filt[df_pl_filt["Subcategoria"] == plat_escolhida]

        renderizar_tabela_hierarquica(df_pl_filt, f"Plataforma — {plat_escolhida}", "Engajamento / Uso (%)")

    # 3. SUB-ABA FREQUÊNCIA
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
