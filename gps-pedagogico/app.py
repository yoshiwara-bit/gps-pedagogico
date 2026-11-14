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
        elif "etapa" in c_low or "segmento" in c_low
