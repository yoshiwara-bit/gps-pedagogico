
import time
import requests
import streamlit as st

@st.cache_data(ttl=60, max_entries=1, show_spinner=False)
def consultar_visitas_lab():

    config = st.secrets["gps_lab"]
    inicio = time.perf_counter()

    resposta = requests.post(
        config["url"],
        json={
            "chave": config["chave"],
            "acao": "consultar_visitas"
        },
        timeout=30
    )

    resposta.raise_for_status()
    dados = resposta.json()

    def falha(mensagem):
        st.error(mensagem)
        raise ValueError("Falha de validacao LAB.")

    if not isinstance(dados, dict):
        falha("D01 - Formato de resposta invalido.")

    if dados.get("status") == "NEGADO":
        falha("D02 - Autenticacao ou acesso negado.")

    if dados.get("status") != "OK":
        falha("D03 - Apps Script retornou erro.")

    if dados.get("ambiente") != "LAB":
        falha("D04 - Ambiente LAB nao identificado.")

    if dados.get("total") != 1:
        falha("D05 - Total de registros inesperado.")

    registros = dados.get("registros", [])

    if not isinstance(registros, list) or len(registros) != 1:
        falha("D06 - Lista de registros invalida.")

    registro = registros[0]

    if not isinstance(registro, dict):
        falha("D07 - Estrutura do registro invalida.")

    if registro.get("ID_REGISTRO") != "LAB-001":
        falha("D08 - Identificador LAB incorreto.")

    if registro.get("ESCOLA") != "Escola Fictícia Alfa":
        falha("D09 - Escola ficticia nao identificada.")

    campos = [
        "ID_REGISTRO",
        "ESCOLA",
        "PEC",
        "DATA_VISITA",
        "ATIVIDADE",
        "OBJETIVO",
        "ENCAMINHAMENTO"
    ]

    if any(campo not in registro for campo in campos):
        falha("D10 - Campos obrigatorios ausentes.")

    return {
        "registros": [
            {campo: registro[campo] for campo in campos}
        ],
        "tempo_externo": time.perf_counter() - inicio
    }

def limpar_cache_lab():
    consultar_visitas_lab.clear()
