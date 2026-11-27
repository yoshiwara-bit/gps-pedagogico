import time
import requests
import streamlit as st

@st.cache_data(ttl=60, max_entries=1, show_spinner=False)
def consultar_visitas_lab():
    config = st.secrets['gps_lab']
    inicio = time.perf_counter()
    resposta = requests.post(config['url'], json={
        'chave': config['chave'], 'acao': 'consultar_visitas'
    }, timeout=30)
    resposta.raise_for_status()
    dados = resposta.json()
    registros = dados.get('registros', [])
    if (dados.get('status') != 'OK' or dados.get('ambiente') != 'LAB'
        or dados.get('total') != 1 or len(registros) != 1
        or registros[0].get('ID_REGISTRO') != 'LAB-001'
        or registros[0].get('ESCOLA') != 'Escola Fictícia Alfa'):
        raise ValueError('Resposta LAB inesperada.')
    campos = ['ID_REGISTRO', 'ESCOLA', 'PEC', 'DATA_VISITA', 'ATIVIDADE', 'OBJETIVO', 'ENCAMINHAMENTO']
    return { 'registros': [{c: registros[0][c] for c in campos}],
             'tempo_externo': time.perf_counter() - inicio }

def limpar_cache_lab():
    consultar_visitas_lab.clear()
