
import time
import requests
import streamlit as st
from gps_integracao import consultar_visitas_lab, limpar_cache_lab

st.set_page_config(page_title="GPS 2.0 - LAB", layout="wide")

st.title("GPS 2.0 | Laboratório de integração")
st.info("Ambiente experimental. Somente dados fictícios.")

if st.button("Atualizar dados do Google Sheets"):
    limpar_cache_lab()

tela = st.sidebar.radio(
    "Navegação",
    ["Resumo", "Registro de visita"]
)

inicio = time.perf_counter()

try:
    config = st.secrets["gps_lab"]

    if not config["url"] or not config["chave"]:
        st.error("E01 - Configuração incompleta.")
        st.stop()

    dados = consultar_visitas_lab()

except (KeyError, FileNotFoundError):
    st.error("E01 - Segredos ausentes ou incompletos.")
    st.stop()

except requests.exceptions.Timeout:
    st.error("E02 - Tempo limite de conexão excedido.")
    st.stop()

except requests.exceptions.ConnectionError:
    st.error("E03 - Falha de conexão com o Apps Script.")
    st.stop()

except requests.exceptions.HTTPError as erro:
    codigo = (
        erro.response.status_code
        if erro.response is not None else "indisponível"
    )
    st.error(f"E04 - Erro HTTP: {codigo}")
    st.stop()

except requests.exceptions.JSONDecodeError:
    st.error("E05 - Resposta recebida não está em JSON.")
    st.stop()

except ValueError:
    st.error("E06 - Resposta do Apps Script não passou na validação.")
    st.stop()

except Exception as erro:
    st.error(f"E07 - Tipo de erro: {type(erro).__name__}")
    st.stop()

tempo = time.perf_counter() - inicio

c1, c2, c3 = st.columns(3)

c1.metric("Registros", len(dados["registros"]))

c2.metric(
    "Consulta externa original",
    f"{dados['tempo_externo']:.3f} s"
)

c3.metric(
    "Consulta nesta execução",
    f"{tempo:.6f} s"
)

if tela == "Resumo":
    st.success("Conexão com a base LAB confirmada.")
else:
    st.dataframe(
        dados["registros"],
        use_container_width=True
    )

st.caption(
    "Cache de 60 segundos. Os tempos medem a consulta, "
    "não a renderização completa da página."
)
