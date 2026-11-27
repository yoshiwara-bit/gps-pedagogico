import time
import streamlit as st
from gps_integracao import consultar_visitas_lab, limpar_cache_lab

st.set_page_config(page_title='GPS 2.0 - LAB', layout='wide')
st.title('GPS 2.0 | Laboratório de integração')
st.info('Ambiente experimental. Somente dados fictícios.')
if st.button('Atualizar dados do Google Sheets'):
    limpar_cache_lab()
tela = st.sidebar.radio('Navegação', ['Resumo', 'Registro de visita'])
inicio = time.perf_counter()
try:
    dados = consultar_visitas_lab()
except Exception:
    st.error('Não foi possível consultar o LAB. Verifique os segredos e a implantação.')
    st.stop()
tempo = time.perf_counter() - inicio
c1, c2, c3 = st.columns(3)
c1.metric('Registros', len(dados['registros']))
c2.metric('Consulta externa original', f"{dados['tempo_externo']:.3f} s")
c3.metric('Consulta nesta execução', f'{tempo:.6f} s')
if tela == 'Resumo':
    st.success('Conexão com a base LAB confirmada.')
else:
    st.dataframe(dados['registros'], use_container_width=True)
st.caption('Cache de 60 segundos. Os tempos medem a consulta, não a renderização completa da página.')
