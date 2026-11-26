# GPS Pedagógico 2.0 — relatório da quarta missão

A implementação partiu de `0be8917` na branch `gps-2-apps-script`, após revisão da missão integral, dos 22 módulos, dos relatórios e dos fluxos do GPS V3. O GPS oficial, `GPS_LAB.gs` e as implantações existentes foram preservados.

| Área | Implementação nesta entrega | Situação ou dependência |
| --- | --- | --- |
| Interface | Menu lateral recolhível, organização por perfil, estados explícitos de módulos parciais, visão CEC com prioridades e alertas, acesso às funções anteriores. | Navegação validada localmente. Painel estratégico completo, agenda EEC própria, metas regionais e administração AT exigem complementação. |
| Cartografia | SVG com projeção de coordenadas, seis divisas GeoJSON, cores municipais, sede opcional, zoom, marcadores, prévia com contexto, calor por intensidade e camadas com filtros. Clique abre ficha. | Demonstração esquemática. Renderizador e validação preparados. A consulta ao IBGE foi bloqueada pelo proxy do ambiente. Faltam GeoJSON autorizado, coordenadas verificadas, sede e ensaio real. |
| Agenda territorial | Consulta autenticada por dia, estado, escola, PEC autorizado, foco e período somente quando informado. Registro de plano oferece manhã ou tarde opcional. | Histórico territorial por intervalo preservado. Alterações sincronizadas, Google Agenda e roteirização rodoviária pendentes. Função de distância linear disponível como base técnica, sem cálculo de itinerário. |
| Ficha escolar | Busca, tabela e marcador levam à ficha. Nove seções sob demanda: cadastro, equipe, agenda, frequência, avaliações, plataformas, acompanhamentos, prioridades e histórico. Importador do cadastro enriquecido a partir de CSV privado com fonte. | Campos ausentes explícitos. Extração automática dos PDFs institucionais, ATPC, docentes acompanhados, termos externos, encaminhamentos e histórico de intervenções incompletos. |
| Frequência | Filtros de recorte e granularidade da fonte, ponderação por denominador, evolução em p.p., comparativo, farol homologado e tabela. Duplicatas ficam fora da ponderação e lacunas são sinalizadas. | Períodos totalmente ausentes exigem um calendário de entregas homologado. Metas, prioridades por turma e roteamento de alertas por pasta ainda dependem de fontes e integração. |
| Avaliações | Cinco painéis com filtros comuns, escala e período preservados, faróis existentes, séries e comparativos onde os dados são compatíveis, e barras empilhadas para distribuições completas. | Sem inventar cortes ou resultados por turma. Boletins, convenções de escala e amostras de cada edição ainda exigem homologação. Projeções anteriores preservadas. |
| Plataformas | Catálogo regional e 12 painéis, dez logotipos já existentes em versões reduzidas, filtros, metas na fonte, resultados, comparativos compatíveis e faróis de registros homologados. | Matific AF e Programação aparecem com nome textual. Adaptadores consolidados disponíveis; metas, métricas por ferramenta, prioridades e deduplicação anual por usuário aguardam contratos de fonte. Acumulados exigem registro consolidado. |
| Exportação | CSV compatível com Excel, impressão para PDF e gráfico de linha em SVG. | CSV exporta o recorte exibido, até 200 linhas. XLSX nativo e relatórios institucionais versionados seguem pendentes. |
| Permissões | CEC recebe os novos indicadores. PEC mantém os registros próprios e escolas atribuídas. AT mantém o recorte escolar de consulta. Identidade e escopo conferidos no servidor. | Regras mais granulares de pasta, componente e escopos administrativos da AT aguardam matriz homologada. A interface sinaliza módulos sem integração. |
| Regressão | Acesso a planos, termos, justificativas, regularidade, revisão, alertas e consultas existentes. | Validação e devolução integrais de planos, conectores, automações, MDI, formações e orientações seguem com pendências documentadas nos módulos. |

## Testes e compatibilidade

- 47 testes Node passaram, incluindo cinco do laboratório original. Cobrem geometria, autorização por escola, estrutura de importação, dados ausentes e duplicados, ponderação, escala e distribuição dos níveis.
- Chromium com RPC simulada verificou CEC em 1280 px e PEC e AT em 390 px. Exercícios incluem menu e recolhimento, pesquisa e ficha, clique e prévia em mapa sintético, agenda, faróis, gráficos, avaliações e plataformas. As capturas são desse ambiente simulado.
- A carga sintética anterior de 20/40 consultas, 500 documentos e 50 mil registros foi reexecutada pela suíte. Ela executa as regras em um processo Node e não reproduz sessões concorrentes no Workspace. Latência, cotas, cache real e HTML Service ainda requerem ensaio institucional.
- A opção de SVG usa DOM nativo e dispensa CDN e carregamento de tiles. Leaflet é viável como alternativa se ativos HTTPS e política do HTML Service forem homologados. A etapa inclui carregamento por seção, cache territorial existente e paginação de consultas. Os painéis ainda leem as respectivas abas em lote e limitam a resposta tabular inicial a 200 linhas.
- Nenhuma implantação, gatilho ou mensagem real foi executada. O pacote novo contém somente código e documentação do piloto.

## Arquivos e dados necessários para homologação

1. GeoJSON das seis divisas, cadastro autorizado com as coordenadas das 40 escolas e localização verificada da sede. Se houver rota real, provedor de malha viária e política de uso.
2. Fichas gerenciais consolidadas com CIE/INEP, contatos e contagens agregadas, cadastro de atribuições e matriz de acesso por papel, escola, componente e pasta.
3. Frequência por período e recorte com numeradores, denominadores, cobertura esperada, metas e documento oficial de farol.
4. Exportações específicas das 12 plataformas, dicionário de métricas, regras de denominador, granularidade, deduplicação e metas por edição.
5. Gabaritos e documentos orientadores de Prova Paulista, AvD, SARESP, SAEB e ENEM. Agenda e calendário oficiais, termos e situação de evidências, medidas de deslocamento, arquivos de ATPC e encaminhamentos autorizados.
6. Contas de validação CEC, PEC e AT em laboratório e testes de carga no Workspace. Guarde IDs e autorizações nas propriedades privadas. [Guia de atualização](GUIA_ATUALIZACAO_ETAPA_4.md)
