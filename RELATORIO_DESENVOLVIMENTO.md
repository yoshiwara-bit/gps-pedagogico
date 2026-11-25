# GPS Pedagógico 2.0 — relatório de desenvolvimento

## Diagnóstico

A branch de desenvolvimento trazia o sistema anterior como referência e o laboratório experimental em `GPS_LAB.gs`. O laboratório já havia demonstrado conversão de arquivos, leitura de planilhas e recuperação parcial de texto por OCR. A reconstrução começou pelo esquema de dados e pelas regras testáveis. Foram preservados o laboratório original e o sistema anterior.

Os experimentos registrados em `AGENTS.md` apresentam divergências de ordem e associação semântica em OCR. Duas imagens do SAEB produziram erro `imageId` mesmo após reconversão. A implementação nova registra a posição de cada falha e prossegue com as demais imagens. A recuperação dessas duas imagens continua em aberto.

## Implementação

| Arquivo | Responsabilidade |
| --- | --- |
| `src/appsscript.json` | Runtime V8, fuso paulista e serviço avançado do Drive v3 |
| `src/Dados.gs` | Esquema do Sheets, criação idempotente e validação de IDs, referências e documentos duplicados |
| `src/Documentos.gs` | Leitores e conversão em pasta isolada, reconhecimento de grades, seleção de cabeçalhos, OCR, classificação conservadora e propostas de revisão |
| `src/Acompanhamento.gs` | Agenda consolidada, conciliação de visitas e comprovantes, metas por escopo e monitoramento de prioridades |
| `src/Pedagogia.gs` | Importação por contexto de escala, critérios sujeitos à aprovação do CEC e projeção agregada de coortes |
| `src/Fluxo.gs` | Núcleo da fila, lotes, novas tentativas, caminho documental por data da visita e recorrência de alertas |
| `src/Web.gs` e `src/Painel.html` | Interface responsiva inicial de consulta paginada, com controle por perfil no servidor |
| `tests/gps.test.cjs` | Verificações sintéticas e exercício local de volume |
| `AGENTS.md` | Instruções da reconstrução, granularidade permitida e isolamento dos testes |

A projeção histórica é regional por série e componente. Retenções, transferências, mudanças de escola e saída da rede exigem dados agregados adicionais antes de qualquer interpretação longitudinal por escola. Os resultados atuais por turma vêm da Prova Paulista consolidada. A referência histórica é registrada separadamente e não altera a escala do resultado corrente.

A agenda recebe atividades já identificadas e datas ISO. Ela separa escolas, ignora tabelas contêiner já sinalizadas e usa tipo, PEC, escola, data e horário para deduplicação. Registros com a mesma chave e escopo pedagógico incompatível geram revisão. Uma visita pode ser comprovada por um documento e ainda ter pendência pelo segundo. O planejamento vencido e a comprovação têm contagens separadas. Os limites de cada semana ou quinzena são fornecidos explicitamente pelo chamador.

Os leitores de PDF e Word retornam texto e tabelas quando o conversor as expõe. A importação estruturada usa grades com cabeçalho identificável e contexto explícito de avaliação, escala e métrica. Boletins que chegam como texto desestruturado ou OCR ainda exigem interpretação humana antes da importação. Valores ausentes e percentuais fora do intervalo esperado permanecem identificáveis. A detecção de cabeçalhos pessoais é uma barreira adicional. A origem autorizada deve ser um consolidado revisado.

## Evidências locais

- `node --test tests/*.test.cjs` executou 17 testes, todos aprovados. Foram cobertos classificação, OCR e `imageId`, escalas, granularidade, coortes, metas, conciliação documental, integridade, perfil, fila e alertas.
- O exercício de volume processou 500 entradas em lotes e verificou 20 e 40 consultas agendadas localmente sobre 50 mil registros sintéticos. O exercício roda em um processo Node com funções simuladas.
- O exercício de volume não mede latência do navegador, cotas do Apps Script, carga concorrente no Sheets ou tempo de conversão no Drive. Os objetivos de p95 de 500 ms e abertura em até 5 s seguem pendentes de medição.

## Roteiro de validação em projeto autorizado

1. Usar um projeto Apps Script de laboratório, com o projeto padrão gerenciado pelo Apps Script. Copiar os arquivos de `src/` e habilitar Drive v3 conforme o manifesto. Manter `GPS_LAB.gs` fora desse projeto.
2. Configurar `GPS_LAB_DB_ID` e `GPS_LAB_FOLDER_ID` nas propriedades privadas do script, apontando para uma planilha vazia e uma pasta sintética isolada. Executar `gpsPrepararBaseLaboratorio_` pelo editor duas vezes e verificar a estrutura.
3. Criar registros sintéticos nas abas e provisionar apenas contas de teste autorizadas na aba `usuarios`. Confirmar que a política do Workspace fornece a identidade ativa no modo escolhido. Validar recusa para identidade ausente, duplicada ou inativa e isolamento entre dois PECs antes de usar a interface com qualquer dado institucional.
4. Exercitar os leitores com documentos sintéticos de cada formato diretamente na pasta experimental. Comparar os resultados com gabaritos, conferir a ordem das imagens e verificar a permanência dos originais e a limpeza dos temporários.
5. Medir 20 e 40 sessões autorizadas em um laboratório isolado com carga sintética, separando latência inicial, navegação interna, conversão e taxa de falhas. Manter e-mails e gatilhos desativados nessa validação.

## Pendências e próximas etapas

- Criar adaptadores de leiaute para boletins SARESP e SAEB, validação das imagens indisponíveis, cabeçalhos compostos e mapeamento de modelos reais autorizados.
- Automatizar extração de agenda e vínculo de termos a partir de documentos brutos. Integrar as funções de validação ao gravador do Sheets, com integridade em lote e auditoria persistente. O esquema inclui revisões, justificativas e encaminhamentos. As telas de edição e aprovação ainda precisam ser implementadas.
- Integrar a fila ao upload e à descoberta incremental no Drive, com persistência, bloqueio, checkpoint e reprocessamento operacional. Avaliar um gatilho de cinco minutos conforme o orçamento de cotas. O caminho documental é calculado localmente. Nenhum arquivo foi movido.
- Conectar alertas a eventos, prazos, justificativas e histórico persistente. Implementar pop-ups e envio de e-mail somente após revisar destinatários e resolver idempotência dos envios.
- Completar fichas, painel estratégico, formulários, gráficos, mapas de calor, comparativos e relatórios. Leaflet e Chart.js são opções para mapas e gráficos. Antes da adoção, validar o carregamento por HTTPS no HTML Service, versão dos ativos e provedor cartográfico. Um SVG local com limites autorizados é alternativa sem serviço de mapas.
- Validar em ambiente Google cache, autorização por perfil, auditoria, desempenho, cotas e concorrência. A consulta inicial usa paginação e leitura de identidade a cada chamada. Os painéis são um ponto de partida, com registros sintéticos até haver autorização para uso dos dados institucionais.
