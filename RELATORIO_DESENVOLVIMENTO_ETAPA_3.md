# GPS Pedagógico 2.0 — terceira missão

## Origem e escopo

Foram lidos os 27 arquivos do pacote de especificações. O segundo ZIP recebido é idêntico ao primeiro. Os documentos permanecem em `docs/gps-2.0/`, com o AGENTS original do pacote identificado como `AGENTS_DO_PACOTE.md` e uma referência no AGENTS da raiz. O documento mestre registra as observações do piloto. A inspeção local partiu de `1094dfb`. A implantação existente fica na mesma versão.

## Matriz de rastreabilidade

| Módulo | Arquivo ou evidência | Resultado e pendências |
| --- | --- | --- |
| 01 Arquitetura | `Dados.gs`, `Demonstracao.gs`, testes de instalação e trava | Esquema ampliado com abas novas. Validação real da migração e ensaio de reversão pendentes. |
| 02 Calendário | `Regularidade.gs`, teste de calendário completo | Cálculo do segundo dia útil com dias oficialmente conferidos. Importadores, revisão de eventos e proposta de atualização pendentes. |
| 03 Agenda e planos | `Acompanhamento.gs`, `Operacao.gs`, teste de conciliação | Formulário e deduplicação anteriores preservados. Word institucional, Google Agenda e sincronização aprovada pendentes. |
| 04 Validação de planos | `Operacao.gs`, teste de prioridades | Avisos prévios preservados. Fluxo de devolução, versões e publicação após aprovação pendentes. |
| 05 Termos e apoios | `Operacao.gs`, `Notificacoes.gs`, testes de conciliação | Termo web encerra a cobrança vinculada. Importação e associação de termos externos e evidências ATPC pendentes. |
| 06 Justificativas | `Atuacao.html`, `Operacao.gs`, teste de decisão | Fluxo PEC após sete dias sem termo. Atraso conserva a cobrança; não ocorrência exige motivo, retira contagem e preserva o plano original. Conflito com evidência recusa a decisão. Faltam ocorrência provisória independente, tratamento de alterações e formulário mensal institucional. |
| 07 Regularidade | `Regularidade.gs`, `Web.gs`, testes de prazos | Consulta CEC e PEC sobre registros preparados. Cores e lembretes D-7 e D-2 somente com prazo e referência homologados. Geração de obrigações e emissão dos lembretes ainda pendentes. |
| 08 Comunicação | `Atuacao.html`, `Notificacoes.gs`, Chromium e teste de reserva | CEC vê estado explícito e atalho para alertas gerais. Integrações externas, notificações administrativas AT e consolidação por escola pendentes. |
| 09 Orientações | Especificação 09 | Pendente, incluindo modelos e aprovação CEC. |
| 10 Indicadores | `Web.gs`, `Painel.html`, testes de farol | Cor separada do texto. PP admite os quatro níveis oficiais; classificação exige regra inequívoca de fonte, ano, escala, métrica e nível explícito, com referência. Nota isolada permanece neutra. Cortes, plataformas e layouts oficiais pendentes. |
| 11 Frequência | `Regularidade.gs`, teste de cobertura parcial | Núcleo de agregação ponderada por denominador e sinalização de lacunas. Importação, séries temporais e alertas por pasta pendentes. |
| 12 Docentes | Especificação 12 | Pendente, com dados pessoais restritos a base autorizada. |
| 13 Formações | Especificação 13 | Pendente, depende de modelos e presença autorizada. |
| 14 Feedback | Especificação 14 | Pendente. Comparações devem preservar escala e contexto. |
| 15 CEC | `Atuacao.html`, `Web.gs`, teste Chromium | A aba antes vazia agora distingue falta de formulários atribuídos de ausência de alertas pessoais. Agenda própria e decisão de intervenção pendentes. |
| 16 MDI e 1x1 | Especificação 16 | Pendente. Metodologia MDI e integração Agenda exigem homologação. |
| 17 Intervenções | `Pedagogia.gs`, testes de coorte e histórico escolar | Projeções anteriores preservadas. Fluxo de decisão, pesos e acompanhamento posterior pendentes. |
| 18 AT e eventos | Especificação 18 e teste de escopo | Consulta restrita anterior preservada. Checklist e integrações administrativas pendentes. |
| 19 Integridade | `Sincronizacao.gs`, testes de fila e checkpoints | Deduplicação e checkpoint anteriores preservados. Central de cobertura e integridade para AT e CEC pendente. |
| 20 Mapas e fichas | `Atuacao.html`, `Web.gs`, teste Chromium | Filtros responsivos, cabeçalhos traduzidos, nomes de cadastros e ficha resumida ao clicar. SVG identificado como esquemático. Divisas, 40 escolas reais, geodados e ficha completa dependem de fontes autorizadas. |
| 21 Relatórios | `Acompanhamento.gs`, `Territorio.gs`, teste de decisão | A não ocorrência sai dos totais operacionais e permanece na agenda de origem. Exportações, snapshots e ATPC pendentes. |
| 22 Segurança e desempenho | `Web.gs`, `Operacao.gs`, testes de perfil, carga local e Chromium | Autorização no servidor, confirmação para registrar decisão e nenhuma alteração em gatilhos ou e-mails. Validação do Workspace, cotas, carga concorrente real e tempo de resposta pendentes. |

## Validação e limitações

- 40 testes locais passaram, incluindo cinco do laboratório original. Chromium passou nos cenários CEC, PEC e AT com chamadas simuladas. Foram verificados o estado vazio do CEC, o farol com texto de nível e cor, a ficha por clique e a contenção dos controles no grid a 390 e 1280 pixels.
- O mapa autorizado, OCR e `imageId`, regras oficiais, calendário, conectores e permissões seguem sem validação real. Nenhum módulo foi homologado institucionalmente.
- A configuração atual de importação preserva o escopo anterior. A incorporação de AvD por turma requer política atualizada de granularidade e validação de base agregada. Nenhum registro individual de estudante foi criado.

## Instalação e reversão

Use o guia experimental e o pacote desta etapa somente em projeto e planilha de laboratório isolados. Em uma cópia sintética existente, salve um backup da planilha, instale os arquivos e execute `gpsPrepararBaseLaboratorio_` para criar as quatro abas novas. Em instalação nova, execute o instalador da demonstração. Mantenha gatilhos e e-mails desativados. Teste o fluxo com casos sintéticos e três identidades autorizadas no Workspace antes de propor qualquer atualização da implantação existente.

Para reverter um teste futuro, restaure a versão anterior do código no projeto de laboratório e a cópia da planilha anterior à migração. As quatro abas são aditivas. A versão anterior ignora as novas abas e, portanto, não aplica decisões registradas nelas. Preserve a cópia pós-migração para auditoria.
