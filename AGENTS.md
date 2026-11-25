# GPS Pedagógico 2.0 — instruções para Codex

Trabalhar somente na branch gps-2-apps-script. Não alterar main, não implantar, não acessar dados reais nem enviar documentos institucionais ao repositório público. Preservar GPS_LAB.gs como referência experimental; desenvolver módulos separados e testes sintéticos. Apps Script V8, Drive API avançada v3, Sheets, sem dependência de Cloud Console. Configurar GPS_LAB_FOLDER_ID privadamente nas propriedades do script; logs detalhados somente com dados sintéticos. Não executar funções contra Drive real sem autorização.

## Evidências do laboratório (amostras, não validação de produção)
- Drive, PDF/DOCX, Excel, CSV e Sheets: leitura técnica bem-sucedida. CSV 36x34, separador ponto e vírgula, primeira linha vazia; detectar cabeçalhos múltiplos e células mescladas.
- AvD Excel: 30 escolas, 120 distribuições, 360 percentuais, consistência matemática aprovada. Três níveis: Abaixo do Básico, Básico, Proficiente. Duas edições, LP/MAT. Validar ano, códigos, cabeçalhos dinâmicos; não calcular regional por média simples.
- OCR de quatro imagens AvD: texto recuperado, tabelas não estruturadas; números e ordem podem estar incorretos. Validar antes de gravar.
- SARESP: cinco boletins lidos, notas e percentuais recuperados sem tabelas; associação semântica pendente. Preservar ausências como nulo, distinguir notas, participação e rankings. Farol oficial de componentes: até 4 vermelho, acima de 4 até 6 amarelo, acima de 6 verde.
- SAEB: um boletim, 17 imagens, 15 com texto, imagens 16 e 17 falharam com Argumento inválido: imageId mesmo após reconversão e espera. OCR mistura anos e valores. Investigar alternativas compatíveis; não inventar resultados. Preservar participação acima de 100% com sinalização, níveis originais e escalas versionadas.
- Plano: tabela externa aninhada foi confundida com agenda; agenda real 22 linhas de atividades, 16 visitas escolares previstas na amostra. Separar escolas na mesma data, excluir feriados, reuniões e deslocamentos.
- Termos: três documentos representaram duas visitas, uma com ficha e termo detalhado e outra somente ficha. Exigir ambos para regularidade documental, sem duplicar visitas. Similaridade de 70% é experimental; validar PEC, escola, data, registro e conflitos. PDF SAEB/SARESP nunca é termo.

## Entregas
1. Importadores modulares e classificação por conteúdo, com fila de revisão e testes sintéticos reproduzíveis. Evitar filtros por nomes de escolas ou linhas fixas.
2. Propor esquema Sheets relacional: escolas, PECs, documentos, visitas, apoios, planos, avaliações, frequência, metas, alertas, usuários, auditoria. IDs únicos, idempotência, integridade, origem e versão. Sem indicadores por turma ou estudante; regional, municipal, escola e série quando disponível.
3. Processamento em lote, cache, filas, locks, sincronização incremental; preservar originais e limpar apenas temporários criados pela execução. Não registrar CPF, RG, RA ou relatos pessoais.
4. Documentar em RELATORIO_DESENVOLVIMENTO.md testes executados, resultados, limitações, decisões e próximos passos. Não declarar resolvido o que depende de execução real no Apps Script. Não fazer merge na main.

Etapas 13–18 (base, cruzamento/metas, alertas, sincronização, desempenho concorrente, especificações finais) ainda pendentes. Priorizar correções dos importadores e proposta de base; não presumir regras de negócio ausentes.
