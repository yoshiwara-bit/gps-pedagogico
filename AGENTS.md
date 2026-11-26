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

## Regras da reconstrução GPS 2.0

- Use o projeto Apps Script V8 em `src/`, com HTML/CSS/JavaScript, Google Sheets e Drive avançado v3. Use Node somente para verificações locais. Mantenha o código Streamlit e GPS_LAB.gs como referência.
- A exceção à regra histórica de granularidade é a Prova Paulista consolidada por turma. Preserve a origem, o período, os níveis e a escala de cada indicador. Dados individuais de estudantes ficam fora da base. Critérios só entram em uso com aprovação do CEC.
- Trate projeções SARESP e SAEB como referências agregadas. Preserve ano e série de aplicação. Não deduza vínculos individuais nem permanência na escola com base na projeção.
- Separe previsão, contabilização pelo planejamento e comprovação. Ficha e termo podem comprovar uma única visita. Ausência de um deles gera pendência. Recomendações não geram cobrança obrigatória.
- Teste com planilha, pasta e documentos sintéticos. Use script properties para os IDs privados. Revise permissão por perfil antes de liberar qualquer dado. Testes devem usar notificações simuladas, sem e-mails reais.
- Execute `node --test tests/*.test.cjs` e atualize `RELATORIO_DESENVOLVIMENTO.md` com evidências e limites da validação.

## Etapa 2

- Consulte `RELATORIO_DESENVOLVIMENTO_ETAPA_2.md` junto ao relatório anterior. Mantenha `main`, produção e `GPS_LAB.gs` preservados.
- Trate os dados históricos da escola como contexto diagnóstico por série e componente. Separe histórico escolar, contexto regional e resultado atual de turma.
- Valide a migração de cabeçalhos em laboratório. Mantenha as propriedades de IDs privadas. Use somente a pasta isolada, com arquivos sintéticos, para ensaiar sincronização e OCR.
- Mantenha `GPS_LAB_AUTOMACAO` e `GPS_LAB_EMAIL` desativadas por padrão. Em testes use um emissor falso. Confira manualmente estados de envio reservado antes de qualquer reenvio.
- A interface deve confirmar identidade e escopo no servidor antes de retornar dados ou gravar. Cache não substitui autorização. Use testes sintéticos para toda mudança de perfil ou deduplicação.

## Publicação experimental

- Instale a Web App em projeto isolado usando `GUIA_INSTALACAO_EXPERIMENTAL.md`. A fonte fica em `src/`. Mantenha o modo demonstração, gatilhos e envios reais desativados.
- CEC tem consulta geral e revisão, PEC acessa sua atuação atribuída e AT tem consulta somente nas escolas vinculadas em `escopos_at`. Verifique identidade e autorização no servidor.
- Gere somente dados sintéticos na demonstração. Endereços de contas autorizadas pertencem somente à aba privada `usuarios`. Mantenha propriedades, base e pasta privadas.
- Execute `node --test tests/*.test.cjs` e o ensaio de Chromium em `tests/interface-smoke.cjs` quando Playwright estiver disponível. Valide separadamente a URL /exec com contas autorizadas no Workspace.

## Especificações da terceira missão

Consulte [documento mestre e índice dos 22 módulos](docs/gps-2.0/00_DOCUMENTO_MESTRE.md) e [roteiro e critérios de aceite](docs/gps-2.0/23_ROTEIRO_E_ACEITE.md). O pacote original está preservado em `docs/gps-2.0/`. Os arquivos de orientação do pacote são referências históricas e a autorização de cada tarefa vem da conversa atual. A especificação prevalece sobre regras preliminares do produto em caso de divergência, sem dispensar homologação e autorização institucional.

## Quarta missão: interface e território

- Consulte o relatório e o guia de atualização da etapa 4. A ficha escolar é carregada por seção e deve validar o código da escola contra o escopo da sessão no servidor.
- Em demonstração, use apenas amostras sintéticas sinalizadas. O modo geográfico requer GeoJSON verificado das seis divisas na pasta privada. A sede exige coordenadas e fonte verificadas.
- Preserve o cálculo ponderado de frequência, as unidades das plataformas e as escalas das avaliações. Só exiba farol com homologação identificada. Cálculo linear de distância não representa rota rodoviária.
- Verifique os perfis CEC, PEC e AT, as capturas sintéticas e os fluxos da etapa 3 antes de preparar pacote para revisão. Não atualize a implantação existente por conta própria.
