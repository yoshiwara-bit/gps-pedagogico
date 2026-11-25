# GPS Pedagógico 2.0 — relatório da etapa 2

## Escopo

A implementação parte do commit `e5c0b26` em `gps-2-apps-script`. O trabalho ocorreu no checkout de desenvolvimento. A branch `main`, a produção e `GPS_LAB.gs` foram preservados. Nenhum gatilho foi ativado e nenhum e-mail foi enviado.

## O que foi implementado

| Área | Implementação |
| --- | --- |
| Boletins | Adaptadores SARESP e SAEB associam dados por colunas rotuladas, reconhecem etapas EF/EM e separam nota, proficiência, participação e nível. A escala é identificada por fonte, métrica e edição. Duplicidades, OCR, valores ambíguos e grades incompatíveis geram revisão. Boletins convertidos de PDF ficam retidos para conferência do CEC antes de alimentar prioridades. |
| OCR | Para `imageId`, o leitor tenta exportar o documento convertido como HTML/ZIP, localizar imagens pela ordem das referências HTML e submetê-las ao OCR. A tentativa exige contagem correspondente e referências distintas. O resultado mantém revisão obrigatória. Quando a exportação é incompatível, o erro original permanece registrado para revisão. |
| Pedagogia | O motor usa escola, série projetada e componente para distinguir fragilidade persistente, emergente e ausência de histórico escolar. A Prova Paulista do ano vigente oferece o recorte por turma. O histórico regional fica em campo separado. A projeção preserva ano e série de origem e trata a conclusão prevista. O painel CEC cruza prioridades com visitas e apoios das atividades planejadas. |
| Sheets e Drive | O gravador em lote valida referências e IDs. A fila descobre arquivos diretamente na pasta de laboratório, fixa a janela de tempo entre páginas, grava antes de avançar o checkpoint e usa checksum ou ID e versão para deduplicação. Há trava, revisão, auditoria, novas tentativas e reabertura manual de erros. Um instalador de gatilho de cinco minutos fica disponível no laboratório. |
| Alertas | Sinais persistentes para PECs e CEC cobrem prioridades sem planejamento ou comprovação, termos pendentes, última meta encerrada, prazos próximos e encaminhamentos vencidos. Há pop-up e registro de justificativa. O emissor, desativado por padrão, reserva o envio antes da chamada externa. Um envio incerto permanece reservado para conferência. |
| Interface | Formulários de Plano de Trabalho e de ficha ou termo de visita verificam a atribuição e exibem avisos pedagógicos. O CEC pode validar candidato numérico extraído do boletim. A consulta territorial tem escopos de URE, município, escola, PEC e pasta, além da consulta estratégica de prioridades. |
| Território | Mapa SVG interativo de pontos com intensidade por volume, calor temporal e gráficos semanais, quinzenais, mensais e anuais. Filtros incluem escola, município, PEC, pasta, datas e atividade. As séries distinguem agendas previstas, contabilizadas pela data e comprovadas. O importador de cadastro utiliza código, nome, município e coordenadas. |
| Segurança e desempenho | Sessão ativa validada por usuário e perfil. Ações de escrita verificam atribuições, resguardam fórmulas e operam sob trava. A consulta territorial usa cache de 30 segundos por escopo autorizado e data. Planos, atividades, fila e alertas deixam trilha de auditoria. |

## Validação local

- `node --test tests/*.test.cjs` executou 25 testes aprovados. Os novos testes cobrem histórico escolar, SARESP e SAEB, fallback de imagem, checkpoints, agregação territorial, calendário de metas, cadastro e supressão de e-mail após reserva ambígua.
- A suíte conserva o ensaio sintético de 500 documentos e 20 e 40 consultas simuladas sobre 50 mil registros. Os resultados medem somente o processo local.
- A sintaxe de todos os módulos Apps Script, do JavaScript embutido no HTML e do JSON do manifesto passou. Os cinco testes locais do laboratório original também passaram.
- O cadastro versionado foi conferido localmente e contém 40 linhas com coordenadas. Nenhum registro foi carregado no Google.

## Roteiro de laboratório autorizado

1. Usar apenas `src/` no projeto Apps Script de laboratório com Drive avançado v3. Configurar `GPS_LAB_DB_ID` e `GPS_LAB_FOLDER_ID` como propriedades privadas. Manter `GPS_LAB_AUTOMACAO` e `GPS_LAB_EMAIL` desativadas. Executar `gpsPrepararBaseLaboratorio_` em uma cópia sintética da base. A migração acrescenta colunas ao fim e recusa cabeçalhos incompatíveis.
2. Popular os cadastros, atribuições, usuários, anos e componentes com dados sintéticos. Para boletins, conferir os códigos e as escalas. Para grades de Prova Paulista e outras fontes, associar o ID do arquivo a uma regra em `fontes_importacao`, com fonte, escala, métrica, unidade e período. Definir `GPS_LAB_CRITERIO_VERSAO` para a versão aprovada pelo CEC.
3. Para ensaiar o mapa, usar `GPS_LAB_CADASTRO_CSV_ID` com um CSV sintético na pasta isolada e executar `gpsCarregarCadastroTerritorial_`. Para as 40 escolas, autorizar e carregar o cadastro existente em separado. A interface mostra quantas escolas têm coordenadas válidas.
4. Ensaiar o sincronizador manualmente, com automação habilitada apenas nesse projeto sintético. Conferir repetição após checkpoint, troca de versão, pastas de testes, auditoria e deduplicação. Somente depois considerar `gpsConfigurarGatilhoLaboratorio_`. Para reabrir uma falha, usar `GPS_LAB_REPROCESSAR_ID` e executar `gpsReabrirErroFilaLaboratorio_`.
5. Usar boletins de teste com gabarito, verificar conversão, ordenação e valores com o CEC. Confirmar se o Drive retorna ZIP compatível e se as imagens que causaram `imageId` estão presentes no HTML exportado. A recuperação dessas imagens continua sem comprovação.
6. Testar duas identidades PEC com atribuições distintas e uma conta CEC em projeto isolado. Verificar restrição de acesso, contenção do cache, formulários, alertas e atas de revisão. Medir cotas e latência com sessões concorrentes antes de adotar a frequência de cinco minutos.

## Limitações e próximos passos

- A leitura cobre layouts tabulares reconhecíveis. Tabelas compostas, extrações de OCR sem correspondência verificável e versões reais dos boletins precisam de gabaritos. A revisão numérica só atua sobre candidatos já associados a rótulos. Dados descartados por ambiguidade estrutural exigem análise da fonte.
- A descoberta cobre filhos diretos da pasta configurada e mudanças por data de modificação. Movimentos com data antiga, subpastas, estado de exclusão, arquivos removidos e reconciliação por feed de mudanças ainda exigem cobertura. Escritas no Sheets usam trava e chaves repetíveis. O Sheets não fornece transação única entre abas.
- A evidência web registra metadados e observação no Sheets. Anexação e vinculação automática de termos externos, fluxo de aprovação do documento inteiro e edições posteriores de planos ainda precisam de trabalho.
- Os recortes de URE, municípios, escolas e PECs usam o painel e os filtros comuns. Fichas institucionais detalhadas, polígonos municipais, exportações e personalizações de identidade visual ainda estão pendentes. O mapa SVG é uma projeção de pontos, sem mapa-base de limites territoriais. O acumulado anual respeita o intervalo selecionado.
- A cadeia de notificações e os gatilhos estão preparados para testes, com a ativação institucional pendente de autorização. Cotas, concorrência real e objetivos de latência não foram medidos. Dependendo do resultado, será necessário separar o lote de OCR do processamento de alertas ou usar índices persistentes para consultas mais frequentes.
