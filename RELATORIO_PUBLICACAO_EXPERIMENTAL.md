# GPS Pedagógico 2.0 — validação para publicação experimental

## Entrega

O projeto em `src/` foi consolidado para instalação em um Apps Script isolado. A interface combina verde-petróleo, turquesa e branco, com menu responsivo de visão geral, território, consultas e registros. Os cartões apresentam escolas no escopo e atividades previstas, contabilizadas e comprovadas. Consultas incluem faróis para notas SARESP revistas e na escala correspondente, com rótulos textuais e cores.

| Situação | Funcionalidades |
| --- | --- |
| Operacionais no piloto | Cadastro sintético, navegação, consultas por perfil, filtros, mapa SVG, gráficos, mapa de calor temporal, registros de plano e evidência e justificativa de alertas. |
| Experimentais | Projeções e faróis pedagógicos, revisão de candidatos extraídos dos boletins e dados de exemplo. |
| Em desenvolvimento | Homologação de boletins reais, recuperação de imagens com `imageId` e validação institucional de cotas, concorrência e permissões. Automação e e-mail ficam bloqueados no modo demonstração. |

O instalador cria uma planilha e uma pasta privadas e preenche a planilha com 40 escolas fictícias. Nenhum dado de usuário é gerado. As contas autorizadas são configuradas diretamente na planilha privada. CEC consulta a base e valida indicadores, PEC opera dentro das atribuições, e AT consulta apenas as escolas vinculadas e suas atividades. Uma identidade vazia, inativa ou duplicada é recusada pelo servidor.

## Testes executados

| Verificação | Resultado |
| --- | --- |
| `node --test tests/*.test.cjs /workspace/.gps-2-dev/gps-lab.test.cjs` | 35 testes passaram, incluindo 30 da reconstrução e cinco do laboratório original. Os casos novos cobrem massa sintética, segurança de perfis, faróis e bloqueio de automação. |
| `node tests/interface-smoke.cjs` | Chromium e Playwright, com RPC simulada. CEC em 1280 px, PEC e AT em 390 px. Foram exercitados menu, recortes do mapa, tabela de consultas e farol, pop-up e registro PEC e ausência de edição para AT. |
| Sintaxe e empacotamento | Validação dos arquivos .gs, JavaScript embutido, manifesto e script consolidado de instalação. |
| Preservação | Diff dos arquivos do GPS oficial e de `GPS_LAB.gs` conferido em relação ao início da missão. |

O ensaio de interface detectou um pop-up que herdava a ocultação da seção de registros. O diálogo foi movido para fora da seção e a navegação voltou a passar nos três perfis.

## Validação no Workspace

A instalação, o consentimento, a disponibilidade da identidade ativa e a emissão da URL dependem da conta autorizada e da política do Workspace. A URL `/exec` aparece na confirmação de implantação da Web App. Nenhum projeto Google foi criado, nenhum e-mail foi enviado, nenhum gatilho foi ativado e nenhuma implantação foi feita durante esta preparação. O guia orienta a instalação, a obtenção da URL e o teste de acesso com as contas autorizadas. [Guia de instalação](GUIA_INSTALACAO_EXPERIMENTAL.md)
