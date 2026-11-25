# 22 — Segurança e desempenho

## Escopo
Perfis CEC, PEC e AT; autorização no servidor em cada operação, escopo por escola e componente.

## Regras de negócio
QA escolas fixas; DC todas com componente próprio; REC escolas atribuídas e AvD todas; EE/Multiplica/Conviva indicadores conforme atribuição; AT administrativo autorizado. CEC visão integral. LGPD, logs sem CPF, segredos em propriedades privadas, sem dados reais no GitHub público. Cache, batch, locks e auditoria.

## Critérios mínimos de aceite
Testar 20 usuários simultâneos, stress 40, 500 documentos/dia e 50 mil registros sintéticos. Metas propostas p95 navegação 500ms e inicial 5s; medir ambiente real, não prometer.

## Implementação
Inspecionar código existente; registrar dependências, migrações, testes e pendências. Não declarar concluído sem evidência.
