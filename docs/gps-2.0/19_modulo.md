# 19 — Integridade de dados

## Escopo
Checklist esperado/recebido/processado/validado por escola, indicador, período e elegibilidade. Exclusivo CEC e AT.

## Regras de negócio
Hash e chave semântica, versões, cobertura de arquivos consolidados, erros e revisão. Drive/GPS mesma fila; processamento incremental, locks, checkpoint, retry, auditoria. Polling de cinco minutos como proposta sujeita a quotas. Nunca apagar originais.

## Critérios mínimos de aceite
Arquivo repetido não duplica dados; cobertura parcial sinalizada; PEC não acessa central administrativa.

## Implementação
Inspecionar código existente; registrar dependências, migrações, testes e pendências. Não declarar concluído sem evidência.
