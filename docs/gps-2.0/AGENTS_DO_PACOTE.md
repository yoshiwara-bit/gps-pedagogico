# Instruções ao Codex — GPS Pedagógico 2.0

Leia os cinco documentos em `docs/` antes de alterar código. São especificações propostas para validação, não evidência de que a arquitetura já funciona.

Primeira tarefa: analisar o repositório existente em modo de leitura; inventariar módulos, dependências, fontes, cálculos, layout, regras, gargalos e testes. Entregar relatório de lacunas e proposta de arquitetura com alternativas, riscos, cotas, custos, permissões, modelo de dados e plano de migração. Não alterar `main` nem publicar sem aprovação. Trabalhar em branch experimental `gps-2-apps-script` após autorização.

Priorizar Apps Script + Drive + Sheets, mantendo Python apenas se necessário. Validar acesso institucional antes de depender de serviços avançados ou CLASP. Implementar primeiro prova de conceito de ingestão documental, depois interface. Separar sincronização da navegação. Não confundir Codex com serviço de execução em produção.

Preservar regras RN01–RN15, identidade visual e funcionalidades atuais; divergências devem ser apresentadas ao CEC. Não inventar indicadores, limiares ou campos. Não colocar termos reais, dados pessoais, credenciais ou tokens no GitHub ou em logs. Usar dados sintéticos/anonimizados. Não contornar restrições administrativas.

Cada alteração deve identificar requisitos atendidos, testes executados, limitações e instruções de reversão. Não afirmar que testes passaram sem evidência. Solicitar aprovação antes de mudanças irreversíveis, permissões, custos ou publicação.
