# GPS Pedagógico 2.0 — instalar e obter o endereço da demonstração

A demonstração usa um projeto Apps Script novo, uma planilha privada e 40 escolas fictícias. Separe uma conta Google Workspace autorizada pela organização para administrar o projeto. Use somente contas de teste autorizadas para experimentar os perfis.

## 1. Baixe e abra os arquivos

Descompacte o pacote de instalação. Ele contém `Codigo.gs`, `Painel.html`, `Atuacao.html` e `appsscript.json`. Abra os arquivos com o Bloco de Notas ou outro editor de texto. Também é possível obter o código pela branch `gps-2-apps-script`. Nesse caso, copie os arquivos da pasta `src/` separadamente, criando um arquivo de mesmo nome no Apps Script para cada arquivo da pasta.

## 2. Crie um projeto Apps Script isolado

1. Abra [Apps Script](https://script.google.com/) com sua conta autorizada e clique em **Novo projeto**.
2. Dê o nome **GPS Pedagógico 2.0 — Demonstração**.
3. Abra o arquivo **Código.gs** criado pelo editor. Apague o exemplo e cole todo o conteúdo de `Codigo.gs` do pacote.
4. No painel **Arquivos**, clique no **+**, escolha **HTML**, digite **Painel** e cole o conteúdo de `Painel.html`. Repita para **Atuacao**. Digite os nomes sem a extensão, que o editor acrescenta.
5. Clique na engrenagem **Configurações do projeto**. Marque **Mostrar o arquivo de manifesto “appsscript.json” no editor**. Volte ao **Editor**, abra `appsscript.json`, substitua seu conteúdo pelo manifesto do pacote e salve. Ele define o runtime V8, o fuso de São Paulo e o serviço avançado Drive v3. A preparação usa o projeto padrão gerenciado pelo Apps Script.
6. Salve todos os arquivos. Confira que o editor contém o script, os dois HTML e o manifesto. Se usar os arquivos da branch em separado, mantenha só uma cópia de cada módulo e remova o exemplo padrão vazio.

## 3. Crie a base sintética

1. Na barra do Editor, selecione a função **gpsInstalarDemonstracao_** e clique em **Executar**.
2. Quando o Google solicitar permissões, leia a tela de consentimento. Autorize com a conta responsável pelo projeto recém-criado, de acordo com a política da sua organização. Se a organização bloquear a autorização, peça ao administrador do Workspace para avaliar o projeto experimental.
3. Aguarde **Execução concluída**. Em **Meu Drive**, procure **GPS 2.0 · Base sintética** e **GPS 2.0 · Pasta sintética**. Se houver arquivos com o mesmo nome, use os IDs mostrados nas propriedades privadas do projeto para encontrar os arquivos criados nesta instalação.
4. Em **Configurações do projeto → Propriedades do script**, confira que `GPS_LAB_MODO` está como `demonstracao` e que `GPS_LAB_AUTOMACAO` e `GPS_LAB_EMAIL` estão como `false`. `GPS_LAB_DB_ID` e `GPS_LAB_FOLDER_ID` são preenchidos pelo instalador. Mantenha essas propriedades privadas e preserve os valores gerados.
5. Confira o menu **Acionadores**. O projeto experimental deve ficar sem acionadores instalados. A demonstração já contém dados para o mapa, gráficos, prioridades, visitas, apoios e alertas de exemplo.

## 4. Autorize as contas de teste

Abra a planilha privada e, na aba **usuarios**, preencha as colunas conforme a tabela. Digite os endereços das contas Google autorizadas **nessa planilha privada**. Digite `=TRUE()` na coluna **ativo** para registrar um valor booleano verdadeiro. Cadastre um endereço por linha e use endereços distintos para testar os três perfis.

| id | email | perfil | pec_id | ativo |
| --- | --- | --- | --- | --- |
| DEMO-CEC | Endereço autorizado da conta CEC | CEC | Deixe vazio | `=TRUE()` |
| DEMO-LOGIN-PEC | Endereço autorizado da conta PEC | PEC | DEMO-PEC | `=TRUE()` |
| DEMO-AT | Endereço autorizado da conta AT | AT | Deixe vazio | `=TRUE()` |

- **CEC** consulta toda a base sintética, vê faróis e prioridades e pode validar indicadores.
- **PEC** consulta a própria atuação e registra planos e evidências para as escolas atribuídas. Há um alerta sintético de exemplo.
- **AT** tem consulta das escolas vinculadas em **escopos_at** e das atividades dessas escolas. O ID `DEMO-AT` já tem oito vínculos de exemplo.

Ao testar com uma única conta, comece pelo CEC e mantenha uma só linha ativa para esse endereço. Para validar o isolamento entre perfis, use contas distintas. Mantenha a planilha e a pasta privadas. A lista de usuários do aplicativo fica na planilha.

## 5. Publique a versão experimental e copie o endereço

1. No projeto isolado do Apps Script, clique em **Implantar → Nova implantação**.
2. Ao lado de **Selecionar tipo**, clique na engrenagem e escolha **Aplicativo da Web**.
3. Na descrição, escreva **GPS 2.0 — demonstração sintética**.
4. Em **Executar como**, selecione **Eu** (a conta responsável pela implantação). Assim o aplicativo usa a base privada do responsável. Para um teste individual, em **Quem pode acessar**, selecione **Somente eu**. Para o piloto com contas autorizadas do mesmo Workspace, escolha a opção restrita **Qualquer pessoa em [nome da organização]**, se oferecida. O código ainda exige uma conta ativa e única cadastrada em **usuarios**. Se a opção do domínio estiver indisponível, mantenha a demonstração individual e consulte o administrador.
5. Clique em **Implantar** e, se solicitado, conclua a autorização. Na confirmação, em **Aplicativo da Web**, clique em **Copiar** ao lado de **URL do aplicativo da Web**.
6. O endereço copiado termina em **`/exec`**. Abra-o com a conta autorizada. Esse é o endereço da demonstração. Para localizá-lo depois, use **Implantar → Gerenciar implantações**, selecione a implantação e copie a URL. O endereço **`/dev`** em **Testar implantações** é destinado aos editores do projeto.
7. Para atualizar a demonstração usando a mesma URL, abra **Implantar → Gerenciar implantações**, edite essa implantação, selecione **Nova versão** e clique em **Implantar**.

Use a opção restrita durante todo o piloto. Se uma conta cadastrada receber **Acesso restrito**, confira o perfil, o valor booleano em **ativo**, a ausência de e-mails duplicados, a conta logada e a opção de acesso. Em algumas configurações o Workspace pode ocultar a identidade ativa na Web App. Nesse caso o sistema recusa o acesso. Mantenha a restrição e peça ao administrador para confirmar a política de identidade do piloto.

## 6. Confira a experiência antes de compartilhar o link

1. Com CEC, confirme **Demonstração · dados sintéticos**, 40 escolas na visão geral, mapa e filtros em **Território**, e faróis em **Consultas → Indicadores**.
2. Com PEC, teste **Registros e alertas** e registre uma atividade com uma escola fictícia atribuída. Use somente texto sintético. Confira o recorte próprio e a evidência de uma visita disponível no formulário.
3. Com AT, verifique o recorte das oito escolas de exemplo e o menu de consulta. Com uma conta do domínio ausente de **usuarios**, verifique que a aplicação retorna **Acesso restrito**.
4. Compartilhe somente o endereço `/exec` da demonstração com as contas do piloto após essas verificações.
