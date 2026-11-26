# GPS Pedagógico 2.0 — guia de atualização da etapa 4

## Atualizar sem trocar a base ou o favorito

1. Faça uma cópia da planilha experimental no Google Drive. Em **Implantar → Gerenciar implantações**, anote a versão atual. Deixe a implantação existente ativa.
2. No editor do projeto experimental existente, substitua o conteúdo de **Código.gs**, **Painel.html** e **Atuacao.html** pelos do ZIP. Crie dois arquivos HTML pelo botão **+** com os nomes **Interface** e **Logos** e cole os conteúdos correspondentes. Se já existirem, atualize-os. O nome de cada HTML precisa coincidir com o nome do pacote.
3. Compare o manifesto. O ZIP documenta Drive v3 e V8. Se o manifesto atual já estiver com essas opções e contiver a seção `webapp` com `access: MYSELF`, preserve a configuração atual. Mantenha `GPS_LAB_DB_ID`, `GPS_LAB_FOLDER_ID` e o modo da instalação. Confira que `GPS_LAB_AUTOMACAO` e `GPS_LAB_EMAIL` estão desativados.
4. No fim de **Código.gs**, cole temporariamente o trecho abaixo. Execute a função **prepararBaseEtapa4** pelo seletor do editor. Ela só completa as abas e os cabeçalhos. Para preencher fichas, frequências e plataformas de exemplo em uma base **sintética**, use a segunda função. Ela verifica o modo demonstração e preserva os registros existentes. Exclua ambas as funções temporárias e salve antes de qualquer implantação.

```javascript
function prepararBaseEtapa4() {
  gpsPrepararBaseLaboratorio_();
}

function carregarExemplosEtapa4() {
  gpsAtualizarAmostraEtapa4_();
}
```

5. Em **Implantar → Testar implantações**, abra o `/dev`. Confira menu, busca, ficha, mapa e consultas. O `/dev` compartilha a base e as propriedades do projeto. Para testar gravações, use uma cópia isolada.
6. Só após validar, abra **Implantar → Gerenciar implantações**, selecione a implantação experimental atual, clique no lápis e escolha **Nova versão**. Mantenha o acesso restrito e confirme. Editar a mesma implantação preserva o endereço `/exec` salvo nos favoritos. Nenhuma atualização dessa implantação foi executada pelo Codex.

Para reverter o código publicado, edite a mesma implantação e escolha a versão anterior. A migração acrescenta abas e colunas. A cópia da planilha preserva o estado de antes da migração.

## Ensaio cartográfico separado e autorizado

- O modo demonstração mantém o mapa esquemático. Use um **projeto de laboratório autorizado separado** para coordenadas e cadastro institucionais. Fonte pública de referência para as divisas: [API de malhas do IBGE](https://servicodados.ibge.gov.br/api/docs/malhas?versao=3).
- Prepare um `FeatureCollection` GeoJSON em longitude e latitude. Ele deve ter seis feições `Polygon` ou `MultiPolygon`, uma para cada município, com `properties.municipio` ou `properties.name` identificando a unidade. Os anéis precisam estar fechados. Se necessário, use simplificação oficial ou validada. O limite de leitura é 1,5 MB e 12 mil vértices. Coloque o arquivo diretamente na pasta privada configurada para o laboratório.
- Configure privadamente `GPS_LAB_GEOJSON_ID` e `GPS_LAB_GEOJSON_FONTE`. Para a sede, configure `GPS_LAB_SEDE_LATITUDE`, `GPS_LAB_SEDE_LONGITUDE` e `GPS_LAB_SEDE_FONTE` a partir de dados conferidos. Sem fonte ou coordenadas, o marcador da sede fica ausente.
- Para cadastro, use `GPS_LAB_CADASTRO_CSV_ID` e a função privada `gpsCarregarCadastroTerritorial_` com o CSV autorizado, em uma base de laboratório própria. Para completar fichas, use `GPS_LAB_FICHAS_CSV_ID` e `gpsImportarFichasLaboratorio_` com CSV consolidado, código SEDUC e fonte por linha. Use o padrão do cadastro enriquecido legado como referência para os cabeçalhos. As duas operações são restritas à pasta privada e bloqueadas na demonstração.
- Preencha `frequencias` com início e fim, granularidade, recorte, presenças, denominador e referência. Preencha `plataformas` com código da plataforma, métrica, unidade, período e fonte. Registros de ano inteiro precisam vir consolidados pela fonte. Preserve a separação entre escola, segmento, série e turma. Mantenha a origem, a escala e a edição na aba de avaliações.

### Importação consolidada de medições

Em laboratório autorizado, os CSVs privados podem ser apontados pelas propriedades `GPS_LAB_FREQUENCIA_CSV_ID` e `GPS_LAB_PLATAFORMAS_CSV_ID`. Use um invólucro temporário no editor para chamar `gpsImportarMedicoesLaboratorio_('frequencias')` ou `gpsImportarMedicoesLaboratorio_('plataformas')`, remova-o depois e preserve o original no Drive.

Os cabeçalhos comuns são `escola_id,inicio,fim,granularidade,fonte`, com `turma_id` opcional. Frequência requer `presencas,denominador` e aceita `segmento,serie`. Plataformas requer `plataforma,metrica,unidade,valor` e aceita `denominador`. Datas devem estar em AAAA-MM-DD e a granularidade deve ser semanal, quinzenal, mensal, acumulado ou anual. A importação preserva as unidades e não atribui faróis automaticamente.
