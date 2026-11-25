const assert = require('node:assert/strict');
const { readFileSync, readdirSync } = require('node:fs');
const { test } = require('node:test');
const vm = require('node:vm');

const arquivos = readdirSync('src').filter(nome => nome.endsWith('.gs'));
const contexto = vm.createContext({});
for (const nome of arquivos) {
  new vm.Script(readFileSync('src/' + nome, 'utf8'), { filename: nome }).runInContext(contexto);
}
const gps = contexto;
const plain = valor => JSON.parse(JSON.stringify(valor));

test('boletins de avaliação ficam fora dos termos de visita', () => {
  assert.equal(gps.gpsClassificarDocumento_('Boletim SAEB\nTermo de visita citado').tipo, 'saeb');
  assert.equal(gps.gpsClassificarDocumento_('Boletim SARESP').tipo, 'saresp');
  assert.equal(gps.gpsClassificarDocumento_('SAEB e SARESP').revisar, true);
  assert.equal(gps.gpsClassificarDocumento_('Termo detalhado de visita').tipo, 'termo_visita');
});

test('OCR mantém ordem e encaminha imageId e valores ambíguos para revisão', () => {
  const imagens = gps.gpsLerImagensSAEB_([0, 1, 2], item => {
    if (item === 1) throw new Error('Argumento inválido: imageId');
    return 'Texto sintético ' + item;
  });
  assert.deepEqual(plain(imagens.map(item => item.ordem)), [1, 2, 3]);
  assert.equal(imagens[1].motivo, 'image_id_invalido');
  assert.equal(imagens[2].revisar, true);
  assert.deepEqual(plain(gps.gpsValorAvaliacao_('102,5%', 'percentual')),
    { original: '102,5%', valor: 102.5, revisar: true });
  assert.equal(gps.gpsValorAvaliacao_('1.234', 'nota').valor, null);
  assert.equal(gps.gpsValorAvaliacao_('—', 'nota').valor, null);
});

test('importa Prova Paulista consolidada por turma com cabeçalho variável', () => {
  const linhas = [
    [], ['Resultados'],
    ['Turma', 'Resultado', 'Componente', 'Série', 'Escola', 'Ano'],
    ['T1', '62,5%', 'LP', 'EF8', 'E1', '2026']
  ];
  const entrada = { fonte: 'prova_paulista', documento_id: 'D', metrica: 'acerto',
    escala: 'pp_percentual', unidade: 'percentual', periodo: '3º bimestre' };
  const resultado = gps.gpsImportarIndicadores_(linhas, entrada);
  assert.equal(resultado.indicadores.length, 1);
  assert.equal(resultado.indicadores[0].turma_id, 'T1');
  assert.equal(resultado.indicadores[0].periodo, '3º bimestre');
  assert.equal(resultado.indicadores[0].valor, 62.5);
  assert.equal(gps.gpsImportarIndicadores_([...linhas, linhas[3]], entrada).indicadores.length, 0);
});

test('preserva os três níveis de Avaliação Diagnóstica e rejeita nível desconhecido', () => {
  const niveis = ['Abaixo do Básico', 'Básico', 'Proficiente', 'Outro'];
  const linhas = [
    ['Ano', 'Série', 'Componente', 'Escola', 'Valor', 'Nível'],
    ...niveis.map(nivel => ['2026', 'EF7', 'LP', 'E1', '30%', nivel])
  ];
  const resultado = gps.gpsImportarIndicadores_(linhas, {
    fonte: 'avd', documento_id: 'D', metrica: 'distribuicao', escala: 'avd_v1', unidade: 'percentual'
  });
  assert.deepEqual(plain(resultado.indicadores.map(item => item.nivel)), niveis.slice(0, 3));
  assert.equal(resultado.revisoes.length, 1);
});

test('coortes avançam entre etapas e preservam o período original', () => {
  assert.equal(gps.gpsProjetarCoorte_({ fonte: 'saresp', ano: 2025, serie: 'EF7' }, 2026).serieAcompanhamento, 'EF8');
  const saeb = { fonte: 'saeb', ano: 2025, serie: 'EF9' };
  assert.equal(gps.gpsProjetarCoorte_(saeb, 2026).serieAcompanhamento, 'EM1');
  assert.equal(gps.gpsProjetarCoorte_(saeb, 2027).serieAcompanhamento, 'EM2');
  assert.equal(gps.gpsProjetarCoorte_(saeb, 2027).serieAplicacao, 'EF9');
  assert.equal(gps.gpsProjetarCoorte_(saeb, 2029).situacao, 'conclusao_prevista');
  assert.equal(gps.gpsProjetarCoorte_({ ...saeb, ano: 2024 }, 2026), null);
});

test('aprovação CEC e regras por escala precedem o cruzamento com a turma atual', () => {
  const base = { componente_id: 'LP', metrica: 'nota' };
  const historico = { ...base, id: 'H', fonte: 'saresp', ano: 2025, serie: 'EF7', valor: 3, escala: 'saresp_X' };
  const atual = { ...base, id: 'A', fonte: 'prova_paulista', ano: 2026, serie: 'EF8',
    escola_id: 'E2', turma_id: 'T2', valor: 50, escala: 'pp_Y' };
  const regras = [historico, atual].map(({ fonte, escala, metrica }) =>
    ({ fonte, escala, metrica, direcao: 'abaixo', limite: fonte === 'saresp' ? 4 : 60 }));
  assert.throws(() => gps.gpsSelecionarPrioridades_([atual], { regras }, 2026), /CEC/);
  const prioridades = gps.gpsSelecionarPrioridades_([historico, atual, {...atual, escala: 'outra'}],
    { aprovadaCEC: true, versao: 'sintetica', regras }, 2026);
  assert.equal(prioridades.length, 1);
  assert.equal(prioridades[0].turma_id, 'T2');
  assert.deepEqual(plain(prioridades[0].referencias_historicas), ['H']);
});

test('visitas separam escolas e PECs, deduplicam agenda e conciliam ficha e termo', () => {
  const linha = { tipo: 'visita', pec_id: 'P1', escolas: ['E1', 'E2'], data: '2026-08-10' };
  const agenda = gps.gpsConsolidarAgenda_([linha, linha, { ...linha, tabelaConteiner: true },
    { ...linha, pec_id: 'P2', escolas: ['E1'] }, { ...linha, data: '2026-08-11', escolas: ['E1'] }]);
  assert.equal(agenda.atividades.length, 4);
  const [primeira, segunda] = agenda.atividades;
  const documentos = [
    { visita_id: primeira.id, tipo: 'ficha_visita', validado: true },
    { visita_id: primeira.id, tipo: 'termo_visita', validado: true },
    { visita_id: segunda.id, tipo: 'ficha_visita', validado: true }
  ];
  const visitas = gps.gpsCruzarVisitas_(agenda.atividades, documentos, '2026-08-10');
  assert.equal(visitas[0].pendenciaDocumental, false);
  assert.equal(visitas[1].comprovada, true);
  assert.equal(visitas[1].aguardandoTermo, true);
  assert.equal(visitas[1].pendenciaDocumental, true);
  assert.equal(visitas[2].situacao, 'aguardando_comprovacao');
  assert.equal(visitas[3].contabilizada, false);
});

test('metas vencidas distinguem obrigação por escola de recomendação', () => {
  const metas = vm.runInContext('GPS_METAS_INICIAIS', contexto);
  const periodo = { inicio: '2026-08-01', fim: '2026-08-15', hoje: '2026-08-16' };
  for (const pasta of ['Qualidade da Aula', 'Recomposição']) {
    const regra = metas.find(item => item.pasta === pasta);
    assert.equal(gps.gpsAvaliarMeta_(regra, [], 'P', periodo, 'E').pendente, regra.obrigatoria);
    assert.throws(() => gps.gpsAvaliarMeta_(regra, [], 'P', periodo), /Escopo/);
  }
});

test('monitoramento distingue divergência e falta de comprovação', () => {
  const prioridade = { id: 'P', escola_id: 'E', ano: 2026, turma_id: 'T', componente_id: 'LP' };
  const visita = { escola_id: 'E', data: '2026-05-05', turma_id: 'T', componente_id: 'MAT', contabilizada: true };
  assert.equal(gps.gpsMonitorarPrioridades_([prioridade], [visita])[0].direcionamento, 'divergente');
  assert.equal(gps.gpsMonitorarPrioridades_([prioridade], [{...visita, componente_id: 'LP'}])[0].evidencias,
    'sem_evidencia_documental');
});

test('fila deduplica por versão, limita lote e isola falhas', () => {
  const entrada = { origem: 'D', chave: 'conteudo', versao: '1' };
  const fila = gps.gpsAtualizarFila_([], [entrada, entrada, { ...entrada, versao: '2' }]);
  assert.equal(fila.length, 2);
  const resultado = gps.gpsProcessarLote_(fila, 1, () => { throw new Error('falha sintética'); }, 1);
  assert.deepEqual(plain(resultado.map(item => item.status)), ['erro', 'pendente']);
  assert.deepEqual(plain(gps.gpsCaminhoTermo_('2026-08-01', 'P1', 'E1')),
    ['Termos de Visita', '2026', '08', 'P1', 'E1']);
});

test('alertas respeitam recorrência e ficam silenciados após resolução', () => {
  const sinal = { tipo: 'pendencia', alvo_id: 'D', periodo: '2026-08', destinatario_id: 'P' };
  const primeira = gps.gpsConciliarAlertas_([], [sinal, sinal], '2026-08-01');
  assert.equal(primeira.notificacoes.length, 1);
  assert.equal(gps.gpsConciliarAlertas_(primeira.alertas, [sinal], '2026-08-02').notificacoes.length, 0);
  assert.equal(gps.gpsConciliarAlertas_(primeira.alertas, [sinal], '2026-08-08').notificacoes.length, 1);
  primeira.alertas[0].status = 'resolvido';
  assert.equal(gps.gpsConciliarAlertas_(primeira.alertas, [sinal], '2026-09-01').notificacoes.length, 0);
});

test('validação detecta referências ausentes e alterações conflitantes', () => {
  const indice = { escolas: new Map([['E', { id: 'E', nome: 'Escola sintética' }]]) };
  assert.equal(gps.gpsValidarRegistro_('escolas', { id: 'E', nome: 'Escola sintética' }, indice), false);
  assert.throws(() => gps.gpsValidarRegistro_('escolas', { id: 'E', nome: 'Outra' }, indice), /Conflito/);
  assert.throws(() => gps.gpsValidarRegistro_('escolas', { id: 'E2', municipio_id: 'M' }, indice), /Referência/);
});

test('escopo PEC deixa de fora terceiros e tabelas estratégicas', () => {
  const perfil = { papel: 'PEC', pec_id: 'P1' };
  assert.equal(gps.gpsPodeVer_(perfil, 'visitas', { pec_id: 'P1' }), true);
  assert.equal(gps.gpsPodeVer_(perfil, 'visitas', { pec_id: 'P2' }), false);
  assert.equal(gps.gpsPodeVer_(perfil, 'prioridades', {}), false);
  assert.equal(gps.gpsPodeVer_({ papel: 'CEC' }, 'prioridades', {}), true);
});

test('carga local de 500 entradas e 50 mil registros com 20 e 40 consultas simuladas', async () => {
  const historico = Array.from({ length: 50_000 }, (_, i) => ({
    id: 'V' + i, pec_id: 'P' + (i % 40), escola_id: 'E' + (i % 40),
    data: '2026-08-01', tipo: 'visita', contabilizada: true
  }));
  const entradas = Array.from({ length: 500 }, (_, i) => ({ chave: 'K' + i, versao: '1', origem: 'D' + i }));
  let fila = gps.gpsAtualizarFila_([], entradas);
  for (let i = 0; i < 10; i++) fila = gps.gpsProcessarLote_(fila, 50, () => ({ revisar: false }));
  assert.equal(fila.filter(item => item.status === 'concluido').length, 500);
  for (const volume of [20, 40]) {
    const consultas = Array.from({ length: volume }, (_, i) => new Promise(resolve => setImmediate(() => {
      resolve(gps.gpsAvaliarMeta_({ atividade: 'visita', quantidade: 1, obrigatoria: true },
        historico, 'P' + i, { inicio: '2026-08-01', fim: '2026-08-15', hoje: '2026-08-16' }));
    })));
    assert.ok((await Promise.all(consultas)).every(item => item.planejamentoContabilizado === 1250));
  }
});

test('bloqueia uma grade com cabeçalho de estudante antes de importar', () => {
  const resultado = gps.gpsImportarIndicadores_([['Nome do estudante', 'Nota']],
    { fonte: 'prova_paulista', documento_id: 'D', escala: 'PP', metrica: 'nota' });
  assert.equal(resultado.indicadores.length, 0);
  assert.equal(resultado.revisoes[0].motivo, 'dados_individuais_bloqueados');
});

test('preparação da base é idempotente e recusa cabeçalhos incompatíveis', () => {
  const abas = new Map();
  const planilha = {
    getSheetByName: nome => abas.get(nome),
    insertSheet(nome) {
      const aba = {
        cabecalho: [],
        getLastRow() { return this.cabecalho.length ? 1 : 0; },
        getLastColumn() { return this.cabecalho.length; },
        getRange() { return { getValues: () => [this.cabecalho], setValues: ([linha]) => { this.cabecalho = linha; } }; },
        setFrozenRows() {}
      };
      abas.set(nome, aba);
      return aba;
    }
  };
  gps.gpsCriarEstrutura_(planilha);
  const escolas = abas.get('escolas');
  gps.gpsCriarEstrutura_(planilha);
  assert.equal(abas.get('escolas'), escolas);
  assert.equal(escolas.cabecalho[0], 'id');
  escolas.cabecalho[0] = 'outra_coluna';
  assert.throws(() => gps.gpsCriarEstrutura_(planilha), /Cabeçalho incompatível/);
});

test('leitura PDF limita a origem e envia apenas temporários criados à lixeira', () => {
  let pasta = 'fora';
  const lixeira = [];
  const origem = {
    getParents: () => ({ hasNext: (() => { let i = 0; return () => i++ === 0; })(),
      next: () => ({ getId: () => pasta }) }),
    getMimeType: () => 'application/pdf',
    getBlob: () => 'pdf'
  };
  gps.PropertiesService = { getScriptProperties: () => ({ getProperty: () => 'lab' }) };
  gps.DriveApp = { getFileById: id => id === 'original' ? origem :
    ({ setTrashed: valor => { assert.equal(valor, true); lixeira.push(id); } }) };
  gps.Drive = { Files: { create: (_, blob) => ({ id: 'temporario-' + blob }) } };
  gps.DocumentApp = { openById: id => ({ getBody: () => ({
    getText: () => id.endsWith('pdf') ? 'Boletim SAEB' : 'OCR sintético',
    getTables: () => [],
    getImages: () => [{ getBlob: () => 'imagem' }, { getBlob() { throw new Error('imageId'); } }]
  }) }) };
  assert.throws(() => gps.gpsLerArquivoLaboratorio_('original'), /fora da pasta/);
  pasta = 'lab';
  const doc = gps.gpsLerArquivoLaboratorio_('original');
  assert.equal(doc.classificacao.tipo, 'saeb');
  assert.equal(doc.imagens[0].texto, 'OCR sintético');
  assert.equal(doc.imagens[1].motivo, 'image_id_invalido');
  assert.deepEqual(lixeira, ['temporario-imagem', 'temporario-pdf']);
});

test('histórico por escola diferencia persistência, emergência e ausência de referência escolar', () => {
  const base = { componente_id: 'LP', metrica: 'nota' };
  const historicos = [2, 7].map((valor, i) => ({ ...base, id: 'H' + i, fonte: 'saresp', ano: 2025,
    serie: 'EF7', escola_id: 'E' + i, valor, escala: 's' }));
  const atuais = [0, 1, 2].map(i => ({ ...base, id: 'A' + i, fonte: 'prova_paulista', ano: 2026,
    serie: 'EF8', escola_id: 'E' + i, turma_id: 'T' + i, valor: 40, escala: 'p' }));
  const regras = [
    { fonte: 'saresp', metrica: 'nota', escala: 's', direcao: 'abaixo', limite: 4 },
    { fonte: 'prova_paulista', metrica: 'nota', escala: 'p', direcao: 'abaixo', limite: 60 }
  ];
  const saida = plain(gps.gpsSelecionarPrioridades_([...historicos, ...atuais], { aprovadaCEC: true, versao: 'v1', regras }, 2026));
  assert.deepEqual(saida.map(r => [r.diagnostico, r.referencias_escola]), [
    ['persistente', ['H0']], ['emergente', ['H1']], ['sem_historico_escolar', []]
  ]);
  assert.deepEqual(saida[2].referencias_historicas, ['H0']);
});

test('adaptadores leem grade em ordem variável e isolam OCR e SAEB par', () => {
  const cabecalho = ['Componente', 'Etapa', 'Código da Escola', 'Ano', 'Participação', 'Proficiência média'];
  const saeb = gps.gpsExtrairBoletim_({ grades: [[cabecalho,
    ['LP', '1ª série do Ensino Médio', 'E', '2025', '102,5%', '260,5'],
    ['MAT', '9º ano', 'E', '2024', '90%', '245']]], imagens: [{ revisar: true }] }, 'saeb', 'D');
  assert.deepEqual(plain(saeb.indicadores.map(r => [r.serie, r.metrica, r.valor, r.revisar])),
    [['EM1', 'proficiencia', 260.5, false], ['EM1', 'participacao', 102.5, true]]);
  assert.ok(saeb.revisoes.some(r => r.motivo === 'ocr_requer_conferencia'));
  const saresp = gps.gpsExtrairBoletim_({ grades: [[['Série', 'Código SEDUC', 'Disciplina', 'Edição', 'Nota média'],
    ['7º ano', 'E', 'LP', '2024', '4,5']]] }, 'saresp', 'S');
  assert.equal(saresp.indicadores[0].escala, 'saresp:nota:2024');
});

test('exportação alternativa mantém revisão humana e não troca imagens de ordem ambígua', () => {
  const imagens = [{ ordem: 1, revisar: true, motivo: 'image_id_invalido' }, { ordem: 2, texto: 'anterior' }];
  gps.Drive = { Files: { export: () => 'zip' } };
  const imagem = { getName: () => 'images/img.png' };
  const html = src => ({ getName: () => 'index.html', getDataAsString: () => src });
  gps.Utilities = { unzip: () => [html('<img src="images/img.png"><img src="images/img.png">'), imagem] };
  assert.equal(gps.gpsRecuperarImagensExportadas_('D', imagens, 'lab'), imagens);
  gps.Utilities.unzip = () => [html('<img src="images/img.png"><img src="images/outro.png">'), imagem];
  const original = gps.gpsOCRBlob_;
  gps.gpsOCRBlob_ = blob => { assert.equal(blob, imagem); return ' recuperada '; };
  try {
    const resultado = gps.gpsRecuperarImagensExportadas_('D', imagens, 'lab');
    assert.equal(resultado[0].texto, 'recuperada');
    assert.equal(resultado[0].revisar, true);
    assert.equal(resultado[1].texto, 'anterior');
  } finally {
    gps.gpsOCRBlob_ = original;
  }
});

test('janela incremental mantém limite entre páginas e exclui temporários', () => {
  const limite = '2026-11-25T10:00:00.000Z';
  const inicio = '2026-11-24T10:00:00.000Z';
  const primeira = gps.gpsDescobrirPagina_('folder', { desde: inicio }, limite, consulta => {
    assert.match(consulta.q, /modifiedTime >= '2026-11-24/);
    assert.match(consulta.q, /modifiedTime < '2026-11-25/);
    return { nextPageToken: 'p2', files: [
      { id: 'origem', version: '1' }, { id: 'temp', version: '1', appProperties: { gpsTemporario: 'true' } }
    ] };
  });
  assert.equal(primeira.entradas.length, 1);
  const segunda = gps.gpsDescobrirPagina_('folder', primeira.proximo, '2026-11-25T10:05:00.000Z', consulta => {
    assert.equal(consulta.pageToken, 'p2');
    assert.ok(consulta.q.includes(limite));
    return { files: [] };
  });
  assert.deepEqual(plain(segunda.proximo), { desde: limite });
  assert.throws(() => gps.gpsDescobrirPagina_('folder', {}, limite, () => ({ incompleteSearch: true })), /incompleta/);
});

test('agregação territorial separa situações e aplica filtros', () => {
  const escolas = [{ id: '1', nome: 'Sintética', latitude: '', longitude: -47, municipio_id: 'M' }];
  const agenda = [
    { escola_id: '1', pasta_id: 'A', data: '2026-11-04', tipo: 'visita', contabilizada: true, comprovada: true },
    { escola_id: '1', pasta_id: 'A', data: '2026-11-24', tipo: 'visita' },
    { escola_id: '1', pasta_id: 'B', data: '2026-11-24', tipo: 'apoio' }
  ];
  const filtro = { municipio_id: 'M', pasta_id: 'A', inicio: '2026-11-01', fim: '2026-11-30', granularidade: 'quinzenal' };
  const resultado = gps.gpsAgregarTerritorio_(escolas, agenda, filtro);
  assert.deepEqual(plain(resultado.resumo), { previstas: 2, contabilizadas: 1, comprovadas: 1 });
  assert.deepEqual(plain(resultado.periodos.map(r => r.periodo)), ['2026-11 / 1', '2026-11 / 2']);
  assert.equal(resultado.pontos[0].latitude, null);
  assert.equal(gps.gpsFaixaGrafico_('2026-11-01', 'semanal'), '2026-10-26');
  assert.equal(gps.gpsFaixaGrafico_('2026-11-01', 'ano'), '2026');
});

test('metas usam semana e quinzena já encerradas inclusive na virada do mês', () => {
  assert.deepEqual(plain(gps.gpsPeriodoFechado_('2026-03-01', 'quinzena')),
    { inicio: '2026-02-16', fim: '2026-02-28', hoje: '2026-03-01' });
  assert.deepEqual(plain(gps.gpsPeriodoFechado_('2026-11-25', 'semana')),
    { inicio: '2026-11-16', fim: '2026-11-22', hoje: '2026-11-25' });
});

test('cadastro territorial descarta endereço e recusa coordenadas ausentes', () => {
  const campos = ['codigo_seduc', 'escola', 'municipio', 'latitude', 'longitude', 'logradouro'];
  const escola = ['11', 'Escola sintética', 'Município sintético', '-24,2', '-47,4', 'endereço de teste'];
  const resultado = gps.gpsCadastroTerritorial_([campos, escola]);
  assert.equal(resultado.escolas[0].latitude, -24.2);
  assert.equal(resultado.escolas[0].longitude, -47.4);
  assert.equal(resultado.escolas[0].logradouro, undefined);
  assert.throws(() => gps.gpsCadastroTerritorial_([campos, escola, escola]), /duplicado/);
  assert.throws(() => gps.gpsCadastroTerritorial_([campos, ['12', 'E', 'M', '', '-47']]), /ausente/);
});

test('alertas persistidos não repetem e-mail reservado após falha ambígua', () => {
  const nomes = ['gpsTabela_', 'gpsSinaisAtuacao_', 'gpsAtualizarEstado_', 'gpsAuditar_'];
  const preservar = Object.fromEntries(nomes.map(nome => [nome, gps[nome]]));
  const planilhaAnterior = gps.SpreadsheetApp;
  const registros = [];
  gps.SpreadsheetApp = { flush() {} };
  gps.gpsTabela_ = (_, tabela) => ({ registros: tabela === 'usuarios'
    ? [{ id: 'U', ativo: true, perfil: 'PEC', pec_id: 'P', email: 'teste@example.invalid' }]
    : registros });
  gps.gpsSinaisAtuacao_ = () => [{ tipo: 'termo_pendente', alvo_id: 'V', periodo: '2026-11-01', destinatario_id: 'P' }];
  gps.gpsAtualizarEstado_ = (_, nome, anterior, item) => {
    assert.equal(nome, 'alertas');
    if (anterior) registros[anterior._linha - 2] = { ...item, _linha: anterior._linha };
    else registros.push({ ...item, _linha: registros.length + 2 });
  };
  gps.gpsAuditar_ = () => {};
  const tentativas = [];
  try {
    const emissor = endereco => { tentativas.push(endereco); throw new Error('Falha sintética após reserva'); };
    gps.gpsExecutarAlertas_({}, '2026-11-01', emissor);
    gps.gpsExecutarAlertas_({}, '2026-11-02', emissor);
    gps.gpsExecutarAlertas_({}, '2026-11-10', emissor);
    assert.deepEqual(tentativas, ['teste@example.invalid']);
    assert.equal(registros[0].envio, 'reservado');
  } finally {
    Object.assign(gps, preservar, { SpreadsheetApp: planilhaAnterior });
  }
});
