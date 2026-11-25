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
