// Execute uma vez no editor do projeto isolado. Os IDs ficam em propriedades privadas.
function gpsInstalarDemonstracao_() {
  gpsComTrava_(() => {
    const props = PropertiesService.getScriptProperties();
    const modo = props.getProperty('GPS_LAB_MODO');
    if (modo !== 'demonstracao' && (props.getProperty('GPS_LAB_DB_ID') || props.getProperty('GPS_LAB_FOLDER_ID'))) {
      throw new Error('Use um projeto novo e isolado para instalar a demonstração.');
    }
    props.setProperties({ GPS_LAB_MODO: 'demonstracao', GPS_LAB_AUTOMACAO: 'false', GPS_LAB_EMAIL: 'false' });
    let data = props.getProperty('GPS_LAB_DEMO_DATA');
    if (!data) {
      data = gpsHoje_();
      props.setProperty('GPS_LAB_DEMO_DATA', data);
    }
    if (!props.getProperty('GPS_LAB_DB_ID')) {
      props.setProperty('GPS_LAB_DB_ID', SpreadsheetApp.create('GPS 2.0 · Base sintética').getId());
    }
    if (!props.getProperty('GPS_LAB_FOLDER_ID')) {
      props.setProperty('GPS_LAB_FOLDER_ID', DriveApp.createFolder('GPS 2.0 · Pasta sintética').getId());
    }
    const base = gpsAbrirBase_();
    gpsCriarEstrutura_(base);
    for (const [tabela, linhas] of Object.entries(gpsDadosDemonstracao_(data))) {
      const existentes = new Set(gpsTabela_(base, tabela).registros.map(r => r.id));
      gpsAnexar_(base, tabela, linhas.filter(r => !existentes.has(r.id)));
    }
    props.setProperty('GPS_LAB_CRITERIO_VERSAO', 'demo-v1');
  });
}

function gpsDadosDemonstracao_(data) {
  if (!gpsDataValida_(data)) throw new Error('Data de demonstração inválida');
  const ano = Number(data.slice(0, 4));
  const anoId = 'DEMO-ANO';
  const pasta = 'DEMO-PASTA';
  const componente = 'DEMO-LP';
  const anterior = deslocamento => {
    const dia = new Date(data + 'T00:00:00Z');
    dia.setUTCDate(dia.getUTCDate() + deslocamento);
    return dia.toISOString().slice(0, 10);
  };
  const municipios = Array.from({ length: 4 }, (_, i) => ({
    id: 'DEMO-M' + (i + 1), nome: 'Município fictício ' + (i + 1)
  }));
  const escolas = Array.from({ length: 40 }, (_, i) => ({
    id: 'DEMO-E' + (i + 1), codigo: 'DEMO-' + (i + 1), nome: 'Escola fictícia ' + (i + 1),
    municipio_id: municipios[i % 4].id,
    latitude: -10 - Math.floor(i / 8) * 0.08 - (i % 3) * 0.014,
    longitude: -39 + (i % 8) * 0.1
  }));
  const pecs = [{ id: 'DEMO-PEC', nome: 'PEC de demonstração' }, { id: 'DEMO-PEC2', nome: 'Outro PEC fictício' }];
  const atribuicoes = escolas.map((escola, i) => ({
    id: 'DEMO-A' + (i + 1), pec_id: pecs[i % 2].id, pasta_id: pasta,
    escola_id: escola.id, ano_letivo_id: anoId
  }));
  const turmas = escolas.map((escola, i) => ({
    id: 'DEMO-T' + (i + 1), escola_id: escola.id, ano_letivo_id: anoId, codigo: '8A', serie: 'EF8'
  }));
  const visitas = escolas.slice(0, 12).map((escola, i) => ({
    id: 'DEMO-V' + (i + 1), escola_id: escola.id, pec_id: pecs[i % 2].id,
    pasta_id: pasta, componente_id: componente, turma_id: turmas[i].id,
    data: anterior((i - 6) * 3), status: 'agendada'
  }));
  const escalaAtual = 'demo-pp';
  const escalaHistorica = 'saresp:nota:' + (ano - 1);
  const avaliacoes = escolas.slice(0, 12).flatMap((escola, i) => [{
    id: 'DEMO-IH' + i, escola_id: escola.id, componente_id: componente, ano: ano - 1,
    serie: 'EF7', fonte: 'saresp', metrica: 'nota', escala: escalaHistorica, valor: 2 + i % 6, revisar: false
  }, {
    id: 'DEMO-IA' + i, escola_id: escola.id, turma_id: turmas[i].id, componente_id: componente, ano,
    serie: 'EF8', fonte: 'prova_paulista', metrica: 'acerto', escala: escalaAtual, valor: 45 + i * 3, revisar: false
  }]);
  return {
    municipios, escolas, pecs, pastas: [{ id: pasta, nome: 'Qualidade da Aula · demonstração' }],
    anos_letivos: [{ id: anoId, ano }], atribuicoes,
    escopos_at: escolas.slice(0, 8).map((escola, i) => ({ id: 'DEMO-AT-E' + i, usuario_id: 'DEMO-AT', escola_id: escola.id })),
    turmas, componentes: [{ id: componente, nome: 'Língua Portuguesa' }],
    visitas,
    apoios: visitas.slice(0, 4).map((r, i) => ({ ...r, id: 'DEMO-AP' + i })),
    termos_web: visitas.slice(0, 3).map((r, i) => ({
      id: 'DEMO-TERMO' + i, visita_id: r.id, pec_id: r.pec_id, tipo: 'termo_visita', observacao: 'Evidência fictícia.'
    })),
    indicadores: avaliacoes,
    criterios: [
      { id: 'DEMO-C1', fonte: 'saresp', metrica: 'nota', escala: escalaHistorica, direcao: 'abaixo', limite: 4, versao: 'demo-v1', aprovado_cec: true },
      { id: 'DEMO-C2', fonte: 'prova_paulista', metrica: 'acerto', escala: escalaAtual, direcao: 'abaixo', limite: 60, versao: 'demo-v1', aprovado_cec: true }
    ],
    metas: [{ id: 'DEMO-META', pasta_id: pasta, atividade: 'visita', periodicidade: 'quinzena', quantidade: 1, por_escola: true, obrigatoria: true }],
    alertas: [{ id: 'DEMO-ALERTA', tipo: 'termo_pendente', criticidade: 'media',
      destinatario_id: 'DEMO-PEC', periodo: data, status: 'aberto', envio: 'popup' }]
  };
}

function gpsModoDemonstracao_() {
  return PropertiesService.getScriptProperties().getProperty('GPS_LAB_MODO') === 'demonstracao';
}
