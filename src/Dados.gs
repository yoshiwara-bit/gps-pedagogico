const GPS_SCHEMA = {
  municipios: ['nome'], escolas: ['municipio_id', 'codigo', 'nome'],
  pecs: ['nome'], pastas: ['nome'], anos_letivos: ['ano'],
  atribuicoes: ['pec_id', 'pasta_id', 'escola_id', 'ano_letivo_id'],
  turmas: ['escola_id', 'ano_letivo_id', 'serie', 'codigo'],
  componentes: ['nome'], planos: ['pec_id', 'documento_id', 'periodo'],
  visitas: ['plano_id', 'pec_id', 'escola_id', 'turma_id', 'componente_id', 'data', 'status'],
  apoios: ['plano_id', 'pec_id', 'escola_id', 'turma_id', 'componente_id', 'data', 'status'],
  documentos: ['chave', 'tipo', 'versao', 'origem'],
  comprovantes: ['visita_id', 'documento_id', 'tipo'],
  indicadores: ['escola_id', 'municipio_id', 'turma_id', 'componente_id', 'ano', 'serie', 'fonte', 'periodo', 'metrica', 'escala', 'nivel', 'valor_original', 'valor', 'revisar', 'documento_id'],
  metas: ['pasta_id', 'atividade', 'periodicidade', 'quantidade', 'por_escola', 'obrigatoria'],
  prioridades: ['escola_id', 'turma_id', 'componente_id', 'ano', 'serie', 'criterio_versao', 'indicador_id', 'referencias_historicas'],
  alertas: ['prioridade_id', 'visita_id', 'tipo', 'status', 'criticidade', 'destinatario_id', 'periodo', 'alvo_id', 'ultima_notificacao', 'ocorrencias'],
  justificativas: ['visita_id', 'motivo', 'status'],
  encaminhamentos: ['alerta_id', 'acao', 'prazo', 'status'],
  revisoes: ['documento_id', 'motivo', 'status'],
  usuarios: ['email', 'perfil', 'pec_id', 'ativo'],
  fila: ['documento_id', 'origem', 'status', 'tentativas'],
  criterios: ['fonte', 'metrica', 'escala', 'nivel', 'direcao', 'limite', 'versao', 'aprovado_cec'],
  auditoria: ['entidade', 'registro_id', 'operacao', 'instante']
};

const GPS_RELACOES = {
  municipio_id: 'municipios', escola_id: 'escolas', pec_id: 'pecs', pasta_id: 'pastas',
  ano_letivo_id: 'anos_letivos', turma_id: 'turmas', componente_id: 'componentes',
  documento_id: 'documentos', plano_id: 'planos', visita_id: 'visitas',
  prioridade_id: 'prioridades', alerta_id: 'alertas', indicador_id: 'indicadores'
};

function gpsChave_(...partes) {
  return JSON.stringify(partes);
}

function gpsValidarRegistro_(tabela, registro, indice) {
  if (!Object.prototype.hasOwnProperty.call(GPS_SCHEMA, tabela) || !registro.id) {
    throw new Error('Tabela ou ID inválido');
  }
  const anterior = indice[tabela]?.get(registro.id);
  const chaves = ['id', ...GPS_SCHEMA[tabela]];
  if (anterior && chaves.some(campo => (anterior[campo] ?? '') !== (registro[campo] ?? ''))) {
    throw new Error('Conflito para o mesmo ID');
  }
  if (tabela === 'documentos' && registro.chave &&
      Array.from(indice.documentos?.values() ?? []).some(item =>
        item.id !== registro.id && item.chave === registro.chave && item.versao === registro.versao)) {
    throw new Error('Documento duplicado');
  }
  for (const campo of GPS_SCHEMA[tabela]) {
    const destino = GPS_RELACOES[campo];
    if (destino && registro[campo] && !indice[destino]?.has(registro[campo])) {
      throw new Error('Referência ausente: ' + campo);
    }
  }
  return !anterior;
}

// Execute manualmente apenas no projeto de laboratório, com a planilha isolada configurada.
function gpsPrepararBaseLaboratorio_() {
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_DB_ID');
  if (!id) throw new Error('Configure GPS_LAB_DB_ID para uma planilha experimental.');
  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    gpsCriarEstrutura_(SpreadsheetApp.openById(id));
    SpreadsheetApp.flush();
  } finally {
    trava.releaseLock();
  }
}

function gpsCriarEstrutura_(planilha) {
  for (const [nome, campos] of Object.entries(GPS_SCHEMA)) {
    const aba = planilha.getSheetByName(nome);
    if (!aba || !aba.getLastRow()) continue;
    const esperados = ['id', ...campos];
    const atuais = aba.getRange(1, 1, 1, esperados.length).getValues()[0];
    if (atuais.some((campo, i) => campo !== esperados[i]) || aba.getLastColumn() !== esperados.length) {
      throw new Error('Cabeçalho incompatível: ' + nome);
    }
  }
  for (const [nome, campos] of Object.entries(GPS_SCHEMA)) {
    const aba = planilha.getSheetByName(nome) || planilha.insertSheet(nome);
    if (aba.getLastRow()) continue;
    const cabecalho = ['id', ...campos];
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.setFrozenRows(1);
  }
}
