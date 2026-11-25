const GPS_SINCRONIZACAO = 'GPS_LAB_CHECKPOINT_DRIVE';

function gpsDescobrirPagina_(pasta, estado, agora, listar) {
  if (!/^[\w-]+$/.test(pasta)) throw new Error('Pasta de laboratório inválida');
  const ate = estado.ate || agora;
  const q = ["'" + pasta + "' in parents", 'trashed = false', "modifiedTime < '" + ate + "'"];
  if (estado.desde) q.push("modifiedTime >= '" + estado.desde + "'");
  const pagina = listar({
    q: q.join(' and '), orderBy: 'modifiedTime', pageSize: 25, pageToken: estado.pagina || undefined,
    fields: 'nextPageToken,incompleteSearch,files(id,version,md5Checksum,mimeType,appProperties)'
  });
  if (pagina.incompleteSearch) throw new Error('Listagem incompleta');
  const entradas = (pagina.files || []).filter(item => item.version &&
    !item.appProperties?.gpsTemporario &&
    !['application/vnd.google-apps.folder', 'application/vnd.google-apps.shortcut'].includes(item.mimeType));
  return { entradas, proximo: pagina.nextPageToken
    ? { desde: estado.desde, ate, pagina: pagina.nextPageToken }
    : { desde: ate } };
}

function gpsImportarItemFila_(base, item) {
  const atual = Drive.Files.get(item.origem, { fields: 'version,md5Checksum' });
  const chave = atual.md5Checksum || item.origem;
  const idAtual = gpsChave_(chave, atual.md5Checksum ? 'conteudo' : String(atual.version));
  if (idAtual !== item.id) return 'substituido';
  const doc = gpsLerArquivoLaboratorio_(item.origem);
  const depois = Drive.Files.get(item.origem, { fields: 'version,md5Checksum' });
  if (String(atual.version) !== String(depois.version)) return 'substituido';
  const fonte = gpsTabela_(base, 'fontes_importacao').registros.filter(r => r.origem === item.origem);
  const contexto = fonte.length === 1 ? fonte[0] : null;
  const identificado = doc.classificacao?.tipo;
  const tipo = identificado || contexto?.fonte || 'indeterminado';
  let avaliacao = { indicadores: [], revisoes: [{ motivo: 'leiaute_pendente' }] };
  if (['saresp', 'saeb'].includes(identificado)) {
    avaliacao = gpsExtrairBoletim_(doc, identificado, item.id);
  } else if (contexto && !identificado) {
    const partes = (doc.grades || []).map(grade => gpsImportarIndicadores_(grade, { ...contexto, documento_id: item.id }));
    const indicadores = partes.flatMap(r => r.indicadores);
    const grupos = new Map();
    for (const r of indicadores) grupos.set(r.id, (grupos.get(r.id) || 0) + 1);
    const revisoes = partes.flatMap(r => r.revisoes);
    if ([...grupos.values()].some(total => total > 1)) revisoes.push({ motivo: 'indicador_duplicado_entre_abas' });
    avaliacao = { indicadores: indicadores.filter(r => grupos.get(r.id) === 1), revisoes };
  }
  gpsAnexar_(base, 'documentos', [{ id: item.id, chave, versao: item.versao, tipo, origem: item.origem }]);
  const escolas = gpsTabela_(base, 'escolas').registros;
  const componentes = gpsTabela_(base, 'componentes').registros;
  const turmas = gpsTabela_(base, 'turmas').registros;
  const anos = new Map(gpsTabela_(base, 'anos_letivos').registros.map(r => [r.id, Number(r.ano)]));
  const resolver = (registros, token, alias) => {
    const encontrados = registros.filter(r => String(r.id) === String(token) ||
      (r[alias] && gpsNormalizar_(r[alias]) === gpsNormalizar_(token)));
    return encontrados.length === 1 ? encontrados[0].id : null;
  };
  const validos = [];
  for (const r of avaliacao.indicadores) {
    const escola_id = resolver(escolas, r.escola_id, 'codigo');
    const componente_id = resolver(componentes, r.componente_id, 'nome');
    const turma_id = r.turma_id ? resolver(turmas.filter(t => t.escola_id === escola_id &&
      anos.get(t.ano_letivo_id) === r.ano && t.serie === r.serie), r.turma_id, 'codigo') : null;
    if (escola_id && componente_id && (!r.turma_id || turma_id)) validos.push({ ...r, escola_id, componente_id, turma_id });
  }
  if (validos.length !== avaliacao.indicadores.length) avaliacao.revisoes.push({ motivo: 'cadastro_nao_encontrado' });
  const requerRevisao = Boolean(doc.requerRevisao || avaliacao.revisoes.length);
  gpsAnexar_(base, 'indicadores', validos.map(r => ({ ...r, revisar: requerRevisao })));
  if (requerRevisao) {
    const motivos = [...new Set(avaliacao.revisoes.map(r => r.motivo).concat(doc.requerRevisao ? ['conferir_conversao'] : []))];
    gpsAnexar_(base, 'revisoes', motivos.map(motivo => ({
      id: gpsChave_(item.id, motivo), documento_id: item.id, motivo, status: 'aberta'
    })));
  }
  return requerRevisao ? 'revisao' : 'concluido';
}

function gpsSincronizarLaboratorio_() {
  const props = PropertiesService.getScriptProperties();
  const pasta = props.getProperty('GPS_LAB_FOLDER_ID');
  if (gpsModoDemonstracao_() || props.getProperty('GPS_LAB_AUTOMACAO') !== 'true' || !pasta) return;
  gpsComTrava_(() => {
    const base = gpsAbrirBase_();
    const chave = GPS_SINCRONIZACAO + '_' + pasta;
    const estado = JSON.parse(props.getProperty(chave) || '{}');
    let pagina;
    try {
      pagina = gpsDescobrirPagina_(pasta, estado, new Date().toISOString(), args => Drive.Files.list(args));
    } catch (erro) {
      // Retoma a janela com deduplicação quando um token de paginação expira.
      if (!estado.pagina || !/page.?token|invalid value/i.test(String(erro.message))) throw erro;
      props.setProperty(chave, JSON.stringify({ desde: estado.desde, ate: estado.ate }));
      return;
    }
    const entradas = pagina.entradas.map(r => {
      const versao = r.md5Checksum ? 'conteudo' : String(r.version);
      return { id: gpsChave_(r.md5Checksum || r.id, versao), origem: r.id, versao };
    });
    const fila = gpsTabela_(base, 'fila').registros;
    const ids = new Set(fila.map(r => r.id));
    const novos = [];
    for (const entrada of entradas) {
      if (ids.has(entrada.id)) continue;
      ids.add(entrada.id);
      novos.push({ ...entrada, status: 'pendente', tentativas: 0 });
    }
    gpsAnexar_(base, 'fila', novos);
    SpreadsheetApp.flush();
    props.setProperty(chave, JSON.stringify(pagina.proximo));
    const pendentes = gpsTabela_(base, 'fila').registros.filter(r => ['pendente', 'repetir'].includes(r.status));
    for (const item of pendentes.slice(0, 2)) {
      const tentativas = Number(item.tentativas) + 1;
      let status;
      try {
        status = gpsImportarItemFila_(base, item);
      } catch {
        status = tentativas >= 3 ? 'erro' : 'repetir';
      }
      gpsAuditar_(base, 'fila', item.id, status + ':' + tentativas, 'automacao');
      gpsAtualizarEstado_(base, 'fila', item, { ...item, tentativas, status });
      SpreadsheetApp.flush();
    }
    const enviar = props.getProperty('GPS_LAB_EMAIL') === 'true'
      ? (endereco, assunto, corpo) => MailApp.sendEmail(endereco, assunto, corpo) : null;
    gpsExecutarAlertas_(base, gpsHoje_(), enviar);
  });
}

function gpsConfigurarGatilhoLaboratorio_() {
  const props = PropertiesService.getScriptProperties();
  if (gpsModoDemonstracao_() || props.getProperty('GPS_LAB_AUTOMACAO') !== 'true') throw new Error('Automação de laboratório desabilitada');
  gpsComTrava_(() => {
    const existe = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'gpsSincronizarLaboratorio_');
    if (!existe) ScriptApp.newTrigger('gpsSincronizarLaboratorio_').timeBased().everyMinutes(5).create();
  });
}

// Reabre um erro confirmado no editor do projeto de laboratório, preservando o histórico.
function gpsReabrirErroFilaLaboratorio_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('GPS_LAB_REPROCESSAR_ID');
  if (gpsModoDemonstracao_() || !id || props.getProperty('GPS_LAB_AUTOMACAO') !== 'true') throw new Error('Reprocessamento indisponível');
  gpsComTrava_(() => {
    const base = gpsAbrirBase_();
    const item = gpsTabela_(base, 'fila').registros.find(r => r.id === id && r.status === 'erro');
    if (!item) throw new Error('Erro da fila indisponível');
    gpsAuditar_(base, 'fila', item.id, 'reaberto:' + item.tentativas, 'operador');
    gpsAtualizarEstado_(base, 'fila', item, { ...item, status: 'repetir' });
    SpreadsheetApp.flush();
    props.deleteProperty('GPS_LAB_REPROCESSAR_ID');
  });
}
