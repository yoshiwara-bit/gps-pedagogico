function doGet() {
  return HtmlService.createTemplateFromFile('Painel').evaluate().setTitle('GPS Pedagógico 2.0');
}

const GPS_PAINEL = {
  indicadores: { titulo: 'Indicadores', campos: ['id', 'fonte', 'ano', 'escola_id', 'serie', 'metrica', 'escala', 'valor_original', 'valor', 'revisar'] },
  revisoes: { titulo: 'Revisões documentais', campos: ['documento_id', 'motivo', 'status'] },
  escolas: { titulo: 'Escolas', campos: ['nome', 'codigo', 'municipio_id'] },
  municipios: { titulo: 'Municípios', campos: ['nome'] },
  pecs: { titulo: 'PECs', campos: ['nome'] },
  prioridades: { titulo: 'Prioridades', campos: ['escola_id', 'serie', 'turma_id', 'componente_id', 'ano', 'diagnostico', 'direcionamento', 'evidencias'] },
  alertas: { titulo: 'Alertas', campos: ['tipo', 'criticidade', 'status'] },
  visitas: { titulo: 'Visitas', campos: ['data', 'escola_id', 'status'] },
  apoios: { titulo: 'Apoios', campos: ['data', 'escola_id', 'status'] },
  metas: { titulo: 'Metas', campos: ['atividade', 'periodicidade', 'quantidade', 'obrigatoria'] }
};

function gpsPodeVer_(perfil, tabela, registro) {
  if (perfil.papel === 'CEC') return true;
  if (perfil.papel !== 'PEC' || !perfil.pec_id) return false;
  if (['visitas', 'apoios'].includes(tabela)) return registro.pec_id === perfil.pec_id;
  if (tabela === 'alertas') return registro.destinatario_id === perfil.pec_id;
  return false;
}

function gpsCarregarPainel(tabela, cursor) {
  if (!Object.prototype.hasOwnProperty.call(GPS_PAINEL, tabela) ||
      !Number.isSafeInteger(cursor) || cursor < 0) throw new Error('Consulta inválida');
  const { base: planilha, perfil } = gpsSessao_();
  if (tabela === 'prioridades') {
    const { titulo, campos } = GPS_PAINEL.prioridades;
    if (perfil.papel !== 'CEC') return { titulo, campos, linhas: [], proximo: null };
    const hoje = gpsHoje_();
    const prioridades = gpsPrioridadesBase_(planilha, Number(hoje.slice(0, 4)));
    const acompanhamento = new Map(gpsMonitorarPrioridades_(prioridades, gpsAtividadesBase_(planilha, hoje))
      .map(r => [r.prioridade_id, r]));
    const linhas = prioridades.slice(cursor, cursor + 100).map(r => {
      const dados = { ...r, ...acompanhamento.get(r.id) };
      return campos.map(campo => dados[campo] ?? '');
    });
    return { titulo, campos, linhas, proximo: cursor + 100 < prioridades.length ? cursor + 100 : null };
  }
  const aba = planilha.getSheetByName(tabela);
  const ultimaLinha = aba.getLastRow();
  const inicio = cursor + 2;
  const quantidade = Math.max(0, Math.min(100, ultimaLinha - inicio + 1));
  const colunas = ['id', ...GPS_SCHEMA[tabela]];
  const campos = GPS_PAINEL[tabela].campos;
  const valores = quantidade ? aba.getRange(inicio, 1, quantidade, colunas.length).getDisplayValues() : [];
  const linhas = valores.map(linha => Object.fromEntries(colunas.map((campo, i) => [campo, linha[i]])))
    .filter(registro => registro.id && gpsPodeVer_(perfil, tabela, registro))
    .map(registro => campos.map(campo => registro[campo]));
  return {
    titulo: GPS_PAINEL[tabela].titulo, campos, linhas,
    proximo: inicio + quantidade <= ultimaLinha ? cursor + quantidade : null
  };
}

function gpsSessao_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!email) throw new Error('Acesso ao laboratório indisponível');
  const base = gpsAbrirBase_();
  const candidatos = gpsTabela_(base, 'usuarios').registros.filter(item => String(item.email).trim().toLowerCase() === email);
  const usuario = candidatos[0];
  if (candidatos.length !== 1 || usuario.ativo !== true ||
      !['CEC', 'PEC'].includes(usuario.perfil) || (usuario.perfil === 'PEC' && !usuario.pec_id)) {
    throw new Error('Acesso ao laboratório indisponível');
  }
  return { base, perfil: { papel: usuario.perfil, pec_id: usuario.pec_id, id: usuario.id } };
}

function gpsHoje_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function gpsIncluir_(arquivo) {
  return HtmlService.createHtmlOutputFromFile(arquivo).getContent();
}
