function doGet() {
  let acesso;
  try {
    acesso = gpsSessao_();
  } catch {
    return HtmlService.createHtmlOutput('<main><h1>Acesso restrito</h1><p>Entre com uma conta autorizada para a demonstração. Se o problema persistir, peça ao responsável que revise a instalação e seu acesso.</p></main>')
      .setTitle('GPS Pedagógico 2.0').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const pagina = HtmlService.createTemplateFromFile('Painel');
  pagina.papel = acesso.perfil.papel;
  pagina.demonstracao = gpsModoDemonstracao_();
  return pagina.evaluate().setTitle('GPS Pedagógico 2.0').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const GPS_PAINEL = {
  indicadores: { titulo: 'Indicadores', campos: ['id', 'fonte', 'ano', 'escola_id', 'serie', 'metrica', 'escala', 'valor_original', 'valor', 'farol', 'revisar'] },
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
  if (perfil.papel === 'AT') {
    if (tabela === 'escolas') return perfil.escolas?.has(registro.id) === true;
    return ['visitas', 'apoios'].includes(tabela) && perfil.escolas?.has(registro.escola_id) === true;
  }
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
  if (perfil.papel === 'AT') perfil.escolas = gpsEscopoEscolas_(planilha, perfil);
  const aba = planilha.getSheetByName(tabela);
  const ultimaLinha = aba.getLastRow();
  const inicio = cursor + 2;
  const quantidade = Math.max(0, Math.min(100, ultimaLinha - inicio + 1));
  const colunas = ['id', ...GPS_SCHEMA[tabela]];
  const campos = GPS_PAINEL[tabela].campos;
  const intervalo = quantidade ? aba.getRange(inicio, 1, quantidade, colunas.length) : null;
  const brutos = tabela === 'indicadores' && intervalo ? intervalo.getValues() : [];
  const valores = intervalo ? intervalo.getDisplayValues() : [];
  const linhas = valores.map((linha, j) => ({
    ...Object.fromEntries(colunas.map((campo, i) => [campo, linha[i]])),
    farolValidado: brutos[j]?.[colunas.indexOf('revisar')] === false
  }))
    .filter(registro => registro.id && gpsPodeVer_(perfil, tabela, registro))
    .map(registro => campos.map(campo => campo === 'farol' ? gpsFarolSaresp_(registro) : registro[campo]));
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
      !['CEC', 'PEC', 'AT'].includes(usuario.perfil) || (usuario.perfil === 'PEC' && !usuario.pec_id)) {
    throw new Error('Acesso ao laboratório indisponível');
  }
  return { base, perfil: { papel: usuario.perfil, pec_id: usuario.perfil === 'AT' ? '' : usuario.pec_id, id: usuario.id } };
}

function gpsHoje_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function gpsIncluir_(arquivo) {
  return HtmlService.createHtmlOutputFromFile(arquivo).getContent();
}


function gpsEscopoEscolas_(base, perfil) {
  if (perfil.papel === 'CEC') return null;
  const vinculos = perfil.papel === 'AT'
    ? gpsTabela_(base, 'escopos_at').registros.filter(r => r.usuario_id === perfil.id)
    : gpsTabela_(base, 'atribuicoes').registros.filter(r => r.pec_id === perfil.pec_id);
  return new Set(vinculos.map(r => r.escola_id));
}

function gpsFarolSaresp_(registro) {
  if (registro.fonte !== 'saresp' || registro.metrica !== 'nota' ||
      registro.escala !== 'saresp:nota:' + registro.ano || !registro.farolValidado || !String(registro.valor ?? '').trim()) return '—';
  const valor = Number(String(registro.valor).replace(',', '.'));
  if (!Number.isFinite(valor) || valor < 0 || valor > 10) return '—';
  return valor <= 4 ? 'Vermelho' : valor <= 6 ? 'Amarelo' : 'Verde';
}
