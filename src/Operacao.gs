function gpsAtividadesBase_(base, hoje) {
  const atividades = [['visitas', 'visita'], ['apoios', 'apoio']].flatMap(([aba, tipo]) =>
    gpsTabela_(base, aba).registros.filter(r => gpsDataValida_(r.data)).map(r => ({ ...r, tipo })));
  const evidencias = gpsTabela_(base, 'comprovantes').registros.concat(
    gpsTabela_(base, 'termos_web').registros.map(r => ({ ...r, validado: true })));
  const decisoes = gpsTabela_(base, 'decisoes_visita').registros;
  return gpsCruzarVisitas_(gpsAplicarDecisoesVisita_(atividades, decisoes), evidencias, hoje);
}

function gpsPrioridadesBase_(base, ano) {
  const versao = PropertiesService.getScriptProperties().getProperty('GPS_LAB_CRITERIO_VERSAO');
  const regras = gpsTabela_(base, 'criterios').registros.filter(r => r.versao === versao && r.aprovado_cec === true);
  if (!versao || !regras.length) return [];
  return gpsSelecionarPrioridades_(gpsTabela_(base, 'indicadores').registros, { versao, regras, aprovadaCEC: true }, ano);
}

function gpsAvisosFormulario(dados) {
  const { base, perfil } = gpsSessao_();
  const plano = gpsValidarAtividadeWeb_(base, perfil, dados);
  const prioridades = gpsPrioridadesBase_(base, Number(plano.data.slice(0, 4)))
    .filter(r => r.escola_id === plano.escola_id);
  const alinhadas = prioridades.filter(r => r.componente_id === plano.componente_id && (!r.turma_id || r.turma_id === plano.turma_id));
  return { prioridades: prioridades.length, alinhadas: alinhadas.length };
}

function gpsValidarAtividadeWeb_(base, perfil, dados) {
  const { escola_id, componente_id, turma_id, pasta_id, data, tipo } = dados || {};
  if (perfil.papel === 'AT' || !perfil.pec_id || !gpsDataValida_(data) || !['visita', 'apoio'].includes(tipo) ||
      !escola_id || !componente_id || !pasta_id) throw new Error('Preencha os campos obrigatórios');
  const ano = gpsTabela_(base, 'anos_letivos').registros.find(r => Number(r.ano) === Number(data.slice(0, 4)));
  const atribuida = ano && gpsTabela_(base, 'atribuicoes').registros.some(r =>
    r.pec_id === perfil.pec_id && r.escola_id === escola_id && r.pasta_id === pasta_id && r.ano_letivo_id === ano.id);
  if (!atribuida) throw new Error('Atividade fora da atribuição');
  if (!gpsTabela_(base, 'componentes').registros.some(r => r.id === componente_id)) throw new Error('Componente inválido');
  if (turma_id && !gpsTabela_(base, 'turmas').registros.some(r =>
    r.id === turma_id && r.escola_id === escola_id && r.ano_letivo_id === ano.id)) throw new Error('Turma fora da escola ou do ano');
  return { escola_id, componente_id, turma_id, pasta_id, data, tipo, pec_id: perfil.pec_id };
}

function gpsTextoWeb_(valor) {
  const texto = typeof valor === 'string' ? valor.trim() : '';
  if (!texto || texto.length > 1000) throw new Error('Informe um texto de até mil caracteres');
  return texto;
}

function gpsSalvarPlano(dados) {
  return gpsComTrava_(() => {
    const { base, perfil } = gpsSessao_();
    const r = gpsValidarAtividadeWeb_(base, perfil, dados);
    const horario = dados.horario || '';
    if (horario && !/^([01]\d|2[0-3]):[0-5]\d$/.test(horario)) throw new Error('Horário inválido');
    const id = gpsChave_(r.tipo, r.pec_id, r.data, r.escola_id, horario);
    const plano_id = gpsChave_('plano_web', id);
    const tabela = r.tipo === 'visita' ? 'visitas' : 'apoios';
    const atividade = { ...r, id, plano_id, horario, status: 'agendada' };
    gpsAnexar_(base, 'planos', [{ id: plano_id, pec_id: r.pec_id, periodo: r.data.slice(0, 7), objetivo: gpsTextoWeb_(dados.objetivo) }]);
    gpsAnexar_(base, tabela, [atividade]);
    gpsAuditar_(base, tabela, id, 'registro_web', perfil.id);
    return { id };
  });
}

function gpsContextoTermo(id) {
  const { base, perfil } = gpsSessao_();
  if (perfil.papel === 'AT') throw new Error('Perfil de consulta');
  const visita = gpsTabela_(base, 'visitas').registros.find(r => r.id === id && r.pec_id === perfil.pec_id);
  if (!visita) throw new Error('Visita indisponível');
  const avisos = gpsPrioridadesBase_(base, Number(visita.data.slice(0, 4))).filter(r =>
    r.escola_id === visita.escola_id && (r.componente_id !== visita.componente_id ||
      (r.turma_id && r.turma_id !== visita.turma_id)));
  return { divergencias: avisos.length };
}

function gpsSalvarTermo(dados) {
  return gpsComTrava_(() => {
    const { base, perfil } = gpsSessao_();
    if (perfil.papel === 'AT') throw new Error('Perfil de consulta');
    const visita = gpsTabela_(base, 'visitas').registros.find(r => r.id === dados.visita_id && r.pec_id === perfil.pec_id);
    const confirmadaNaoRealizada = gpsTabela_(base, 'decisoes_visita').registros.some(r =>
      r.visita_id === dados.visita_id && r.decisao === 'nao_ocorreu');
    if (!visita || visita.data > gpsHoje_() || visita.status === 'cancelada' || confirmadaNaoRealizada) throw new Error('Visita indisponível');
    if (!['termo_visita', 'ficha_visita'].includes(dados.tipo)) throw new Error('Tipo de evidência inválido');
    const id = gpsChave_('termo_web', visita.id, dados.tipo);
    gpsAnexar_(base, 'termos_web', [{ id, visita_id: visita.id, pec_id: perfil.pec_id,
      tipo: dados.tipo, observacao: gpsTextoWeb_(dados.observacao) }]);
    gpsAuditar_(base, 'termos_web', id, 'registro_web', perfil.id);
    if (dados.tipo === 'termo_visita') gpsResolverAlertaTermo_(base, visita);
    return { id };
  });
}

function gpsMinhasOpcoes() {
  const { base, perfil } = gpsSessao_();
  if (perfil.papel === 'AT') return { papel: 'AT', escolas: [], pastas: [], componentes: [], turmas: [], visitas: [] };
  const atribuicoes = gpsTabela_(base, 'atribuicoes').registros.filter(r => r.pec_id === perfil.pec_id);
  const opcoes = (tabela, ids) => gpsTabela_(base, tabela).registros.filter(r => !ids || ids.has(r.id))
    .map(r => ({ id: r.id, nome: r.nome || r.codigo || r.id }));
  const escolas = new Set(atribuicoes.map(r => r.escola_id));
  const escolasOpcoes = opcoes('escolas', escolas);
  const nomesEscolas = new Map(escolasOpcoes.map(r => [r.id, r.nome]));
  return {
    papel: perfil.papel, escolas: escolasOpcoes,
    pastas: opcoes('pastas', new Set(atribuicoes.map(r => r.pasta_id))),
    componentes: opcoes('componentes'),
    turmas: gpsTabela_(base, 'turmas').registros.filter(r => escolas.has(r.escola_id))
      .map(r => ({ id: r.id, nome: r.codigo || r.id, escola_id: r.escola_id })),
    visitas: gpsTabela_(base, 'visitas').registros.filter(r => r.pec_id === perfil.pec_id &&
      r.status !== 'cancelada' && r.data <= gpsHoje_()).slice(-100)
      .map(r => ({ id: r.id, nome: r.data + ' · ' + (nomesEscolas.get(r.escola_id) || 'Escola sem cadastro na atribuição') }))
  };
}

function gpsRevisarIndicador(dados) {
  return gpsComTrava_(() => {
    const { base, perfil } = gpsSessao_();
    if (perfil.papel !== 'CEC') throw new Error('Acesso restrito ao CEC');
    const item = gpsTabela_(base, 'indicadores').registros.find(r => r.id === dados.id);
    if (!item || item.revisar !== true) throw new Error('Indicador indisponível para revisão');
    const valor = dados.valor;
    if (valor !== null && (typeof valor !== 'number' || !Number.isFinite(valor))) throw new Error('Valor inválido');
    gpsAuditar_(base, 'indicadores', item.id, 'validado:' + gpsChave_(valor), perfil.id);
    gpsAtualizarEstado_(base, 'indicadores', item, { ...item, valor, revisar: false });
    return { id: item.id };
  });
}


function gpsAplicarDecisoesVisita_(atividades, decisoes) {
  const porVisita = new Map(decisoes.map(r => [r.visita_id, r.decisao]));
  return atividades.map(r => {
    if (r.tipo !== 'visita') return r;
    const decisao = porVisita.get(r.id);
    return decisao === 'nao_ocorreu' ? { ...r, status: 'cancelada', decisao } : decisao ? { ...r, decisao } : r;
  });
}

function gpsRegistrarDecisaoVisita(dados) {
  return gpsComTrava_(() => {
    const { base, perfil } = gpsSessao_();
    if (perfil.papel !== 'PEC' || !['termo_atrasado', 'nao_ocorreu'].includes(dados?.decisao)) {
      throw new Error('Decisão de visita indisponível');
    }
    const hoje = gpsHoje_();
    const visita = gpsAtividadesBase_(base, hoje).find(r => r.tipo === 'visita' &&
      r.id === dados.visita_id && r.pec_id === perfil.pec_id);
    if (!visita) throw new Error('Visita indisponível para esta decisão.');
    const id = gpsChave_('decisao_visita', visita.id, dados.decisao);
    const existente = gpsTabela_(base, 'decisoes_visita').registros.find(r => r.id === id);
    if (existente) {
      gpsAuditar_(base, 'decisoes_visita', id, dados.decisao, perfil.id);
      if (dados.decisao === 'nao_ocorreu') gpsResolverAlertaTermo_(base, visita);
      return { id };
    }
    if (!visita.aguardandoTermo || (Date.parse(hoje) - Date.parse(visita.data)) < 7 * 86400000 ||
        (dados.decisao === 'nao_ocorreu' && visita.comprovada)) {
      throw new Error('Visita indisponível para esta decisão. Confirme a data e as evidências.');
    }
    const motivo = dados.decisao === 'nao_ocorreu' ? gpsTextoWeb_(dados.motivo) : '';
    gpsAnexar_(base, 'decisoes_visita', [{ id, visita_id: visita.id, pec_id: perfil.pec_id,
      decisao: dados.decisao, motivo, instante: new Date().toISOString() }]);
    gpsAuditar_(base, 'decisoes_visita', id, dados.decisao, perfil.id);
    if (dados.decisao === 'nao_ocorreu') gpsResolverAlertaTermo_(base, visita);
    return { id };
  });
}
