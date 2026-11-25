function gpsSinaisAtuacao_(base, hoje) {
  const ano = Number(hoje.slice(0, 4));
  const atividades = gpsAtividadesBase_(base, hoje);
  const prioridades = gpsPrioridadesBase_(base, ano);
  const porId = new Map(prioridades.map(r => [r.id, r]));
  const anoId = gpsTabela_(base, 'anos_letivos').registros.find(r => Number(r.ano) === ano)?.id;
  const atribuicoes = gpsTabela_(base, 'atribuicoes').registros.filter(r => r.ano_letivo_id === anoId);
  const sinais = [];
  function sinal(tipo, alvo_id, periodo, destinatario_id, criticidade = 'media') {
    sinais.push({ tipo, alvo_id, periodo, destinatario_id, criticidade });
  }
  for (const resultado of gpsMonitorarPrioridades_(prioridades, atividades)) {
    if (resultado.direcionamento === 'alinhado' && resultado.evidencias !== 'sem_evidencia_documental') continue;
    const prioridade = porId.get(resultado.prioridade_id);
    const responsaveis = atribuicoes.filter(item => item.escola_id === prioridade.escola_id);
    for (const pec of responsaveis.length ? responsaveis.map(r => r.pec_id) : ['cec']) {
      sinal(resultado.evidencias === 'sem_evidencia_documental' ? 'prioridade_sem_evidencia' : 'prioridade_nao_contemplada',
        prioridade.id, String(ano), pec, 'alta');
    }
  }
  for (const r of atividades.filter(item => item.pendenciaDocumental)) sinal('termo_pendente', r.id, r.data, r.pec_id);
  const alertas = new Map(gpsTabela_(base, 'alertas').registros.map(r => [r.id, r]));
  for (const r of gpsTabela_(base, 'encaminhamentos').registros) {
    const alerta = alertas.get(r.alerta_id);
    if (!alerta || !gpsDataValida_(r.prazo) || ['resolvido', 'concluido'].includes(r.status)) continue;
    const dias = (Date.parse(r.prazo) - Date.parse(hoje)) / 86400000;
    if (dias <= 3) sinal(dias < 0 ? 'encaminhamento_nao_retomado' : 'prazo_proximo', r.id, r.prazo, alerta.destinatario_id);
  }
  const metas = gpsTabela_(base, 'metas').registros.filter(r => r.obrigatoria === true);
  for (const meta of metas) {
    const periodo = gpsPeriodoFechado_(hoje, meta.periodicidade);
    if (!periodo) continue;
    const atribuicoesMeta = atribuicoes.filter(r => r.pasta_id === meta.pasta_id);
    for (const r of atribuicoesMeta) {
      const relevantes = atividades.filter(a => a.pasta_id === r.pasta_id);
      const escola = meta.por_escola ? r.escola_id : '';
      if (gpsAvaliarMeta_(meta, relevantes, r.pec_id, periodo, escola).pendente) {
        sinal('meta_nao_atingida', gpsChave_(meta.id, escola), periodo.fim, r.pec_id);
      }
    }
  }
  return sinais;
}

function gpsPeriodoFechado_(hoje, tipo) {
  const data = new Date(hoje + 'T00:00:00Z');
  const iso = d => d.toISOString().slice(0, 10);
  const inicio = new Date(data);
  const fim = new Date(data);
  if (tipo === 'semana') {
    inicio.setUTCDate(data.getUTCDate() - ((data.getUTCDay() + 6) % 7) - 7);
    fim.setTime(inicio.getTime() + 6 * 86400000);
  } else if (tipo === 'quinzena') {
    if (data.getUTCDate() > 15) {
      inicio.setUTCDate(1);
      fim.setUTCDate(15);
    } else {
      inicio.setUTCMonth(data.getUTCMonth() - 1, 16);
      fim.setUTCDate(0);
    }
  } else return null;
  return { inicio: iso(inicio), fim: iso(fim), hoje };
}

function gpsExecutarAlertas_(base, hoje, enviar) {
  const usuarios = gpsTabela_(base, 'usuarios').registros.filter(r => r.ativo === true);
  const sinaisPec = gpsSinaisAtuacao_(base, hoje);
  const cecs = usuarios.filter(r => r.perfil === 'CEC');
  const sinais = sinaisPec.filter(s => s.destinatario_id !== 'cec').concat(cecs.flatMap(u => sinaisPec.map(s => ({
    ...s, alvo_id: gpsChave_(s.alvo_id, s.destinatario_id), destinatario_id: 'cec:' + u.id
  }))));
  const anteriores = gpsTabela_(base, 'alertas').registros;
  const porId = new Map(anteriores.map(item => [item.id, item]));
  const resultado = gpsConciliarAlertas_(anteriores, sinais, hoje);
  for (const alerta of resultado.notificacoes) {
    const anterior = porId.get(alerta.id);
    const envio = anterior?.envio === 'reservado' ? 'reservado' : enviar ? 'pendente' : 'popup';
    gpsAtualizarEstado_(base, 'alertas', anterior, { ...alerta, envio });
    gpsAuditar_(base, 'alertas', alerta.id, 'notificacao:' + alerta.ocorrencias, 'automacao');
  }
  if (!enviar) return;
  const pendentes = gpsTabela_(base, 'alertas').registros.filter(r => r.envio === 'pendente' && r.status === 'aberto');
  for (const alerta of pendentes.slice(0, 10)) {
    const destinatarios = usuarios.filter(u => u.perfil === 'CEC' ? 'cec:' + u.id === alerta.destinatario_id :
      u.perfil === 'PEC' && u.pec_id === alerta.destinatario_id);
    if (destinatarios.length !== 1 || !destinatarios[0].email) continue;
    // O envio é externo. Uma interrupção após a reserva exige conferência do operador.
    gpsAtualizarEstado_(base, 'alertas', alerta, { ...alerta, envio: 'reservado' });
    SpreadsheetApp.flush();
    try {
      enviar(destinatarios[0].email, 'GPS Pedagógico · aviso', 'Há um aviso de acompanhamento disponível no painel do GPS.');
    } catch {
      continue;
    }
    gpsAtualizarEstado_(base, 'alertas', alerta, { ...alerta, envio: 'enviado' });
  }
}

function gpsMeusAlertas() {
  const { base, perfil } = gpsSessao_();
  return gpsTabela_(base, 'alertas').registros.filter(r => r.status === 'aberto' &&
    r.destinatario_id === (perfil.papel === 'CEC' ? 'cec:' + perfil.id : perfil.pec_id))
    .slice(-30).map(({ id, tipo, criticidade, periodo }) => ({ id, tipo, criticidade, periodo }));
}

function gpsJustificarAlerta(dados) {
  return gpsComTrava_(() => {
    const { base, perfil } = gpsSessao_();
    const alerta = gpsTabela_(base, 'alertas').registros.find(r => r.id === dados.id && r.destinatario_id ===
      (perfil.papel === 'CEC' ? 'cec:' + perfil.id : perfil.pec_id));
    if (!alerta || !['aberto', 'justificado'].includes(alerta.status)) throw new Error('Alerta indisponível');
    const motivo = gpsTextoWeb_(dados.motivo);
    const digest = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, motivo));
    const id = gpsChave_(alerta.id, perfil.id, digest);
    gpsAnexar_(base, 'justificativas', [{ id, alerta_id: alerta.id, pec_id: perfil.pec_id, motivo, status: 'registrada' }]);
    gpsAtualizarEstado_(base, 'alertas', alerta, { ...alerta, status: 'justificado' });
    gpsAuditar_(base, 'justificativas', id, 'registro_web', perfil.id);
    return { id };
  });
}
