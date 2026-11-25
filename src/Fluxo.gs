function gpsAtualizarFila_(fila, entradas) {
  const itens = new Map(fila.map(item => [item.id, item]));
  for (const entrada of entradas) {
    if (!entrada.origem || !entrada.chave || !entrada.versao) throw new Error('Origem documental incompleta');
    const id = gpsChave_(entrada.chave, entrada.versao);
    if (!itens.has(id)) itens.set(id, { id, origem: entrada.origem, status: 'pendente', tentativas: 0 });
  }
  return [...itens.values()];
}

function gpsProcessarLote_(fila, limite, processar, maxTentativas = 3) {
  if (!Number.isInteger(limite) || limite < 1 || !Number.isInteger(maxTentativas) || maxTentativas < 1) {
    throw new Error('Limites de processamento inválidos');
  }
  let processados = 0;
  return fila.map(item => {
    if (processados >= limite || !['pendente', 'repetir'].includes(item.status)) return item;
    processados++;
    const tentativas = item.tentativas + 1;
    try {
      const resultado = processar(item.origem);
      return { ...item, tentativas, status: resultado.revisar ? 'revisao' : 'concluido' };
    } catch {
      return { ...item, tentativas, status: tentativas >= maxTentativas ? 'erro' : 'repetir' };
    }
  });
}

function gpsCaminhoTermo_(dataVisita, pecId, escolaId) {
  const segmento = valor => typeof valor === 'string' && /^[\p{L}\p{N}_-]+$/u.test(valor);
  if (!gpsDataValida_(dataVisita) || !segmento(pecId) || !segmento(escolaId)) {
    throw new Error('Destino documental inválido');
  }
  return ['Termos de Visita', dataVisita.slice(0, 4), dataVisita.slice(5, 7), pecId, escolaId];
}

function gpsConciliarAlertas_(historico, sinais, hoje, intervaloDias = 7) {
  if (!gpsDataValida_(hoje) || !Number.isInteger(intervaloDias) || intervaloDias < 1) {
    throw new Error('Configuração de recorrência inválida');
  }
  const estado = new Map(historico.map(item => [item.id, item]));
  const notificacoes = [];
  for (const sinal of sinais) {
    if (!sinal.tipo || !sinal.alvo_id || !sinal.periodo || !sinal.destinatario_id) continue;
    const id = gpsChave_(sinal.tipo, sinal.alvo_id, sinal.periodo, sinal.destinatario_id);
    const anterior = estado.get(id);
    if (anterior?.status === 'resolvido' || anterior?.status === 'justificado') continue;
    const dias = anterior?.ultima_notificacao
      ? (Date.parse(hoje) - Date.parse(anterior.ultima_notificacao)) / 86400000 : Infinity;
    if (dias < intervaloDias) continue;
    const alerta = {
      ...anterior, id, tipo: sinal.tipo, alvo_id: sinal.alvo_id, periodo: sinal.periodo,
      destinatario_id: sinal.destinatario_id, criticidade: sinal.criticidade || 'media',
      status: 'aberto', ultima_notificacao: hoje,
      ocorrencias: (anterior?.ocorrencias || 0) + 1
    };
    estado.set(id, alerta);
    notificacoes.push(alerta);
  }
  return { alertas: [...estado.values()], notificacoes };
}
