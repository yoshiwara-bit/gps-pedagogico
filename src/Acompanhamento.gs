const GPS_METAS_INICIAIS = [
  ['Desenvolvimento Curricular', 'visita', 'semana', 2, false, true],
  ['Desenvolvimento Curricular', 'apoio', 'quinzena', 2, false, true],
  ['Qualidade da Aula', 'visita', 'quinzena', 1, true, true],
  ['Qualidade da Aula', 'apoio', 'quinzena', 1, true, true],
  ['Educação Especial', 'visita', 'semana', 2, false, true],
  ['Multiplica', 'visita', 'semana', 2, false, true],
  ['Recomposição', 'visita', 'quinzena', 1, true, false]
].map(([pasta, atividade, periodicidade, quantidade, por_escola, obrigatoria]) =>
  ({ pasta, atividade, periodicidade, quantidade, por_escola, obrigatoria }));

function gpsDataValida_(data) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) return false;
  const valor = new Date(data + 'T00:00:00Z');
  return Number.isFinite(valor.getTime()) && valor.toISOString().slice(0, 10) === data;
}

function gpsConsolidarAgenda_(linhas) {
  const atividades = new Map();
  const revisoes = [];
  for (const linha of linhas) {
    if (linha.tabelaConteiner || linha.tipo === 'interna') continue;
    const escolas = [...new Set((linha.escolas || []).filter(Boolean))];
    if (!linha.pec_id || !gpsDataValida_(linha.data) || !escolas.length ||
        !['visita', 'apoio'].includes(linha.tipo)) {
      revisoes.push({ origem: linha.origem, motivo: 'atividade_indeterminada' });
      continue;
    }
    for (const escola_id of escolas) {
      const id = gpsChave_(linha.tipo, linha.pec_id, linha.data, escola_id, linha.horario || '');
      if (!atividades.has(id)) {
        atividades.set(id, {
          id, escola_id, data: linha.data, pec_id: linha.pec_id, tipo: linha.tipo,
          horario: linha.horario || '', turma_id: linha.turma_id, componente_id: linha.componente_id,
          origens: [linha.origem]
        });
        continue;
      }
      const registro = atividades.get(id);
      if (registro.turma_id !== linha.turma_id || registro.componente_id !== linha.componente_id) {
        revisoes.push({ id, origem: linha.origem, motivo: 'atividades_conflitantes' });
      }
      if (!registro.origens.includes(linha.origem)) registro.origens.push(linha.origem);
    }
  }
  return { atividades: Array.from(atividades.values()), revisoes };
}

function gpsCruzarVisitas_(atividades, comprovantes, hoje) {
  if (!gpsDataValida_(hoje)) throw new Error('Data de referência inválida');
  const documentos = new Map();
  for (const documento of comprovantes) {
    if (!documento.validado || !['ficha_visita', 'termo_visita'].includes(documento.tipo)) continue;
    const tipos = documentos.get(documento.visita_id) || new Set();
    tipos.add(documento.tipo);
    documentos.set(documento.visita_id, tipos);
  }
  const unicas = Array.from(new Map(atividades.map(atividade => [atividade.id, atividade])).values());
  return unicas.map(atividade => {
    const contabilizada = atividade.data <= hoje && atividade.status !== 'cancelada';
    const tipos = documentos.get(atividade.id) || new Set();
    const comprovada = contabilizada && tipos.size > 0;
    return {
      ...atividade, contabilizada, comprovada, pendenciaDocumental: contabilizada && atividade.tipo === 'visita' && tipos.size < 2,
      aguardandoTermo: contabilizada && atividade.tipo === 'visita' && !tipos.has('termo_visita'),
      situacao: atividade.status === 'cancelada' ? 'cancelada' :
        !contabilizada ? 'prevista' : comprovada ? 'comprovada' : 'aguardando_comprovacao'
    };
  });
}

function gpsAvaliarMeta_(regra, atividades, pec_id, periodo, escola_id) {
  if (!gpsDataValida_(periodo.inicio) || !gpsDataValida_(periodo.fim) ||
      periodo.inicio > periodo.fim || !gpsDataValida_(periodo.hoje) ||
      (regra.por_escola && !escola_id)) throw new Error('Escopo da meta inválido');
  const itens = new Map(atividades.filter(item =>
    item.pec_id === pec_id && item.tipo === regra.atividade && item.status !== 'cancelada' &&
    item.data >= periodo.inicio && item.data <= periodo.fim &&
    (!regra.por_escola || item.escola_id === escola_id)).map(item => [item.id, item]));
  const registros = [...itens.values()];
  const planejamentoContabilizado = registros.filter(item => item.contabilizada).length;
  const comprovado = registros.filter(item => item.comprovada).length;
  return {
    previsto: registros.length, planejamentoContabilizado, comprovado,
    meta: regra.quantidade, obrigatoria: regra.obrigatoria,
    pendente: regra.obrigatoria && periodo.hoje > periodo.fim && planejamentoContabilizado < regra.quantidade
  };
}

function gpsMonitorarPrioridades_(prioridades, atividades) {
  const porEscola = new Map();
  for (const atividade of atividades) {
    if (atividade.status === 'cancelada') continue;
    const chave = gpsChave_(atividade.escola_id, atividade.data.slice(0, 4));
    if (!porEscola.has(chave)) porEscola.set(chave, []);
    porEscola.get(chave).push(atividade);
  }
  return prioridades.map(prioridade => {
    const candidatas = porEscola.get(gpsChave_(prioridade.escola_id, String(prioridade.ano))) || [];
    const alinhadas = candidatas.filter(item =>
      (!prioridade.turma_id || item.turma_id === prioridade.turma_id) &&
      (!prioridade.componente_id || item.componente_id === prioridade.componente_id));
    return {
      prioridade_id: prioridade.id,
      quantidade: alinhadas.length,
      direcionamento: !candidatas.length ? 'sem_planejamento' :
        alinhadas.length ? 'alinhado' : 'divergente',
      intencionalidade: 'requer_revisao_pedagogica',
      evidencias: alinhadas.some(item => item.comprovada) ? 'comprovado' :
        alinhadas.some(item => item.contabilizada) ? 'sem_evidencia_documental' : 'pendente',
      efetividade: 'requer_avaliacao_pedagogica'
    };
  });
}
