function gpsFarolRegistro_(r) {
  return r.homologado === true && r.referencia_regra && ['verde', 'amarelo', 'vermelho', 'azul'].includes(r.farol)
    ? r.farol : 'cinza';
}

function gpsSerieFrequencia_(registros, escolasEsperadas = []) {
  const grupos = new Map();
  for (const r of registros) {
    const id = gpsChave_(r.inicio, r.fim);
    if (!grupos.has(id)) grupos.set(id, []);
    grupos.get(id).push(r);
  }
  return [...grupos.values()].map(itens => {
    const duplicados = new Set();
    const porUnidade = new Map();
    for (const r of itens) {
      const chave = gpsChave_(r.escola_id, r.turma_id, r.segmento, r.serie, r.granularidade);
      if (porUnidade.has(chave)) duplicados.add(chave);
      porUnidade.set(chave, r);
    }
    const validos = [...porUnidade].filter(([chave]) => !duplicados.has(chave)).map(([, r]) => r);
    const agregado = gpsAgregarFrequencia_(validos);
    const recebidas = new Set(itens.map(r => r.escola_id));
    const ausentes = escolasEsperadas.filter(id => !recebidas.has(id)).length;
    return { periodo: itens[0].inicio + ' – ' + itens[0].fim, valor: agregado.percentual,
      semCobertura: agregado.semCobertura + duplicados.size + ausentes, denominador: agregado.denominador };
  }).sort((a, b) => a.periodo.localeCompare(b.periodo)).map((r, i, lista) => ({
    ...r, variacao: r.valor !== null && i && lista[i - 1].valor !== null ? r.valor - lista[i - 1].valor : null
  }));
}

function gpsConsultaIndicadores(tipo, filtro) {
  const { base, perfil } = gpsSessao_();
  return gpsDadosIndicadores_(base, perfil, tipo, filtro);
}

function gpsDadosIndicadores_(base, perfil, tipo, filtro) {
  if (perfil.papel !== 'CEC') throw new Error('Consulta de indicadores disponível para o CEC.');
  const { fonte, escola_id, municipio_id, ano, granularidade, segmento, serie, turma_id, metrica } = filtro || {};
  const fontes = GPS_AVALIACOES_UI.concat([['frequencia', 'Frequência']]);
  if (!['avaliacoes', 'frequencia', 'plataformas'].includes(tipo) ||
      (tipo === 'avaliacoes' && !fontes.some(([id]) => id === fonte)) ||
      (tipo === 'plataformas' && fonte && !GPS_PLATAFORMAS.some(([id]) => id === fonte)) ||
      (ano && !/^\d{4}$/.test(String(ano)))) throw new Error('Filtro inválido.');
  const cadastro = gpsConsultaEscolas_(base, perfil);
  const escolas = new Map(cadastro.filter(r => (!escola_id || r.id === escola_id) &&
    (!municipio_id || r.municipio_id === municipio_id)).map(r => [r.id, r.nome]));
  const turmas = new Map(gpsTabela_(base, 'turmas').registros.map(r => [r.id, r.codigo]));
  const componentes = new Map(gpsTabela_(base, 'componentes').registros.map(r => [r.id, r.nome]));
  const formato = tipo === 'avaliacoes' ? 'indicadores' : tipo === 'frequencia' ? 'frequencias' : 'plataformas';
  const fonteNome = new Map(GPS_PLATAFORMAS);
  const todos = gpsTabela_(base, formato).registros.filter(r => escolas.has(r.escola_id) &&
    (!ano || String(r.ano || String(r.inicio).slice(0, 4)) === String(ano)) &&
    (!turma_id || r.turma_id === turma_id) &&
    (!granularidade || tipo === 'avaliacoes' || r.granularidade === granularidade) &&
    (!segmento || r.segmento === segmento) && (!serie || r.serie === serie) &&
    (tipo === 'avaliacoes' ? r.fonte === fonte : tipo !== 'plataformas' || !fonte || r.plataforma === fonte));
  const faixa = new Map();
  const linhas = todos.filter(r => !metrica || r.metrica === metrica).sort((a, b) =>
    String(a.inicio || a.ano).localeCompare(String(b.inicio || b.ano), 'pt-BR'));
  for (const r of linhas) {
    const dado = tipo === 'avaliacoes'
      ? gpsFarolIndicador_({ ...r, farolValidado: r.revisar === false }, gpsRegrasAvaliacao_(base, faixa))
      : { texto: gpsValorMedicao_(r, tipo), cor: gpsValorMedicao_(r, tipo) === 'Sem dado' ? 'cinza' : gpsFarolRegistro_(r) };
    faixa.set('cor:' + dado.cor, (faixa.get('cor:' + dado.cor) || 0) + 1);
    r._tela = [{ texto: escolas.get(r.escola_id), escola_id: r.escola_id }, turmas.get(r.turma_id) || '', tipo === 'plataformas' ? fonteNome.get(r.plataforma) || '' :
      componentes.get(r.componente_id) || r.segmento || '',
      tipo === 'avaliacoes' ? [r.ano, r.periodo].filter(Boolean).join(' · ') : r.inicio + ' – ' + r.fim,
      r.metrica ? gpsValorPainel_(r, 'metrica', {}) : 'Frequência', dado, r.meta ?? '', r.escala || r.unidade || '%'];
  }
  // Agregação usa só o recorte mais específico solicitado e a granularidade declarada na fonte.
  const serieFrequencia = tipo === 'frequencia' && granularidade
    ? gpsSerieFrequencia_(linhas.filter(r => Boolean(r.turma_id) === Boolean(turma_id) &&
        Boolean(r.segmento) === Boolean(segmento) && Boolean(r.serie) === Boolean(serie)),
        !turma_id && !segmento && !serie ? [...escolas.keys()] : []) : [];
  const evolucao = tipo === 'frequencia' ? serieFrequencia :
    gpsSerieIndicador_(tipo, linhas, escola_id, metrica, turma_id);
  return { colunas: ['Escola', 'Turma', tipo === 'plataformas' ? 'Plataforma' : 'Componente ou segmento',
    'Período', 'Indicador', 'Resultado', 'Meta na fonte', 'Escala ou unidade'],
  linhas: linhas.slice(0, 200).map(r => r._tela), total: linhas.length, evolucao,
  niveis: tipo === 'avaliacoes' ? gpsDistribuicaoNiveis_(linhas, filtro || {}) : [],
  comparativo: gpsComparativosIndicador_(tipo, linhas, filtro || {}),
  distribuicao: [...faixa].filter(([id]) => id.startsWith('cor:')).map(([id, quantidade]) => ({ cor: id.slice(4), quantidade })),
  filtros: { metricas: [...new Set(todos.map(r => r.metrica).filter(Boolean))],
    segmentos: [...new Set(todos.map(r => r.segmento).filter(Boolean))], series: [...new Set(todos.map(r => r.serie).filter(Boolean))],
    turmas: [...new Set(todos.map(r => r.turma_id).filter(Boolean))].map(id => ({ id, nome: turmas.get(id) || 'Turma sem cadastro' })) },
  catalogo: tipo === 'plataformas' ? GPS_PLATAFORMAS.map(([id, nome]) => ({
    id, nome, registros: todos.filter(r => r.plataforma === id).length
  })) : [],
  aviso: tipo === 'plataformas' ? 'Resultados conforme a fonte. O acumulado depende de registros anuais consolidados.' :
    tipo === 'frequencia' ? 'Evolução calculada apenas com denominadores válidos e granularidade explícita.' :
      'As escalas de cada avaliação são preservadas. Séries só comparam registros da mesma escala.' };
}

function gpsRegrasAvaliacao_(base, cache) {
  if (!cache.has('regras')) cache.set('regras', gpsTabela_(base, 'farois').registros);
  return cache.get('regras');
}

function gpsSerieIndicador_(tipo, linhas, escola, metrica, turma) {
  if (!escola || !metrica) return [];
  const registros = linhas.filter(r => r.turma_id === (turma || '') || !r.turma_id && !turma)
    .filter(r => r.revisar !== true && Number.isFinite(r.valor));
  const grupos = new Set(registros.map(r => gpsChave_(r.escala || r.unidade, r.componente_id, r.plataforma, r.nivel, r.serie)));
  if (grupos.size !== 1) return [];
  const periodo = r => tipo === 'avaliacoes' ? [r.ano, r.periodo].filter(Boolean).join(' · ') : r.inicio + ' – ' + r.fim;
  if (new Set(registros.map(periodo)).size !== registros.length) return [];
  return registros.map(r => ({ periodo: periodo(r), valor: r.valor }));
}

function gpsValorMedicao_(r, tipo) {
  if (tipo === 'frequencia') {
    if (Number.isFinite(r.percentual)) return r.percentual.toLocaleString('pt-BR', { maximumFractionDigits: 2 }) + '%';
    if (Number.isFinite(r.presencas) && r.denominador > 0) return (r.presencas * 100 / r.denominador).toFixed(2).replace('.', ',') + '%';
  } else if (Number.isFinite(r.valor)) return r.valor.toLocaleString('pt-BR') + (r.unidade ? ' ' + r.unidade : '');
  return 'Sem dado';
}

function gpsFichaIndicadores_(base, perfil, id, secao) {
  if (perfil.papel !== 'CEC') return { aviso: 'A visualização de indicadores permanece restrita ao CEC nesta versão.' };
  const tipos = secao === 'historico' ? ['avaliacoes', 'frequencia', 'plataformas'] : [secao];
  const resultados = tipos.flatMap(tipo => {
    const fontes = tipo === 'avaliacoes' ? GPS_AVALIACOES_UI.map(([fonte]) => fonte) : [''];
    return fontes.flatMap(fonte => gpsDadosIndicadores_(base, perfil, tipo, { escola_id: id, fonte }).linhas
      .map(linha => [(new Map(GPS_AVALIACOES_UI).get(fonte) || (tipo === 'frequencia' ? 'Frequência' : 'Plataformas')), ...linha.slice(1)]));
  });
  return { colunas: ['Fonte', 'Turma', 'Componente, segmento ou plataforma', 'Período', 'Indicador',
    'Resultado', 'Meta na fonte', 'Escala ou unidade'], linhas: resultados.slice(0, 100),
  aviso: secao === 'historico' ? 'Histórico de indicadores. Intervenções e orientações aguardam integração.' :
    'Até 100 registros. Use o painel específico para selecionar a edição e os filtros.' };
}


function gpsComparativosIndicador_(tipo, linhas, filtro) {
  if (!linhas.length || (tipo !== 'frequencia' && !filtro.metrica) ||
      (tipo === 'frequencia' && !filtro.granularidade) || (tipo === 'plataformas' && !filtro.fonte)) return [];
  const periodo = r => tipo === 'avaliacoes' ? gpsChave_(r.ano, r.periodo) : gpsChave_(r.inicio, r.fim);
  const ultimo = periodo(linhas[linhas.length - 1]);
  const recorte = linhas.filter(r => periodo(r) === ultimo && r.revisar !== true &&
    (filtro.turma_id ? r.turma_id === filtro.turma_id : !r.turma_id));
  if (tipo === 'frequencia') {
    const elegiveis = recorte.filter(r => Boolean(r.segmento) === Boolean(filtro.segmento) &&
      Boolean(r.serie) === Boolean(filtro.serie));
    return elegiveis.slice(0, 20).map(r => ({ nome: r._tela[0].texto,
      valor: Number.isFinite(r.percentual) ? r.percentual : gpsAgregarFrequencia_([r]).percentual })).filter(r => r.valor !== null);
  }
  const grupos = new Set(recorte.map(r => gpsChave_(r.escala || r.unidade, r.componente_id, r.nivel, r.serie)));
  if (grupos.size !== 1 || new Set(recorte.map(r => r.escola_id)).size !== recorte.length) return [];
  return recorte.filter(r => Number.isFinite(r.valor) && r.valor >= 0).slice(0, 20)
    .map(r => ({nome:r._tela[0].texto, valor:r.valor}));
}

function gpsDistribuicaoNiveis_(linhas, filtro) {
  if (!filtro.escola_id || filtro.metrica !== 'percentual_nivel') return [];
  const niveis = linhas.filter(r => Boolean(r.turma_id) === Boolean(filtro.turma_id) &&
    r.revisar !== true && r.nivel && Number.isFinite(r.valor) && r.valor >= 0 && r.valor <= 100);
  const grupo = r => gpsChave_(r.ano, r.periodo, r.escala, r.componente_id, r.serie, r.turma_id);
  if (new Set(niveis.map(grupo)).size !== 1 || new Set(niveis.map(r => r.nivel)).size !== niveis.length ||
      Math.abs(niveis.reduce((soma, r) => soma + r.valor, 0) - 100) > .11) return [];
  return niveis.map(r => ({ nivel: r.nivel, valor: r.valor }));
}

function gpsResumoEstrategico() {
  const { base, perfil } = gpsSessao_();
  if (perfil.papel !== 'CEC') return {};
  const hoje = gpsHoje_();
  const prioridades = gpsPrioridadesBase_(base, Number(hoje.slice(0, 4)));
  const abertos = gpsTabela_(base, 'alertas').registros.filter(r => r.status === 'aberto');
  const pendencias = gpsTabela_(base, 'obrigacoes_documentais').registros.filter(r => gpsSituacaoDocumento_(r, hoje).cor === 'vermelho');
  return {
    escolasPrioritarias: new Set(prioridades.map(r => r.escola_id).filter(Boolean)).size,
    turmasPrioritarias: new Set(prioridades.map(r => r.turma_id).filter(Boolean)).size,
    alertasAbertos: abertos.length, pendenciasDocumentais: pendencias.length
  };
}

function gpsPrepararMedicoesCSV_(grade, tipo) {
  if (!['frequencias','plataformas'].includes(tipo)) throw new Error('Tipo de medição desconhecido.');
  const comuns = ['escola_id','inicio','fim','granularidade','fonte'];
  const exigidos = tipo === 'frequencias' ? comuns.concat('presencas','denominador') :
    comuns.concat('plataforma','metrica','unidade','valor');
  const aceita = ['semanal','quinzenal','mensal','acumulado','anual'];
  const campos = (grade[0] || []).map(gpsNormalizar_);
  if (exigidos.some(k => campos.filter(c => c === k).length !== 1)) throw new Error('Cabeçalho incompleto ou duplicado.');
  const ids = new Set();
  return grade.slice(1).filter(linha => linha.some(Boolean)).map(linha => {
    const ler = k => String(linha[campos.indexOf(k)] ?? '').trim();
    const registro = Object.fromEntries(comuns.map(k => [k, ler(k)]));
    registro.turma_id = ler('turma_id');
    if (!registro.escola_id || !registro.fonte || !aceita.includes(registro.granularidade) ||
        !gpsDataValida_(registro.inicio) || !gpsDataValida_(registro.fim) || registro.inicio > registro.fim) {
      throw new Error('Medição sem escola, período, granularidade ou fonte válida.');
    }
    const num = k => {
      const valor = ler(k);
      return /^\d+(?:[.,]\d+)?$/.test(valor) ? Number(valor.replace(',', '.')) : NaN;
    };
    if (tipo === 'frequencias') {
      const presencas = num('presencas');
      const denominador = num('denominador');
      if (!Number.isSafeInteger(denominador) || denominador <= 0 ||
          !Number.isSafeInteger(presencas) || presencas > denominador) throw new Error('Denominadores ou presenças inválidos.');
      Object.assign(registro, { presencas, denominador, segmento: ler('segmento'), serie: ler('serie') });
    } else {
      const plataforma = ler('plataforma');
      const metrica = ler('metrica');
      const unidade = ler('unidade');
      const valor = num('valor');
      if (!GPS_PLATAFORMAS.some(([id]) => id === plataforma) || !metrica || !unidade ||
          !Number.isFinite(valor)) throw new Error('Plataforma, unidade, métrica ou valor inválidos.');
      const denominador = num('denominador');
      if (ler('denominador') && (!Number.isFinite(denominador) || denominador <= 0)) throw new Error('Denominador inválido.');
      Object.assign(registro, { plataforma, metrica, unidade, valor, denominador: Number.isFinite(denominador) ? denominador : null });
    }
    const id = gpsChave_(tipo, registro.escola_id, registro.turma_id, registro.plataforma, registro.metrica, registro.unidade,
      registro.segmento, registro.serie, registro.inicio, registro.fim, registro.granularidade);
    if (ids.has(id)) throw new Error('Medição duplicada no arquivo.');
    ids.add(id);
    return { id, ...registro };
  });
}

function gpsImportarMedicoesLaboratorio_(tipo) {
  if (gpsModoDemonstracao_()) throw new Error('Importação institucional indisponível em demonstração.');
  const propriedades = { frequencias: 'GPS_LAB_FREQUENCIA_CSV_ID', plataformas: 'GPS_LAB_PLATAFORMAS_CSV_ID' };
  if (!Object.prototype.hasOwnProperty.call(propriedades, tipo)) throw new Error('Tipo de medição desconhecido.');
  const id = PropertiesService.getScriptProperties().getProperty(propriedades[tipo]);
  if (!id) throw new Error('Configure o CSV autorizado na pasta privada.');
  const documento = gpsLerArquivoLaboratorio_(id);
  if (!documento.grades || documento.grades.length !== 1) throw new Error('Use um CSV consolidado.');
  const registros = gpsPrepararMedicoesCSV_(documento.grades[0], tipo);
  gpsComTrava_(() => {
    const base = gpsAbrirBase_();
    gpsAnexar_(base, tipo, registros);
    gpsAuditar_(base, tipo, id, 'importacao_medicoes', 'operador');
  });
}
