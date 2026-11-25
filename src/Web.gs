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
  obrigacoes_documentais: { titulo: 'Regularidade documental', campos: ['tipo', 'prazo', 'entrega', 'regularidade'] },
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
  if (tabela === 'obrigacoes_documentais') return registro.pec_id === perfil.pec_id;
  if (tabela === 'alertas') return registro.destinatario_id === perfil.pec_id;
  return false;
}

function gpsCarregarPainel(tabela, cursor) {
  if (!Object.prototype.hasOwnProperty.call(GPS_PAINEL, tabela) ||
      !Number.isSafeInteger(cursor) || cursor < 0) throw new Error('Consulta inválida');
  const { base: planilha, perfil } = gpsSessao_();
  const camposVisiveis = GPS_PAINEL[tabela].campos.map(gpsRotuloCampo_);
  const permitidas = perfil.papel === 'AT' ? ['escolas', 'visitas', 'apoios'] : ['visitas', 'apoios', 'alertas', 'obrigacoes_documentais'];
  if (perfil.papel !== 'CEC' && !permitidas.includes(tabela)) {
    return { titulo: GPS_PAINEL[tabela].titulo, campos: camposVisiveis, linhas: [], proximo: null };
  }
  const nomes = gpsNomesPainel_(planilha, tabela);
  if (tabela === 'prioridades') {
    const { titulo, campos } = GPS_PAINEL.prioridades;
    const hoje = gpsHoje_();
    const prioridades = gpsPrioridadesBase_(planilha, Number(hoje.slice(0, 4)));
    const acompanhamento = new Map(gpsMonitorarPrioridades_(prioridades, gpsAtividadesBase_(planilha, hoje))
      .map(r => [r.prioridade_id, r]));
    const linhas = prioridades.slice(cursor, cursor + 100).map(r => {
      const dados = { ...r, ...acompanhamento.get(r.id) };
      return campos.map(campo => gpsValorPainel_(dados, campo, nomes));
    });
    return { titulo, campos: camposVisiveis, linhas, proximo: cursor + 100 < prioridades.length ? cursor + 100 : null };
  }
  if (tabela === 'obrigacoes_documentais') {
    const hoje = gpsHoje_();
    const autorizadas = gpsTabela_(planilha, tabela).registros.filter(r => gpsPodeVer_(perfil, tabela, r));
    const linhas = autorizadas.slice(cursor, cursor + 100).map(r => {
      const situacao = gpsSituacaoDocumento_(r, hoje);
      return [gpsValorPainel_(r, 'tipo', {}), r.prazo || '—', r.entrega || '—', { texto: situacao.situacao, cor: situacao.cor }];
    });
    return { titulo: GPS_PAINEL[tabela].titulo, campos: camposVisiveis, linhas,
      proximo: cursor + 100 < autorizadas.length ? cursor + 100 : null };
  }
  if (perfil.papel === 'AT') perfil.escolas = gpsEscopoEscolas_(planilha, perfil);
  const aba = planilha.getSheetByName(tabela);
  const ultimaLinha = aba.getLastRow();
  const inicio = cursor + 2;
  const quantidade = Math.max(0, Math.min(100, ultimaLinha - inicio + 1));
  const colunas = ['id', ...GPS_SCHEMA[tabela]];
  const campos = GPS_PAINEL[tabela].campos;
  const intervalo = quantidade ? aba.getRange(inicio, 1, quantidade, colunas.length) : null;
  const regrasFarol = tabela === 'indicadores' ? gpsTabela_(planilha, 'farois').registros : [];
  const brutos = tabela === 'indicadores' && intervalo ? intervalo.getValues() : [];
  const valores = intervalo ? intervalo.getDisplayValues() : [];
  const linhas = valores.map((linha, j) => ({
    ...Object.fromEntries(colunas.map((campo, i) => [campo, linha[i]])),
    farolValidado: brutos[j]?.[colunas.indexOf('revisar')] === false,
    nivel: linha[colunas.indexOf('nivel')]
  }))
    .filter(registro => registro.id && gpsPodeVer_(perfil, tabela, registro))
    .map(registro => campos.map(campo => campo === 'farol' ? gpsFarolIndicador_(registro, regrasFarol) : gpsValorPainel_(registro, campo, nomes)));
  return {
    titulo: GPS_PAINEL[tabela].titulo, campos: camposVisiveis, linhas,
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

function gpsFarolIndicador_(registro, regras) {
  const valor = String(registro.valor ?? '').trim();
  if (!registro.farolValidado) return { texto: valor ? 'Em revisão' : 'Sem dado', cor: 'cinza' };
  if (!valor) return { texto: 'Sem dado', cor: 'cinza' };
  const nivel = String(registro.nivel ?? '').trim();
  const niveisPP = ['Abaixo do Básico', 'Básico', 'Adequado', 'Avançado'];
  const nivelAceito = nivel && (registro.fonte !== 'prova_paulista' || niveisPP.includes(nivel));
  const correspondencias = nivelAceito ? regras.filter(r => r.homologado === true && r.referencia &&
    r.fonte === registro.fonte && String(r.ano) === String(registro.ano) && r.escala === registro.escala &&
    r.metrica === registro.metrica && r.nivel === nivel && ['vermelho', 'amarelo', 'verde', 'azul'].includes(r.cor)) : [];
  const regra = correspondencias.length === 1 ? correspondencias[0] : null;
  return { texto: regra ? nivel : valor, cor: regra ? regra.cor : 'cinza' };
}

function gpsRotuloCampo_(campo) {
  const rotulos = { id: 'Identificador', nome: 'Nome', codigo: 'Código', municipio_id: 'Município',
    escola_id: 'Escola', turma_id: 'Turma', componente_id: 'Componente curricular', fonte: 'Avaliação',
    ano: 'Ano', serie: 'Ano ou série', metrica: 'Indicador', escala: 'Escala', valor_original: 'Valor na fonte',
    valor: 'Valor conferido', farol: 'Farol', revisar: 'Revisão necessária', diagnostico: 'Diagnóstico',
    direcionamento: 'Direcionamento', evidencias: 'Evidências', documento_id: 'Identificador do documento',
    prazo: 'Prazo', entrega: 'Entrega', regularidade: 'Regularidade', motivo: 'Motivo', status: 'Situação', tipo: 'Tipo', criticidade: 'Criticidade', data: 'Data',
    atividade: 'Atividade', periodicidade: 'Periodicidade', quantidade: 'Quantidade', obrigatoria: 'Obrigatória' };
  return rotulos[campo] || campo;
}

function gpsNomesPainel_(base, tabela) {
  const nomes = {};
  const cadastros = { escola_id: 'escolas', municipio_id: 'municipios', turma_id: 'turmas', componente_id: 'componentes' };
  for (const campo of GPS_PAINEL[tabela].campos) {
    if (!cadastros[campo]) continue;
    nomes[campo] = new Map(gpsTabela_(base, cadastros[campo]).registros.map(r => [String(r.id), r.nome || r.codigo]));
  }
  return nomes;
}

function gpsValorPainel_(registro, campo, nomes) {
  const valor = registro[campo] ?? '';
  if (nomes[campo]) return valor ? nomes[campo].get(String(valor)) || 'Cadastro indisponível' : '—';
  const rotulos = { saresp: 'SARESP', saeb: 'SAEB', prova_paulista: 'Prova Paulista', avd: 'Avaliação Diagnóstica',
    frequencia: 'Frequência', participacao: 'Participação', proficiencia: 'Proficiência', percentual_nivel: 'Percentual no nível',
    aguardando_comprovacao: 'Aguardando comprovação', sem_historico_escolar: 'Sem histórico escolar',
    sem_planejamento: 'Sem planejamento', sem_evidencia_documental: 'Sem evidência documental', termo_pendente: 'Termo pendente',
    prioridade_nao_contemplada: 'Prioridade não contemplada', meta_nao_atingida: 'Meta não atingida' };
  return ['fonte', 'metrica', 'status', 'tipo', 'diagnostico', 'direcionamento', 'evidencias', 'motivo'].includes(campo)
    ? rotulos[valor] || String(valor).replaceAll('_', ' ') : valor;
}
