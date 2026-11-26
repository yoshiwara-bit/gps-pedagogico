const GPS_PLATAFORMAS = [
  ['matific_ai', 'Matific – Anos Iniciais'], ['matific_af', 'Matific – Anos Finais'],
  ['khan', 'Khan Academy'], ['leia', 'Leia SP'], ['redacao', 'Redação Paulista'],
  ['tarefa', 'Tarefa SP'], ['elefante', 'Elefante Letrado'], ['speak', 'SPEAK'],
  ['open_english', 'Open English'], ['robotica', 'Robótica'], ['alura', 'Alura'], ['programacao', 'Programação']
];
const GPS_AVALIACOES_UI = [
  ['prova_paulista', 'Prova Paulista'], ['avd', 'Avaliação Diagnóstica'], ['saresp', 'SARESP'],
  ['saeb', 'SAEB'], ['enem', 'Simulado ENEM']
];

function gpsConsultaEscolas_(base, perfil) {
  const escopo = gpsEscopoEscolas_(base, perfil);
  return gpsTabela_(base, 'escolas').registros.filter(r => !escopo || escopo.has(r.id));
}

function gpsListarEscolas() {
  const { base, perfil } = gpsSessao_();
  const municipios = new Map(gpsTabela_(base, 'municipios').registros.map(r => [r.id, r.nome]));
  return gpsConsultaEscolas_(base, perfil).map(r => ({
    id: r.id, nome: r.nome, codigo: r.codigo, municipio: municipios.get(r.municipio_id) || ''
  })).sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function gpsAtividadesVisiveis_(base, perfil, hoje) {
  const escopo = gpsEscopoEscolas_(base, perfil);
  return gpsAtividadesBase_(base, hoje).filter(r => (!escopo || escopo.has(r.escola_id)) &&
    (perfil.papel !== 'PEC' || r.pec_id === perfil.pec_id));
}

function gpsAgendaRecorte_(base, perfil, atividades) {
  const escolas = new Map(gpsConsultaEscolas_(base, perfil).map(r => [r.id, r.nome]));
  const pecs = new Map(gpsTabela_(base, 'pecs').registros.map(r => [r.id, r.nome]));
  const planos = perfil.papel === 'AT' ? new Map() :
    new Map(gpsTabela_(base, 'planos').registros.map(r => [r.id, r.objetivo]));
  return atividades.sort((a, b) => a.data.localeCompare(b.data)).map(r => ({
    data: r.data, escola: escolas.get(r.escola_id) || 'Escola sem cadastro',
    pec: perfil.papel === 'AT' ? '' : pecs.get(r.pec_id) || '', tipo: r.tipo,
    situacao: r.decisao === 'nao_ocorreu' ? 'Visita não realizada' : r.situacao,
    periodo: ['manhã', 'tarde'].includes(r.periodo_dia) ? r.periodo_dia : '',
    foco: planos.get(r.plano_id) || ''
  }));
}

function gpsAgendaDoDia(data) {
  const { base, perfil } = gpsSessao_();
  if (!gpsDataValida_(data)) throw new Error('Informe uma data válida.');
  return gpsAgendaRecorte_(base, perfil, gpsAtividadesVisiveis_(base, perfil, gpsHoje_()).filter(r => r.data === data));
}

function gpsFichaEscola(id, secao) {
  const { base, perfil } = gpsSessao_();
  const escola = gpsConsultaEscolas_(base, perfil).find(r => r.id === id);
  if (!escola) throw new Error('Escola indisponível para seu perfil.');
  const nome = escola.nome;
  const complemento = () => gpsTabela_(base, 'fichas_escolas').registros.find(r => r.escola_id === id) || {};
  if (secao === 'cadastro') {
    const detalhes = complemento();
    const municipio = gpsTabela_(base, 'municipios').registros.find(r => r.id === escola.municipio_id)?.nome;
    const coords = gpsCoordenadas_(escola.latitude, escola.longitude);
    return { nome, dados: {
      Município: municipio, CIE: detalhes.cie, INEP: detalhes.inep, Endereço: detalhes.endereco,
      CEP: detalhes.cep, Telefone: detalhes.telefone, 'E-mail institucional': detalhes.email,
      Coordenadas: coords?.join(', '), Segmentos: detalhes.segmentos, Estudantes: detalhes.estudantes,
      Professores: detalhes.professores, Turmas: detalhes.turmas, Fonte: detalhes.fonte
    }, localizacao: !gpsModoDemonstracao_() && coords ? coords : null };
  }
  if (secao === 'equipe') {
    const detalhes = perfil.papel === 'CEC' ? complemento() : {};
    const atribuicoes = gpsTabela_(base, 'atribuicoes').registros.filter(r => r.escola_id === id &&
      (perfil.papel === 'CEC' || perfil.papel === 'PEC' && r.pec_id === perfil.pec_id));
    const ids = new Set(atribuicoes.map(r => r.pec_id));
    const responsaveis = gpsTabela_(base, 'pecs').registros.filter(r => ids.has(r.id)).map(r => r.nome);
    return { nome, dados: { Direção: detalhes.diretor, 'Vice-direção': detalhes.vice,
      'Coordenação pedagógica': detalhes.coordenadores, Supervisão: detalhes.supervisor,
      'PECs responsáveis': responsaveis.join(', ') },
    aviso: perfil.papel !== 'CEC' ? 'Dados nominais da equipe disponíveis para o CEC.' : '' };
  }
  if (['agenda', 'acompanhamentos'].includes(secao)) {
    const hoje = gpsHoje_();
    const atividades = gpsAtividadesVisiveis_(base, perfil, hoje).filter(r => r.escola_id === id);
    const selecao = secao === 'agenda' ? atividades.filter(r => r.data >= hoje).slice(0, 100) :
      atividades.sort((a, b) => b.data.localeCompare(a.data)).slice(0, 100);
    const agenda = gpsAgendaRecorte_(base, perfil, selecao);
    return { nome, colunas: ['Data', 'Atividade', 'PEC', 'Período informado', 'Situação', 'Foco pedagógico'],
      linhas: agenda.map(r => [r.data, r.tipo, r.pec, r.periodo, r.situacao.replaceAll('_', ' '), r.foco]),
      aviso: secao === 'agenda' ? 'Atividades de hoje e futuras. Horários não informados ficam em branco.' :
        'Até 100 atividades. ATPC, docentes e encaminhamentos dependem de fonte e integração.' };
  }
  if (secao === 'prioridades') {
    if (perfil.papel !== 'CEC') return { nome, aviso: 'Prioridades estratégicas disponíveis para o CEC.' };
    const componentes = new Map(gpsTabela_(base, 'componentes').registros.map(r => [r.id, r.nome]));
    const turmas = new Map(gpsTabela_(base, 'turmas').registros.map(r => [r.id, r.codigo]));
    const linhas = gpsPrioridadesBase_(base, Number(gpsHoje_().slice(0, 4))).filter(r => r.escola_id === id)
      .map(r => [r.serie, turmas.get(r.turma_id), componentes.get(r.componente_id), r.diagnostico.replaceAll('_', ' ')]);
    return { nome, colunas: ['Série', 'Turma', 'Componente', 'Diagnóstico'], linhas };
  }
  if (['frequencia', 'avaliacoes', 'plataformas', 'historico'].includes(secao)) {
    return { nome, ...gpsFichaIndicadores_(base, perfil, id, secao) };
  }
  throw new Error('Seção indisponível.');
}

function gpsPrepararFichasCSV_(grade) {
  const mapa = {
    escola_id: 'codigo_seduc', cie: 'cie', inep: 'inep', endereco: 'logradouro_pdf',
    cep: 'cep', telefone: 'telefone', email: 'email_gestor', segmentos: 'segmentos',
    estudantes: 'total_alunos_pdf', professores: 'total_professores', turmas: 'total_turmas',
    diretor: 'diretor', vice: 'vice_diretores', coordenadores: 'coordenadores_pedagogicos',
    supervisor: 'supervisor', fonte: 'fonte_pdf'
  };
  const cabecalho = grade.findIndex(r => r.some(celula => gpsNormalizar_(celula) === 'codigo_seduc'));
  if (cabecalho < 0) throw new Error('Cabeçalho sem código SEDUC.');
  const campos = grade[cabecalho].map(gpsNormalizar_);
  const ids = new Set();
  return grade.slice(cabecalho + 1).filter(r => r.some(Boolean)).map(r => {
    const registro = Object.fromEntries(Object.entries(mapa).map(([campo, coluna]) =>
      [campo, String(r[campos.indexOf(coluna)] ?? '').trim()]));
    if (!registro.escola_id || ids.has(registro.escola_id) || !registro.fonte) {
      throw new Error('Ficha sem código ou fonte, ou com código duplicado.');
    }
    ids.add(registro.escola_id);
    return { id: registro.escola_id, ...registro };
  });
}

function gpsImportarFichasLaboratorio_() {
  if (gpsModoDemonstracao_()) throw new Error('Importação institucional indisponível em demonstração.');
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_FICHAS_CSV_ID');
  if (!id) throw new Error('Configure um CSV autorizado na pasta privada.');
  const documento = gpsLerArquivoLaboratorio_(id);
  if (!documento.grades || documento.grades.length !== 1) throw new Error('Use o cadastro consolidado em uma única grade.');
  const dados = gpsPrepararFichasCSV_(documento.grades[0]);
  gpsComTrava_(() => {
    const base = gpsAbrirBase_();
    gpsAnexar_(base, 'fichas_escolas', dados);
    gpsAuditar_(base, 'fichas_escolas', 'importacao', 'cadastro', 'operador');
  });
}

function gpsResumoMapaEscola(id, inicio, fim) {
  const { base, perfil } = gpsSessao_();
  const escola = gpsConsultaEscolas_(base, perfil).find(r => r.id === id);
  if (!escola || !gpsDataValida_(inicio) || !gpsDataValida_(fim) || inicio > fim) throw new Error('Escola ou período indisponível.');
  const hoje = gpsHoje_();
  const atividades = gpsAtividadesVisiveis_(base, perfil, hoje).filter(r => r.escola_id === id && r.data >= inicio && r.data <= fim);
  const registro = atividades.filter(r => r.tipo === 'visita' && r.status !== 'cancelada')
    .sort((a,b) => b.data.localeCompare(a.data))[0];
  const resumo = { visita: registro ? { data: registro.data, situacao: registro.situacao.replaceAll('_',' '),
    termo: registro.aguardandoTermo ? 'Aguardando termo' : registro.pendenciaDocumental ? 'Documentação parcial' : '' } : null };
  if (perfil.papel !== 'CEC') return resumo;
  const indicadores = gpsTabela_(base, 'indicadores').registros.filter(r => r.escola_id === id && r.revisar === false &&
    !r.turma_id && String(r.ano) >= inicio.slice(0,4) && String(r.ano) <= fim.slice(0,4));
  const pp = indicadores.filter(r => r.fonte === 'prova_paulista' && r.metrica === 'acerto_geral');
  if (pp.length === 1 && Number.isFinite(pp[0].valor)) {
    resumo.provaPaulista = [String(pp[0].valor), pp[0].escala, pp[0].periodo].filter(Boolean).join(' · ');
  }
  const saresp = indicadores.filter(r => r.fonte === 'saresp' && r.metrica === 'nota_geral');
  if (saresp.length === 1) resumo.saresp = gpsFarolIndicador_({ ...saresp[0], farolValidado:true }, gpsTabela_(base, 'farois').registros);
  const frequencias = gpsTabela_(base, 'frequencias').registros.filter(r => r.escola_id === id &&
    !r.turma_id && !r.segmento && !r.serie && r.inicio >= inicio && r.fim <= fim);
  const ultima = frequencias.sort((a,b) => b.fim.localeCompare(a.fim))[0];
  if (ultima && frequencias.filter(r => r.fim === ultima.fim).length === 1) {
    resumo.frequencia = gpsValorMedicao_(ultima, 'frequencia') + ' · ' + ultima.inicio + ' a ' + ultima.fim;
  }
  return resumo;
}
