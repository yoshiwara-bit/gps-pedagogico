const GPS_SERIES = Array.from({length: 9}, (_, i) => 'EF' + (i + 1)).concat('EM1', 'EM2', 'EM3');
const GPS_NIVEIS_AVD = ['Abaixo do Básico', 'Básico', 'Proficiente'];
const GPS_FONTES = ['saresp', 'saeb', 'prova_paulista', 'avd', 'frequencia'];

function gpsProjetarCoorte_(indicador, ano) {
  const { fonte, ano: aplicacao, serie } = indicador;
  const posicao = GPS_SERIES.indexOf(serie);
  if (!['saresp', 'saeb'].includes(fonte) || !Number.isInteger(aplicacao) ||
      !Number.isInteger(ano) || ano < aplicacao || posicao < 0 ||
      (fonte === 'saeb' && aplicacao % 2 !== 1)) return null;
  const destino = GPS_SERIES[posicao + ano - aplicacao];
  return {
    anoAplicacao: aplicacao, serieAplicacao: serie, anoAcompanhamento: ano,
    serieAcompanhamento: destino || null,
    situacao: destino ? 'referencia_regional_sem_vinculo_individual' : 'conclusao_prevista'
  };
}

function gpsSelecionarPrioridades_(indicadores, politica, ano) {
  if (!politica?.aprovadaCEC || !politica.versao) throw new Error('Critérios precisam de aprovação do CEC');
  const resultados = [];
  for (const indicador of indicadores) {
    if (!GPS_FONTES.includes(indicador.fonte) ||
        !Number.isFinite(indicador.valor) || indicador.revisar ||
        (indicador.turma_id && indicador.fonte !== 'prova_paulista') ||
        (indicador.fonte === 'avd' && indicador.nivel && !GPS_NIVEIS_AVD.includes(indicador.nivel))) continue;
    const regra = politica.regras.find(item => item.fonte === indicador.fonte &&
      item.escala === indicador.escala && item.metrica === indicador.metrica &&
      (item.nivel ?? '') === (indicador.nivel ?? ''));
    if (!regra || !Number.isFinite(regra.limite) ||
        !((regra.direcao === 'abaixo' && indicador.valor < regra.limite) ||
          (regra.direcao === 'acima' && indicador.valor > regra.limite))) continue;
    const historico = indicador.ano !== ano;
    const coorte = historico ? gpsProjetarCoorte_(indicador, ano) : null;
    if (historico && !coorte?.serieAcompanhamento) continue;
    resultados.push({ indicador, coorte });
  }
  const referencias = new Map();
  for (const { indicador, coorte } of resultados) {
    if (!coorte || !indicador.componente_id) continue;
    const chave = gpsChave_(coorte.serieAcompanhamento, indicador.componente_id);
    const itens = referencias.get(chave) || [];
    itens.push(indicador.id);
    referencias.set(chave, itens);
  }
  return resultados.filter(item => !item.coorte).map(({ indicador }) => ({
    id: gpsChave_(politica.versao, indicador.id, ano),
    escola_id: indicador.escola_id, turma_id: indicador.turma_id,
    componente_id: indicador.componente_id, serie: indicador.serie, ano,
    fonte: indicador.fonte, indicador_id: indicador.id, criterio_versao: politica.versao,
    referencias_historicas: indicador.fonte === 'prova_paulista' && indicador.componente_id
      ? referencias.get(gpsChave_(indicador.serie, indicador.componente_id)) || [] : []
  }));
}
