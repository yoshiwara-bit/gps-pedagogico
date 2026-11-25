function gpsSituacaoDocumento_(obrigacao, hoje) {
  if (!gpsDataValida_(hoje)) throw new Error('Data de referência inválida');
  if (obrigacao.aplicavel === false) return { situacao: 'não aplicável', cor: 'cinza', lembrete: null };
  const prazo = obrigacao.prazo;
  if (!gpsDataValida_(prazo) || obrigacao.homologado !== true || !obrigacao.referencia) {
    return { situacao: 'prazo pendente de homologação', cor: 'cinza', lembrete: null };
  }
  const entrega = obrigacao.entrega;
  if (entrega && (!gpsDataValida_(entrega) || entrega > hoje)) {
    return { situacao: 'entrega em revisão', cor: 'cinza', lembrete: null };
  }
  if (entrega) return entrega > prazo
    ? { situacao: 'entregue com atraso', cor: 'amarelo', lembrete: null }
    : { situacao: 'entregue no prazo', cor: 'verde', lembrete: null };
  const dias = (Date.parse(prazo) - Date.parse(hoje)) / 86400000;
  if (dias < 0) return { situacao: 'vencido e pendente', cor: 'vermelho', lembrete: 'apos_prazo' };
  return { situacao: 'a vencer', cor: 'azul', lembrete: dias === 7 ? 'D-7' : dias === 2 ? 'D-2' : null };
}

function gpsSegundoDiaUtil_(mes, dias) {
  if (!/^\d{4}-\d{2}$/.test(mes) || !gpsDataValida_(mes + '-01')) throw new Error('Mês inválido');
  const porData = new Map();
  for (const item of dias.filter(r => String(r.data).slice(0, 7) === mes)) {
    if (porData.has(item.data)) return null;
    porData.set(item.data, item);
  }
  let encontrados = 0;
  for (let dia = 1; dia <= 31; dia++) {
    const data = mes + '-' + String(dia).padStart(2, '0');
    const item = porData.get(data);
    if (!gpsDataValida_(data) || item?.homologado !== true || !item.referencia || typeof item.util !== 'boolean') return null;
    if (item.util && ++encontrados === 2) return data;
  }
  return null;
}

function gpsAgregarFrequencia_(linhas) {
  let total = 0;
  let presencas = 0;
  let semCobertura = 0;
  for (const linha of linhas) {
    if (!Number.isSafeInteger(linha.denominador) || linha.denominador <= 0 ||
        !Number.isSafeInteger(linha.presencas) || linha.presencas < 0 || linha.presencas > linha.denominador) {
      semCobertura++;
      continue;
    }
    total += linha.denominador;
    presencas += linha.presencas;
  }
  return { percentual: total ? presencas * 100 / total : null, presencas, denominador: total, semCobertura };
}
