function doGet() {
  return HtmlService.createHtmlOutputFromFile('Painel').setTitle('GPS Pedagógico 2.0');
}

const GPS_PAINEL = {
  escolas: { titulo: 'Escolas', campos: ['nome', 'codigo', 'municipio_id'] },
  municipios: { titulo: 'Municípios', campos: ['nome'] },
  pecs: { titulo: 'PECs', campos: ['nome'] },
  prioridades: { titulo: 'Prioridades', campos: ['escola_id', 'serie', 'turma_id', 'componente_id', 'ano'] },
  alertas: { titulo: 'Alertas', campos: ['tipo', 'criticidade', 'status'] },
  visitas: { titulo: 'Visitas', campos: ['data', 'escola_id', 'status'] },
  apoios: { titulo: 'Apoios', campos: ['data', 'escola_id', 'status'] },
  metas: { titulo: 'Metas', campos: ['atividade', 'periodicidade', 'quantidade', 'obrigatoria'] }
};

function gpsPodeVer_(perfil, tabela, registro) {
  if (perfil.papel === 'CEC') return true;
  if (perfil.papel !== 'PEC' || !perfil.pec_id) return false;
  if (['visitas', 'apoios'].includes(tabela)) return registro.pec_id === perfil.pec_id;
  if (tabela === 'alertas') return registro.destinatario_id === perfil.pec_id;
  return false;
}

function gpsCarregarPainel(tabela, cursor) {
  if (!Object.prototype.hasOwnProperty.call(GPS_PAINEL, tabela) ||
      !Number.isSafeInteger(cursor) || cursor < 0) throw new Error('Consulta inválida');
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_DB_ID');
  if (!email || !id) throw new Error('Acesso ao laboratório indisponível');
  const planilha = SpreadsheetApp.openById(id);
  const usuarios = planilha.getSheetByName('usuarios').getDataRange().getValues().slice(1);
  const candidatos = usuarios.filter(([, endereco]) => String(endereco).trim().toLowerCase() === email);
  if (candidatos.length !== 1 || candidatos[0][4] !== true) throw new Error('Acesso ao laboratório indisponível');
  const [, , papel, pec_id] = candidatos[0];
  const perfil = { papel, pec_id };
  const aba = planilha.getSheetByName(tabela);
  const ultimaLinha = aba.getLastRow();
  const inicio = cursor + 2;
  const quantidade = Math.max(0, Math.min(100, ultimaLinha - inicio + 1));
  const colunas = ['id', ...GPS_SCHEMA[tabela]];
  const campos = GPS_PAINEL[tabela].campos;
  const valores = quantidade ? aba.getRange(inicio, 1, quantidade, colunas.length).getDisplayValues() : [];
  const linhas = valores.map(linha => Object.fromEntries(colunas.map((campo, i) => [campo, linha[i]])))
    .filter(registro => registro.id && gpsPodeVer_(perfil, tabela, registro))
    .map(registro => campos.map(campo => registro[campo]));
  return {
    titulo: GPS_PAINEL[tabela].titulo, campos, linhas,
    proximo: inicio + quantidade <= ultimaLinha ? cursor + quantidade : null
  };
}
