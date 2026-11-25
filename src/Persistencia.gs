function gpsAbrirBase_() {
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_DB_ID');
  if (!id) throw new Error('Base de laboratório indisponível');
  return SpreadsheetApp.openById(id);
}

function gpsComTrava_(executar) {
  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    return executar();
  } finally {
    try { SpreadsheetApp.flush(); } finally { trava.releaseLock(); }
  }
}

function gpsTabela_(base, nome) {
  const aba = base.getSheetByName(nome);
  const campos = ['id', ...GPS_SCHEMA[nome]];
  const n = aba.getLastRow() - 1;
  const valores = n > 0 ? aba.getRange(2, 1, n, campos.length).getValues() : [];
  return {
    aba, campos,
    registros: valores.map((linha, i) => ({
      ...Object.fromEntries(campos.map((campo, j) => [campo, gpsLerCelula_(linha[j])])),
      _linha: i + 2
    })).filter(item => item.id)
  };
}

function gpsLerCelula_(valor) {
  if (valor instanceof Date) return Utilities.formatDate(valor, 'America/Sao_Paulo', 'yyyy-MM-dd');
  // Remove o escape apenas quando devolvido literalmente pelo adaptador.
  return typeof valor === 'string' && /^'[=+@-]/.test(valor) ? valor.slice(1) : valor;
}

function gpsCelula_(valor) {
  if (Array.isArray(valor)) return JSON.stringify(valor);
  return typeof valor === 'string' && /^[=+@-]/.test(valor) ? "'" + valor : valor ?? '';
}

function gpsAnexar_(base, nome, registros) {
  if (!registros.length) return;
  const destino = gpsTabela_(base, nome);
  const indice = { [nome]: new Map(destino.registros.map(item => [item.id, item])) };
  const refs = new Set(GPS_SCHEMA[nome].map(campo => GPS_RELACOES[campo]).filter(Boolean));
  for (const ref of refs) {
    if (!indice[ref]) indice[ref] = new Map(gpsTabela_(base, ref).registros.map(item => [item.id, item]));
  }
  const linhas = [];
  for (const registro of registros) {
    const normalizado = Object.fromEntries(destino.campos.map(campo => [
      campo, Array.isArray(registro[campo]) ? JSON.stringify(registro[campo]) : registro[campo]
    ]));
    if (!gpsValidarRegistro_(nome, normalizado, indice)) continue;
    indice[nome].set(registro.id, normalizado);
    linhas.push(destino.campos.map(campo => gpsCelula_(normalizado[campo])));
  }
  if (linhas.length) destino.aba.getRange(destino.aba.getLastRow() + 1, 1, linhas.length, destino.campos.length).setValues(linhas);
}

// Apenas para estado mutável de fila e notificações, sob a trava do script.
function gpsAtualizarEstado_(base, nome, anterior, registro) {
  const aba = base.getSheetByName(nome);
  const campos = ['id', ...GPS_SCHEMA[nome]];
  const linha = anterior?._linha || aba.getLastRow() + 1;
  aba.getRange(linha, 1, 1, campos.length).setValues([campos.map(campo => gpsCelula_(registro[campo]))]);
}

function gpsAuditar_(base, entidade, registro_id, operacao, ator_id) {
  const id = gpsChave_(entidade, registro_id, operacao);
  const aba = base.getSheetByName('auditoria');
  const ultima = aba.getLastRow();
  if (ultima >= 2 && aba.getRange(2, 1, ultima - 1, 1).createTextFinder(id).matchEntireCell(true).findNext()) return;
  gpsAnexar_(base, 'auditoria', [{ id, entidade, registro_id, operacao, ator_id, instante: new Date().toISOString() }]);
}
