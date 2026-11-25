function gpsNormalizar_(texto) {
  return String(texto ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function gpsClassificarDocumento_(texto) {
  const corpo = gpsNormalizar_(texto);
  const saeb = /\b(saeb|sistema de avaliacao da educacao basica)\b/.test(corpo);
  const saresp = /\b(saresp|sistema de avaliacao de rendimento escolar do estado de sao paulo)\b/.test(corpo);
  if (saeb || saresp) {
    const abertura = corpo.split('\n').slice(0, 5).join('\n');
    const temBoletim = /\b(boletim|resultado|proficiencia)\b/.test(abertura);
    return !temBoletim || saeb && saresp
      ? { tipo: 'avaliacao_indeterminada', revisar: true }
      : { tipo: saeb ? 'saeb' : 'saresp', revisar: false };
  }
  const titulo = corpo.split('\n').find(linha => linha.trim()) || '';
  const tipos = [
    [/ficha (oficial )?de visita/, 'ficha_visita'],
    [/termo (detalhado )?de visita/, 'termo_visita'],
    [/plano de trabalho/, 'plano'],
    [/avaliacao diagnostica/, 'avd'],
    [/prova paulista/, 'prova_paulista'],
    [/frequencia escolar/, 'frequencia']
  ].filter(([padrao]) => padrao.test(titulo));
  return tipos.length === 1 ? { tipo: tipos[0][1], revisar: false } : { tipo: 'indeterminado', revisar: true };
}

function gpsDetectarCabecalho_(linhas, aliases) {
  const localizados = [];
  linhas.forEach((linha, i) => {
    const campos = {};
    for (const [campo, nomes] of Object.entries(aliases)) {
      const correspondencias = linha.map(gpsNormalizar_)
        .flatMap((celula, coluna) => nomes.includes(celula) ? [coluna] : []);
      if (correspondencias.length === 1) campos[campo] = correspondencias[0];
    }
    if (Object.keys(campos).length === Object.keys(aliases).length &&
        new Set(Object.values(campos)).size === Object.keys(campos).length) {
      localizados.push({ linha: i, campos });
    }
  });
  return localizados;
}

function gpsLerCSV_(texto) {
  const corpo = String(texto).replace(/^\uFEFF/, '').trimStart();
  const alternativas = [',', ';', '\t'].map(separador => {
    const linhas = Utilities.parseCsv(corpo, separador);
    const tamanhos = linhas.filter(linha => linha.some(celula => celula.trim())).map(linha => linha.length);
    const largura = tamanhos[0] || 0;
    return { linhas, valido: largura > 1 && tamanhos.every(tamanho => tamanho === largura) };
  }).filter(item => item.valido);
  if (alternativas.length !== 1) throw new Error('Separador CSV ausente ou ambíguo');
  return alternativas[0].linhas;
}

function gpsValorAvaliacao_(original, unidade) {
  const texto = String(original ?? '').trim();
  if (['', '-', '—'].includes(texto)) return { original, valor: null, revisar: false };
  const numero = texto.replace(/\s*%$/, '');
  if (!/^[-+]?\d+(?:[.,]\d+)?$/.test(numero) ||
      (/[.,]\d{3}$/.test(numero) && !texto.endsWith('%') && unidade !== 'percentual') ||
      (texto.endsWith('%') && unidade !== 'percentual')) return { original, valor: null, revisar: true };
  const valor = Number(numero.replace(',', '.'));
  if (!Number.isFinite(valor)) return { original, valor: null, revisar: true };
  return { original, valor, revisar: unidade === 'percentual' && (valor < 0 || valor > 100) };
}

function gpsLerImagensSAEB_(imagens, obterTexto) {
  return Array.from(imagens, (imagem, indice) => {
    const ordem = indice + 1;
    try {
      const texto = String(obterTexto(imagem) ?? '').trim();
      return { ordem, texto, revisar: true, motivo: texto ? 'conferir_ordem_e_valores_ocr' : 'ocr_sem_texto' };
    } catch (erro) {
      const motivo = /imageId/i.test(String(erro.message)) ? 'image_id_invalido' : 'falha_ocr';
      return { ordem, texto: '', revisar: true, motivo };
    }
  });
}

function gpsImportarIndicadores_(linhas, contexto) {
  const { fonte, documento_id, escala, metrica, unidade } = contexto;
  if (!GPS_FONTES.includes(fonte) || !documento_id || !escala || !metrica) {
    throw new Error('Contexto da avaliação incompleto');
  }
  if (linhas.some(linha => linha.some(celula => /^(cpf|rg|ra|nome do aluno|nome do estudante|registro do aluno)$/.test(gpsNormalizar_(celula))))) {
    return { indicadores: [], revisoes: [{ motivo: 'dados_individuais_bloqueados' }] };
  }
  const aliases = {
    ano: ['ano', 'ano de aplicacao'], serie: ['serie', 'ano/série'].map(gpsNormalizar_),
    componente: ['componente', 'componente curricular'], escola: ['codigo escola', 'escola'],
    valor: ['valor', 'resultado', 'percentual', 'nota']
  };
  if (fonte === 'avd') aliases.nivel = ['nivel', 'nivel de proficiencia'];
  const indicadores = [];
  const revisoes = [];
  const encontrados = new Map(gpsDetectarCabecalho_(linhas, aliases).map(item => [item.linha, item.campos]));
  let colunas;
  linhas.forEach((linha, indice) => {
    if (encontrados.has(indice)) {
      colunas = { ...encontrados.get(indice), turma: linha.map(gpsNormalizar_).indexOf('turma') };
      return;
    }
    if (!linha.some(item => String(item ?? '').trim())) return;
    if (!colunas) return;
    const ler = nome => String(linha[colunas[nome]] ?? '').trim();
    const ano = Number(ler('ano'));
    const serie = ler('serie').toUpperCase();
    const escola_id = ler('escola');
    const componente_id = ler('componente');
    const turma_id = ler('turma') || null;
    const nivel = ler('nivel') || null;
    const resultado = gpsValorAvaliacao_(ler('valor'), unidade);
    if (!/^\d{4}$/.test(ler('ano')) || !GPS_SERIES.includes(serie) ||
        !escola_id || !componente_id || (turma_id && fonte !== 'prova_paulista') ||
        (fonte === 'saeb' && ano % 2 !== 1) || (fonte === 'avd' && !GPS_NIVEIS_AVD.includes(nivel))) {
      revisoes.push({ linha: indice + 1, motivo: 'estrutura_ou_escopo_invalido' });
      return;
    }
    const periodo = contexto.periodo ?? null;
    const id = gpsChave_(documento_id, fonte, ano, periodo, escola_id, turma_id, serie, componente_id, metrica, nivel);
    indicadores.push({
      id, ano, serie, escola_id, componente_id, turma_id, nivel, documento_id,
      fonte, escala, metrica, periodo, valor_original: resultado.original,
      valor: resultado.valor, revisar: resultado.revisar
    });
    if (resultado.revisar) revisoes.push({ linha: indice + 1, motivo: 'valor_requer_revisao' });
  });
  if (!encontrados.size) revisoes.push({ motivo: 'cabecalho_indeterminado' });
  const contagem = new Map();
  for (const item of indicadores) contagem.set(item.id, (contagem.get(item.id) || 0) + 1);
  const idsDuplicados = [...contagem].filter(([, quantidade]) => quantidade > 1).map(([id]) => id);
  for (const id of idsDuplicados) revisoes.push({ id, motivo: 'indicador_duplicado' });
  return { indicadores: indicadores.filter(item => contagem.get(item.id) === 1), revisoes };
}

function gpsLerGrade_(aba) {
  const intervalo = aba.getDataRange();
  const linhas = intervalo.getDisplayValues();
  for (const mescla of intervalo.getMergedRanges()) {
    const linha = mescla.getRow() - intervalo.getRow();
    const coluna = mescla.getColumn() - intervalo.getColumn();
    for (let i = linha; i < Math.min(linhas.length, linha + mescla.getNumRows()); i++) {
      for (let j = coluna; j < Math.min(linhas[i].length, coluna + mescla.getNumColumns()); j++) {
        linhas[i][j] = linhas[linha][coluna];
      }
    }
  }
  return linhas;
}

// A origem deve pertencer ao laboratório. O serviço avançado usa o projeto padrão do Apps Script.
function gpsOCRBlob_(blob, pastaId) {
  const temporario = Drive.Files.create(
    { name: 'GPS LAB OCR temporário', mimeType: 'application/vnd.google-apps.document', parents: [pastaId] },
    blob, { fields: 'id', ocrLanguage: 'pt' }
  );
  try {
    return DocumentApp.openById(temporario.id).getBody().getText();
  } finally {
    DriveApp.getFileById(temporario.id).setTrashed(true);
  }
}

function gpsLerArquivoLaboratorio_(arquivoId) {
  const pastaId = PropertiesService.getScriptProperties().getProperty('GPS_LAB_FOLDER_ID');
  if (!pastaId) throw new Error('Configure a pasta privada de laboratório');
  const arquivo = DriveApp.getFileById(arquivoId);
  const pais = arquivo.getParents();
  let autorizado = false;
  while (pais.hasNext()) if (pais.next().getId() === pastaId) autorizado = true;
  if (!autorizado) throw new Error('Arquivo fora da pasta direta de laboratório');
  const mime = arquivo.getMimeType();
  const sheets = 'application/vnd.google-apps.spreadsheet';
  const docs = 'application/vnd.google-apps.document';
  if (mime === 'text/csv') return { grades: [gpsLerCSV_(arquivo.getBlob().getDataAsString('UTF-8'))] };
  if (mime === sheets) return { grades: SpreadsheetApp.openById(arquivoId).getSheets().map(gpsLerGrade_) };
  const destino = ({
    'application/pdf': docs, 'application/msword': docs,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': docs,
    'application/vnd.ms-excel': sheets,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': sheets
  })[mime];
  if (!destino) throw new Error('Formato documental não suportado');
  const temporario = Drive.Files.create(
    { name: 'GPS LAB temporário', mimeType: destino, parents: [pastaId] }, arquivo.getBlob(),
    { fields: 'id', ...(destino === docs ? { ocrLanguage: 'pt' } : {}) }
  );
  try {
    if (destino === sheets) return { grades: SpreadsheetApp.openById(temporario.id).getSheets().map(gpsLerGrade_) };
    const corpo = DocumentApp.openById(temporario.id).getBody();
    const texto = corpo.getText();
    const classificacao = gpsClassificarDocumento_(texto);
    const grades = corpo.getTables().map(tabela => Array.from(
      { length: tabela.getNumRows() }, (_, i) => {
        const linha = tabela.getRow(i);
        return Array.from({ length: linha.getNumCells() }, (_, j) => linha.getCell(j).getText());
      }
    ));
    let imagens = [];
    if (mime === 'application/pdf' && classificacao.tipo !== 'saresp') {
      try {
        imagens = gpsLerImagensSAEB_(corpo.getImages(), imagem => gpsOCRBlob_(imagem.getBlob(), pastaId));
      } catch (erro) {
        imagens.push({ ordem: null, texto: '', revisar: true,
          motivo: /imageId/i.test(String(erro.message)) ? 'image_id_invalido' : 'imagens_indisponiveis' });
      }
    }
    return { texto, classificacao, grades, requerRevisao: mime === 'application/pdf' || classificacao.revisar, imagens };
  } finally {
    DriveApp.getFileById(temporario.id).setTrashed(true);
  }
}
