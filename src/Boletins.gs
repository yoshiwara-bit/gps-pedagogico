const GPS_LAYOUT_BOLETINS = {
  saresp: {
    ano: ['ano', 'ano de aplicacao', 'edicao'],
    nota: ['nota', 'nota media'],
    proficiencia: ['proficiencia', 'proficiencia media'],
    participacao: ['participacao', 'taxa de participacao']
  },
  saeb: {
    ano: ['ano', 'ano de aplicacao', 'edicao saeb'],
    proficiencia: ['proficiencia', 'proficiencia media', 'media de desempenho'],
    participacao: ['participacao', 'taxa de participacao'],
    percentual_nivel: ['percentual no nivel', 'percentual por nivel']
  }
};

function gpsSerieBoletim_(original) {
  const texto = gpsNormalizar_(original);
  if (/^(ef[1-9]|em[1-3])$/.test(texto)) return texto.toUpperCase();
  const resultado = texto.match(/^([1-9])\s*[ºª°o.]?\s*(ano(?: do ensino fundamental)?|serie do ensino medio)$/);
  if (!resultado) return '';
  if (resultado[2].startsWith('serie')) return Number(resultado[1]) <= 3 ? 'EM' + resultado[1] : '';
  return 'EF' + resultado[1];
}

function gpsExtrairBoletim_(documento, fonte, documento_id) {
  const perfil = GPS_LAYOUT_BOLETINS[fonte];
  if (!perfil || !documento_id) throw new Error('Boletim sem contexto');
  const indicadores = [];
  const revisoes = [];
  const grades = documento.grades || [];
  for (const [tabela, grade] of grades.entries()) {
    if (grade.some(linha => linha.some(celula => /^(cpf|rg|ra|nome do aluno|nome do estudante)$/.test(gpsNormalizar_(celula))))) {
      revisoes.push({ tabela, motivo: 'dados_individuais_bloqueados' });
      continue;
    }
    const base = { ano: perfil.ano, serie: ['ano/serie', 'serie', 'etapa'],
      escola: ['codigo escola', 'codigo da escola', 'codigo seduc', 'codigo inep'],
      componente: ['componente', 'componente curricular', 'disciplina'] };
    for (const [metrica, nomes] of Object.entries(perfil)) {
      if (metrica === 'ano') continue;
      const aliases = { ...base, valor: nomes };
      if (metrica === 'percentual_nivel') aliases.nivel = ['nivel', 'nivel de proficiencia'];
      const cabecalhos = gpsDetectarCabecalho_(grade, aliases);
      for (const [i, cabecalho] of cabecalhos.entries()) {
        const fim = cabecalhos[i + 1]?.linha ?? grade.length;
        const { campos } = cabecalho;
        for (let linha = cabecalho.linha + 1; linha < fim; linha++) {
          const valores = grade[linha];
          if (!valores.some(item => String(item ?? '').trim())) continue;
          const ler = campo => String(valores[campos[campo]] ?? '').trim();
          const ano = Number(ler('ano'));
          const serie = gpsSerieBoletim_(ler('serie'));
          const escola_id = ler('escola');
          const componente_id = ler('componente');
          const nivel = ler('nivel');
          const percentual = ['participacao', 'percentual_nivel'].includes(metrica);
          const resultado = gpsValorAvaliacao_(ler('valor'), percentual ? 'percentual' : metrica);
          if (!/^\d{4}$/.test(ler('ano')) || (fonte === 'saeb' && ano % 2 !== 1) || !serie ||
              !escola_id || !componente_id || (aliases.nivel && !nivel)) {
            revisoes.push({ tabela, linha: linha + 1, motivo: 'rotulo_valor_ou_escala_indeterminados' });
            continue;
          }
          if (resultado.revisar) revisoes.push({ tabela, linha: linha + 1, motivo: 'valor_requer_conferencia' });
          // A versão identifica a métrica e a edição sem inferir equivalência entre boletins.
          const escala = [fonte, metrica, ano].join(':');
          indicadores.push({
            id: gpsChave_(documento_id, fonte, ano, escola_id, serie, componente_id, metrica, nivel),
            ano, periodo: String(ano), escola_id, serie, componente_id, metrica, escala, nivel,
            documento_id, fonte, valor_original: resultado.original, valor: resultado.valor, revisar: resultado.revisar
          });
        }
      }
    }
  }
  const porId = new Map();
  const duplicados = new Set();
  for (const item of indicadores) {
    if (porId.has(item.id)) duplicados.add(item.id);
    porId.set(item.id, item);
  }
  if (duplicados.size) revisoes.push({ motivo: 'rotulo_repetido_ou_conflitante' });
  if (!porId.size) revisoes.push({ motivo: 'sem_tabela_identificavel' });
  if ((documento.imagens || []).some(item => item.revisar)) revisoes.push({ motivo: 'ocr_requer_conferencia' });
  return { indicadores: [...porId.values()].filter(item => !duplicados.has(item.id)), revisoes };
}

function gpsRecuperarImagensExportadas_(id, imagens, pastaId) {
  if (!imagens.some(item => item.motivo === 'image_id_invalido')) return imagens;
  try {
    const zip = Drive.Files.export(id, 'application/zip');
    const arquivos = Utilities.unzip(zip);
    const html = arquivos.find(item => /\.html?$/i.test(item.getName()));
    const caminhos = [...(html?.getDataAsString('UTF-8') || '').matchAll(/<img\b[^>]*?\bsrc=["']([^"']+)["']/gi)]
      .map(item => item[1].replace(/^\.\//, ''));
    if (caminhos.length !== imagens.length || new Set(caminhos).size !== caminhos.length) return imagens;
    const porNome = new Map(arquivos.map(item => [item.getName(), item]));
    return imagens.map((item, i) => {
      const blob = porNome.get(caminhos[i]);
      if (item.motivo !== 'image_id_invalido' || !blob ||
          !/\.(png|jpe?g|gif|bmp|webp)$/i.test(caminhos[i])) return item;
      try {
        const texto = gpsOCRBlob_(blob, pastaId).trim();
        return texto ? { ...item, texto, motivo: 'ocr_exportado_ordem_requer_conferencia' } : item;
      } catch {
        return item;
      }
    });
  } catch {
    return imagens;
  }
}
