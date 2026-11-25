function gpsFaixaGrafico_(iso, granularidade) {
  if (granularidade === 'ano') return iso.slice(0, 4);
  if (granularidade === 'mensal') return iso.slice(0, 7);
  if (granularidade === 'quinzenal') return iso.slice(0, 7) + (Number(iso.slice(8)) <= 15 ? ' / 1' : ' / 2');
  const data = new Date(iso + 'T00:00:00Z');
  data.setUTCDate(data.getUTCDate() - ((data.getUTCDay() + 6) % 7));
  return data.toISOString().slice(0, 10);
}

function gpsAgregarTerritorio_(escolas, atividades, filtro) {
  const apenasEscolas = escolas.filter(r => (!filtro.escola_id || r.id === filtro.escola_id) &&
    (!filtro.municipio_id || r.municipio_id === filtro.municipio_id));
  const pontos = new Map(apenasEscolas.map(r => {
    const latitude = r.latitude === '' || r.latitude == null ? null : Number(r.latitude);
    const longitude = r.longitude === '' || r.longitude == null ? null : Number(r.longitude);
    return [r.id, { id: r.id, nome: r.nome, latitude: Math.abs(latitude) <= 90 ? latitude : null,
      longitude: Math.abs(longitude) <= 180 ? longitude : null, previstas: 0, contabilizadas: 0, comprovadas: 0 }];
  }));
  const resumo = { previstas: 0, contabilizadas: 0, comprovadas: 0 };
  const periodos = new Map();
  const calor = new Map();
  for (const r of atividades) {
    if (!pontos.has(r.escola_id) || r.status === 'cancelada' || r.data < filtro.inicio || r.data > filtro.fim ||
        (filtro.pec_id && r.pec_id !== filtro.pec_id) || (filtro.pasta_id && r.pasta_id !== filtro.pasta_id) ||
        (filtro.tipo && r.tipo !== filtro.tipo)) continue;
    const periodo = gpsFaixaGrafico_(r.data, filtro.granularidade);
    const celula = gpsChave_(r.escola_id, periodo);
    if (!periodos.has(periodo)) periodos.set(periodo, { periodo, previstas: 0, contabilizadas: 0, comprovadas: 0 });
    if (!calor.has(celula)) calor.set(celula, { escola_id: r.escola_id, periodo, previstas: 0, contabilizadas: 0, comprovadas: 0 });
    for (const total of [resumo, pontos.get(r.escola_id), periodos.get(periodo), calor.get(celula)]) {
      total.previstas++;
      if (r.contabilizada) total.contabilizadas++;
      if (r.comprovada) total.comprovadas++;
    }
  }
  return { resumo, pontos: [...pontos.values()], periodos: [...periodos.values()].sort((a, b) => a.periodo.localeCompare(b.periodo)),
    calor: [...calor.values()] };
}

function gpsDadosTerritorio(filtro) {
  const { base, perfil } = gpsSessao_();
  if (!filtro || !gpsDataValida_(filtro.inicio) || !gpsDataValida_(filtro.fim) || filtro.fim < filtro.inicio ||
      Date.parse(filtro.fim) - Date.parse(filtro.inicio) > 366 * 86400000 ||
      !['semanal', 'quinzenal', 'mensal', 'ano'].includes(filtro.granularidade) ||
      !['', 'visita', 'apoio'].includes(filtro.tipo)) throw new Error('Filtro inválido');
  const hoje = gpsHoje_();
  const escopo = perfil.papel === 'CEC' ? '' : perfil.pec_id;
  const consulta = { ...filtro, pec_id: perfil.papel === 'AT' ? '' : escopo || filtro.pec_id,
    pasta_id: perfil.papel === 'AT' ? '' : filtro.pasta_id };
  const visiveis = gpsEscopoEscolas_(base, perfil);
  const chave = 'territorio:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify([base.getId(), perfil.papel, escopo, hoje, consulta, visiveis ? [...visiveis].sort() : null])));
  const cache = CacheService.getScriptCache();
  const hit = cache.get(chave);
  if (hit) return JSON.parse(hit);
  const escolas = gpsTabela_(base, 'escolas').registros.filter(r => !visiveis || visiveis.has(r.id));
  const atividades = gpsAtividadesBase_(base, hoje).filter(r => !escopo || r.pec_id === escopo);
  const dados = gpsAgregarTerritorio_(escolas, atividades, consulta);
  const json = JSON.stringify(dados);
  if (Utilities.newBlob(json).getBytes().length < 95000) cache.put(chave, json, 30);
  return dados;
}

function gpsOpcoesTerritorio() {
  const { base, perfil } = gpsSessao_();
  const escopo = perfil.papel === 'CEC' ? '' : perfil.pec_id;
  const atribuicoes = escopo ? gpsTabela_(base, 'atribuicoes').registros.filter(r => r.pec_id === escopo) : [];
  const ids = campo => new Set(atribuicoes.map(r => r[campo]));
  const escolasIds = gpsEscopoEscolas_(base, perfil);
  const escolas = gpsTabela_(base, 'escolas').registros.filter(r => !escolasIds || escolasIds.has(r.id));
  const municipios = new Set(escolas.map(r => r.municipio_id));
  const pastas = ids('pasta_id');
  const opcoes = registros => registros.map(({ id, nome }) => ({ id, nome }));
  return {
    hoje: gpsHoje_(), escolas: opcoes(escolas),
    municipios: opcoes(gpsTabela_(base, 'municipios').registros.filter(r => !escolasIds || municipios.has(r.id))),
    pecs: perfil.papel === 'AT' ? [] : opcoes(gpsTabela_(base, 'pecs').registros.filter(r => !escopo || r.id === escopo)),
    pastas: perfil.papel === 'AT' ? [] : opcoes(gpsTabela_(base, 'pastas').registros.filter(r => !escopo || pastas.has(r.id)))
  };
}

// Execute no editor de laboratório com um CSV autorizado na pasta privada do laboratório.
function gpsCarregarCadastroTerritorial_() {
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_CADASTRO_CSV_ID');
  if (!id) throw new Error('Configure o CSV de cadastro territorial');
  const doc = gpsLerArquivoLaboratorio_(id);
  if (!doc.grades || doc.grades.length !== 1) throw new Error('Cadastro deve ser um CSV');
  const cadastro = gpsCadastroTerritorial_(doc.grades[0]);
  gpsComTrava_(() => {
    const base = gpsAbrirBase_();
    gpsAnexar_(base, 'municipios', cadastro.municipios);
    gpsAnexar_(base, 'escolas', cadastro.escolas);
    gpsAuditar_(base, 'escolas', 'cadastro_territorial', 'importacao', 'operador');
  });
}

function gpsCadastroTerritorial_(grade) {
  const colunas = { codigo: ['codigo_seduc'], nome: ['escola'], municipio: ['municipio'],
    latitude: ['latitude'], longitude: ['longitude'] };
  const cabecalhos = gpsDetectarCabecalho_(grade, colunas);
  if (cabecalhos.length !== 1) throw new Error('Cabeçalho territorial indeterminado');
  const { linha, campos } = cabecalhos[0];
  const municipios = new Map();
  const escolas = new Map();
  for (const r of grade.slice(linha + 1)) {
    if (!r.some(valor => String(valor ?? '').trim())) continue;
    const ler = campo => String(r[campos[campo]] ?? '').trim();
    const codigo = ler('codigo');
    const nome = ler('nome');
    const municipio = ler('municipio');
    const latitude = Number(ler('latitude').replace(',', '.'));
    const longitude = Number(ler('longitude').replace(',', '.'));
    if (!/^[\d]+$/.test(codigo) || !nome || !municipio || !ler('latitude') || !ler('longitude') ||
        !Number.isFinite(latitude) || !Number.isFinite(longitude) || Math.abs(latitude) > 90 || Math.abs(longitude) > 180 ||
        escolas.has(codigo)) throw new Error('Registro territorial ausente ou duplicado');
    const municipio_id = gpsChave_('municipio', gpsNormalizar_(municipio));
    if (!municipios.has(municipio_id)) municipios.set(municipio_id, { id: municipio_id, nome: municipio });
    escolas.set(codigo, { id: codigo, codigo, nome, municipio_id, latitude, longitude });
  }
  if (!escolas.size) throw new Error('Cadastro territorial vazio');
  return { municipios: [...municipios.values()], escolas: [...escolas.values()] };
}
