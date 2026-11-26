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
    return [r.id, { id: r.id, nome: r.nome, municipio_id: r.municipio_id, visitas: 0, apoios: 0, latitude: Math.abs(latitude) <= 90 ? latitude : null,
      longitude: Math.abs(longitude) <= 180 ? longitude : null, previstas: 0, contabilizadas: 0, comprovadas: 0 }];
  }));
  const resumo = { previstas: 0, contabilizadas: 0, comprovadas: 0 };
  const periodos = new Map();
  const calor = new Map();
  for (const r of atividades) {
    if (!pontos.has(r.escola_id) || r.status === 'cancelada' || r.data < filtro.inicio || r.data > filtro.fim ||
        (filtro.pec_id && r.pec_id !== filtro.pec_id) || (filtro.pasta_id && r.pasta_id !== filtro.pasta_id) ||
        (filtro.tipo && r.tipo !== filtro.tipo)) continue;
    if (r.contabilizada) {
      if (r.tipo === 'visita') pontos.get(r.escola_id).visitas++;
      if (r.tipo === 'apoio') pontos.get(r.escola_id).apoios++;
    }
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
  const camada = filtro.camada || 'atividades';
  if (!['atividades','calor','visitas','apoios','frequencia','avaliacoes','plataformas','prioridades'].includes(camada)) throw new Error('Camada inválida');
  if (['frequencia','avaliacoes','plataformas','prioridades'].includes(camada) && perfil.papel !== 'CEC') {
    throw new Error('Camada disponível para o CEC.');
  }
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
  dados.camada = camada;
  if (['frequencia','avaliacoes','plataformas','prioridades'].includes(camada)) {
    const tabelas = {frequencia:'frequencias', avaliacoes:'indicadores', plataformas:'plataformas'};
    const registros = camada === 'prioridades' ? gpsPrioridadesBase_(base, Number(filtro.fim.slice(0, 4))) :
      gpsTabela_(base, tabelas[camada]).registros;
    const contagem = new Map();
    for (const r of registros) {
      const ano = String(r.ano);
      const noPeriodo = r.inicio && r.fim ? r.inicio >= filtro.inicio && r.fim <= filtro.fim :
        /^\d{4}$/.test(ano) && ano >= filtro.inicio.slice(0,4) && ano <= filtro.fim.slice(0,4);
      if (noPeriodo && r.revisar !== true) contagem.set(r.escola_id, (contagem.get(r.escola_id) || 0) + 1);
    }
    for (const p of dados.pontos) p.cobertura = contagem.get(p.id) || 0;
  }
  const municipios = new Map(gpsTabela_(base, 'municipios').registros.map(r => [r.id, r.nome]));
  for (const p of dados.pontos) p.municipio = municipios.get(p.municipio_id) || '';
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
    hoje: gpsHoje_(), camadasPedagogicas: perfil.papel === 'CEC', escolas: opcoes(escolas),
    municipios: opcoes(gpsTabela_(base, 'municipios').registros.filter(r => !escolasIds || municipios.has(r.id))),
    pecs: perfil.papel === 'AT' ? [] : opcoes(gpsTabela_(base, 'pecs').registros.filter(r => !escopo || r.id === escopo)),
    pastas: perfil.papel === 'AT' ? [] : opcoes(gpsTabela_(base, 'pastas').registros.filter(r => !escopo || pastas.has(r.id)))
  };
}

// Execute no editor de laboratório com um CSV autorizado na pasta privada do laboratório.
function gpsCarregarCadastroTerritorial_() {
  const id = PropertiesService.getScriptProperties().getProperty('GPS_LAB_CADASTRO_CSV_ID');
  if (gpsModoDemonstracao_()) throw new Error('Importação institucional indisponível em demonstração.');
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

const GPS_MUNICIPIOS_MAPA = ['miracatu', 'juquia', 'iguape', 'ilha comprida', 'pedro de toledo', 'itariri'];

function gpsCoordenadas_(latitude, longitude) {
  if (latitude === '' || longitude === '' || latitude == null || longitude == null) return null;
  const lat = Number(latitude);
  const lon = Number(longitude);
  return Math.abs(lat) <= 90 && Math.abs(lon) <= 180 ? [lat, lon] : null;
}

function gpsValidarDivisas_(geojson) {
  if (geojson?.type !== 'FeatureCollection' || geojson.features?.length !== GPS_MUNICIPIOS_MAPA.length) {
    throw new Error('GeoJSON deve conter os seis municípios da URE.');
  }
  const vistos = new Set();
  let total = 0;
  return geojson.features.map(f => {
    const nome = String(f.properties?.municipio || f.properties?.name || '');
    const chave = gpsNormalizar_(nome);
    if (!GPS_MUNICIPIOS_MAPA.includes(chave) || vistos.has(chave)) throw new Error('Município ausente ou duplicado.');
    vistos.add(chave);
    const { type, coordinates } = f.geometry || {};
    const p = type === 'Polygon' ? [coordinates] : type === 'MultiPolygon' ? coordinates : null;
    if (!Array.isArray(p) || !p.length || !p.every(poligono => Array.isArray(poligono) && poligono.length &&
      poligono.every(anel => Array.isArray(anel) && anel.length >= 4 && anel.every(pos => {
        total++;
        return total <= 12000 && Array.isArray(pos) && Number.isFinite(pos[0]) &&
          Number.isFinite(pos[1]) && gpsCoordenadas_(pos[1], pos[0]);
      }) && anel[0][0] === anel[anel.length - 1][0] && anel[0][1] === anel[anel.length - 1][1]))) {
      throw new Error('Divisas inválidas ou acima de 12 mil vértices.');
    }
    return { nome, chave, cor: GPS_MUNICIPIOS_MAPA.indexOf(chave), poligonos: p };
  });
}

function gpsCartografia() {
  gpsSessao_();
  if (gpsModoDemonstracao_()) return { modo: 'esquematico', motivo: 'Mapa esquemático com dados sintéticos.' };
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('GPS_LAB_GEOJSON_ID');
  const fonte = props.getProperty('GPS_LAB_GEOJSON_FONTE');
  if (!id || !fonte) return { modo: 'pendente', motivo: 'Divisas oficiais aguardam configuração de GeoJSON autorizado.' };
  try {
    const pasta = props.getProperty('GPS_LAB_FOLDER_ID');
    const arquivo = DriveApp.getFileById(id);
    const pais = arquivo.getParents();
    let autorizado = false;
    while (pais.hasNext()) if (pais.next().getId() === pasta) autorizado = true;
    if (!autorizado || arquivo.getSize() > 1500000) throw new Error('Arquivo cartográfico fora do escopo permitido.');
    return { modo: 'geografico', fonte, divisas: gpsValidarDivisas_(JSON.parse(arquivo.getBlob().getDataAsString('UTF-8'))),
      sede: props.getProperty('GPS_LAB_SEDE_FONTE')
        ? gpsCoordenadas_(props.getProperty('GPS_LAB_SEDE_LATITUDE'), props.getProperty('GPS_LAB_SEDE_LONGITUDE')) : null };
  } catch {
    return { modo: 'pendente', motivo: 'GeoJSON indisponível ou inválido. Confira a pasta, o formato e a fonte.' };
  }
}

function gpsDistanciaLinearKm_(a, b) {
  const rad = graus => graus * Math.PI / 180;
  const dLat = rad(b[0] - a[0]);
  const dLon = rad(b[1] - a[1]);
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(rad(a[0])) * Math.cos(rad(b[0])) * Math.sin(dLon / 2) ** 2;
  return 6371 * 2 * Math.asin(Math.sqrt(Math.min(1, h)));
}
