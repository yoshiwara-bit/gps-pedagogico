const assert = require('node:assert/strict');
const fs = require('node:fs');
const { chromium } = require('playwright');

function renderizar(papel) {
  const painel = fs.readFileSync('src/Painel.html', 'utf8');
  const codigo = [];
  let ultimo = 0;
  for (const trecho of painel.matchAll(/<\?([!=]*)([\s\S]*?)\?>/g)) {
    codigo.push('saida += ' + JSON.stringify(painel.slice(ultimo, trecho.index)));
    if (trecho[1]) codigo.push('saida += (' + trecho[2] + ')');
    else codigo.push(trecho[2]);
    ultimo = trecho.index + trecho[0].length;
  }
  codigo.push('saida += ' + JSON.stringify(painel.slice(ultimo)), 'return saida');
  return new Function('papel', 'demonstracao', 'gpsIncluir_', 'let saida = "";\n' + codigo.join('\n'))(
    papel, true, nome => fs.readFileSync('src/' + nome + '.html', 'utf8'));
}

function instalarRpcFalso(papel) {
  const escolas = Array.from({ length: papel === 'CEC' ? 40 : papel === 'AT' ? 8 : 20 }, (_, i) => ({
    id: 'E' + i, nome: 'Escola fictícia ' + (i + 1), latitude: -10 - i / 20, longitude: -40 + (i % 4) / 6,
    contabilizadas: i ? 0 : 1
  }));
  const resposta = {
    gpsResumoMapaEscola: {visita:{data:'2026-11-25',situacao:'aguardando comprovação',termo:'Aguardando termo'}},
    gpsResumoEstrategico: {escolasPrioritarias:4,turmasPrioritarias:7,pendenciasDocumentais:2,alertasAbertos:3},
    gpsCartografia: papel === 'AT' ? {
      modo:'geografico', fonte:'Geometria sintética para teste',
      sede:[-10.2,-39.7], divisas:[{nome:'Área sintética',chave:'area',cor:0,
        poligonos:[[[[-40,-10],[-39,-10],[-39,-11],[-40,-10]]]]}]
    } : { modo:'esquematico', motivo:'Mapa esquemático com dados sintéticos.' },
    gpsListarEscolas: escolas.map(r => ({...r, municipio:'Município fictício'})),
    gpsFichaEscola: { nome:'Escola fictícia 1', dados:{ Município:'Município fictício', Estudantes:150 } },
    gpsAgendaDoDia: [{ escola:'Escola fictícia 1', pec:'PEC fictício', tipo:'visita', periodo:'', situacao:'prevista', foco:'Exemplo' }],
    gpsConsultaIndicadores: { colunas:['Escola','Resultado'],
      linhas:[[{texto:'Escola fictícia 1',escola_id:'E0'},{texto:'90%',cor:'verde'}]], total:1,
      evolucao:[{periodo:'Período 1',valor:88},{periodo:'Período 2',valor:90,variacao:2}],
      distribuicao:[{cor:'verde',quantidade:1}], filtros:{metricas:[],segmentos:[],series:[],turmas:[]},
      catalogo:[], aviso:'Dados sintéticos' },
    gpsOpcoesTerritorio: { escolas, municipios: [], pecs: [], pastas: [], hoje: '2026-11-25' },
    gpsMinhasOpcoes: { papel, escolas: papel === 'PEC' ? escolas : [], pastas: [{ id: 'P', nome: 'Pasta fictícia' }],
      componentes: [{ id: 'LP', nome: 'Português' }], turmas: [], visitas: [{ id: 'V', nome: 'Visita fictícia' }] },
    gpsMeusAlertas: papel === 'PEC' ? [{ id: 'L', tipo: 'termo_pendente', periodo: '2026-11-25' }] : [],
    gpsCarregarPainel: { campos: ['Farol'], linhas: [[{ cor: 'vermelho', texto: 'Abaixo do Básico' }]], proximo: null },
    gpsAvisosFormulario: { prioridades: 1, alinhadas: 1 },
    gpsSalvarPlano: { id: 'V' }
  };
  const chamadas = [];
  window.chamadasRpc = chamadas;
  window.google = { script: { run: {
    withSuccessHandler(ok) {
      return { withFailureHandler(falha) {
        return new Proxy({}, { get(_, metodo) {
          return (...args) => {
            chamadas.push(metodo);
            const [filtro] = args;
            const pontos = filtro?.escola_id ? escolas.filter(item => item.id === filtro.escola_id) : escolas;
            const dados = metodo === 'gpsDadosTerritorio' ? {
              pontos, resumo: { previstas: 2, contabilizadas: 1, comprovadas: 1 },
              periodos: [{ periodo: '2026-11', previstas: 2, contabilizadas: 1, comprovadas: 1 }],
              calor: [{ escola_id: 'E0', periodo: '2026-11', previstas: 2, contabilizadas: 1, comprovadas: 1 }]
            } : resposta[metodo];
            setTimeout(() => dados ? ok(dados) : falha(new Error('RPC sem simulação')), 0);
          };
        } });
      } };
    }
  } } };
}

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_BIN || '/usr/bin/chromium', args: ['--no-sandbox'] });
  try {
    for (const [papel, largura, total] of [['CEC', 1280, 40], ['PEC', 390, 20], ['AT', 390, 8]]) {
      const page = await browser.newPage({ viewport: { width: largura, height: 800 } });
      const erros = [];
      page.on('pageerror', erro => erros.push(erro.message));
      await page.evaluate(instalarRpcFalso, papel);
      await page.setContent(renderizar(papel));
      await page.waitForFunction(q => document.getElementById('gps-total-escolas').textContent === String(q), total);
      if (papel === 'PEC') {
        await page.getByRole('dialog').waitFor();
        await page.getByRole('button', { name: 'Fechar', exact: true }).click();
      }
      async function abrir(nome) {
        const menu = page.locator('#gps-menu');
        if (await menu.getAttribute('aria-expanded') === 'false') await menu.click();
        await page.locator('#gps-navegacao').getByRole('button', { name: nome }).click();
      }
      async function captura(nome) {
        if (process.env.GPS_CAPTURAS) await page.screenshot({ path:process.env.GPS_CAPTURAS + '/' + nome + '.png', fullPage:true });
      }
      if (papel === 'CEC') {
        await page.locator('[data-resumo=escolasPrioritarias]').getByText('4', {exact:true}).waitFor();
        await captura('painel-estrategico');
      }
      await abrir('Mapa territorial');
      await page.waitForFunction(q => document.querySelectorAll('#gps-mapa circle').length === q, total);
      if (papel === 'CEC') await captura('mapa-demonstracao');
      if (papel === 'AT') {
        assert.equal(await page.locator('#gps-mapa path').count(),1);
        await page.locator('#gps-mapa rect').getByText('Sede da URE').waitFor({state:'attached'});
        await page.getByRole('button',{name:'Aproximar'}).click();
        assert.notEqual(await page.locator('#gps-mapa').getAttribute('viewBox'),'0 0 600 320');
        await page.getByRole('button',{name:'Ver território inteiro'}).click();
      }
      const ponto = page.locator('#gps-mapa circle').first();
      const previa = page.locator('#gps-previa');
      await ponto.hover();
      await previa.getByText('Escola fictícia 1').waitFor();
      await previa.getByText('0 previstas · 1 contabilizadas · 0 comprovadas').waitFor();
      await previa.getByText(/Aguardando termo/).waitFor();
      assert.equal(await page.locator('#gps-mapa circle').count(), total);
      await page.locator('#gps-mapa-estado').hover();
      await previa.waitFor({ state: 'hidden' });
      await ponto.focus();
      await previa.waitFor();
      await ponto.press('Escape');
      await previa.waitFor({ state: 'hidden' });
      await ponto.click();
      await page.waitForFunction(() => document.querySelectorAll('#gps-mapa circle').length === 1);
      await page.locator('#gps-ficha-titulo').getByText('Escola fictícia 1').waitFor();
      await page.getByText('Município fictício', {exact:true}).waitFor();
      if (papel === 'CEC') await captura('ficha-escola');
      await abrir('Unidades escolares');
      await page.locator('#gps-busca-escola').fill('Escola fictícia 1');
      await page.locator('#gps-escolas-lista').getByRole('button', { name: /^Escola fictícia 1Município|^Escola fictícia 1 Município/ }).first().click();
      await page.locator('#gps-ficha-titulo').getByText('Escola fictícia 1').waitFor();
      await abrir('Agenda pedagógica');
      await page.getByText('1 atividades na data.').waitFor();
      await abrir('Relatórios');
      await page.locator('.farol-vermelho').getByText('Abaixo do Básico').waitFor();
      if (papel === 'CEC') {
        await abrir('Central de alertas');
        await page.getByText('Nenhum alerta pessoal em aberto.').waitFor();
        await page.getByRole('button', { name: 'Consultar todos os alertas' }).click();
        assert.equal(await page.locator('#secao').inputValue(), 'alertas');
        await abrir('Frequência');
        await page.getByText('Evolução no período').waitFor();
        await page.locator('.farol-verde').getByText('90%', {exact:true}).waitFor();
        await captura('frequencia');
        await abrir('Avaliações');
        await page.getByRole('button', {name:'Simulado ENEM',exact:true}).click();
        await page.getByRole('heading', {name:'Simulado ENEM'}).waitFor();
        await captura('avaliacoes');
        await abrir('Plataformas digitais');
        assert.equal(await page.locator('#gps-indicador-catalogo button').count(), 13);
        await captura('plataformas');
      }
      if (papel === 'PEC') {
        await abrir('Planos de Trabalho');
        await page.getByText('Plano de Trabalho', { exact: true }).click();
        const plano = page.locator('#gps-plano');
        await plano.locator('[name=escola_id]').selectOption('E0');
        await plano.locator('[name=pasta_id]').selectOption('P');
        await plano.locator('[name=componente_id]').selectOption('LP');
        await plano.locator('[name=data]').fill('2026-11-25');
        await plano.locator('[name=objetivo]').fill('Atividade sintética');
        await plano.getByRole('button', { name: 'Verificar prioridades' }).click();
        await page.getByText('1 de 1 prioridades da escola contempladas.').waitFor();
        await plano.getByRole('button', { name: 'Salvar atividade' }).click();
        await page.getByText('Registro salvo.').waitFor();
      }
      if (papel === 'AT') assert.equal(await page.getByRole('button', { name: 'Planos de Trabalho' }).count(), 0);
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
      assert.deepEqual(erros, []);
      console.log(papel + ': navegação validada a ' + largura + ' px (RPC simulada)');
      await page.close();
    }
  } finally {
    await browser.close();
  }
})().catch(erro => { console.error(erro); process.exitCode = 1; });
