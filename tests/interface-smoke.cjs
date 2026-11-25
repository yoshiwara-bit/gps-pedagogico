const assert = require('node:assert/strict');
const fs = require('node:fs');
const { chromium } = require('playwright');

function renderizar(papel) {
  const atuacao = fs.readFileSync('src/Atuacao.html', 'utf8');
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
    papel, true, () => atuacao);
}

function instalarRpcFalso(papel) {
  const escolas = Array.from({ length: papel === 'CEC' ? 40 : papel === 'AT' ? 8 : 20 }, (_, i) => ({
    id: 'E' + i, nome: 'Escola fictícia ' + (i + 1), latitude: -10 - i / 20, longitude: -40 + (i % 4) / 6,
    contabilizadas: i ? 0 : 1
  }));
  const resposta = {
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
      await page.getByRole('button', { name: 'Território', exact: true }).click();
      assert.equal(await page.locator('#gps-mapa circle').count(), total);
      await page.locator('#gps-mapa circle').first().click();
      await page.waitForFunction(() => document.querySelectorAll('#gps-mapa circle').length === 1);
      await page.locator('#gps-ficha').getByText('Escola fictícia 1').waitFor();
      const filtrosSobrepostos = await page.locator('#gps-filtros label').evaluateAll(rotulos =>
        rotulos.some(rotulo => rotulo.firstElementChild.getBoundingClientRect().right > rotulo.getBoundingClientRect().right + 1));
      assert.equal(filtrosSobrepostos, false);
      await page.getByRole('button', { name: 'Consultas', exact: true }).click();
      await page.locator('.farol-vermelho').getByText('Abaixo do Básico').waitFor();
      if (papel === 'CEC') {
        await page.getByRole('button', { name: 'Registros e alertas' }).click();
        await page.getByText('Nenhum alerta pessoal em aberto.').waitFor();
        await page.getByRole('button', { name: 'Consultar todos os alertas' }).click();
        assert.equal(await page.locator('#secao').inputValue(), 'alertas');
      }
      if (papel === 'PEC') {
        await page.getByRole('button', { name: 'Registros e alertas' }).click();
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
      if (papel === 'AT') assert.equal(await page.getByRole('button', { name: 'Registros e alertas' }).count(), 0);
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
      assert.deepEqual(erros, []);
      console.log(papel + ': navegação validada a ' + largura + ' px (RPC simulada)');
      await page.close();
    }
  } finally {
    await browser.close();
  }
})().catch(erro => { console.error(erro); process.exitCode = 1; });
