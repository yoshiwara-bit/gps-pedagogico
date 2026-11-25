// Configuração privada: definir GPS_LAB_FOLDER_ID nas propriedades do script.
function obterPastaGPSLAB() {
  const id = PropertiesService.getScriptProperties().getProperty("GPS_LAB_FOLDER_ID");
  if (!id) throw new Error("Configure GPS_LAB_FOLDER_ID em uma pasta de testes isolada.");
  return DriveApp.getFolderById(id);
}
// Logs detalhados desativados: podem conter dados pessoais. Usar somente dados sintéticos.
function registrarGPSLAB(mensagem) {
  if (PropertiesService.getScriptProperties().getProperty("GPS_LAB_LOG_SINTETICO") === "true") {
    Logger.log(mensagem);
  }
}
function obterPastasGPSLAB() {
  const pasta = obterPastaGPSLAB();
  let consumida = false;
  return {hasNext: () => !consumida, next: () => { consumida = true; return pasta; }};
}

function testarIntegracaoDrive() {
  const pasta = DriveApp.createFolder("GPS LAB - TESTE");

  const arquivo = pasta.createFile(
    "teste_integracao.txt",
    "GPS Pedagogico - URE Miracatu"
  );

  registrarGPSLAB("Acesso ao Drive: OK");
  registrarGPSLAB("Criacao de pasta: OK");
  registrarGPSLAB("Criacao de arquivo: OK");
  registrarGPSLAB("Leitura: " + arquivo.getBlob().getDataAsString());
  registrarGPSLAB("Pasta: " + pasta.getUrl());
  registrarGPSLAB("TESTE CONCLUIDO COM SUCESSO!");
}

function testarIdentificacaoDocumentos() {
  const pastas = obterPastasGPSLAB();

  if (!pastas.hasNext()) {
    registrarGPSLAB("Pasta de teste nao encontrada.");
    return;
  }

  const pasta = pastas.next();
  const arquivos = pasta.getFiles();
  let total = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();

    registrarGPSLAB("--------------------------");
    registrarGPSLAB("Nome: " + arquivo.getName());
    registrarGPSLAB("Tipo: " + arquivo.getMimeType());
    registrarGPSLAB("Tamanho: " + arquivo.getSize() + " bytes");
    registrarGPSLAB("ID: " + arquivo.getId());

    total++;
  }

  registrarGPSLAB("--------------------------");
  registrarGPSLAB("Total de arquivos: " + total);
  registrarGPSLAB("TESTE CONCLUIDO!");
}



function testarLeituraDocumentos() {
  const pastas = obterPastasGPSLAB();

  if (!pastas.hasNext()) {
    registrarGPSLAB("Pasta de teste nao encontrada.");
    return;
  }

  const pasta = pastas.next();
  const arquivos = pasta.getFiles();

  let sucessos = 0;
  let erros = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const tipo = arquivo.getMimeType();

    if (
      tipo !== "application/pdf" &&
      tipo !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      continue;
    }

    let temporario = null;
    const inicio = Date.now();

    registrarGPSLAB("==========================");
    registrarGPSLAB("DOCUMENTO: " + arquivo.getName());

    try {
      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_LAB_" + arquivo.getName(),
          mimeType: "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const doc = DocumentApp.openById(temporario.id);
      const corpo = doc.getBody();
      const texto = corpo.getText();

      registrarGPSLAB("Caracteres extraidos: " + texto.length);
      registrarGPSLAB("Tabelas identificadas: " + corpo.getTables().length);

      if (!texto.trim()) {
        throw new Error("Nenhum texto extraido.");
      }

      registrarGPSLAB("LEITURA: OK");
      sucessos++;

    } catch (erro) {
      registrarGPSLAB("ERRO: " + erro.message);
      erros++;

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id).setTrashed(true);
          registrarGPSLAB("Copia temporaria removida.");
        } catch (erroLimpeza) {
          registrarGPSLAB("Aviso: copia temporaria nao removida.");
        }
      }

      registrarGPSLAB("Tempo: " + (Date.now() - inicio) + " ms");
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");
  registrarGPSLAB("Sucessos: " + sucessos);
  registrarGPSLAB("Erros: " + erros);
}

function testarLeituraPlanilhas() {
  const pastas = obterPastasGPSLAB();

  if (!pastas.hasNext()) {
    registrarGPSLAB("Pasta de teste nao encontrada.");
    return;
  }

  const pasta = pastas.next();
  const arquivos = pasta.getFiles();

  let sucessos = 0;
  let erros = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const tipo = arquivo.getMimeType();
    const nome = arquivo.getName();
    const extensao = nome.toLowerCase();

    const excel =
      extensao.endsWith(".xlsx") ||
      extensao.endsWith(".xls");

    const csv = extensao.endsWith(".csv");

    const sheets =
      tipo === "application/vnd.google-apps.spreadsheet";

    if (!excel && !csv && !sheets) {
      continue;
    }

    let temporario = null;
    const inicio = Date.now();

    registrarGPSLAB("==========================");
    registrarGPSLAB("ARQUIVO: " + nome);
    registrarGPSLAB("FORMATO: " + tipo);

    try {
      if (csv) {
        const texto = arquivo.getBlob()
          .getDataAsString("UTF-8")
          .replace(/^\uFEFF/, "");

        const primeiraLinha = texto.split(/\r?\n/)[0];

        const delimitadores = [",", ";", "\t"];
        const delimitador = delimitadores.sort(
          (a, b) =>
            primeiraLinha.split(b).length -
            primeiraLinha.split(a).length
        )[0];

        const dados = Utilities.parseCsv(
          texto, delimitador
        );

        registrarGPSLAB("Linhas: " + dados.length);
        registrarGPSLAB("Colunas: " + dados[0].length);
        registrarGPSLAB("Delimitador: " +
          JSON.stringify(delimitador));

      } else {
        let id = arquivo.getId();

        if (excel) {
          temporario = Drive.Files.create(
            {
              name: "TEMP_GPS_LAB_" + nome,
              mimeType:
                "application/vnd.google-apps.spreadsheet",
              parents: [pasta.getId()]
            },
            arquivo.getBlob(),
            { fields: "id,name" }
          );

          id = temporario.id;
        }

        const planilha = SpreadsheetApp.openById(id);
        const abas = planilha.getSheets();

        registrarGPSLAB("Total de abas: " + abas.length);

        abas.forEach(function(aba) {
          registrarGPSLAB("Aba: " + aba.getName());
          registrarGPSLAB("Linhas: " + aba.getLastRow());
          registrarGPSLAB("Colunas: " + aba.getLastColumn());

          if (aba.getLastRow() > 0 &&
              aba.getLastColumn() > 0) {

            const amostra = aba.getRange(
              1, 1,
              Math.min(5, aba.getLastRow()),
              Math.min(10, aba.getLastColumn())
            ).getDisplayValues();

            registrarGPSLAB("Leitura de celulas: OK");
          }
        });
      }

      registrarGPSLAB("LEITURA: OK");
      sucessos++;

    } catch (erro) {
      registrarGPSLAB("ERRO: " + erro.message);
      erros++;

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB("Copia temporaria removida.");
        } catch (e) {
          registrarGPSLAB("Aviso: falha na limpeza.");
        }
      }

      registrarGPSLAB("Tempo: " +
        (Date.now() - inicio) + " ms");
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");
  registrarGPSLAB("Sucessos: " + sucessos);
  registrarGPSLAB("Erros: " + erros);
}

function testarExtracaoIndicadoresExcel() {
  const pastas = obterPastasGPSLAB();

  if (!pastas.hasNext()) {
    registrarGPSLAB("Pasta nao encontrada.");
    return;
  }

  const pasta = pastas.next();
  const arquivos = pasta.getFiles();

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (!nome.toLowerCase().endsWith(".xlsx")) {
      continue;
    }

    let temporario = null;
    const inicio = Date.now();

    try {
      registrarGPSLAB("==========================");
      registrarGPSLAB("ARQUIVO: " + nome);

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_LAB_" + nome,
          mimeType:
            "application/vnd.google-apps.spreadsheet",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        { fields: "id,name" }
      );

      const planilha = SpreadsheetApp.openById(
        temporario.id
      );

      planilha.getSheets().forEach(function(aba) {
        const intervalo = aba.getDataRange();
        const dados = intervalo.getValues();
        const exibidos = intervalo.getDisplayValues();
        const formatos = intervalo.getNumberFormats();

        registrarGPSLAB("ABA: " + aba.getName());
        registrarGPSLAB("TOTAL DE LINHAS: " + dados.length);

        const cabecalhos = exibidos[0];

        cabecalhos.forEach(function(campo, coluna) {
          registrarGPSLAB(
            "COLUNA " + (coluna + 1) +
            ": " + campo
          );

          if (dados.length > 1) {
            const valor = dados[1][coluna];

            // Nao exibir textos ou identificadores
            // pessoais no registro de execucao.
            const sensivel =
              /nome|aluno|estudante|cpf|rg|ra\b|matricula|matrícula|codigo|código|id\b/i
                .test(campo);

            const seguro =
              !sensivel &&
              typeof valor === "number";

            if (seguro) {
              registrarGPSLAB("  Valor bruto: " + valor);
              registrarGPSLAB(
                "  Valor exibido: " +
                exibidos[1][coluna]
              );
              registrarGPSLAB(
                "  Formato: " +
                formatos[1][coluna]
              );
            } else {
              registrarGPSLAB(
                "  Tipo de dado: " + typeof valor
              );
            }
          }
        });
      });

      registrarGPSLAB("EXTRACAO: OK");

    } catch (erro) {
      registrarGPSLAB("ERRO: " + erro.message);

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id)
            .setTrashed(true);
          registrarGPSLAB("Copia temporaria removida.");
        } catch (e) {
          registrarGPSLAB("Falha na limpeza.");
        }
      }

      registrarGPSLAB(
        "Tempo: " + (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("TESTE FINALIZADO");
}

function testarLeituraImagensAVD() {
  const pastas = obterPastasGPSLAB();

  if (!pastas.hasNext()) {
    registrarGPSLAB("Pasta nao encontrada.");
    return;
  }

  const pasta = pastas.next();
  const arquivos = pasta.getFiles();

  let sucessos = 0;
  let erros = 0;
  let total = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (
  !/^image\/(png|jpeg)$/i.test(arquivo.getMimeType())
) {
  continue;
}

    total++;

    let temporario = null;
    const inicio = Date.now();

    try {
      registrarGPSLAB("==========================");
      registrarGPSLAB("IMAGEM: " + nome);
      registrarGPSLAB(
        "TAMANHO: " + arquivo.getSize() + " bytes"
      );

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_LAB_OCR_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const documento = DocumentApp.openById(
        temporario.id
      );

      const corpo = documento.getBody();
      const texto = corpo.getText();
      const tabelas = corpo.getTables();

      registrarGPSLAB(
        "CARACTERES EXTRAIDOS: " + texto.length
      );

      registrarGPSLAB(
        "TABELAS IDENTIFICADAS: " + tabelas.length
      );

      registrarGPSLAB("----- INICIO DO TEXTO -----");

      // Divide o texto para evitar truncamento
      // de mensagens individuais no registro.
      for (
        let i = 0;
        i < texto.length;
        i += 3000
      ) {
        registrarGPSLAB(texto.substring(i, i + 3000));
      }

      registrarGPSLAB("----- FIM DO TEXTO -----");

      if (texto.trim().length > 0) {
        sucessos++;
        registrarGPSLAB("OCR: TEXTO EXTRAIDO");
      } else {
        erros++;
        registrarGPSLAB("OCR: NENHUM TEXTO EXTRAIDO");
      }

    } catch (erro) {
      erros++;
      registrarGPSLAB("ERRO: " + erro.message);

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB("Copia temporaria removida.");

        } catch (e) {
          registrarGPSLAB(
            "Falha na limpeza: " + e.message
          );
        }
      }

      registrarGPSLAB(
        "TEMPO: " + (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE OCR");
  registrarGPSLAB("IMAGENS ENCONTRADAS: " + total);
  registrarGPSLAB("IMAGENS COM TEXTO: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);
  registrarGPSLAB("TESTE FINALIZADO");
}

function verificarArquivosGPSLAB() {
  const pastas = obterPastasGPSLAB();

  let totalPastas = 0;

  while (pastas.hasNext()) {
    const pasta = pastas.next();
    totalPastas++;

    registrarGPSLAB("==========================");
    registrarGPSLAB("PASTA: " + pasta.getName());
    registrarGPSLAB("ID: " + pasta.getId());

    const arquivos = pasta.getFiles();
    let total = 0;

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      total++;

      registrarGPSLAB("ARQUIVO: " + arquivo.getName());
      registrarGPSLAB("TIPO: " + arquivo.getMimeType());
      registrarGPSLAB("--------------------------");
    }

    registrarGPSLAB("TOTAL DE ARQUIVOS: " + total);
  }

  registrarGPSLAB("TOTAL DE PASTAS: " + totalPastas);
  registrarGPSLAB("VERIFICACAO FINALIZADA");
}

function testarExtracaoPlanoTrabalho() {
  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let sucessos = 0;
  let erros = 0;
  let total = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (
      !/plano\s+de\s+trabalho/i.test(nome) ||
      !/\.docx$/i.test(nome)
    ) {
      continue;
    }

    total++;

    let temporario = null;
    const inicio = Date.now();

    try {
      registrarGPSLAB("==========================");
      registrarGPSLAB("PLANO: " + nome);

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_LAB_PLANO_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          fields: "id,name"
        }
      );

      const documento = DocumentApp.openById(
        temporario.id
      );

      const corpo = documento.getBody();
      const texto = corpo.getText();
      const tabelas = corpo.getTables();

      registrarGPSLAB(
        "CARACTERES: " + texto.length
      );

      registrarGPSLAB(
        "TOTAL DE TABELAS: " + tabelas.length
      );

      registrarGPSLAB("----- IDENTIFICACAO -----");

      registrarGPSLAB(texto.substring(0, 1500));

      let agendas = 0;
      let registros = 0;

      tabelas.forEach(function(tabela, indice) {
        registrarGPSLAB("==========================");

        registrarGPSLAB("TABELA: " + (indice + 1));
        registrarGPSLAB(
          "LINHAS: " + tabela.getNumRows()
        );

        let linhaCabecalho = -1;

        for (
          let i = 0;
          i < tabela.getNumRows();
          i++
        ) {
          const linha = tabela.getRow(i);
          const campos = [];

          for (
            let j = 0;
            j < linha.getNumCells();
            j++
          ) {
            campos.push(
              linha.getCell(j).getText().trim()
            );
          }

          const normalizados = campos.map(
            function(campo) {
              return campo.toLowerCase()
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "");
            }
          );

          if (
            normalizados.some(
              c => /\bdata\b/.test(c)
            ) &&
            normalizados.some(
              c => /\blocal\b/.test(c)
            ) &&
            normalizados.some(
              c => /\bacao\b/.test(c)
            )
          ) {
            linhaCabecalho = i;

            registrarGPSLAB(
              "CABECALHO: " + campos.join(" | ")
            );

            break;
          }
        }

        if (linhaCabecalho === -1) {
          registrarGPSLAB(
            "Tabela nao identificada como agenda."
          );
          return;
        }

        agendas++;

        registrarGPSLAB("AGENDA IDENTIFICADA");

        for (
          let i = linhaCabecalho + 1;
          i < tabela.getNumRows();
          i++
        ) {
          const linha = tabela.getRow(i);
          const campos = [];

          for (
            let j = 0;
            j < linha.getNumCells();
            j++
          ) {
            campos.push(
              linha.getCell(j).getText()
                .replace(/\n/g, " / ")
                .trim()
            );
          }

          if (!campos.some(c => c.length > 0)) {
            continue;
          }

          registros++;

          registrarGPSLAB(
            "REGISTRO " + registros +
            ": " + campos.join(" | ")
          );
        }
      });

      registrarGPSLAB("==========================");
      registrarGPSLAB("AGENDAS IDENTIFICADAS: " + agendas);
      registrarGPSLAB("REGISTROS EXTRAIDOS: " + registros);

      if (agendas > 0 && registros > 0) {
        sucessos++;
        registrarGPSLAB("EXTRACAO: OK");
      } else {
        erros++;
        registrarGPSLAB("EXTRACAO: AGENDA NAO LOCALIZADA");
      }

    } catch (erro) {
      erros++;
      registrarGPSLAB("ERRO: " + erro.message);

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB("Copia temporaria removida.");

        } catch (e) {
          registrarGPSLAB("Falha na limpeza: " + e.message);
        }
      }

      registrarGPSLAB(
        "TEMPO: " + (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");
  registrarGPSLAB("PLANOS ENCONTRADOS: " + total);
  registrarGPSLAB("SUCESSOS: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);
  registrarGPSLAB("TESTE FINALIZADO");
}

function testarExtracaoTermosVisita() {
  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let total = 0;
  let sucessos = 0;
  let erros = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();
    const tipo = arquivo.getMimeType();

    if (
      !/jaire|jairê|sebastiana/i.test(nome) ||
      !/\.(pdf|docx)$/i.test(nome)
    ) {
      continue;
    }

    total++;

    let temporario = null;
    const inicio = Date.now();

    try {
      registrarGPSLAB("==========================");
      registrarGPSLAB("DOCUMENTO: " + nome);
      registrarGPSLAB("FORMATO: " + tipo);

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_LAB_TERMO_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const documento = DocumentApp.openById(
        temporario.id
      );

      const corpo = documento.getBody();
      const texto = corpo.getText();
      const tabelas = corpo.getTables();

      registrarGPSLAB(
        "CARACTERES: " + texto.length
      );

      registrarGPSLAB(
        "TABELAS: " + tabelas.length
      );

      registrarGPSLAB("----- TEXTO EXTRAIDO -----");

      for (
        let i = 0;
        i < texto.length;
        i += 2500
      ) {
        registrarGPSLAB(
          texto.substring(i, i + 2500)
        );
      }

      registrarGPSLAB("----- ESTRUTURA DAS TABELAS -----");

      tabelas.forEach(function(tabela, indice) {
        registrarGPSLAB(
          "TABELA " + (indice + 1)
        );

        for (
          let i = 0;
          i < tabela.getNumRows();
          i++
        ) {
          const linha = tabela.getRow(i);
          const campos = [];

          for (
            let j = 0;
            j < linha.getNumCells();
            j++
          ) {
            campos.push(
              linha.getCell(j).getText()
                .replace(/\n/g, " / ")
                .trim()
            );
          }

          registrarGPSLAB(
            "LINHA " + (i + 1) +
            ": " + campos.join(" | ")
          );
        }
      });

      if (texto.trim().length > 0) {
        sucessos++;
        registrarGPSLAB("EXTRACAO: OK");
      } else {
        erros++;
        registrarGPSLAB("EXTRACAO: TEXTO VAZIO");
      }

    } catch (erro) {
      erros++;
      registrarGPSLAB("ERRO: " + erro.message);

    } finally {
      if (temporario && temporario.id) {
        try {
          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB("Copia temporaria removida.");

        } catch (e) {
          registrarGPSLAB("Falha na limpeza: " + e.message);
        }
      }

      registrarGPSLAB(
        "TEMPO: " + (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");
  registrarGPSLAB("TERMOS ENCONTRADOS: " + total);
  registrarGPSLAB("SUCESSOS: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);
  registrarGPSLAB("TESTE FINALIZADO");
}

function testarDuplicidadeVisitas() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();
  const documentos = [];
  const visitas = [];
  const pendencias = [];

  function normalizar(texto) {
    return texto.toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9 ]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (
      !/jaire|jairê|sebastiana/i.test(nome) ||
      !/\.(pdf|docx)$/i.test(nome)
    ) {
      continue;
    }

    let temporario = null;

    try {

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_DUP_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const texto = DocumentApp
        .openById(temporario.id)
        .getBody()
        .getText();

      const registro = texto.match(
        /COMPROVANTE\s+REGISTRO\s+DE\s+VISITA\s*[-–]\s*(\d+)/i
      );

      const pec = texto.match(
        /NOME:\s*(.*?)\s*CPF:/i
      );

      const data = texto.match(
        /DATA\s*VISITA:\s*(\d{2}\/\d{2}\/\d{4})/i
      );

      const escola = texto.match(
        /ESCOLA:\s*(.*?)\s*STATUS:/i
      );

      const relato = texto.match(
        /AÇÃO\s+E\s+ENCAMINHAMENTO:\s*([\s\S]*?)(?=TEMA\(S\)|$)/i
      );

      documentos.push({
        id: arquivo.getId(),
        nome: nome,
        tipo: registro ? "FICHA" : "DETALHADO",
        registro: registro ? registro[1] : "",
        pec: pec ? pec[1].trim() : "",
        data: data ? data[1] : "",
        escola: escola ? escola[1].trim() : "",
        texto: normalizar(texto),
        relato: normalizar(
          relato ? relato[1] : texto
        )
      });

      registrarGPSLAB("DOCUMENTO LIDO: " + nome);

    } catch (erro) {

      registrarGPSLAB(
        "ERRO: " + nome + " - " + erro.message
      );

    } finally {

      if (temporario && temporario.id) {
        DriveApp.getFileById(temporario.id)
          .setTrashed(true);
      }
    }
  }

  // Primeiro: registrar as fichas oficiais.

  documentos
    .filter(d => d.tipo === "FICHA")
    .forEach(function(d) {

      if (!d.pec || !d.data || !d.escola) {
        pendencias.push(d.nome);
        return;
      }

      const existente = visitas.find(
        v => v.registro === d.registro
      );

      if (existente) {

        if (
          normalizar(existente.pec) ===
            normalizar(d.pec) &&
          existente.data === d.data &&
          normalizar(existente.escola) ===
            normalizar(d.escola)
        ) {
          existente.documentos.push(d);
        } else {
          pendencias.push(d.nome);
        }

      } else {

        visitas.push({
          registro: d.registro,
          pec: d.pec,
          data: d.data,
          escola: d.escola,
          relato: d.relato,
          documentos: [d]
        });
      }
    });

  // Segundo: localizar os termos correspondentes.

  documentos
    .filter(d => d.tipo === "DETALHADO")
    .forEach(function(d) {

      const candidatos = visitas.filter(
        function(v) {

          const mesmaEscola = d.texto.includes(
            normalizar(v.escola)
          );

          const partes = v.data.split("/");

          const meses = [
            "janeiro", "fevereiro", "marco",
            "abril", "maio", "junho",
            "julho", "agosto", "setembro",
            "outubro", "novembro", "dezembro"
          ];

          const dataExtenso =
            Number(partes[0]) + " de " +
            meses[Number(partes[1]) - 1] +
            " de " + partes[2];

          const mesmaData =
            d.texto.includes(
              normalizar(v.data)
            ) ||
            d.texto.includes(dataExtenso);

          const palavras = d.relato.split(" ");

          let trechos = 0;
          let coincidencias = 0;

          for (
            let i = 0;
            i + 12 <= palavras.length;
            i += 12
          ) {

            const trecho = palavras
              .slice(i, i + 12)
              .join(" ");

            trechos++;

            if (v.relato.includes(trecho)) {
              coincidencias++;
            }
          }

          const similaridade = trechos > 0
            ? coincidencias / trechos
            : 0;

          registrarGPSLAB(
            "COMPARACAO: " + d.nome +
            " x " + v.escola +
            " | Escola: " + mesmaEscola +
            " | Data: " + mesmaData +
            " | Similaridade: " +
            (similaridade * 100).toFixed(1) +
            "%"
          );

          return mesmaEscola &&
            mesmaData &&
            similaridade >= 0.70;
        }
      );

      if (candidatos.length === 1) {

        candidatos[0].documentos.push(d);

        registrarGPSLAB(
          "VINCULACAO: " + d.nome +
          " -> Registro " +
          candidatos[0].registro
        );

      } else {

        pendencias.push(d.nome);

        registrarGPSLAB(
          "VALIDACAO MANUAL NECESSARIA: " +
          d.nome
        );
      }
    });

  // Terceiro: apresentar o resultado.

  let completos = 0;
  let incompletos = 0;

  visitas.forEach(function(v) {

    const temFicha = v.documentos.some(
      d => d.tipo === "FICHA"
    );

    const temDetalhado = v.documentos.some(
      d => d.tipo === "DETALHADO"
    );

    const completo = temFicha && temDetalhado;

    if (completo) {
      completos++;
    } else {
      incompletos++;
    }

    registrarGPSLAB("==========================");
    registrarGPSLAB("PEC: " + v.pec);
    registrarGPSLAB("ESCOLA: " + v.escola);
    registrarGPSLAB("DATA: " + v.data);
    registrarGPSLAB("REGISTRO: " + v.registro);
    registrarGPSLAB("VISITAS CONTABILIZADAS: 1");

    registrarGPSLAB(
      "DOCUMENTOS VINCULADOS: " +
      v.documentos.length
    );

    v.documentos.forEach(function(d) {
      registrarGPSLAB(d.tipo + ": " + d.nome);
    });

    registrarGPSLAB(
      "DOCUMENTACAO: " +
      (completo ? "COMPLETA" : "INCOMPLETA")
    );
  });

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");

  registrarGPSLAB(
    "DOCUMENTOS PROCESSADOS: " +
    documentos.length
  );

  registrarGPSLAB(
    "VISITAS DISTINTAS: " + visitas.length
  );

  registrarGPSLAB(
    "DOCUMENTACAO COMPLETA: " + completos
  );

  registrarGPSLAB(
    "DOCUMENTACAO INCOMPLETA: " + incompletos
  );

  registrarGPSLAB(
    "PENDENCIAS DE VALIDACAO: " +
    pendencias.length
  );

  pendencias.forEach(function(p) {
    registrarGPSLAB("PENDENTE: " + p);
  });

  registrarGPSLAB("TESTE FINALIZADO");
}

function testarExtracaoSARESP() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let total = 0;
  let sucessos = 0;
  let erros = 0;

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (
  !/2025/i.test(nome) ||
  !/\.pdf$/i.test(nome) ||
  /resultado final|saeb|temp_/i.test(nome)
) {
  continue;
}

    total++;

    let temporario = null;
    const inicio = Date.now();

    try {

      registrarGPSLAB("==========================");
      registrarGPSLAB("BOLETIM SARESP: " + nome);

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_SARESP_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const documento = DocumentApp.openById(
        temporario.id
      );

      const corpo = documento.getBody();
      const texto = corpo.getText();
      const tabelas = corpo.getTables();

      registrarGPSLAB(
        "CARACTERES: " + texto.length
      );

      registrarGPSLAB(
        "TABELAS: " + tabelas.length
      );

      registrarGPSLAB("----- TEXTO INTEGRAL -----");

      for (
        let i = 0;
        i < texto.length;
        i += 2500
      ) {
        registrarGPSLAB(
          texto.substring(i, i + 2500)
        );
      }

      registrarGPSLAB("----- IDENTIFICACAO -----");

      const escola = texto.match(
        /\b(\d{6})\s*[-–]\s*([^\n]+)/i
      );

      if (escola) {

        registrarGPSLAB(
          "CODIGO ESCOLA: " + escola[1]
        );

        registrarGPSLAB(
          "NOME ESCOLA: " + escola[2].trim()
        );

      } else {

        registrarGPSLAB(
          "ESCOLA: NAO IDENTIFICADA"
        );
      }

      registrarGPSLAB("----- LINHAS EXTRAIDAS -----");

      const linhas = texto.split(/\r?\n/);

      linhas.forEach(function(linha, indice) {

        if (linha.trim()) {

          registrarGPSLAB(
            "LINHA " +
            (indice + 1) +
            ": " + linha.trim()
          );
        }
      });

      registrarGPSLAB("----- NOTAS CANDIDATAS -----");

      const notas = texto.match(
        /\b(?:10[,.]0|[0-9][,.][0-9])\b(?![\d%])/g
      ) || [];

      notas.forEach(function(nota, indice) {

        registrarGPSLAB(
          "NOTA " + (indice + 1) +
          ": " + nota
        );
      });

      registrarGPSLAB("----- PERCENTUAIS -----");

      const percentuais = texto.match(
        /\b\d{1,3}(?:[,.]\d+)?\s*%/g
      ) || [];

      percentuais.forEach(
        function(percentual, indice) {

          registrarGPSLAB(
            "PERCENTUAL " +
            (indice + 1) +
            ": " + percentual
          );
        }
      );

      registrarGPSLAB("----- GRUPO DA ESCOLA -----");

      const grupo = texto.match(
        /Escola de per[ií]odo[^\n]*/i
      );

      registrarGPSLAB(
        grupo
          ? grupo[0]
          : "GRUPO NAO IDENTIFICADO"
      );

      registrarGPSLAB("----- ESTRUTURA DOCUMENTAL -----");

      for (
        let i = 0;
        i < corpo.getNumChildren();
        i++
      ) {

        const elemento = corpo.getChild(i);

        registrarGPSLAB(
          "ELEMENTO " + (i + 1) +
          ": " + elemento.getType()
        );
      }

      tabelas.forEach(
        function(tabela, indice) {

          registrarGPSLAB(
            "TABELA " + (indice + 1)
          );

          for (
            let i = 0;
            i < tabela.getNumRows();
            i++
          ) {

            const linha = tabela.getRow(i);
            const campos = [];

            for (
              let j = 0;
              j < linha.getNumCells();
              j++
            ) {

              campos.push(
                linha.getCell(j)
                  .getText()
                  .trim()
              );
            }

            registrarGPSLAB(
              "LINHA " + (i + 1) +
              ": " + campos.join(" | ")
            );
          }
        }
      );

      registrarGPSLAB("==========================");
      registrarGPSLAB(
        "NOTAS CANDIDATAS: " + notas.length
      );

      registrarGPSLAB(
        "PERCENTUAIS ENCONTRADOS: " +
        percentuais.length
      );

      if (texto.trim().length > 0) {

        sucessos++;

        registrarGPSLAB("LEITURA: OK");

        registrarGPSLAB(
          "VALIDACAO SEMANTICA: PENDENTE"
        );

      } else {

        erros++;

        registrarGPSLAB("LEITURA: TEXTO VAZIO");
      }

    } catch (erro) {

      erros++;

      registrarGPSLAB("ERRO: " + erro.message);

    } finally {

      if (temporario && temporario.id) {

        try {

          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB(
            "Copia temporaria removida."
          );

        } catch (e) {

          registrarGPSLAB(
            "Falha na limpeza: " + e.message
          );
        }
      }

      registrarGPSLAB(
        "TEMPO: " +
        (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE");
  registrarGPSLAB("BOLETINS ENCONTRADOS: " + total);
  registrarGPSLAB("SUCESSOS DE LEITURA: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);
  registrarGPSLAB("TESTE FINALIZADO");
}

function testarExtracaoSAEB() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let total = 0;
  let sucessos = 0;
  let erros = 0;

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (
      !/resultado final.*sylas/i.test(nome) ||
      !/\.pdf$/i.test(nome)
    ) {
      continue;
    }

    total++;

    let temporario = null;
    const inicio = Date.now();

    try {

      registrarGPSLAB("==========================");
      registrarGPSLAB("BOLETIM SAEB: " + nome);

      temporario = Drive.Files.create(
        {
          name: "TEMP_GPS_SAEB_" + nome,
          mimeType:
            "application/vnd.google-apps.document",
          parents: [pasta.getId()]
        },
        arquivo.getBlob(),
        {
          ocrLanguage: "pt",
          fields: "id,name"
        }
      );

      const documento = DocumentApp.openById(
        temporario.id
      );

      const corpo = documento.getBody();
      const texto = corpo.getText();
      const tabelas = corpo.getTables();

      registrarGPSLAB(
        "CARACTERES: " + texto.length
      );

      registrarGPSLAB(
        "TABELAS: " + tabelas.length
      );

      registrarGPSLAB("----- TEXTO INTEGRAL -----");

      for (
        let i = 0;
        i < texto.length;
        i += 2500
      ) {

        registrarGPSLAB(
          texto.substring(i, i + 2500)
        );
      }

      const linhas = texto.split(/\r?\n/);

      const categorias = [
        {
          nome: "IDENTIFICACAO DA ESCOLA",
          regex: /inep|c[oó]digo da escola|sylas/i
        },
        {
          nome: "PARTICIPACAO",
          regex: /participa[cç][aã]o|presentes|previstos|censo escolar/i
        },
        {
          nome: "PROFICIENCIA",
          regex: /profici[eê]ncia|m[eé]dia de desempenho/i
        },
        {
          nome: "LINGUA PORTUGUESA",
          regex: /l[ií]ngua portuguesa/i
        },
        {
          nome: "MATEMATICA",
          regex: /matem[aá]tica/i
        },
        {
          nome: "NIVEIS DE PROFICIENCIA",
          regex: /n[ií]vel|n[ií]veis|escala/i
        },
        {
          nome: "SERIE HISTORICA",
          regex: /hist[oó]ric|evolu[cç][aã]o|2011|2013|2015|2017|2019|2021|2023|2025/i
        },
        {
          nome: "INDICADORES CONTEXTUAIS",
          regex: /socioecon[oô]mico|inse|forma[cç][aã]o docente|adequa[cç][aã]o/i
        },
        {
          nome: "COMPARATIVOS",
          regex: /munic[ií]pio|estadual|estado|brasil|escolas similares/i
        }
      ];

      categorias.forEach(function(categoria) {

        registrarGPSLAB("==========================");
        registrarGPSLAB(categoria.nome);

        let encontrados = 0;

        linhas.forEach(function(linha, indice) {

          if (categoria.regex.test(linha)) {

            encontrados++;

            registrarGPSLAB(
              "LINHA " + (indice + 1) +
              ": " + linha.trim()
            );

            if (indice + 1 < linhas.length) {

              registrarGPSLAB(
                "CONTEXTO: " +
                linhas[indice + 1].trim()
              );
            }
          }
        });

        registrarGPSLAB(
          "OCORRENCIAS: " + encontrados
        );
      });

      registrarGPSLAB("----- PERCENTUAIS -----");

      const percentuais = texto.match(
        /\b\d{1,3}(?:[,.]\d+)?\s*%/g
      ) || [];

      registrarGPSLAB(
        "TOTAL: " + percentuais.length
      );

      registrarGPSLAB(
        percentuais.join(" | ")
      );

      registrarGPSLAB("----- VALORES DECIMAIS -----");

      const decimais = texto.match(
        /\b\d+[,.]\d+\b/g
      ) || [];

      registrarGPSLAB(
        "TOTAL: " + decimais.length
      );

      registrarGPSLAB(
        decimais.join(" | ")
      );

      registrarGPSLAB("----- ESTRUTURA DOCUMENTAL -----");

      const estrutura = {};

      for (
        let i = 0;
        i < corpo.getNumChildren();
        i++
      ) {

        const tipo = String(
          corpo.getChild(i).getType()
        );

        estrutura[tipo] =
          (estrutura[tipo] || 0) + 1;
      }

      registrarGPSLAB(
        JSON.stringify(estrutura)
      );

      registrarGPSLAB("----- IMAGENS INCORPORADAS -----");

      registrarGPSLAB(
        "IMAGENS: " + corpo.getImages().length
      );

      registrarGPSLAB("----- TABELAS EXTRAIDAS -----");

      tabelas.forEach(
        function(tabela, indice) {

          registrarGPSLAB(
            "TABELA " + (indice + 1)
          );

          for (
            let i = 0;
            i < tabela.getNumRows();
            i++
          ) {

            const linha = tabela.getRow(i);
            const campos = [];

            for (
              let j = 0;
              j < linha.getNumCells();
              j++
            ) {

              campos.push(
                linha.getCell(j)
                  .getText()
                  .trim()
              );
            }

            registrarGPSLAB(
              "LINHA " + (i + 1) +
              ": " + campos.join(" | ")
            );
          }
        }
      );

      if (texto.trim().length > 0) {

        sucessos++;

        registrarGPSLAB("LEITURA: OK");

        registrarGPSLAB(
          "VALIDACAO DOS INDICADORES: PENDENTE"
        );

      } else {

        erros++;

        registrarGPSLAB("LEITURA: TEXTO VAZIO");
      }

    } catch (erro) {

      erros++;

      registrarGPSLAB("ERRO: " + erro.message);

    } finally {

      if (temporario && temporario.id) {

        try {

          DriveApp.getFileById(temporario.id)
            .setTrashed(true);

          registrarGPSLAB(
            "Copia temporaria removida."
          );

        } catch (e) {

          registrarGPSLAB(
            "Falha na limpeza: " + e.message
          );
        }
      }

      registrarGPSLAB(
        "TEMPO: " +
        (Date.now() - inicio) + " ms"
      );
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DO TESTE SAEB");
  registrarGPSLAB("BOLETINS ENCONTRADOS: " + total);
  registrarGPSLAB("SUCESSOS DE LEITURA: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);
  registrarGPSLAB("TESTE FINALIZADO");
}

function testarOCRImagensSAEB() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let original = null;

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();

    if (
      /resultado final.*sylas/i.test(
        arquivo.getName()
      ) &&
      /\.pdf$/i.test(arquivo.getName())
    ) {

      original = arquivo;
      break;
    }
  }

  if (!original) {

    registrarGPSLAB("BOLETIM NAO ENCONTRADO");
    return;
  }

  let documentoTemporario = null;

  let total = 0;
  let sucessos = 0;
  let vazios = 0;
  let erros = 0;

  const inicio = Date.now();

  try {

    documentoTemporario = Drive.Files.create(
      {
        name: "TEMP_SAEB_IMAGENS",
        mimeType:
          "application/vnd.google-apps.document",
        parents: [pasta.getId()]
      },
      original.getBlob(),
      {
        ocrLanguage: "pt",
        fields: "id,name"
      }
    );

    const documento = DocumentApp.openById(
      documentoTemporario.id
    );

    const imagens = documento
      .getBody()
      .getImages();

    total = imagens.length;

    registrarGPSLAB("==========================");
    registrarGPSLAB("BOLETIM: " + original.getName());
    registrarGPSLAB("IMAGENS ENCONTRADAS: " + total);

    for (let i = 0; i < imagens.length; i++) {

      let tempOCR = null;

      const inicioImagem = Date.now();

      try {

        const imagem = imagens[i];

        const blob = imagem.getBlob();

        registrarGPSLAB("==========================");
        registrarGPSLAB("IMAGEM: " + (i + 1));
        registrarGPSLAB("LARGURA: " + imagem.getWidth());
        registrarGPSLAB("ALTURA: " + imagem.getHeight());
        registrarGPSLAB("TIPO: " + blob.getContentType());
        registrarGPSLAB("BYTES: " + blob.getBytes().length);

        tempOCR = Drive.Files.create(
          {
            name: "TEMP_SAEB_OCR_" + (i + 1),
            mimeType:
              "application/vnd.google-apps.document",
            parents: [pasta.getId()]
          },
          blob,
          {
            ocrLanguage: "pt",
            fields: "id,name"
          }
        );

        const texto = DocumentApp
          .openById(tempOCR.id)
          .getBody()
          .getText()
          .trim();

        registrarGPSLAB(
          "CARACTERES: " + texto.length
        );

        registrarGPSLAB("----- TEXTO OCR -----");

        for (
          let j = 0;
          j < texto.length;
          j += 2500
        ) {

          registrarGPSLAB(
            texto.substring(j, j + 2500)
          );
        }

        const percentuais = texto.match(
          /\d+(?:[,.]\d+)?\s*%/g
        ) || [];

        const decimais = texto.match(
          /\b\d+[,.]\d+\b/g
        ) || [];

        registrarGPSLAB(
          "PERCENTUAIS: " +
          percentuais.join(" | ")
        );

        registrarGPSLAB(
          "DECIMAIS: " +
          decimais.join(" | ")
        );

        if (texto.length > 0) {

          sucessos++;
          registrarGPSLAB("OCR: TEXTO RECUPERADO");

        } else {

          vazios++;
          registrarGPSLAB("OCR: SEM TEXTO");
        }

      } catch (erro) {

        erros++;

        registrarGPSLAB(
          "ERRO NA IMAGEM: " + erro.message
        );

      } finally {

        if (tempOCR && tempOCR.id) {

          try {

            DriveApp.getFileById(tempOCR.id)
              .setTrashed(true);

          } catch (e) {

            registrarGPSLAB(
              "ERRO NA LIMPEZA: " + e.message
            );
          }
        }

        registrarGPSLAB(
          "TEMPO IMAGEM: " +
          (Date.now() - inicioImagem) + " ms"
        );
      }
    }

  } catch (erro) {

    registrarGPSLAB(
      "ERRO GERAL: " + erro.message
    );

  } finally {

    if (
      documentoTemporario &&
      documentoTemporario.id
    ) {

      try {

        DriveApp.getFileById(
          documentoTemporario.id
        ).setTrashed(true);

        registrarGPSLAB(
          "DOCUMENTO TEMPORARIO REMOVIDO"
        );

      } catch (e) {

        registrarGPSLAB(
          "ERRO NA LIMPEZA: " + e.message
        );
      }
    }
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO OCR SAEB");
  registrarGPSLAB("IMAGENS ENCONTRADAS: " + total);
  registrarGPSLAB("IMAGENS COM TEXTO: " + sucessos);
  registrarGPSLAB("IMAGENS SEM TEXTO: " + vazios);
  registrarGPSLAB("ERROS OCR: " + erros);
  registrarGPSLAB(
    "TEMPO TOTAL: " + (Date.now() - inicio) + " ms"
  );
  registrarGPSLAB("VALIDACAO SEMANTICA: PENDENTE");
  registrarGPSLAB("TESTE FINALIZADO");
}

function testarImagensRestantesSAEB() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let original = null;

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();

    if (
      /resultado final.*sylas/i.test(
        arquivo.getName()
      ) &&
      /\.pdf$/i.test(arquivo.getName())
    ) {

      original = arquivo;
      break;
    }
  }

  if (!original) {
    registrarGPSLAB("BOLETIM NAO ENCONTRADO");
    return;
  }

  let temporario = null;

  try {

    temporario = Drive.Files.create(
      {
        name: "TEMP_SAEB_RETESTE",
        mimeType:
          "application/vnd.google-apps.document",
        parents: [pasta.getId()]
      },
      original.getBlob(),
      {
        ocrLanguage: "pt",
        fields: "id,name"
      }
    );

    Utilities.sleep(2000);

    const imagens = DocumentApp
      .openById(temporario.id)
      .getBody()
      .getImages();

    registrarGPSLAB(
      "IMAGENS ENCONTRADAS: " + imagens.length
    );

    const indices = [16, 15];

    for (const indice of indices) {

      let tempOCR = null;

      registrarGPSLAB("==========================");
      registrarGPSLAB("IMAGEM: " + (indice + 1));

      try {

        if (indice >= imagens.length) {
          registrarGPSLAB("IMAGEM INDISPONIVEL");
          continue;
        }

        const imagem = imagens[indice];

        const blob = imagem.getBlob();

        registrarGPSLAB(
          "BYTES: " + blob.getBytes().length
        );

        tempOCR = Drive.Files.create(
          {
            name: "TEMP_SAEB_RETESTE_" +
              (indice + 1),
            mimeType:
              "application/vnd.google-apps.document",
            parents: [pasta.getId()]
          },
          blob,
          {
            ocrLanguage: "pt",
            fields: "id,name"
          }
        );

        const texto = DocumentApp
          .openById(tempOCR.id)
          .getBody()
          .getText();

        registrarGPSLAB("TEXTO OCR:");
        registrarGPSLAB(texto);

        const percentuais = texto.match(
          /\d+(?:[,.]\d+)?\s*%/g
        ) || [];

        registrarGPSLAB(
          "PERCENTUAIS: " +
          percentuais.join(" | ")
        );

      } catch (erro) {

        registrarGPSLAB("ERRO: " + erro.message);

      } finally {

        if (tempOCR && tempOCR.id) {

          try {

            DriveApp.getFileById(tempOCR.id)
              .setTrashed(true);

          } catch (e) {

            registrarGPSLAB(
              "ERRO NA LIMPEZA: " + e.message
            );
          }
        }
      }
    }

  } catch (erro) {

    registrarGPSLAB("ERRO GERAL: " + erro.message);

  } finally {

    if (temporario && temporario.id) {

      try {

        DriveApp.getFileById(temporario.id)
          .setTrashed(true);

        registrarGPSLAB("TEMPORARIO REMOVIDO");

      } catch (e) {

        registrarGPSLAB(
          "ERRO NA LIMPEZA: " + e.message
        );
      }
    }
  }

  registrarGPSLAB("TESTE FINALIZADO");
}

function testarCSVGoogleSheets() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let totalCSV = 0;
  let totalSheets = 0;
  let sucessos = 0;
  let erros = 0;
  let avdEncontrada = false;

  const inicioGeral = Date.now();

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();
    const nome = arquivo.getName();
    const tipo = arquivo.getMimeType();

    if (/avd\s*2\.xlsx$/i.test(nome)) {

      avdEncontrada = true;

      registrarGPSLAB(
        "PLANILHA AVD ENCONTRADA: " + nome
      );
    }

    const ehCSV =
      /\.csv$/i.test(nome) ||
      tipo === "text/csv";

    const ehSheets =
      tipo ===
      "application/vnd.google-apps.spreadsheet";

    if (!ehCSV && !ehSheets) {
      continue;
    }

    const inicio = Date.now();

    registrarGPSLAB("==========================");
    registrarGPSLAB("ARQUIVO: " + nome);
    registrarGPSLAB("TIPO: " + tipo);

    try {

      if (ehCSV) {

        totalCSV++;

        const texto = arquivo
          .getBlob()
          .getDataAsString("UTF-8")
          .replace(/^\uFEFF/, "");

        if (!texto.trim()) {
          throw new Error("CSV vazio");
        }

        const separadores = [",", ";", "\t"];

        let melhor = null;

        for (const separador of separadores) {

          try {

            const dados = Utilities.parseCsv(
              texto,
              separador
            );

            const colunas = dados[0].length;

            const consistentes = dados.filter(
              linha => linha.length === colunas
            ).length;

            const proporcao =
              consistentes / dados.length;

            const pontuacao =
              colunas * proporcao;

            if (
              !melhor ||
              pontuacao > melhor.pontuacao
            ) {

              melhor = {
                separador: separador,
                dados: dados,
                pontuacao: pontuacao,
                proporcao: proporcao
              };
            }

          } catch (e) {
            // Testa o proximo separador.
          }
        }

        if (!melhor) {
          throw new Error(
            "Nao foi possivel interpretar o CSV"
          );
        }

        const dados = melhor.dados;

        registrarGPSLAB("FORMATO: CSV");

        registrarGPSLAB(
          "SEPARADOR: " +
          JSON.stringify(melhor.separador)
        );

        registrarGPSLAB(
          "LINHAS: " + dados.length
        );

        registrarGPSLAB(
          "COLUNAS: " + dados[0].length
        );

        registrarGPSLAB(
          "CONSISTENCIA: " +
          (melhor.proporcao * 100).toFixed(2) +
          "%"
        );

        registrarGPSLAB(
          "PRIMEIRA LINHA: " +
          dados[0].join(" | ")
        );

        if (dados[0].length === 1) {

          registrarGPSLAB(
            "ATENCAO: Apenas uma coluna detectada"
          );
        }
      }

      if (ehSheets) {

        totalSheets++;

        const planilha = SpreadsheetApp
          .openById(arquivo.getId());

        const abas = planilha.getSheets();

        registrarGPSLAB("FORMATO: GOOGLE SHEETS");

        registrarGPSLAB(
          "TOTAL DE ABAS: " + abas.length
        );

        for (const aba of abas) {

          const linhas = aba.getLastRow();
          const colunas = aba.getLastColumn();

          registrarGPSLAB("--------------------------");
          registrarGPSLAB("ABA: " + aba.getName());
          registrarGPSLAB("LINHAS: " + linhas);
          registrarGPSLAB("COLUNAS: " + colunas);

          if (linhas === 0 || colunas === 0) {

            registrarGPSLAB("ABA VAZIA");
            continue;
          }

          const intervalo = aba.getRange(
            1,
            1,
            Math.min(linhas, 5),
            Math.min(colunas, 30)
          );

          const valores = intervalo.getValues();

          const exibidos =
            intervalo.getDisplayValues();

          const formatos =
            intervalo.getNumberFormats();

          registrarGPSLAB(
            "PRIMEIRA LINHA: " +
            exibidos[0].join(" | ")
          );

          const tipos = {};

          valores.forEach(function(linha) {

            linha.forEach(function(valor) {

              if (valor === "") {
                return;
              }

              const tipoValor =
                valor instanceof Date
                  ? "data"
                  : typeof valor;

              tipos[tipoValor] =
                (tipos[tipoValor] || 0) + 1;
            });
          });

          registrarGPSLAB(
            "TIPOS NA AMOSTRA: " +
            JSON.stringify(tipos)
          );

          const formatosUnicos = [
            ...new Set(formatos.flat())
          ];

          registrarGPSLAB(
            "FORMATOS NA AMOSTRA: " +
            formatosUnicos.join(" | ")
          );
        }
      }

      sucessos++;

      registrarGPSLAB("LEITURA: OK");

    } catch (erro) {

      erros++;

      registrarGPSLAB("ERRO: " + erro.message);
    }

    registrarGPSLAB(
      "TEMPO: " + (Date.now() - inicio) + " ms"
    );
  }

  registrarGPSLAB("==========================");
  registrarGPSLAB("RESUMO DA ETAPA 12");

  registrarGPSLAB(
    "ARQUIVOS CSV: " + totalCSV
  );

  registrarGPSLAB(
    "ARQUIVOS GOOGLE SHEETS: " + totalSheets
  );

  registrarGPSLAB("SUCESSOS: " + sucessos);
  registrarGPSLAB("ERROS: " + erros);

  registrarGPSLAB(
    "AVD 2.XLSX: " +
    (avdEncontrada ? "ENCONTRADA" : "NAO ENCONTRADA")
  );

  registrarGPSLAB(
    "TEMPO TOTAL: " +
    (Date.now() - inicioGeral) + " ms"
  );

  registrarGPSLAB("TESTE FINALIZADO");
}

function testarExtracaoAVDExcel() {

  const pasta = obterPastaGPSLAB();

  const arquivos = pasta.getFiles();

  let original = null;

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();

    if (/^avd\s*2\.xlsx$/i.test(
      arquivo.getName()
    )) {

      original = arquivo;
      break;
    }
  }

  if (!original) {

    registrarGPSLAB("PLANILHA AVD NAO ENCONTRADA");
    return;
  }

  const inicio = Date.now();

  let temporario = null;

  let escolas = 0;
  let validas = 0;
  let inconsistentes = 0;

  const codigos = new Set();

  try {

    temporario = Drive.Files.create(
      {
        name: "TEMP_GPS_AVD",
        mimeType:
          "application/vnd.google-apps.spreadsheet",
        parents: [pasta.getId()]
      },
      original.getBlob(),
      {
        fields: "id,name"
      }
    );

    const planilha = SpreadsheetApp
      .openById(temporario.id);

    const aba = planilha.getSheetByName("Export");

    if (!aba) {

      throw new Error("ABA EXPORT NAO ENCONTRADA");
    }

    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();

    registrarGPSLAB("==========================");
    registrarGPSLAB("ARQUIVO: " + original.getName());
    registrarGPSLAB("ABA: " + aba.getName());
    registrarGPSLAB("LINHAS: " + ultimaLinha);
    registrarGPSLAB("COLUNAS: " + ultimaColuna);

    if (ultimaColuna < 22 || ultimaLinha < 8) {

      throw new Error(
        "ESTRUTURA DIFERENTE DO MODELO AVD"
      );
    }

    const dados = aba.getRange(
      1, 1, ultimaLinha, 22
    ).getValues();

    registrarGPSLAB("----- CABECALHOS -----");

    [1, 3, 5].forEach(function(i) {

      registrarGPSLAB(
        "LINHA " + (i + 1) + ": " +
        dados[i].join(" | ")
      );
    });

    const normalizar = texto => String(texto)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();

    const cabecalho = normalizar(
      dados.slice(0, 7).flat().join(" ")
    );

    const termosEsperados = [
      "participacao",
      "basico",
      "proficiente"
    ];

    if (!termosEsperados.every(
      termo => cabecalho.includes(termo)
    )) {

      throw new Error(
        "CABECALHO NAO COMPATIVEL"
      );
    }

    const percentual = valor => {

      if (
        typeof valor !== "number" ||
        !Number.isFinite(valor)
      ) {
        return "INVALIDO";
      }

      return (valor * 100)
        .toFixed(2)
        .replace(".", ",") + "%";
    };

    const validarPercentual = valor =>
      typeof valor === "number" &&
      Number.isFinite(valor) &&
      valor >= 0 &&
      valor <= 1;

    const validarDistribuicao = valores => {

      if (!valores.every(validarPercentual)) {
        return false;
      }

      const soma = valores.reduce(
        (a, b) => a + b, 0
      );

      return Math.abs(soma - 1) <= 0.005;
    };

    const mostrarDistribuicao = valores =>
      valores.map(percentual).join(" | ");

    for (let i = 7; i < dados.length; i++) {

      const linha = dados[i];

      const identificacao =
        String(linha[1]).trim();

      const escola = identificacao.match(
        /^(.*?)\s*-\s*(\d+)\s*$/
      );

      if (!escola) {
        continue;
      }

      escolas++;

      const nome = escola[1].trim();
      const codigo = escola[2];

      const matriculados = linha[2];
      const participacao = linha[3];

      const lpAvd1 = [
        linha[6],
        linha[7],
        linha[8]
      ];

      const lpAvd2 = [
        linha[9],
        linha[12],
        linha[13]
      ];

      const matAvd1 = [
        linha[14],
        linha[15],
        linha[16]
      ];

      const matAvd2 = [
        linha[17],
        linha[20],
        linha[21]
      ];

      const duplicado = codigos.has(codigo);

      codigos.add(codigo);

      const valido =
        !duplicado &&
        validarPercentual(participacao) &&
        validarDistribuicao(lpAvd1) &&
        validarDistribuicao(lpAvd2) &&
        validarDistribuicao(matAvd1) &&
        validarDistribuicao(matAvd2);

      registrarGPSLAB("==========================");
      registrarGPSLAB("ESCOLA: " + nome);
      registrarGPSLAB("CODIGO: " + codigo);
      registrarGPSLAB("MATRICULADOS: " + matriculados);

      registrarGPSLAB(
        "PARTICIPACAO AVD 2: " +
        percentual(participacao)
      );

      registrarGPSLAB(
        "EVOLUCAO PARTICIPACAO: " + linha[5]
      );

      registrarGPSLAB(
        "LP AVD 1: " +
        mostrarDistribuicao(lpAvd1)
      );

      registrarGPSLAB(
        "LP AVD 2: " +
        mostrarDistribuicao(lpAvd2)
      );

      registrarGPSLAB(
        "EVOLUCAO ABAIXO DO BASICO LP: " +
        linha[11]
      );

      registrarGPSLAB(
        "MAT AVD 1: " +
        mostrarDistribuicao(matAvd1)
      );

      registrarGPSLAB(
        "MAT AVD 2: " +
        mostrarDistribuicao(matAvd2)
      );

      registrarGPSLAB(
        "EVOLUCAO ABAIXO DO BASICO MAT: " +
        linha[19]
      );

      if (duplicado) {

        registrarGPSLAB("ATENCAO: CODIGO DUPLICADO");
      }

      if (valido) {

        validas++;
        registrarGPSLAB("CONSISTENCIA: OK");

      } else {

        inconsistentes++;
        registrarGPSLAB("CONSISTENCIA: REVISAR");
      }
    }

    registrarGPSLAB("==========================");
    registrarGPSLAB("RESUMO DA EXTRACAO AVD");

    registrarGPSLAB("ESCOLAS: " + escolas);

    registrarGPSLAB(
      "CODIGOS DISTINTOS: " + codigos.size
    );

    registrarGPSLAB(
      "ESCOLAS CONSISTENTES: " + validas
    );

    registrarGPSLAB(
      "ESCOLAS PARA REVISAO: " + inconsistentes
    );

    registrarGPSLAB(
      "DISTRIBUICOES EXTRAIDAS: " + escolas * 4
    );

    registrarGPSLAB(
      "PERCENTUAIS DE PROFICIENCIA: " + escolas * 12
    );

    registrarGPSLAB(
      "ANO DA AVALIACAO: VALIDAR METADADOS"
    );

    registrarGPSLAB(
      "CONSOLIDADO REGIONAL: NAO CALCULADO"
    );

    registrarGPSLAB(
      "VALIDACAO SEMANTICA: PENDENTE"
    );

  } catch (erro) {

    registrarGPSLAB("ERRO: " + erro.message);

  } finally {

    if (temporario && temporario.id) {

      try {

        DriveApp.getFileById(
          temporario.id
        ).setTrashed(true);

        registrarGPSLAB("TEMPORARIO REMOVIDO");

      } catch (e) {

        registrarGPSLAB(
          "ERRO NA LIMPEZA: " + e.message
        );
      }
    }
  }

  registrarGPSLAB(
    "TEMPO TOTAL: " +
    (Date.now() - inicio) + " ms"
  );

  registrarGPSLAB("TESTE FINALIZADO");
}
