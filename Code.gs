// Retorna os dados da linha selecionada manualmente na planilha
function obterLinhaSelecionada() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const linha = aba.getActiveCell().getRow();
  
  // Evita selecionar cabeçalhos (assume dados a partir da linha 3)
  const linhaValida = linha < 3 ? 3 : linha;
  return obterDadosLinha(linhaValida);
}

// Busca dados da linha (A = Nome, B = Username, C = Perfil/Link, D = Status)
function obterDadosLinha(linha) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const nome = aba.getRange(linha, 1).getValue();     // Coluna A
  const username = aba.getRange(linha, 2).getValue(); // Coluna B
  const link = aba.getRange(linha, 3).getValue();     // Coluna C
  const status = aba.getRange(linha, 4).getValue();   // Coluna D (STATUS DO DIA)

  // Monta link do Instagram se houver username
  let urlInsta = "";
  if (username) {
    const userClean = username.toString().replace('@', '').trim();
    urlInsta = "https://instagram.com/" + userClean;
  }

  return {
    linha: linha,
    nome: nome,
    username: username,
    url: urlInsta,
    statusAtual: status
  };
}

/**
 * Grava o status e já retorna os dados do próximo influenciador em uma única operação.
 */
function salvarEAvancar(novoStatus, linha) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const celulaStatus = aba.getRange(linha, 4); // Coluna D
  const statusAtual = celulaStatus.getValue().toString().trim();

  let statusFinal = novoStatus;

  // Unifica status na mesma célula se já houver registro prévio
  if (statusAtual && statusAtual !== "--" && statusAtual !== "Não Postou" && novoStatus !== "Não Postou") {
    if (!statusAtual.includes(novoStatus)) {
      statusFinal = statusAtual + " / " + novoStatus;
    } else {
      statusFinal = statusAtual;
    }
  }

  celulaStatus.setValue(statusFinal);

  // Busca direto os dados da próxima linha (+1)
  const proximaLinha = linha + 1;
  const proximoInfluenciador = navegarParaLinha(linha, 1);

  return {
    sucesso: true,
    statusAtualizado: statusFinal,
    proximoInfluenciador: proximoInfluenciador
  };
}

// Obtém a linha que está ativa no momento em que abre o painel
function obterDadosLinhaAtiva() {
  return obterLinhaSelecionada();
}
/**
 /**
 * Abre a janela pop-up para seleção do período.
 */
function abrirPopUpPeriodo() {
  const html = HtmlService.createHtmlOutputFromFile('FiltroDataModal')
      .setWidth(350)
      .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Filtrar por Período');
}

/**
 * Gera a tabela organizando:
 * Coluna A: Influenciador | Colunas B, C, D...: Datas Selecionadas
 * Agrupado verticalmente por Tipo de Postagem (STORY + LINK, STORY SEM LINK, etc.)
 */
function gerarMatrizPorDatasECategorias(dataInicioStr, dataFimStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) throw new Error("Aba BANCO_DE_DADOS não encontrada.");

  const [anoI, mesI, diaI] = dataInicioStr.split("-");
  const dataInicio = new Date(anoI, mesI - 1, diaI);

  const [anoF, mesF, diaF] = dataFimStr.split("-");
  const dataFim = new Date(anoF, mesF - 1, diaF);

  if (dataInicio > dataFim) throw new Error("A data inicial não pode ser maior que a data final.");

  // Gerar o array com todas as datas do intervalo no formato DD/MM/YYYY
  const listaDatas = [];
  let dataAtual = new Date(dataInicio);
  while (dataAtual <= dataFim) {
    listaDatas.push(Utilities.formatDate(dataAtual, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"));
    dataAtual.setDate(dataAtual.getDate() + 1);
  }

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) throw new Error("Banco de dados vazio.");

  // Mapeamento: { Categoria: { Influenciador: { Data: Status } } }
  const mapaCategorias = {};
  const influenciadoresPorCategoria = {};

  for (let i = 1; i < dados.length; i++) {
    const rawData = dados[i][0];
    const categoria = String(dados[i][1] || "OUTROS").trim().toUpperCase();
    const user = String(dados[i][2] || "").trim();
    const status = String(dados[i][3] || "").trim();

    if (!rawData || !user) continue;

    let dataLinhaFmt = "";
    if (rawData instanceof Date) {
      dataLinhaFmt = Utilities.formatDate(rawData, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    } else {
      dataLinhaFmt = String(rawData).trim();
    }

    // Verifica se a data da linha está dentro da lista de datas do filtro
    if (listaDatas.includes(dataLinhaFmt)) {
      if (!mapaCategorias[categoria]) {
        mapaCategorias[categoria] = {};
        influenciadoresPorCategoria[categoria] = new Set();
      }

      if (!mapaCategorias[categoria][user]) {
        mapaCategorias[categoria][user] = {};
      }

      influenciadoresPorCategoria[categoria].add(user);
      mapaCategorias[categoria][user][dataLinhaFmt] = status || "Pendente";
    }
  }

  // Prepara ou limpa a aba 'RELATORIO_PERIODO'
  let sheetRelatorio = ss.getSheetByName("RELATORIO_PERIODO");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("RELATORIO_PERIODO");
  } else {
    sheetRelatorio.clear();
  }

  // Título e Botão de Exportação PDF
  sheetRelatorio.getRange("A1").setValue(`📊 Relatório por Formato e Data (${listaDatas[0]} até ${listaDatas[listaDatas.length - 1]})`)
    .setFontWeight("bold").setFontSize(13);

  let linhaAtual = 3;
  const numColunas = listaDatas.length + 1;

  for (const cat in mapaCategorias) {
    const listaUsers = Array.from(influenciadoresPorCategoria[cat]).sort();

    // Cabeçalho da Categoria
    sheetRelatorio.getRange(linhaAtual, 1, 1, numColunas).merge()
      .setValue(` ${cat}`)
      .setBackground("#1b5e20").setFontColor("#FFFFFF").setFontWeight("bold");
    linhaAtual++;

    // Cabeçalho de Colunas: Influenciador + Cada Data
    const cabecalho = ["Influenciador", ...listaDatas];
    sheetRelatorio.getRange(linhaAtual, 1, 1, numColunas).setValues([cabecalho])
      .setBackground("#e8f5e9").setFontWeight("bold");
    linhaAtual++;

    // Monta a matriz de dados do influenciador para cada data
    const blocoDados = [];
    listaUsers.forEach(user => {
      const linhaUser = [user];
      listaDatas.forEach(dt => {
        const statusData = mapaCategorias[cat][user][dt] || "Não postou";
        linhaUser.push(statusData);
      });
      blocoDados.push(linhaUser);
    });

    if (blocoDados.length > 0) {
      sheetRelatorio.getRange(linhaAtual, 1, blocoDados.length, numColunas).setValues(blocoDados);
      linhaAtual += blocoDados.length;
    }

    linhaAtual += 2; // Espaçamento entre blocos
  }

  sheetRelatorio.autoResizeColumns(1, numColunas);
  ss.setActiveSheet(sheetRelatorio);
}

/**
 * Exporta a aba RELATORIO_PERIODO para PDF no formato Paisagem.
 */
function exportarRelatorioPeriodoParaPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RELATORIO_PERIODO");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("⚠️ A aba RELATORIO_PERIODO não existe. Gere o relatório primeiro.");
    return;
  }

  const idPlanilha = ss.getId();
  const idAba = sheet.getSheetId();
  const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HHmm");
  const nomeArquivo = `Relatorio_Betesporte_${dataHoje}.pdf`;

  const urlExport = `https://docs.google.com/spreadsheets/d/${idPlanilha}/export?` +
    `exportFormat=pdf` +
    `&gid=${idAba}` +
    `&size=A4` +
    `&portrait=false` +            // Modo Paisagem (Horizontal) para caber todas as colunas de datas
    `&fitw=true` +                 // Ajusta na largura da folha
    `&gridlines=true` +
    `&printtitle=false` +
    `&sheetnames=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(urlExport, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert("❌ Erro ao gerar o PDF.");
    return;
  }

  const blob = response.getBlob().setName(nomeArquivo);
  const arquivoCriado = DriveApp.createFile(blob);

  SpreadsheetApp.getUi().alert(
    "✅ PDF Gerado com Sucesso!",
    `O arquivo foi salvo no seu Google Drive:\n"${nomeArquivo}"\n\nLink:\n${arquivoCriado.getUrl()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
/**
 * Abre a janela pop-up para seleção do mês e ano.
 */
function abrirPopUpMes() {
  const html = HtmlService.createHtmlOutputFromFile('FiltroMesModal')
      .setWidth(320)
      .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecione o Mês');
}

/**
 * Contabiliza o número total de postagens válidas feitas por cada influenciador no mês.
 */
/**
 * Contabiliza o número total de postagens VÁLIDAS feitas por cada influenciador no mês.
 * Correção: ignora frases com "não postou".
 */
/**
 * Contabiliza o número total de postagens feitas por cada influenciador no mês.
 * Versão robusta com validação flexível de datas e status.
 */
function gerarRelatorioMensal(mes, ano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) throw new Error("Aba BANCO_DE_DADOS não encontrada.");

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) throw new Error("Banco de dados vazio.");

  const targetMes = parseInt(mes, 10);
  const targetAno = parseInt(ano, 10);
  const contagemInfluenciadores = {};

  for (let i = 1; i < dados.length; i++) {
    const rawData = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    const status = String(dados[i][3] || "").trim().toLowerCase();

    if (!rawData || !user) continue;

    // Extração robusta do Mês e Ano da célula
    let mesLinha = null;
    let anoLinha = null;

    if (rawData instanceof Date) {
      mesLinha = rawData.getMonth() + 1; // getMonth() é 0-indexed
      anoLinha = rawData.getFullYear();
    } else {
      // Tenta interpretar texto como DD/MM/YYYY ou YYYY-MM-DD
      const strData = String(rawData).trim();
      const partesBarra = strData.split("/");
      const partesTraco = strData.split("-");

      if (partesBarra.length === 3) {
        mesLinha = parseInt(partesBarra[1], 10);
        anoLinha = parseInt(partesBarra[2], 10);
      } else if (partesTraco.length === 3) {
        anoLinha = parseInt(partesTraco[0], 10);
        mesLinha = parseInt(partesTraco[1], 10);
      }
    }

    // Se a linha for do mês e ano selecionados
    if (mesLinha === targetMes && anoLinha === targetAno) {
      if (!contagemInfluenciadores[user]) {
        contagemInfluenciadores[user] = 0;
      }

      // Se o status estiver vazio ou for explicitamente "não postou/pendente", ignora.
      // Qualquer outro texto preenchido (ex: "postou", "ok", "stories link", "entregue") conta como postado.
      const NaoPostou = status === "" || 
                         status.includes("não postou") || 
                         status.includes("nao postou") || 
                         status.includes("pendente") || 
                         status.includes("não feito") || 
                         status.includes("nao feito");

      if (!NaoPostou) {
        contagemInfluenciadores[user]++;
      }
    }
  }

  // Prepara a aba DESEMPENHO_MENSAL
  let sheetRelatorio = ss.getSheetByName("DESEMPENHO_MENSAL");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("DESEMPENHO_MENSAL");
  } else {
    sheetRelatorio.clear();
  }

  // Formatação visual da aba
  const mesFormatado = String(targetMes).padStart(2, '0');
  sheetRelatorio.getRange("A1:B1").merge()
    .setValue(`📈 Desempenho do mês - ${mesFormatado}/${targetAno}`)
    .setFontWeight("bold").setFontSize(14).setFontColor("#1b5e20");

  const headers = [["Influenciador / Perfil", "Total de Postagens Válidas"]];
  sheetRelatorio.getRange(3, 1, 1, 2).setValues(headers)
    .setBackground("#1b5e20").setFontColor("#FFFFFF").setFontWeight("bold");

  // Ordena por maior número de postagens
  const listaOrdenada = Object.keys(contagemInfluenciadores)
    .map(user => [user, contagemInfluenciadores[user]])
    .sort((a, b) => b[1] - a[1]);

  if (listaOrdenada.length > 0) {
    sheetRelatorio.getRange(4, 1, listaOrdenada.length, 2).setValues(listaOrdenada);
    
    // Zebra de cores nas linhas
    for (let r = 0; r < listaOrdenada.length; r++) {
      if (r % 2 === 0) {
        sheetRelatorio.getRange(4 + r, 1, 1, 2).setBackground("#f9f9f9");
      }
    }
  } else {
    sheetRelatorio.getRange(4, 1, 1, 2).setValues([["Nenhum registro encontrado para este mês.", "-"]]);
  }

  sheetRelatorio.autoResizeColumns(1, 2);
  ss.setActiveSheet(sheetRelatorio);
}

/**
 * Exporta a aba DESEMPENHO_MENSAL formatada para PDF.
 */
/**
 * Gera o PDF no Drive e abre o download direto no navegador do usuário.
 */
/**
 * Exporta a aba DESEMPENHO_MENSAL e dispara o download do PDF no navegador.
 */
function exportarRelatorioMensalParaPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DESEMPENHO_MENSAL");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("⚠️ A aba DESEMPENHO_MENSAL não existe. Gere o relatório primeiro.");
    return;
  }

  const idPlanilha = ss.getId();
  const idAba = sheet.getSheetId();
  const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HHmm");
  const nomeArquivo = `Relatorio_Mensal_BETesporte_${dataHoje}.pdf`;

  // URL para exportar a aba em formato PDF
  const urlExport = `https://docs.google.com/spreadsheets/d/${idPlanilha}/export?` +
    `exportFormat=pdf` +
    `&gid=${idAba}` +
    `&size=A4` +
    `&portrait=true` +
    `&fitw=true` +
    `&gridlines=false` +
    `&printtitle=false` +
    `&sheetnames=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(urlExport, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert("❌ Erro ao gerar o PDF. Verifique se o script tem permissão.");
    return;
  }

  // 1. Salva o arquivo no seu Google Drive
  const blob = response.getBlob().setName(nomeArquivo);
  const arquivoCriado = DriveApp.createFile(blob);

  // 2. Torna o arquivo acessível e gera a URL de download direto
  arquivoCriado.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const urlDownloadDireto = `https://drive.google.com/uc?export=download&id=${arquivoCriado.getId()}`;

  // 3. Janela Pop-up com o redirecionamento automático + botão manual (caso o navegador bloqueie pop-up)
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: Arial, sans-serif; text-align: center; padding: 15px;">
      <p style="font-size: 14px; color: #1b5e20; font-weight: bold;">✅ PDF Gerado com Sucesso!</p>
      <p style="font-size: 12px; color: #555;">Se o download não começar em instantes, clique no botão abaixo:</p>
      <br>
      <a href="${urlDownloadDireto}" target="_blank" style="
        background-color: #1b5e20;
        color: white;
        padding: 10px 20px;
        text-decoration: none;
        font-weight: bold;
        border-radius: 5px;
        display: inline-block;
      ">⬇️ Baixar PDF Agora</a>
      <script>
        // Dispara a abertura do link de download em nova aba imediatamente
        window.onload = function() {
          const a = document.createElement('a');
          a.href = "${urlDownloadDireto}";
          a.target = "_blank";
          a.click();
        };
      </script>
    </div>`
  ).setWidth(360).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download do Relatório');
}
/**
 * Abre o Pop-up para consultar Influenciador com filtro de datas.
 */
function abrirPopUpConsultaInfluenciadorData() {
  const html = HtmlService.createHtmlOutputFromFile('FiltroInfluenciadorDataModal')
      .setWidth(360)
      .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Consulta Personalizada');
}

/**
 * Processa a busca do influenciador no intervalo de datas selecionado.
 */
function executarConsultaInfluenciadorPorData(termoBusca, dataInicioStr, dataFimStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) throw new Error("Aba BANCO_DE_DADOS não encontrada.");

  const [anoI, mesI, diaI] = dataInicioStr.split("-");
  const dataInicio = new Date(anoI, mesI - 1, diaI, 0, 0, 0);

  const [anoF, mesF, diaF] = dataFimStr.split("-");
  const dataFim = new Date(anoF, mesF - 1, diaF, 23, 59, 59);

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) throw new Error("Banco de dados vazio.");

  const termoLower = termoBusca.trim().toLowerCase();
  const resultados = [];
  let nomeOficial = "";

  for (let i = 1; i < dados.length; i++) {
    const rawData = dados[i][0];
    const categoria = String(dados[i][1] || "OUTROS").trim();
    const user = String(dados[i][2] || "").trim();
    const status = String(dados[i][3] || "").trim();

    if (!rawData || !user) continue;

    if (user.toLowerCase().includes(termoLower)) {
      if (!nomeOficial) nomeOficial = user;

      let dataLinha;
      if (rawData instanceof Date) {
        dataLinha = rawData;
      } else {
        const partes = String(rawData).split("/");
        if (partes.length === 3) {
          dataLinha = new Date(partes[2], partes[1] - 1, partes[0]);
        } else {
          dataLinha = new Date(rawData);
        }
      }

      // Filtra apenas dentro do intervalo de datas
      if (dataLinha >= dataInicio && dataLinha <= dataFim) {
        const dataFmt = Utilities.formatDate(dataLinha, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        resultados.push([dataFmt, categoria, status || "Sem registro"]);
      }
    }
  }

  if (resultados.length === 0) {
    throw new Error(`Nenhum registro encontrado para "${termoBusca}" no período selecionado.`);
  }

  // Prepara ou atualiza a aba CONSULTA_INFLUENCIADOR
  let sheetConsulta = ss.getSheetByName("CONSULTA_INFLUENCIADOR");
  if (!sheetConsulta) {
    sheetConsulta = ss.insertSheet("CONSULTA_INFLUENCIADOR");
  } else {
    sheetConsulta.clear();
  }

  const dtIniFmt = Utilities.formatDate(dataInicio, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  const dtFimFmt = Utilities.formatDate(dataFim, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  sheetConsulta.getRange("A1:C1").merge()
    .setValue(`👤 Influenciador: ${nomeOficial} | Período: ${dtIniFmt} até ${dtFimFmt}`)
    .setFontWeight("bold").setFontSize(12);

  const headers = [["Data", "Formato / Categoria", "Status"]];
  sheetConsulta.getRange(2, 1, 1, 3).setValues(headers)
    .setBackground("#1b5e20")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  sheetConsulta.getRange(3, 1, resultados.length, 3).setValues(resultados);
  sheetConsulta.autoResizeColumns(1, 3);

  ss.setActiveSheet(sheetConsulta);
}
/**
 * Abre o pop-up para geração da cobrança individual.
 */
function abrirPopUpCobrancaIndividual() {
  const html = HtmlService.createHtmlOutputFromFile('CobrancaSemanalModal')
      .setWidth(380)
      .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cobrança Individual');
}

/**
 * Calcula a assiduidade e gera o texto formatado para o WhatsApp
 * removendo qualquer texto entre parênteses (ex: BABADOS X1, Nomes, etc.)
 */
/**
 * Calcula a assiduidade e gera o texto formatado para o WhatsApp,
 * suportando até 2 status/conteúdos no mesmo dia e o formato "Feed".
 */
function calcularEGerarCobrancaIndividual(termoBusca, dataInicioStr, dataFimStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) throw new Error("Aba BANCO_DE_DADOS não encontrada.");

  const [anoI, mesI, diaI] = dataInicioStr.split("-");
  const dataInicio = new Date(anoI, mesI - 1, diaI);

  const [anoF, mesF, diaF] = dataFimStr.split("-");
  const dataFim = new Date(anoF, mesF - 1, diaF);

  if (dataInicio > dataFim) throw new Error("A data inicial não pode ser maior que a data final.");

  // Mapeia os dias do período
  const listaDatas = [];
  let dataAtual = new Date(dataInicio);
  while (dataAtual <= dataFim) {
    listaDatas.push(Utilities.formatDate(dataAtual, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"));
    dataAtual.setDate(dataAtual.getDate() + 1);
  }

  const totalDias = listaDatas.length;
  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) throw new Error("Banco de dados vazio.");

  const termoLower = termoBusca.trim().toLowerCase();
  let handleOficial = "";
  
  // Mapeamento por data: armazena uma lista de status por dia -> { "DD/MM/YYYY": ["Story + Link", "Feed"] }
  const relatorioDiario = {};

  // Função interna para padronizar o texto do status
  const formatarTextoPadrao = (texto) => {
    if (!texto) return "";
    let limpo = texto.replace(/\s*\([^)]*\)/g, "").trim();
    let norm = limpo
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\w\s+]/gi, "")
      .trim();

    if (norm.includes("nao postou") || norm.includes("naopostou") || norm.includes("pendente")) {
      return "Não Postou";
    } else if (norm.includes("branding") || norm.includes("betesporte") || norm.includes("item da bet")) {
      return "Branding";
    } else if (norm.includes("feed") || norm.includes("reels") || norm.includes("post")) {
      return "Feed";
    } else if (norm.includes("story") && (norm.includes("sem link") || norm.includes("venda"))) {
      return "Story Sem Link";
    } else if (norm.includes("story") && norm.includes("link")) {
      return "Story + Link";
    } else if (norm.includes("story")) {
      return "Story";
    }
    return limpo;
  };

  for (let i = 1; i < dados.length; i++) {
    const rawData = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    let status = String(dados[i][3] || "").trim();

    if (!rawData || !user) continue;

    if (user.toLowerCase().includes(termoLower)) {
      if (!handleOficial) {
        handleOficial = user.startsWith("@") ? user : "@" + user;
      }

      let dataLinhaFmt = "";
      if (rawData instanceof Date) {
        dataLinhaFmt = Utilities.formatDate(rawData, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      } else {
        const partes = String(rawData).split("/");
        if (partes.length === 3) {
          dataLinhaFmt = `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${partes[2]}`;
        }
      }

      status = formatarTextoPadrao(status);

      if (listaDatas.includes(dataLinhaFmt)) {
        if (!relatorioDiario[dataLinhaFmt]) {
          relatorioDiario[dataLinhaFmt] = [];
        }
        if (status && !relatorioDiario[dataLinhaFmt].includes(status)) {
          relatorioDiario[dataLinhaFmt].push(status);
        }
      }
    }
  }

  if (!handleOficial) handleOficial = termoBusca.startsWith("@") ? termoBusca : "@" + termoBusca;

 numEntregas = 0;
detalhamentoDias = "";

 // 1. LEITURA E COLETA DOS DADOS DA PLANILHA
  for (let i = 1; i < dados.length; i++) {
    const rawData = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    let status = String(dados[i][3] || "").trim();

    if (!rawData || !user) continue;

    if (user.toLowerCase().includes(termoLower)) {
      if (!handleOficial) {
        handleOficial = user.startsWith("@") ? user : "@" + user;
      }

      let dataLinhaFmt = "";
      if (rawData instanceof Date) {
        dataLinhaFmt = Utilities.formatDate(rawData, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      } else {
        const partes = String(rawData).split("/");
        if (partes.length === 3) {
          dataLinhaFmt = `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${partes[2]}`;
        }
      }

      status = formatarTextoPadrao(status);

      if (listaDatas.includes(dataLinhaFmt)) {
        if (!relatorioDiario[dataLinhaFmt]) {
          relatorioDiario[dataLinhaFmt] = [];
        }
        
        // Em vez de salvar a string inteira direto, divide por "/" caso o status já venha unificado
        if (status) {
          const subStatus = status.split("/").map(s => s.trim());
          subStatus.forEach(st => {
            if (st && !relatorioDiario[dataLinhaFmt].includes(st)) {
              relatorioDiario[dataLinhaFmt].push(st);
            }
          });
        }
      }
    }
  }

  if (!handleOficial) handleOficial = termoBusca.startsWith("@") ? termoBusca : "@" + termoBusca;

  let numEntregas = 0;
  let detalhamentoDias = "";

  // 2. PROCESSAMENTO DIA A DIA E MONTAGEM DO RELATÓRIO
  listaDatas.forEach(dt => {
    const listaStatusBrutos = relatorioDiario[dt] || ["Não Postou"];

    // Filtra apenas os status válidos do dia (desconsiderando "Não Postou" ou "Pendente")
    const statusValidos = listaStatusBrutos.filter(st => {
      const stLower = st.toLowerCase();
      return !stLower.includes("não postou") && !stLower.includes("nao postou") && !stLower.includes("pendente") && st !== "";
    });

    if (statusValidos.length > 0) {
      numEntregas++; // Contabiliza a entrega do dia

      const textoStatusUnificado = statusValidos.join(" / ");
      
      // Checa se existe a palavra "Branding" entre os status do dia
      const contemBranding = statusValidos.some(st => st.toLowerCase().includes("branding"));
      
      // Define o emoji: ⚠️ se tiver Branding, caso contrário ✅
      const emojiStatus = contemBranding ? "⚠️" : "✅";

      detalhamentoDias += `${emojiStatus} *${dt}:* ${textoStatusUnificado}\n`;
    } else {
      detalhamentoDias += `❌ *${dt}:* Não Postou\n`;
    }
  });

  const dtIniFmt = Utilities.formatDate(dataInicio, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  const dtFimFmt = Utilities.formatDate(dataFim, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  // Montagem da mensagem formatada
  let mensagem = `Olá, ${handleOficial}! 👋\n\n`;
  mensagem += `Segue o resumo detalhado do seu acompanhamento de postagens (${dtIniFmt} a ${dtFimFmt}):\n\n`;
  mensagem += `📊 *Resumo de Entregas:* ${numEntregas} de ${totalDias} (${numEntregas}/${totalDias})\n\n`;
  mensagem += `📋 *Detalhamento por dia:*\n${detalhamentoDias}`;

  if (numEntregas === totalDias) {
    mensagem += `\n🎉 Parabéns! Você cumpriu 100% das postagens do período!`;
  }

  return mensagem;
}
/**
 * Limpa e padroniza TODOS os formatos e status da coluna D na aba BANCO_DE_DADOS.
 */
function padronizarBancoDeDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANCO_DE_DADOS");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Erro: Aba BANCO_DE_DADOS não encontrada.");
    return;
  }

  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha <= 1) {
    SpreadsheetApp.getUi().alert("O banco de dados está vazio.");
    return;
  }

  const colStatus = 4;
  const rangeStatus = sheet.getRange(2, colStatus, ultimaLinha - 1, 1);
  const valores = rangeStatus.getValues();

  let contagemAlteracoes = 0;

  for (let i = 0; i < valores.length; i++) {
    let textoOriginal = String(valores[i][0] || "").trim();
    if (!textoOriginal) continue;

    let textoLimpo = textoOriginal.replace(/\s*\([^)]*\)/g, "").trim();
    let textoNormalizado = textoLimpo
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\w\s+]/gi, "")
      .trim();

    let textoPadronizado = textoOriginal;

    if (textoNormalizado.includes("nao postou") || textoNormalizado.includes("naopostou") || textoNormalizado.includes("pendente")) {
      textoPadronizado = "Não Postou";
    } else if (textoNormalizado.includes("branding") || textoNormalizado.includes("betesporte") || textoNormalizado.includes("item da bet")) {
      textoPadronizado = "Branding";
    } else if (textoNormalizado.includes("feed") || textoNormalizado.includes("reels") || textoNormalizado.includes("post")) {
      textoPadronizado = "Feed";
    } else if (textoNormalizado.includes("story") && (textoNormalizado.includes("sem link") || textoNormalizado.includes("venda"))) {
      textoPadronizado = "Story Sem Link";
    } else if (textoNormalizado.includes("story") && textoNormalizado.includes("link")) {
      textoPadronizado = "Story + Link";
    } else if (textoNormalizado.includes("story")) {
      textoPadronizado = "Story";
    }

    if (textoOriginal !== textoPadronizado) {
      valores[i][0] = textoPadronizado;
      contagemAlteracoes++;
    }
  }

  if (contagemAlteracoes > 0) {
    rangeStatus.setValues(valores);
    SpreadsheetApp.getUi().alert(`✅ Padronização concluída!\n\n${contagemAlteracoes} registro(s) foram corrigidos e padronizados.`);
  } else {
    SpreadsheetApp.getUi().alert("Nenhum registro precisou ser alterado. O banco já está 100% padronizado!");
  }
}
/**
 * Salva múltiplos status na Coluna D de forma assíncrona.
 */
function salvarEAvancarConjunto(listaStatus, linha) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const celulaStatus = aba.getRange(linha, 4); // Coluna D
  let statusAtual = celulaStatus.getValue().toString().trim();

  listaStatus.forEach(function(novoStatus) {
    if (!statusAtual || statusAtual === "--" || statusAtual === "Não Postou" || novoStatus === "Não Postou") {
      statusAtual = novoStatus;
    } else if (!statusAtual.includes(novoStatus)) {
      statusAtual += " / " + novoStatus;
    }
  });

  celulaStatus.setValue(statusAtual);
  return true;
}

/**
 * Exibe uma janela com o resumo diário para cópia rápida.
 */
function exibirRelatorioDiarioGeral() {
  const textoRelatorio = gerarRelatorioDiarioGeral();
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; padding: 10px;">
      <textarea id="txt" style="width: 100%; height: 140px; font-size: 13px; padding: 8px; box-sizing: border-box;" readonly>${textoRelatorio}</textarea>
      <br><br>
      <button onclick="copiar()" style="background-color: #10b981; color: white; border: none; padding: 10px 16px; border-radius: 6px; font-weight: bold; cursor: pointer; width: 100%;">📋 Copiar Relatório</button>
      <script>
        function copiar() {
          var copyText = document.getElementById("txt");
          copyText.select();
          document.execCommand("copy");
          alert("Relatório copiado!");
        }
      </script>
    </div>`
  )
  .setWidth(360)
  .setHeight(230)
  .setTitle('📊 Relatório Geral Diário');

  SpreadsheetApp.getUi().showModalDialog(html, '📊 Relatório Geral Diário');
}
/**
 * Calcula o resumo geral de postagens da data de hoje lendo a aba ativa.
 */
function gerarRelatorioDiarioGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getActiveSheet();
  const dados = aba.getDataRange().getValues();

  if (dados.length < 2) return "Aba vazia ou sem dados suficientes.";

  const hojeFmt = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  let colunaData = -1;
  const linhasCabecalho = [0, 1];

  for (let r of linhasCabecalho) {
    if (dados[r]) {
      for (let col = 0; col < dados[r].length; col++) {
        let val = dados[r][col];
        let valFmt = "";

        if (val instanceof Date) {
          valFmt = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        } else {
          valFmt = String(val).trim();
        }

        if (valFmt === hojeFmt) {
          colunaData = col;
          break;
        }
      }
    }
    if (colunaData !== -1) break;
  }

  if (colunaData === -1) {
    return `⚠️ Não foi encontrada uma coluna para a data de hoje (${hojeFmt}) na aba ativa.`;
  }

  let totalMonitorados = 0;
  let totalPostaram = 0;

  for (let i = 2; i < dados.length; i++) {
    const nomeOuUser = String(dados[i][2] || dados[i][1] || dados[i][0] || "").trim();

    if (!nomeOuUser || nomeOuUser.toLowerCase().includes("influenciador") || nomeOuUser.toLowerCase().includes("nome")) {
      continue;
    }

    totalMonitorados++;

    const statusCelula = String(dados[i][colunaData] || "").trim().toLowerCase();

    const postouValido = statusCelula !== "" && 
                         statusCelula !== "--" &&
                         !statusCelula.includes("não postou") && 
                         !statusCelula.includes("nao postou") && 
                         !statusCelula.includes("pendente");

    if (postouValido) {
      totalPostaram++;
    }
  }

  const totalNaoPostaram = totalMonitorados - totalPostaram;

  let taxaAssiduidade = "0,0";
  if (totalMonitorados > 0) {
    taxaAssiduidade = ((totalPostaram / totalMonitorados) * 100).toFixed(1).replace(".", ",");
  }

  return `✅ *Resumo (${hojeFmt})*\n\n` +
         `*${totalPostaram}* influenciadores postaram (de ${totalMonitorados} monitorados)\n` +
         `*${totalNaoPostaram}* não postaram\n` +
         `*Assiduidade:* ${taxaAssiduidade}%`;
}
