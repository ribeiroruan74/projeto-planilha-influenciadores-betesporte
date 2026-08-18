// @ts-nocheck
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚽ BETEsporte - Painel')
    .addItem('📱 Abrir Painel de Controle (Sidebar)', 'abrirSidebar')
    .addSeparator()
    .addItem('📊 Atualizar Dashboard Geral', 'gerarDashboard')
    .addItem('🏆 Atualizar Ranking de Assiduidade', 'atualizarRanking')
    .addItem('📋 Gerar Relatório Geral Diário', 'exibirRelatorioDiarioGeral')
    .addSeparator()
    .addItem('📅 Filtrar Período (Datas em Colunas)', 'abrirPopUpPeriodo')
    .addItem('📈 Gerar Relatório Mensal de Postagens', 'abrirPopUpMes')
    .addSeparator()
    .addItem('🔍 Consultar por Data', 'consultarPorData')
    .addItem('👤 Consultar por Influenciador', 'consultarPorInfluenciador')
    .addSeparator()
    .addItem('🚨 Checar Quem Não Postou Hoje', 'listarInadimplentes')
    .addItem('📲 Gerar Cobrança Individual (WhatsApp)', 'abrirPopUpCobrancaIndividual')
    .addItem('📌 Salvar Dia no Histórico', 'finalizarDiaEMoverHistorico')
    .addItem('🔤 Reordenar Influenciadores (A-Z)', 'ordenarInfluenciadoresAlfabetico')
    .addSeparator()
    .addItem('🛠️ Manutenção: Recriar Aba HISTÓRICO', 'montarAbaHistoricoAutomatico')
    .addItem('🧹 Padronizar Banco de Dados', 'padronizarBancoDeDados')
    .addToUi();
}

function abrirPainel() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setWidth(700)
      .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, 'BETEsporte - Painel');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();

  // Monitora alterações na data (B1 ou B2) da aba HISTORICO
  if (sheet.getName() === "HISTORICO" && (range.getA1Notation() === "B1" || range.getA1Notation() === "B2")) {
    let selecao = range.getValue();

    if (selecao instanceof Date) {
      selecao = Utilities.formatDate(selecao, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    if (typeof carregarHistoricoPorData === 'function') {
      carregarHistoricoPorData(String(selecao));
    } else if (typeof registrarHistorico === 'function') {
      registrarHistorico();
    }
  }
}

// ==========================================
// FUNÇÕES DE APOIO E COBRANÇA RÁPIDA
// ==========================================

function listarInadimplentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("HISTORICO");
  if (!sheet) return;
  
  const ultLinha = sheet.getLastRow();
  if (ultLinha < 4) return;

  const dados = sheet.getRange(4, 1, ultLinha - 3, 3).getValues();
  const pendentes = [];

  dados.forEach(linha => {
    const nome = linha[0];
    const user = linha[1];
    const status = String(linha[2]).toLowerCase();
    
    if (user && (status.includes("não postou") || status.includes("pendente") || status === "")) {
      pendentes.push(`${nome ? nome + ' ' : ''}(${user})`);
    }
  });

  if (pendentes.length > 0) {
    SpreadsheetApp.getUi().alert("🚨 PENDENTES DE HOJE (" + pendentes.length + "):\n\n" + pendentes.join("\n"));
  } else {
    SpreadsheetApp.getUi().alert("🎉 Todos os influenciadores postaram até o momento!");
  }
}

function copiarListaCobranca() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("HISTORICO");
  if (!sheet) return;

  const ultLinha = sheet.getLastRow();
  if (ultLinha < 4) return;

  const dados = sheet.getRange(4, 2, ultLinha - 3, 2).getValues();
  const pendentesUsernames = [];

  dados.forEach(linha => {
    const user = linha[0];
    const status = String(linha[1]).toLowerCase();
    if (user && (status.includes("não postou") || status.includes("pendente") || status === "")) {
      pendentesUsernames.push(user.startsWith("@") ? user : "@" + user);
    }
  });

  if (pendentesUsernames.length > 0) {
    const mensagem = "⚠️ *Seguem os perfis que não realizaram nenhuma publicação na data de hoje:\n\n" + pendentesUsernames.join("\n");
    
    const htmlOutput = HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif; font-size:13px;">Copie e cole a mensagem abaixo no grupo:</p>' +
      '<textarea style="width:96%; height:140px; font-family:monospace; padding:8px;" readonly onClick="this.select()">' + mensagem + '</textarea>'
    ).setWidth(400).setHeight(220);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📲 Lista para WhatsApp');
  } else {
    SpreadsheetApp.getUi().alert("Nenhum influenciador pendente para cobrança.");
  }
}

/**
 * Exporta a aba de consulta ativa (CONSULTA_DATA ou CONSULTA_INFLUENCIADOR) para PDF no Google Drive.
 */
/**
 * CONSULTA POR DATA
 * Cria ou atualiza a aba 'CONSULTA_DATA' com o status de todos os influenciadores em um dia específico.
 */
function consultarPorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) {
    SpreadsheetApp.getUi().alert("❌ Aba BANCO_DE_DADOS não encontrada.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt("📅 Consulta por Data", "Digite a data desejada (ex: 16/06/2026 ou 2026-06-16):", ui.ButtonSet.OK_CANCEL);
  
  if (resposta.getSelectedButton() !== ui.Button.OK) return;
  
  const dataBuscada = resposta.getResponseText().trim();
  if (!dataBuscada) {
    ui.alert("⚠️ Data inválida.");
    return;
  }

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) {
    ui.alert("⚠️ Banco de dados vazio.");
    return;
  }

  const resultados = [];

  // Percorre o banco (Coluna A = Data, Coluna C = Influenciador, Coluna D = Status)
  for (let i = 1; i < dados.length; i++) {
    const dataLinha = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    const status = String(dados[i][3] || "").trim();

    if (!dataLinha || !user) continue;

    let dataFormatada = "";
    if (dataLinha instanceof Date) {
      dataFormatada = Utilities.formatDate(dataLinha, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    } else {
      dataFormatada = String(dataLinha).trim();
    }

    if (dataFormatada === dataBuscada || String(dataLinha).includes(dataBuscada)) {
      resultados.push([user, status || "Sem registro"]);
    }
  }

  if (resultados.length === 0) {
    ui.alert(`⚠️ Nenhum registro encontrado para a data "${dataBuscada}".`);
    return;
  }

  let sheetConsulta = ss.getSheetByName("CONSULTA_DATA");
  if (!sheetConsulta) {
    sheetConsulta = ss.insertSheet("CONSULTA_DATA");
  } else {
    sheetConsulta.clear();
  }

  sheetConsulta.getRange("A1:B1").setValues([[`📅 Consulta: ${dataBuscada}`, ""]])
    .setFontWeight("bold").setFontSize(12);
  
  const headers = [["Influenciador", "Status"]];
  sheetConsulta.getRange(2, 1, 1, 2).setValues(headers)
    .setBackground("#1b5e20")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  sheetConsulta.getRange(3, 1, resultados.length, 2).setValues(resultados);
  sheetConsulta.autoResizeColumns(1, 2);

  ss.setActiveSheet(sheetConsulta);
  ui.alert(`✅ Consulta concluída! ${resultados.length} influenciadores encontrados para ${dataBuscada}.`);
}


/**
 * CONSULTA POR INFLUENCIADOR
 * Cria ou atualiza a aba 'CONSULTA_INFLUENCIADOR' com todo o histórico de um influenciador.
 */
function consultarPorInfluenciador() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheetBanco) {
    SpreadsheetApp.getUi().alert("❌ Aba BANCO_DE_DADOS não encontrada.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt("👤 Consulta por Influenciador", "Digite o nome ou @ do influenciador:", ui.ButtonSet.OK_CANCEL);
  
  if (resposta.getSelectedButton() !== ui.Button.OK) return;
  
  const termoBusca = resposta.getResponseText().trim().toLowerCase();
  if (!termoBusca) {
    ui.alert("⚠️ Nome inválido.");
    return;
  }

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) {
    ui.alert("⚠️ Banco de dados vazio.");
    return;
  }

  const resultados = [];
  let nomeEncontradoOficial = "";

  for (let i = 1; i < dados.length; i++) {
    const dataLinha = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    const status = String(dados[i][3] || "").trim();

    if (!user) continue;

    if (user.toLowerCase().includes(termoBusca)) {
      if (!nomeEncontradoOficial) nomeEncontradoOficial = user;

      let dataFormatada = "";
      if (dataLinha instanceof Date) {
        dataFormatada = Utilities.formatDate(dataLinha, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      } else {
        dataFormatada = String(dataLinha).trim();
      }

      resultados.push([dataFormatada, status || "Sem registro"]);
    }
  }

  if (resultados.length === 0) {
    ui.alert(`⚠️ Nenhum influenciador encontrado com o termo "${termoBusca}".`);
    return;
  }

  let sheetConsulta = ss.getSheetByName("CONSULTA_INFLUENCIADOR");
  if (!sheetConsulta) {
    sheetConsulta = ss.insertSheet("CONSULTA_INFLUENCIADOR");
  } else {
    sheetConsulta.clear();
  }

  sheetConsulta.getRange("A1:B1").setValues([[`👤 Influenciador: ${nomeEncontradoOficial}`, ""]])
    .setFontWeight("bold").setFontSize(12);

  const headers = [["Data", "Status"]];
  sheetConsulta.getRange(2, 1, 1, 2).setValues(headers)
    .setBackground("#1b5e20")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  sheetConsulta.getRange(3, 1, resultados.length, 2).setValues(resultados);
  sheetConsulta.autoResizeColumns(1, 2);

  ss.setActiveSheet(sheetConsulta);
  ui.alert(`✅ Histórico de ${nomeEncontradoOficial} gerado com sucesso! (${resultados.length} registros).`);
}
