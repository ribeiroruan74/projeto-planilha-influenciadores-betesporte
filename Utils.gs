
function abrirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('BETEsporte - Painel')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


// =========================================================
// ⚙️ SEU BACKEND ORIGINAL E FUNÇÕES DO SISTEMA
// =========================================================

// 1. Retorna a aba principal de acompanhamento
function getAbaAcompanhamento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ACOMPANHAMENTO");
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

// 2. Garante que a aba BANCO_DE_DADOS exista
function getAbaBancoDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("BANCO_DE_DADOS");
  if (!sheet) {
    sheet = ss.insertSheet("BANCO_DE_DADOS");
    sheet.getRange("A1:D1")
      .setValues([["DATA", "INFLUENCIADOR", "@USERNAME", "STATUS"]])
      .setFontWeight("bold")
      .setBackground("#3c4043")
      .setFontColor("#ffffff");
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 220);
  }
  return sheet;
}

// 3. Ordena os influenciadores em ordem alfabética (Coluna A)
function ordenarInfluenciadoresAlfabetico() {
  const sheet = getAbaAcompanhamento();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3) return;

  const range = sheet.getRange(3, 1, lastRow - 2, lastCol);
  range.sort({ column: 1, ascending: true });

  if (SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Lista organizada de A a Z!", "⚽ BETEsporte", 3);
  }
}

// 4. Obtém as datas disponíveis
function obterDatasCabecalho() {
  const sheetAcomp = getAbaAcompanhamento();
  const dataAcomp = sheetAcomp.getRange("D2").getDisplayValue().trim();

  const datas = [];
  if (dataAcomp) {
    datas.push({ coluna: 4, dataTexto: dataAcomp });
  }

  return datas;
}

// 5. Lê os dados da linha selecionada (URL Web formatada)
function obterDadosLinhaAtiva(colunaData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  if (!cell || cell.getRow() < 3) {
    return {
      linha: 0,
      nome: "Clique na linha de um influenciador",
      username: "@username",
      url: "",
      statusAtual: ""
    };
  }

  const row = cell.getRow();
  const nome = String(sheet.getRange(row, 1).getValue()).trim() || "Sem Nome";
  let username = String(sheet.getRange(row, 2).getValue()).trim() || "";

  if (!username) username = "@" + nome.toLowerCase().replace(/\s+/g, "");

  let cleanUser = username.replace("@", "").replace(/https?:\/\/(www\.)?instagram\.com\//, "").replace(/\//g, "").trim();
  let url = "https://www.instagram.com/" + cleanUser + "/?hl=pt";

  const col = colunaData ? Number(colunaData) : 4;
  const statusAtual = String(sheet.getRange(row, col).getDisplayValue()).trim();

  return {
    linha: row,
    nome: nome,
    username: username.startsWith("@") ? username : "@" + username,
    url: url,
    statusAtual: statusAtual
  };
}

// 6. Salva status na linha
function salvarStatusLinhaAtiva(statusTexto, colunaData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  if (!cell || cell.getRow() < 3) {
    return { sucesso: false, mensagem: "Selecione uma linha de influenciador." };
  }

  const row = cell.getRow();
  const col = colunaData ? Number(colunaData) : 4;

  sheet.getRange(row, col).setValue(statusTexto);

  return { sucesso: true, mensagem: "Status salvo!" };
}

// 7. Calcula o progresso do dia atual
function calcularProgressoData(colunaData) {
  const sheet = getAbaAcompanhamento();
  const lastRow = sheet.getLastRow();
  const col = colunaData ? Number(colunaData) : 4;

  if (lastRow < 3) {
    return { concluidos: 0, total: 0, restantes: 0, percentual: 0, dataTexto: "" };
  }

  const dataTexto = sheet.getRange(2, col).getDisplayValue();
  const status = sheet.getRange(3, col, lastRow - 2, 1).getDisplayValues();

  let concluidos = 0;
  let total = status.length;

  for (let i = 0; i < status.length; i++) {
    if (String(status[i][0]).trim() !== "") {
      concluidos++;
    }
  }

  return {
    concluidos: concluidos,
    total: total,
    restantes: total - concluidos,
    percentual: total > 0 ? Math.round((concluidos / total) * 100) : 0,
    dataTexto: dataTexto
  };
}

// 8. FINALIZAR O DIA: Arquiva os registros no Banco de Dados
function finalizarDiaEMoverHistorico() {
  const sheetAcomp = getAbaAcompanhamento();
  const sheetBd = getAbaBancoDados();
  const ui = SpreadsheetApp.getUi();

  const dataHoje = sheetAcomp.getRange("D2").getDisplayValue().trim();
  const lastRow = sheetAcomp.getLastRow();

  if (!dataHoje || lastRow < 3) {
    ui.alert("Não há dados válidos para arquivar.");
    return;
  }

  const confirmacao = ui.alert(
    "Finalizar e Arquivar Dia",
    `Deseja arquivar todos os registros do dia ${dataHoje} no histórico e preparar a aba para o próximo dia?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmacao !== ui.Button.YES) return;

  const dados = sheetAcomp.getRange(3, 1, lastRow - 2, 4).getValues();
  const novosRegistros = [];

  for (let i = 0; i < dados.length; i++) {
    const nome = String(dados[i][0]).trim();
    const user = String(dados[i][1]).trim();
    let status = String(dados[i][3]).trim();

    if (!nome) continue;
    if (!status) status = "Não postou";

    novosRegistros.push([dataHoje, nome, user, status]);
  }

  if (novosRegistros.length > 0) {
    const lastRowBd = Math.max(1, sheetBd.getLastRow());
    sheetBd.getRange(lastRowBd + 1, 1, novosRegistros.length, 4).setValues(novosRegistros);
  }

  sheetAcomp.getRange(3, 4, lastRow - 2, 1).clearContent();

  const partesData = dataHoje.split("/");
  let novaDataTexto = dataHoje;
  if (partesData.length === 3) {
    const objData = new Date(partesData[2], partesData[1] - 1, partesData[0]);
    objData.setDate(objData.getDate() + 1);
    novaDataTexto = Utilities.formatDate(objData, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  sheetAcomp.getRange("D2").setValue(novaDataTexto);

  ui.alert(`✅ Dia ${dataHoje} arquivado com sucesso!\nNova data configurada: ${novaDataTexto}`);
}

// 9. GERAR DASHBOARD EXECUTIVO
function gerarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBd = getAbaBancoDados();
  let sheetDash = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");

  sheetDash.clear();
  sheetDash.clearFormats();
  sheetDash.gridlinesEnabled = false;

  sheetDash.setColumnWidth(1, 30);
  sheetDash.setColumnWidth(2, 220);
  sheetDash.setColumnWidth(3, 180);
  sheetDash.setColumnWidth(4, 40);
  sheetDash.setColumnWidth(5, 220);
  sheetDash.setColumnWidth(6, 180);

  sheetDash.getRange("B2:F2").merge()
    .setValue("📊 DASHBOARD DE DESEMPENHO E ASSIDUIDADE")
    .setFontWeight("bold")
    .setFontSize(16)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheetDash.setRowHeight(2, 45);

  const lastRowBd = sheetBd.getLastRow();

  if (lastRowBd < 2) {
    sheetDash.getRange("B4").setValue("Ainda não há dados no BANCO_DE_DADOS.").setFontStyle("italic");
    SpreadsheetApp.getActiveSpreadsheet().toast("Dashboard criado! Garanta que haja dados no histórico.", "⚽ BETEsporte", 4);
    return;
  }

  const dadosBd = sheetBd.getRange(2, 1, lastRowBd - 1, 4).getDisplayValues();

  const datasUnicas = new Set();
  const influenciadoresUnicos = new Set();
  let totalPostsRealizados = 0;
  let totalNaoPostou = 0;

  const contagemTipos = {
    "Story com item da BETEsporte": 0,
    "Story + Link": 0,
    "Story de venda (sem link)": 0
  };

  const mapaInf = {};

  for (let i = 0; i < dadosBd.length; i++) {
    const dt = dadosBd[i][0];
    const nome = dadosBd[i][1];
    const user = dadosBd[i][2];
    const status = String(dadosBd[i][3]).trim();

    if (!nome) continue;

    datasUnicas.add(dt);
    influenciadoresUnicos.add(nome);

    if (!mapaInf[nome]) {
      mapaInf[nome] = { username: user, total: 0, postou: 0 };
    }
    mapaInf[nome].total++;

    const stLower = status.toLowerCase();
    if (stLower !== "" && stLower !== "não postou") {
      totalPostsRealizados++;
      mapaInf[nome].postou++;

      if (status.includes("Story com item da BETEsporte")) contagemTipos["Story com item da BETEsporte"]++;
      if (status.includes("Story + Link")) contagemTipos["Story + Link"]++;
      if (status.includes("Story de venda (sem link)")) contagemTipos["Story de venda (sem link)"]++;
    } else {
      totalNaoPostou++;
    }
  }

  const totalDias = datasUnicas.size;
  const totalInf = influenciadoresUnicos.size;
  const totalOportunidades = totalDias * totalInf;
  const taxaGeral = totalOportunidades > 0 ? (totalPostsRealizados / totalOportunidades) : 0;

  const criarCard = (rangeStr, titulo, valor, corFundo, corTexto, ePercentual) => {
    const rng = sheetDash.getRange(rangeStr);
    rng.merge();
    rng.setValue(`${titulo}\n${ePercentual ? (valor * 100).toFixed(1) + "%" : valor}`)
      .setFontWeight("bold")
      .setBackground(corFundo)
      .setFontColor(corTexto)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontSize(12);
  };

  criarCard("B4:C4", "📅 DIAS MONITORADOS", totalDias, "#f1f3f4", "#202124", false);
  criarCard("E4:F4", "👥 INFLUENCIADORES ATIVOS", totalInf, "#f1f3f4", "#202124", false);
  criarCard("B5:C5", "🟢 DIAS CUMPRIU POSTAGEM", totalPostsRealizados, "#e6f4ea", "#137333", false);
  criarCard("E5:F5", "📈 TAXA DE CUMPRIMENTO GERAL", taxaGeral, "#e8f0fe", "#1967d2", true);

  sheetDash.setRowHeight(4, 35);
  sheetDash.setRowHeight(5, 35);

  sheetDash.getRange("B7:C7").merge()
    .setValue("TOTAL DE FORMATOS EXECUTADOS")
    .setFontWeight("bold")
    .setBackground("#3c4043")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheetDash.getRange("B8:C11").setValues([
    ["🟢 Story c/ item BETEsporte", contagemTipos["Story com item da BETEsporte"]],
    ["🔵 Story + Link", contagemTipos["Story + Link"]],
    ["🟡 Story de venda (s/ link)", contagemTipos["Story de venda (sem link)"]],
    ["🔴 Não Postou / Pendente", totalNaoPostou]
  ]);

  sheetDash.getRange("B8:B11").setFontWeight("bold");
  sheetDash.getRange("C8:C11").setHorizontalAlignment("center");

  const listaRank = [];
  for (const nome in mapaInf) {
    const p = mapaInf[nome].postou;
    const t = mapaInf[nome].total;
    listaRank.push({
      nome: nome,
      username: mapaInf[nome].username,
      postou: p,
      taxa: t > 0 ? (p / t) : 0
    });
  }

  listaRank.sort((a, b) => b.taxa - a.taxa || b.postou - a.postou);

  sheetDash.getRange("B13:C13").merge()
    .setValue("🏆 TOP 5 MAIS ASSÍDUOS")
    .setFontWeight("bold")
    .setBackground("#27ae60")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const top5Melhores = [];
  for (let k = 0; k < Math.min(5, listaRank.length); k++) {
    top5Melhores.push([listaRank[k].nome, (listaRank[k].taxa * 100).toFixed(0) + "% de taxa"]);
  }
  if (top5Melhores.length > 0) {
    sheetDash.getRange(14, 2, top5Melhores.length, 2).setValues(top5Melhores);
  }

  sheetDash.getRange("E13:F13").merge()
    .setValue("⚠️ TOP 5 MENOS ASSÍDUOS (ALERTAS)")
    .setFontWeight("bold")
    .setBackground("#c0392b")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const listaPiores = [...listaRank].reverse();
  const top5Piores = [];
  for (let m = 0; m < Math.min(5, listaPiores.length); m++) {
    top5Piores.push([listaPiores[m].nome, (listaPiores[m].taxa * 100).toFixed(0) + "% de taxa"]);
  }
  if (top5Piores.length > 0) {
    sheetDash.getRange(14, 5, top5Piores.length, 2).setValues(top5Piores);
  }

  const charts = sheetDash.getCharts();
  for (let i = 0; i < charts.length; i++) {
    sheetDash.removeChart(charts[i]);
  }

  const rangeFormatos = sheetDash.getRange("B7:C10");
  const pieChart = sheetDash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(rangeFormatos)
    .setPosition(7, 5, 0, 0)
    .setOption("title", "Distribuição de Formatos Executados")
    .setOption("width", 400)
    .setOption("height", 240)
    .setOption("is3D", true)
    .build();

  sheetDash.insertChart(pieChart);

  SpreadsheetApp.getActiveSpreadsheet().toast("Dashboard atualizado com sucesso!", "⚽ BETEsporte", 3);
}

// 10. NAVEGAR PARA A LINHA ANTERIOR OU PRÓXIMA
function navegarParaLinha(linhaAtual, direcao) {
  const sheet = getAbaAcompanhamento();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 3) return obterDadosLinhaAtiva();

  let base = Number(linhaAtual);
  if (isNaN(base) || base < 3) {
    base = sheet.getActiveCell().getRow();
    if (base < 3) base = 3;
  }

  let novaLinha = base + Number(direcao);

  if (novaLinha < 3) novaLinha = 3;
  if (novaLinha > lastRow) novaLinha = lastRow;

  sheet.getRange(novaLinha, 1).activate();
  SpreadsheetApp.flush();

  return obterDadosLinhaAtiva();
}

// 11. BUSCAR E SELEÇÃO RÁPIDA DE INFLUENCIADOR
function buscarEAtivarInfluenciador(termo) {
  if (!termo) return obterDadosLinhaAtiva();

  const sheet = getAbaAcompanhamento();
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return obterDadosLinhaAtiva();

  const dados = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
  const busca = termo.toLowerCase().replace("@", "").trim();

  for (let i = 0; i < dados.length; i++) {
    const nome = String(dados[i][0]).toLowerCase();
    const user = String(dados[i][1]).toLowerCase();

    if (nome.includes(busca) || user.includes(busca)) {
      const linhaEncontrada = i + 3;
      sheet.getRange(linhaEncontrada, 1).activate();
      SpreadsheetApp.flush();
      return obterDadosLinhaAtiva();
    }
  }

  return obterDadosLinhaAtiva();
}

// 12. BUSCAR PRÓXIMO INFLUENCIADOR PENDENTE
function buscarProximoPendente(linhaAtual) {
  const sheet = getAbaAcompanhamento();
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return obterDadosLinhaAtiva();

  let base = Number(linhaAtual);
  if (isNaN(base) || base < 3) base = 3;

  const inicio = base >= 3 ? base - 2 : 0;
  const statusValues = sheet.getRange(3, 4, lastRow - 2, 1).getDisplayValues();

  for (let i = inicio; i < statusValues.length; i++) {
    const st = String(statusValues[i][0]).trim().toLowerCase();
    if ((st === "" || st === "não postou") && (i + 3) !== base) {
      const linhaEncontrada = i + 3;
      sheet.getRange(linhaEncontrada, 1).activate();
      SpreadsheetApp.flush();
      return obterDadosLinhaAtiva();
    }
  }

  for (let i = 0; i < inicio; i++) {
    const st = String(statusValues[i][0]).trim().toLowerCase();
    if (st === "" || st === "não postou") {
      const linhaEncontrada = i + 3;
      sheet.getRange(linhaEncontrada, 1).activate();
      SpreadsheetApp.flush();
      return obterDadosLinhaAtiva();
    }
  }

  return obterDadosLinhaAtiva();
}

/**
 * Sincroniza e copia automaticamente todos os dados do BANCO_DE_DADOS para a aba Histórico Detalhado.
 */
function montarAbaHistoricoAutomaticoDetalhes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBd = ss.getSheetByName("BANCO_DE_DADOS");
  let abaHistorico = ss.getSheetByName("Histórico Detalhado");

  if (!sheetBd) {
    ss.toast("Aba BANCO_DE_DADOS não encontrada.", "BETEsporte");
    return;
  }

  if (!abaHistorico) {
    abaHistorico = ss.insertSheet("Histórico Detalhado");
  }

  const lastRowBd = sheetBd.getLastRow();
  if (lastRowBd < 2) {
    ss.toast("Ainda não há registros no BANCO_DE_DADOS.", "BETEsporte");
    return;
  }

  // Pega todos os dados do BANCO_DE_DADOS (incluindo o cabeçalho)
  const dadosBd = sheetBd.getRange(1, 1, lastRowBd, 4).getValues();

  // Limpa o Histórico Detalhado e reescreve com a cópia exata do BANCO_DE_DADOS
  abaHistorico.clear();
  abaHistorico.getRange(1, 1, dadosBd.length, 4).setValues(dadosBd);

  // Formatação visual do cabeçalho
  abaHistorico.getRange("A1:D1")
    .setFontWeight("bold")
    .setBackground("#1e293b")
    .setFontColor("#ffffff");

  abaHistorico.setColumnWidth(1, 120);
  abaHistorico.setColumnWidth(2, 220);
  abaHistorico.setColumnWidth(3, 180);
  abaHistorico.setColumnWidth(4, 220);

  ss.toast("✅ Histórico Detalhado sincronizado com sucesso!", "BETEsporte", 4);
}
// 14. GERA RANKING DE ASSIDUIDADE
function gerarRankingAssiduidadeDetalhes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaRanking = ss.getSheetByName("Ranking Assiduidade");
  
  if (!abaRanking) {
    abaRanking = ss.insertSheet("Ranking Assiduidade");
  }

  const abaAcompanhamento = getAbaAcompanhamento();
  if (!abaAcompanhamento) return;

  const lastRow = abaAcompanhamento.getLastRow();
  if (lastRow < 3) return;

  const dados = abaAcompanhamento.getRange(3, 1, lastRow - 2, 2).getValues();

  abaRanking.clear();
  abaRanking.appendRow(["Posição", "Influenciador", "Username", "Status de Monitoramento"]);
  abaRanking.getRange("A1:D1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");

  dados.forEach((linha, idx) => {
    if (linha[0]) {
      abaRanking.appendRow([idx + 1, linha[0], linha[1] || "", "Ativo"]);
    }
  });

  ss.toast("Ranking de assiduidade gerado!", "BETEsporte", 3);
}
