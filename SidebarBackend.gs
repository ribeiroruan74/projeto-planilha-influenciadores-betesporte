/**
 * Executa Salvar Status + Calcular Progresso + Avançar Linha em 1 ÚNICA chamada!
 */
function salvarEAvancar(statusTexto, linhaAtual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  let linha = Number(linhaAtual) || sheet.getActiveCell().getRow();
  if (linha < 3) linha = 3;

  // 1. Grava o status na linha ativa
  sheet.getRange(linha, 3).setValue(statusTexto);

  // 2. Calcula a próxima linha (Próximo registro)
  const ultLinha = Math.max(sheet.getLastRow(), 3);
  let novaLinha = linha + 1;
  if (novaLinha > ultLinha) novaLinha = 3;

  sheet.getRange(novaLinha, 1).activate();

  // 3. Retorna os dados da nova linha + o progresso recalculado em 1 único payload
  return {
    influenciador: obterDadosDaLinha(sheet, novaLinha),
    progresso: calcularProgressoRapido(sheet)
  };
}

/**
 * Retorna os dados de uma linha específica sem reler toda a planilha
 */
function obterDadosDaLinha(sheet, linha) {
  const dados = sheet.getRange(linha, 1, 1, 3).getValues()[0];
  let username = String(dados[1] || "");
  let userClean = username.replace(/@/g, "").trim();

  return {
    linha: linha,
    nome: String(dados[0] || ""),
    username: username,
    statusAtual: String(dados[2] || ""),
    url: userClean ? "https://instagram.com/" + userClean : ""
  };
}

/**
 * Obtém os dados da linha atualmente ativa
 */
function obterDadosLinhaAtiva() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let linha = sheet.getActiveCell().getRow();
  if (linha < 3) linha = 3;
  return obterDadosDaLinha(sheet, linha);
}

/**
 * Versão ultra-rápida de cálculo de progresso
 */
function calcularProgressoRapido(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultLinha = sheet.getLastRow();
  
  if (ultLinha < 3) {
    return { dataTexto: "Hoje", percentual: 0, concluidos: 0, restantes: 0, total: 0 };
  }

  const dataVal = sheet.getRange("D2").getDisplayValue() || sheet.getRange("B1").getDisplayValue();
  const valoresStatus = sheet.getRange(3, 3, ultLinha - 2, 1).getValues();
  
  let concluidos = 0;
  let total = valoresStatus.length;

  for (let i = 0; i < total; i++) {
    let st = String(valoresStatus[i][0]).trim().toLowerCase();
    if (st !== "" && st !== "pendente" && !st.includes("não postou")) {
      concluidos++;
    }
  }

  return {
    dataTexto: dataVal || "Hoje",
    percentual: total > 0 ? Math.round((concluidos / total) * 100) : 0,
    concluidos: concluidos,
    restantes: total - concluidos,
    total: total
  };
}

/**
 * Trigger automático ao mudar D2
 */
function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getA1Notation().indexOf('D2') !== -1) {
    carregarHistoricoPorData();
  }
}

/**
 * Busca ultra-rápida de histórico na aba "BANCO_DE_DADOS"
 */
function carregarHistoricoPorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAtual = ss.getActiveSheet();
  
  const sheetBD = ss.getSheetByName('BANCO_DE_DADOS') || ss.getSheetByName('banco_de_dados');

  if (!sheetBD) {
    ss.toast('Aba "BANCO_DE_DADOS" não encontrada.', 'Erro ❌');
    return;
  }

  const dataBuscada = sheetAtual.getRange('D2').getDisplayValue().trim();
  if (!dataBuscada) {
    ss.toast('Célula D2 está vazia.', 'Aviso ⚠️');
    return;
  }

  const ultLinhaBD = sheetBD.getLastRow();
  if (ultLinhaBD < 2) {
    ss.toast('A aba BANCO_DE_DADOS está sem registros.', 'Aviso ℹ️');
    return;
  }

  // Pega 4 colunas: Col A (Data), Col B (Nome), Col C (@username), Col D (Status)
  const dadosBD = sheetBD.getRange(2, 1, ultLinhaBD - 1, 4).getDisplayValues();
  const mapaStatus = new Map();
  let encontrouData = false;

  for (let i = 0; i < dadosBD.length; i++) {
    let dataLinha = dadosBD[i][0].trim();

    if (dataLinha === dataBuscada || normalizarData(dataLinha) === normalizarData(dataBuscada)) {
      encontrouData = true;
      
      let nome = dadosBD[i][1].trim().toLowerCase(); // Coluna B
      let user = dadosBD[i][2].trim().toLowerCase(); // Coluna C (@username)
      let status = dadosBD[i][3].trim();             // Coluna D (Status Real)

      // Vincula o Status tanto ao Nome quanto ao @username
      if (nome) mapaStatus.set(nome, status);
      if (user) mapaStatus.set(user, status);
    }
  }

  const ultLinhaAtual = sheetAtual.getLastRow();
  if (ultLinhaAtual < 3) return;

  const rangeAtual = sheetAtual.getRange(3, 1, ultLinhaAtual - 2, 3);
  const valoresAtual = rangeAtual.getValues();

  for (let i = 0; i < valoresAtual.length; i++) {
    let nome = String(valoresAtual[i][0]).trim().toLowerCase();
    let user = String(valoresAtual[i][1]).trim().toLowerCase();

    if (mapaStatus.has(nome)) {
      valoresAtual[i][2] = mapaStatus.get(nome);
    } else if (mapaStatus.has(user)) {
      valoresAtual[i][2] = mapaStatus.get(user);
    } else if (encontrouData) {
      valoresAtual[i][2] = ""; // Reseta caso a data exista mas não haja registro deste influenciador
    }
  }

  rangeAtual.setValues(valoresAtual);

  if (encontrouData) {
    ss.toast(`Status de ${dataBuscada} carregados com sucesso!`, 'Sucesso ⚡');
  } else {
    ss.toast(`Data "${dataBuscada}" não encontrada no BANCO_DE_DADOS.`, 'Aviso ℹ️');
  }
}

/**
 * Auxiliar para normalizar formatação de datas (ex: 5/8/2026 -> 05/08/2026)
 */
function normalizarData(str) {
  if (!str) return "";
  let partes = str.split('/');
  if (partes.length === 3) {
    return `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${partes[2]}`;
  }
  return str;
}
/**
 * Gera o Ranking de Assiduidade detalhado processando os registros da aba BANCO_DE_DADOS
 */
function gerarRankingAssiduidadeDetalhes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBD = ss.getSheetByName('BANCO_DE_DADOS') || ss.getSheetByName('banco_de_dados');

  if (!sheetBD) {
    ss.toast('Aba "BANCO_DE_DADOS" não encontrada!', 'Erro ❌');
    return;
  }

  const ultLinhaBD = sheetBD.getLastRow();
  if (ultLinhaBD < 2) {
    ss.toast('Nenhum dado encontrado no Banco de Dados.', 'Aviso ⚠️');
    return;
  }

  // Lê Coluna A (Data), Coluna B (Nome), Coluna C (@username), Coluna D (Status)
  const dadosBD = sheetBD.getRange(2, 1, ultLinhaBD - 1, 4).getDisplayValues();
  const estatisticas = {};

  dadosBD.forEach(linha => {
    const nome = linha[1].trim();
    const username = linha[2].trim();
    const status = linha[3].trim().toLowerCase();

    if (!nome && !username) return;

    const chave = username || nome;
    if (!estatisticas[chave]) {
      estatisticas[chave] = {
        nome: nome || username,
        username: username,
        totalDias: 0,
        postou: 0,
        naoPostou: 0
      };
    }

    estatisticas[chave].totalDias++;
    
    if (status.includes('não postou')) {
      estatisticas[chave].naoPostou++;
    } else if (status !== "" && status !== "pendente") {
      estatisticas[chave].postou++;
    }
  });

  // Localiza ou cria a aba de relatório "Ranking Assiduidade"
  let sheetRanking = ss.getSheetByName('Ranking Assiduidade');
  if (!sheetRanking) {
    sheetRanking = ss.insertSheet('Ranking Assiduidade');
  } else {
    sheetRanking.clearContents();
  }

  // Cabeçalho do relatório
  const relatorio = [
    ['Posição', 'Influenciador', 'Username', 'Dias Monitorados', 'Postagens Confirmadas', 'Não Postou', 'Taxa de Assiduidade']
  ];

  // Ordena por maior porcentagem de assiduidade e maior volume de postagens
  const listaRanking = Object.values(estatisticas).map(item => {
    const taxa = item.totalDias > 0 ? (item.postou / item.totalDias) : 0;
    return { ...item, taxa };
  });

  listaRanking.sort((a, b) => b.taxa - a.taxa || b.postou - a.postou);

  listaRanking.forEach((item, index) => {
    relatorio.push([
      index + 1,
      item.nome,
      item.username,
      item.totalDias,
      item.postou,
      item.naoPostou,
      item.taxa // Valor numérico para permitir formatação de porcentagem no Sheets
    ]);
  });

  // Escreve os dados no relatório
  sheetRanking.getRange(1, 1, relatorio.length, 7).setValues(relatorio);
  
  // Estilização do Relatório
  sheetRanking.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff');
    
  if (relatorio.length > 1) {
    sheetRanking.getRange(2, 7, relatorio.length - 1, 1).setNumberFormat('0.0%');
  }

  sheetRanking.autoResizeColumns(1, 7);
  sheetRanking.activate();

  ss.toast('Ranking de Assiduidade gerado com sucesso!', 'Sucesso 📊');
}
