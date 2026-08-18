function registrarHistorico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHist = ss.getSheetByName("HISTORICO");
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  
  if (!sheetHist || !sheetBanco) return "Abas não encontradas.";
  
  // Pega a data de B2 (se vazia, tenta B1)
  let dataRaw = sheetHist.getRange("B2").getValue();
  if (!dataRaw) dataRaw = sheetHist.getRange("B1").getValue();
  
  const dataAlvo = normalizarData(dataRaw);
  
  // Mapeia o Banco de Dados (Data + Username => Status)
  const dadosBanco = sheetBanco.getDataRange().getValues();
  const mapaStatus = {};
  
  for (let i = 1; i < dadosBanco.length; i++) {
    const dataLinha = normalizarData(dadosBanco[i][0]);  // Coluna A
    const userLinha = normalizarTexto(dadosBanco[i][2]); // Coluna C (@username)
    const statusLinha = dadosBanco[i][3];               // Coluna D (Status)
    
    if (dataLinha && userLinha) {
      mapaStatus[dataLinha + "_" + userLinha] = statusLinha;
    }
  }
  
  // Lê influenciadores do Histórico (linha 4 até o fim)
  const ultLinha = sheetHist.getLastRow();
  if (ultLinha < 4) return "Histórico sem dados.";
  
  const usernamesHist = sheetHist.getRange(4, 2, ultLinha - 3, 1).getValues();
  const resultados = [];
  
  for (let i = 0; i < usernamesHist.length; i++) {
    const uRaw = usernamesHist[i][0];
    if (!uRaw) {
      resultados.push([""]);
      continue;
    }
    
    const userLimpo = normalizarTexto(uRaw);
    const chave = dataAlvo + "_" + userLimpo;
    
    if (mapaStatus[chave] !== undefined && mapaStatus[chave] !== "") {
      resultados.push([mapaStatus[chave]]);
    } else {
      resultados.push(["Não postou"]);
    }
  }
  
  // Atualiza a Coluna C da aba HISTORICO (linha 4 em diante)
  sheetHist.getRange(4, 3, resultados.length, 1).setValues(resultados);
  return "Histórico atualizado para a data: " + dataAlvo;
}

function gerarRelatorio() {
  return registrarHistorico();
}

// Gatilho automático ao alterar a data na planilha
function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() === "HISTORICO" && (range.getA1Notation() === "B2" || range.getA1Notation() === "B1")) {
    registrarHistorico();
  }
}

// Função para padronizar datas (DD/MM/YYYY)
function normalizarData(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT-0300", "dd/MM/yyyy");
  }
  const str = String(val).trim();
  const partes = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (partes) {
    return partes[1].padStart(2, '0') + "/" + partes[2].padStart(2, '0') + "/" + partes[3];
  }
  return str;
}

// Função para remover @, maiúsculas e caracteres invisíveis
function normalizarTexto(val) {
  if (!val) return "";
  return String(val)
    .toLowerCase()
    .replace(/@/g, "")
    .replace(/[\s\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff\u00a0]/g, "")
    .trim();
}
