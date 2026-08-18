function inicializarEstruturaCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Criar ou Obter Aba ACOMPANHAMENTO
  let sheetAcomp = ss.getSheetByName("ACOMPANHAMENTO");
  if (!sheetAcomp) {
    sheetAcomp = ss.insertSheet("ACOMPANHAMENTO");
  }
  sheetAcomp.getRange("A1:C2").setValues([
    ["PAINEL DE VERIFICAÇÃO", "", ""],
    ["INFLUENCIADOR", "@USERNAME", "LINK/PERFIL"]
  ]);
  sheetAcomp.getRange("A1:C2").setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");

  // 2. Criar ou Obter Aba HISTORICO
  let sheetHist = ss.getSheetByName("HISTORICO");
  if (!sheetHist) {
    sheetHist = ss.insertSheet("HISTORICO");
  }
  sheetHist.getRange("A1").setValue("SELECIONE O INFLUENCIADOR:").setFontWeight("bold");
  sheetHist.getRange("A4:B4").setValues([["DATA", "STATUS DA POSTAGEM"]]).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");

  // 3. Criar ou Obter Aba RANKING
  let sheetRank = ss.getSheetByName("RANKING");
  if (!sheetRank) {
    sheetRank = ss.insertSheet("RANKING");
  }
  sheetRank.getRange("A1").setValue("RANKING DE ASSIDUIDADE").setFontWeight("bold");
  sheetRank.getRange("A3:D3").setValues([["POSIÇÃO", "INFLUENCIADOR", "POSTAGENS REALIZADAS", "TAXA DE CUMPRIMENTO"]]).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");

  // Atualizar Menus
  atualizarMenusDropdown();
  SpreadsheetApp.getUi().alert("Estrutura completa das 4 abas configurada com sucesso!");
}

function atualizarMenusDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAcomp = ss.getSheetByName("ACOMPANHAMENTO");
  const sheetHist = ss.getSheetByName("HISTORICO");

  if (!sheetAcomp || !sheetHist) return;

  const lastRow = sheetAcomp.getLastRow();
  if (lastRow < 3) return;

  const nomes = sheetAcomp.getRange(3, 1, lastRow - 2, 1).getValues().flat().filter(String);

  if (nomes.length > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(nomes, true).setAllowInvalid(false).build();
    sheetHist.getRange("B2").setDataValidation(rule);
  }
}
