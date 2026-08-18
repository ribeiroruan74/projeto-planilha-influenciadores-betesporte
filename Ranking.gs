function atualizarRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBanco = ss.getSheetByName("BANCO_DE_DADOS");
  let sheetRanking = ss.getSheetByName("RANKING");
  
  if (!sheetBanco) return "Aba BANCO_DE_DADOS não encontrada.";
  if (!sheetRanking) {
    sheetRanking = ss.insertSheet("RANKING");
  }

  const dados = sheetBanco.getDataRange().getValues();
  if (dados.length <= 1) return "Banco de dados vazio.";

  const totalDiasSet = new Set();
  const contagem = {};

  // Processa os registros (Data na Coluna A, Username na C, Status na D)
  for (let i = 1; i < dados.length; i++) {
    const data = dados[i][0];
    const user = String(dados[i][2] || "").trim();
    const statusBruto = dados[i][3];

    // Pula linhas onde o influenciador está em branco
    if (!user) continue;

    if (data) totalDiasSet.add(String(data));

    if (!contagem[user]) {
      contagem[user] = { postou: 0, total: 0 };
    }

    // 1. Garante que a célula de status não está em branco ou nula
    if (statusBruto !== null && statusBruto !== undefined && String(statusBruto).trim() !== "") {

      // 2. Transforma em string, converte para minúsculas e remove acentos
      const statusNormalizado = String(statusBruto)
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, ""); // Transforma "não" em "nao"

      // 3. Define os termos que indicam ausência de postagem (sem acentos)
      const termosInvalidos = ["nao postou", "pendente"];

      // 4. Verifica se o status contém algum dos termos inválidos
      const naoPostou = termosInvalidos.some(termo => statusNormalizado.includes(termo));

      // Contabiliza o total apenas se a célula tiver verificação válida
      contagem[user].total += 1;

      // Se NÃO for "não postou" / "pendente", contabiliza a postagem
      if (!naoPostou) {
        contagem[user].postou += 1;
      }
    }
  } // <-- Chave do loop 'for' corrigida aqui

  // Monta a matriz de resultados
  const rankingArray = [];
  for (const user in contagem) {
    const postou = contagem[user].postou;
    const total = contagem[user].total;
    const taxa = total > 0 ? (postou / total) : 0;
    rankingArray.push([user, postou, total, taxa]);
  }

  // Ordena por Taxa de Assiduidade (%) em ordem decrescente; empate decide por número de dias postados
  rankingArray.sort((a, b) => {
    if (b[3] !== a[3]) {
      return b[3] - a[3];
    }
    return b[1] - a[1];
  });

  // Prepara o cabeçalho e limpa a aba RANKING
  sheetRanking.clear();
  const headers = [["Posição", "Influenciador", "Dias Postados", "Total Oportunidades", "Assiduidade (%)"]];
  sheetRanking.getRange(1, 1, 1, 5).setValues(headers)
    .setBackground("#1b5e20")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  const finalData = rankingArray.map((row, index) => [index + 1, row[0], row[1], row[2], row[3]]);

  if (finalData.length > 0) {
    sheetRanking.getRange(2, 1, finalData.length, 5).setValues(finalData);
    sheetRanking.getRange(2, 5, finalData.length, 1).setNumberFormat("0.0%");
    sheetRanking.autoResizeColumns(1, 5);
  }

  SpreadsheetApp.getUi().alert("🏆 Ranking atualizado com sucesso!");
}
