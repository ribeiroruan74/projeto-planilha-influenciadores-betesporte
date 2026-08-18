function enviarRelatorioDiarioEmail() {
  const sheet = getAbaAcompanhamento();
  
  // Garante que a lista está ordenada em ordem alfabética antes de gerar o relatório
  ordenarInfluenciadoresAlfabetico();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3 || lastCol < 4) return;

  const colData = lastCol;
  const dataTexto = sheet.getRange(2, colData).getDisplayValue();

  const dados = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  let totalInf = 0;
  let postaram = 0;
  let naoPostaram = 0;
  let semRegistro = 0;

  let listaDetalhadaHtml = "";

  for (let i = 0; i < dados.length; i++) {
    const nome = String(dados[i][0]).trim();
    let username = String(dados[i][1]).trim();
    const status = String(dados[i][colData - 1]).trim();

    if (!nome) continue;
    totalInf++;

    if (!username) username = "@" + nome.toLowerCase().replace(/\s+/g, "");

    let statusBadge = "";
    let statusLower = status.toLowerCase();

    if (statusLower === "" || statusLower === "aguardando") {
      semRegistro++;
      statusBadge = "<span style='color: #7f8c8d; font-weight: bold;'>⚪ Pendente</span>";
    } else if (statusLower === "não postou") {
      naoPostaram++;
      statusBadge = "<span style='color: #c0392b; font-weight: bold;'>🔴 Não postou</span>";
    } else {
      postaram++;
      statusBadge = `<span style='color: #27ae60; font-weight: bold;'>🟢 ${status}</span>`;
    }

    listaDetalhadaHtml += `
      <tr style="border-bottom: 1px solid #eee;">
        <td style="padding: 8px; font-weight: bold;">${nome}</td>
        <td style="padding: 8px; color: #555;">${username}</td>
        <td style="padding: 8px;">${statusBadge}</td>
      </tr>`;
  }

  const pctPostaram = totalInf > 0 ? ((postaram / totalInf) * 100).toFixed(1) : 0;
  const pctNaoPostaram = totalInf > 0 ? ((naoPostaram / totalInf) * 100).toFixed(1) : 0;

  const emailDestino = Session.getActiveUser().getEmail();
  const assunto = `📊 Relatório Diário BETEsporte - ${dataTexto}`;

  const corpoHtml = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; color: #333; line-height: 1.5;">
      <h2 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 8px;">
        ⚽ Resumo de Acompanhamento - ${dataTexto}
      </h2>

      <div style="display: flex; gap: 10px; margin-bottom: 20px;">
        <div style="background: #f8f9fa; border: 1px solid #e0e0e0; padding: 12px; border-radius: 8px; flex: 1; text-align: center;">
          <span style="font-size: 11px; color: #666; font-weight: bold; display: block;">TOTAL</span>
          <span style="font-size: 20px; font-weight: bold; color: #333;">${totalInf}</span>
        </div>
        <div style="background: #e6f4ea; border: 1px solid #b7e1cd; padding: 12px; border-radius: 8px; flex: 1; text-align: center;">
          <span style="font-size: 11px; color: #137333; font-weight: bold; display: block;">POSTARAM</span>
          <span style="font-size: 20px; font-weight: bold; color: #137333;">${postaram} (${pctPostaram}%)</span>
        </div>
        <div style="background: #fce8e6; border: 1px solid #f5c6cb; padding: 12px; border-radius: 8px; flex: 1; text-align: center;">
          <span style="font-size: 11px; color: #c5221f; font-weight: bold; display: block;">NÃO POSTARAM</span>
          <span style="font-size: 20px; font-weight: bold; color: #c5221f;">${naoPostaram} (${pctNaoPostaram}%)</span>
        </div>
      </div>

      <h3>📋 Lista Organizada por Ordem Alfabética:</h3>
      <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 13px;">
        <thead>
          <tr style="background-color: #f1f3f4; color: #3c4043;">
            <th style="padding: 10px;">Influenciador</th>
            <th style="padding: 10px;">Username</th>
            <th style="padding: 10px;">Status (${dataTexto})</th>
          </tr>
        </thead>
        <tbody>
          ${listaDetalhadaHtml}
        </tbody>
      </table>
    </div>
  `;

  MailApp.sendEmail({
    to: emailDestino,
    subject: assunto,
    htmlBody: corpoHtml
  });
}
