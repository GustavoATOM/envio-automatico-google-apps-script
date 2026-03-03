function envioAutomatico() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaSolic = planilha.getSheetByName("Solicitações");
  const abaFrases = planilha.getSheetByName("Frases");

  const dadosSolic = abaSolic.getDataRange().getValues();
  const dadosFrases = abaFrases.getDataRange().getValues();

  // ==============================
  // DICIONÁRIO DE FRASES
  // ==============================
  let frases = {};
  for (let i = 1; i < dadosFrases.length; i++) {
    frases[dadosFrases[i][0]] = dadosFrases[i][1];
  }

  // ==============================
  // MESES EM PORTUGUÊS
  // ==============================
  const meses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
  ];

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  // ==============================
  // LOOP DAS SOLICITAÇÕES
  // ==============================
  for (let i = 1; i < dadosSolic.length; i++) {

    let status = dadosSolic[i][0];        // Coluna A
    let condominio = dadosSolic[i][1];    // Coluna B
    let dataLeitura = dadosSolic[i][2];   // Coluna C
    let dataEnvio = dadosSolic[i][3];     // Coluna D
    let emails = dadosSolic[i][4];        // Coluna E
    let codigoFrase = dadosSolic[i][5];   // Coluna F

    if (status !== "Pendente") continue;
    if (!dataEnvio || !emails || !codigoFrase) continue;

    let envio = new Date(dataEnvio);
    envio.setHours(0, 0, 0, 0);

    if (envio.getTime() !== hoje.getTime()) continue;

    let textoBase = frases[codigoFrase];
    if (!textoBase) continue;

    // ==============================
    // FORMATAR DATA DA LEITURA
    // ==============================
    let dataLeituraObj = new Date(dataLeitura);

    let dataLeituraFormatada =
      String(dataLeituraObj.getDate()).padStart(2, "0") + "/" +
      String(dataLeituraObj.getMonth() + 1).padStart(2, "0") + "/" +
      dataLeituraObj.getFullYear();

    let textoFinal = textoBase
      .replace("{{CONDOMINIO}}", condominio)
      .replace("{{DATA_LEITURA}}", dataLeituraFormatada);

    // ==============================
    // MÊS/ANO EM PORTUGUÊS
    // ==============================
    let mesNome = meses[dataLeituraObj.getMonth()];
    let ano = dataLeituraObj.getFullYear();
    let mesAno = mesNome + "/" + ano;

    let assunto = "Solicitação de conta – " + condominio + " - " + mesAno;

    // ==============================
    // CONVERTE QUEBRA DE LINHA
    // ==============================
    let textoHtml = textoFinal.replace(/\n/g, "<br>");

    // ==============================
    // ENVIO COM HTML + VERDANA
    // ==============================
    GmailApp.sendEmail(
      emails,
      assunto,
      textoFinal, // corpo texto simples (fallback)
      {
        htmlBody: `
          <div style="
            font-family: Verdana, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #000000;
          ">
            ${textoHtml}
          </div>
        `
      }
    );

    // ==============================
    // MARCA COMO ENVIADO
    // ==============================
    abaSolic.getRange(i + 1, 1).setValue("Enviado");
    abaSolic.getRange(i + 1, 7).setValue(new Date());
  }
}
