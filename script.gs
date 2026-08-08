function doPost(e) {
  try {
    // Abre a folha de cálculo onde o script está vinculado
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    // Processar o ficheiro (Base64)
    const fileBlob = Utilities.newBlob(Utilities.base64Decode(data.fileBase64), data.fileMimeType, data.fileName);
    
    // Guardar o ficheiro no Google Drive (na raiz)
    // Se quiser guardar numa pasta específica, troque a linha abaixo por:
    // const folder = DriveApp.getFolderById("ID_DA_SUA_PASTA");
    const folder = DriveApp.getRootFolder(); 
    
    const file = folder.createFile(fileBlob);
    const fileUrl = file.getUrl();
    
    // Criar uma data/hora amigável
    const timestamp = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    
    // Adicionar linha à folha de cálculo com todas as colunas novas
    sheet.appendRow([
      timestamp,
      data.nome,
      data.whatsapp,
      data.email,
      data.cidade,
      data.idade,
      data.objetivo,
      data.dificuldade,
      data.restricao,
      data.acompanhamento,
      data.conheceu,
      data.confirmacao,
      fileUrl
    ]);
    
    // === ENVIAR NOTIFICAÇÃO POR E-MAIL ===
    const emailNotificacao = "gabes.kerkhoff@gmail.com";
    const assunto = "Nova Inscrição - Cuidando do Templo: " + data.nome;
    const corpoEmail = "Olá!\n\n" +
                       "Acabou de receber uma nova inscrição no grupo Cuidando do Templo.\n\n" +
                       "Resumo dos Dados:\n" +
                       "- Nome: " + data.nome + "\n" +
                       "- WhatsApp: " + data.whatsapp + "\n" +
                       "- E-mail: " + data.email + "\n" +
                       "- Objetivo: " + data.objetivo + "\n\n" +
                       "Ver comprovante de pagamento: " + fileUrl + "\n\n" +
                       "Todos os outros detalhes já estão na Folha de Cálculo.";
                       
    GmailApp.sendEmail(emailNotificacao, assunto, corpoEmail);
    // =====================================
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "fileUrl": fileUrl }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch(error) {
    // SE ALGO FALHAR (ex: erro ao enviar e-mail), ESCREVE NA PLANILHA PARA SABERMOS O QUE FOI!
    const sheetError = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheetError.appendRow(["ERRO NO SISTEMA", error.toString()]);
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function autorizarEmail() {
  GmailApp.sendEmail("gabes.kerkhoff@gmail.com", "Teste de Autorização", "Autorização concluída com sucesso!");
}

function doOptions(e) {
  // Lidar com pedidos CORS (se aplicável)
  return ContentService.createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}
