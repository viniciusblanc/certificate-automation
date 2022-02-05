function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Envio automático')
      .addItem('Enviar e-mail', 'sendEmails')
      .addToUi();
}

var EMAIL_SENT = 'Sim';

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // Linha inicial
  var numRows = 62; // Número de linhas para processar
  
  var dataRange = sheet.getRange(startRow, 1, numRows, 6); // Atenção aqui! Índices de colunas e de linhas começam em 1, vai até a coluna 6 - F

  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[2]; // Coluna do endereço 2 - C

    var mensagem = 'Boa tarde, Prof(a). ' + row[1] + '.<br/><br/>'+
                   'Segue anexa sua declaração por ter orientado o Trabalho de Conclusão de Curso da Graduação em Enfermagem (EPE/UNIFESP) <b>"'+row[3]+'"</b>, de <b>'+row[0]+'</b>, em 2021.<br/><br/>' + 
                   'Este é um reenvio porque em alguns casos acabei mandando declarações do ano passado.<br/>'+
                   'Peço que confira se a declaração está correta e confirme o recebimento desta mensagem.<br/><br/>'+
                   'Obrigado,<br/>--<br/>Vinicius Farias<br/>Secretaria de Graduação<br/>Escola Paulista de Enfermagem<br/>Universidade Federal de São Paulo<br/>(11) 5576-4430 Voip 2552';
    
    var emailSent = row[5]; // Coluna da confirmação de envio

    if (emailSent !== EMAIL_SENT) { // Evita envios duplicados
      var subject = 'Declaração de Orientação - TCC ENF/EPE 2021 - ERRATA';
      
      files=DriveApp.getFilesByName('DecOrient - ' + row[4] + '.pdf');
      while (files.hasNext())file=files.next();

      MailApp.sendEmail(emailAddress, subject, mensagem, { htmlBody: mensagem, attachments: [file.getAs(MimeType.PDF)] });
      sheet.getRange(startRow + i, 6).setValue(EMAIL_SENT); // Posição do Enviado 6 - F

      SpreadsheetApp.flush();
    }
  }
}