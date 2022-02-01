# certificate-automation
Documentation of my procedure to generate and delivery certificates

- The 1st step is compile the data to be used.
- The 2nd step is to create a direct mail.
- - If the direct mail is in MSWord, you can use a litle VBA script to create multiple PDFs individually named:

```
Sub GeraPDFOrient()
'
' GeraPDFOrient Macro
'
'
    ActiveDocument.MailMerge.DataSource.ActiveRecord = wdFirstRecord
    For i = 1 To ActiveDocument.MailMerge.DataSource.RecordCount

        Nome = "DecOrient - " & ActiveDocument.MailMerge.DataSource.DataFields("NomeDecOrient").Value

        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "C:\Users\Vinicius\Desktop\DecOrient\" & Nome & ".pdf" _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False

        ActiveDocument.MailMerge.DataSource.ActiveRecord = wdNextRecord

    Next i
End Sub
```

- - If you already have a single PDF file, you can use a PDF-Spliter like mine[^1]:

[^1]: For more details, see: https://github.com/viniciusblanc/pdf-splitter

```
from pikepdf import Pdf

pdf = Pdf.open('Certificados.pdf')

for n, page in enumerate(pdf.pages):
    dst = Pdf.new()
    dst.pages.append(page)
    dst.save(f'Certificado{n+1:02d}.pdf')
```

- The 3rd step is to upload all these files to a Google Drive.
- Once there, you must create a Google Sheet with the information that will be used to automate the emails.
- - On the Google Sheet, go to Extensions/App Script to put the following code:

```
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
```
  
