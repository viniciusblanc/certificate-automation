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
