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