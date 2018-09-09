Sub report()
'Copy informations from one file (registry) to another (report)
    
Application.ScreenUpdating = False

Dim wb_report As Workbook, ws_report As Worksheet
Dim wb_reg As Workbook, ws_reg As Worksheet
Dim filename As String, filenamePDF As String, reportName As String, recordsFolder As String
Dim test_info(4) As Variant, checklist(12) As Variant, values(4) As Double
Dim row As Integer, i As Integer

Set wb_reg = ActiveWorkbook
Set ws_reg = ActiveSheet

'cartella di salvataggio dei records, definita come parametreo nel foglio di calcolo
recordsFolder = Foglio2.Range("E10")
recordsFolder = Right(recordsFolder, Len(recordsFolder) - 1)

'Raccoglie le informazioni dal registro in 3 array:
'test_info -> info generali del test
'checklist -> controlli prima di iniziare
'values -> rilevamento variabili / pressoine acqua, pressione atmosferica, conducibilità, umidità relativa

row = ActiveCell.row

For i = 1 To 4
    test_info(i) = Cells(row, i)
Next i

For i = 1 To 12
    checklist(i) = Cells(row, i + 4)
Next i

For i = 1 To 4
    values(i) = Cells(row, i + 16)
Next i

'-------passaggio altro file-------
'apre il template per la l'emissione del report
Workbooks.Open filename:=ThisWorkbook.Path & "\report_tar.xlsx"
' nome del file di report (ID Prova)

reportName = "report_" & test_info(1)
filename = ThisWorkbook.Path & recordsFolder & reportName & ".xlsx"
filenamePDF = ThisWorkbook.Path & recordsFolder & reportName & ".pdf"

Set wb_report = Workbooks("report_tar.xlsx")
'copia i valori sul file di report
For i = 1 To 4
    wb_report.Worksheets("Foglio1").Cells(i + 2, 2) = test_info(i)
Next i
For i = 1 To 12
    wb_report.Worksheets("Foglio1").Cells(i + 7, 2) = checklist(i)
Next i
For i = 1 To 4
    wb_report.Worksheets("Foglio1").Cells(i + 21, 2) = values(i)
Next i

wb_report.Worksheets("Report").Activate

wb_report.SaveAs filename:=filename
wb_report.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        filename:=filenamePDF, _
        from:=1, to:=1
        'Quality:=xlQualityStandard, _
        'IncludeDocProperties:=True, _
        'IgnorePrintAreas:=False, _
        'OpenAfterPublish:=False
        
wb_report.Close
ws_reg.Cells(row, 21) = reportName & ".pdf"
MsgBox ("Report creato: " & reportName & ".pdf")

Application.ScreenUpdating = True
End Sub
