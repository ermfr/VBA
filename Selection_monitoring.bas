'Selection monitoring
'Modifica il contenuto di una cella alternativamente quando questa viene selezionata
'Si applica al foglio / Selezionare dall'albero il foglio da monitorare, nella finestra del codice selezionare Worksheet / SelectionChange
'Nell'esempio viene monitorato l'intervallo: "E6:P100"

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Not Intersect(Target, Range("E6:P100")) Is Nothing Then
If Target.Count > 1 Then Exit Sub 'esci se viene fatta una selezione multipla
If Target.Value = "þ" Then Target.Value = "¨" Else: Target.Value = "þ"
End If
End Sub
