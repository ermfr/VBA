' ******************************************************************************
' MACRO per salvataggio di PDF e DXF disegno
' Salvataggio PDF in cartella disegno SW, con stesso nome
' Salvataggio in cartella ARCHIVIO (DrwPath1) Numero disegno + revisione
' Salvataggio in cartella ARCHIVIO (DrwPath1) Numero disegno + revisione - formato DXF
' Salvataggio in cartella GESTIONALE (DrwPath2) Numero disegno
'
'Versione provvisoria senza recupero numero disegno dal modello
'
' ******************************************************************************

'Per recupero proprietà da modello
Option Explicit


Dim swApp           As SldWorks.SldWorks
Dim swmodel         As SldWorks.ModelDoc2
Dim swmod           As SldWorks.ModelDoc2
Dim swdraw          As SldWorks.DrawingDoc
Dim swview          As SldWorks.View
Dim swCustPropMgr   As SldWorks.CustomPropertyManager
Dim model           As SldWorks.ModelDoc2
Dim swConfigMgr     As SldWorks.ConfigurationManager
Dim swConfig        As SldWorks.Configuration
'Dim swCustPropMgr   As SldWorks.CustomPropertyManager

Dim Part As Object
Dim v As Variant
Dim NumDis As String
Dim config As Variant
Dim comp As SldWorks.Component2

Dim FilePath As String
Dim PathSize As Long
Dim PathNoExtention As String
Dim NewFilePath As String
Dim Rev As String
Dim NumDisProp As String
Dim nResponse As String
Dim DrwPath1 As String
Dim DrwPath2 As String
Dim FilePathName1 As String
Dim FilePathName2 As String
Dim FilePathName1_dxf As String

'Dim vPropNames          As Variant
'Dim valOut              As String
'Dim resolvedValOut      As String

Dim valOut              As String


Sub main()


'Recupero numero disegno
Set swApp = Application.SldWorks
'Set swmodel = swApp.ActiveDoc
Set swdraw = swApp.ActiveDoc
'Set swdraw = swmodel
Set swview = swdraw.GetFirstView
Set swview = swview.GetNextView

NumDisProp = "Numero disegno"

'Configurazione corrente
Set swConfigMgr = swview.ReferencedDocument.ConfigurationManager
Set swConfig = swConfigMgr.ActiveConfiguration

Debug.Print "Name of this configuration: " & swConfig.Name

NumDis = swview.ReferencedDocument.CustomInfo2(swConfig.Name, NumDisProp)

Debug.Print "    Name, swCustomInfoType_e value, and resolved value:  " & NumDisProp & ", " & NumDis

'v = swview.GetVisibleComponents
'Set comp = v(0)
'Set swmod = comp.GetModelDoc2
'Set swCustPropMgr = swmodel.Extension.CustomPropertyManager("")


'Configurazione corrente
'Set swConfigMgr = swmodel.ConfigurationManager
'Set swConfig = swConfigMgr.ActiveConfiguration

'Debug.Print "Name of this configuration: " & swConfig.Name

 'Set swCustPropMgr = swConfig.CustomPropertyManager

'swCustPropMgr.Get2 NumDisProp, valOut, NumDis
'Debug.Print "    Name, swCustomInfoType_e value, and resolved value:  " & NumDisProp & ", " & NumDis
'NumDis = swmod.GetCustomInfoValue(config, NumDisProp)

'Fine recupero numero disegno


'Salvataggio PDF

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

Set model = swApp.ActiveDoc

'cartella archivio

DrwPath1 = "I:\DOCUMENTAZONE_TECNICA\20-ARCHIVIO\10-DISEGNI\"
'vecchia posizione - DrwPath1 = "\\EUROMAGLINUX\dati\Ufficio_Tecnico\DISEGNI\Disegni_PDF\"

'cartella gestionale
DrwPath2 = "\\EUROMAGLINUX\dati\Ufficio_Tecnico\Disegni_PDF\"

' forzatura numero disegno da disegno
If NumDis = "" Then
    NumDis = model.CustomInfo("Numero disegno")
    Else

End If


Rev = model.CustomInfo("Revisione numero")

'Aggiungere anche numero parte
nResponse = MsgBox("Numero disegno: " & NumDis & Chr(13) & "Revisione: " & Rev & Chr(13) & "do you want to Continue?", vbYesNo)

If nResponse = vbYes Then
       
    FilePath = Part.GetPathName
    PathSize = Strings.Len(FilePath)
    PathNoExtention = Strings.Left(FilePath, PathSize - 7)
    NewFilePath = PathNoExtention & ".pdf"

    'assegnazione nome per disegno in local folder
    FilePathName1 = DrwPath1 & NumDis & "-" & Rev & ".pdf"
    FilePathName2 = DrwPath2 & NumDis & ".pdf"

    'salvataggio PDF files
    Part.SaveAs2 NewFilePath, 0, True, False
    Part.SaveAs2 FilePathName1, 0, True, False
    Part.SaveAs2 FilePathName2, 0, True, False
    
    nResponse = MsgBox("do you want to Archivie DXF?", vbYesNo)
    If nResponse = vbYes Then
        FilePathName1_dxf = DrwPath1 & NumDis & "-" & Rev & ".dxf"
        Part.SaveAs2 FilePathName1_dxf, 0, True, False
        Else
    End If

    'Fine salvataggio file
    'Stampa Messaggi di stato
    '(Manca condizione di verifica anche per salvataggio o meno DXF)

    MsgBox "Disegno Numero: " & NumDis & "Saved following files: " & Chr(13) & NewFilePath & Chr(13) & FilePathName1 & Chr(13) & FilePathName1_dxf & Chr(13) & FilePathName2
       
    Else
    Exit Sub
End If

End Sub
