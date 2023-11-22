Attribute VB_Name = "BrGroup_Epson"
'1. Sub Epson():
'Esta sub-rotina realiza a consolidação de dados de várias planilhas, incluindo "ZSTOK Epson BASE.xlsx", "MB51 Epson BASE.xlsx" e "FUP Epson BASE.xlsx", bem como a planilha "PSI Epson.xlsm".
'
'Passo a passo:
'1. Declaração de várias variáveis para armazenar as pastas de trabalho e planilhas que serão usadas na consolidação.
'2. As variáveis `filePaths` são usadas para armazenar os nomes dos arquivos que serão abertos: "ZSTOK Epson BASE.xlsx", "MB51 Epson BASE.xlsx" e "FUP Epson BASE.xlsx".
'3. Um loop `For` é usado para abrir os arquivos no SharePoint e salvar as pastas de trabalho correspondentes em suas respectivas variáveis.
'4. Um loop `For` é usado novamente para abrir a planilha "PSI Epson.xlsm" no SharePoint e salvar a pasta de trabalho em uma variável chamada `PSI`.
'5. Os objetos de planilha relevantes são definidos para facilitar o acesso aos dados de cada planilha.
'6. Em seguida, os dados são copiados das planilhas "FUP", "MB51" e "ZSTOK" para as respectivas planilhas "BASE FUP", "BASE MB51" e "BASE ZSTOK" na pasta de trabalho "PSI Epson.xlsm".
'7. Depois de copiar os dados, é aplicada a função `TextToColumns` para processar algumas colunas específicas em cada planilha.
'8. A pasta de trabalho "PSI Epson.xlsm" é atualizada usando o método `RefreshAll` para garantir que os dados consolidados estejam atualizados.
'9. Finalmente, todas as pastas de trabalho abertas são fechadas, salvando as alterações feitas.
'
'2. Sub BrGroup():
'Esta sub-rotina realiza a consolidação de dados de várias planilhas, incluindo "ZSTOK BRGroup BASE", "MB51 BRGroup BASE" e "FUP BRGroup BASE", bem como a planilha "PSI BR GROUP".
'
'Passo a passo:
'1. Declaração de várias variáveis para armazenar as pastas de trabalho e planilhas que serão usadas na consolidação.
'2. As pastas de trabalho "ZSTOK BRGroup BASE", "MB51 BRGroup BASE" e "FUP BRGroup BASE" são abertas do local especificado na área de trabalho (desktop).
'3. A planilha "PSI BR GROUP" é aberta do local especificado na área de trabalho (desktop).
'4. Os objetos de planilha relevantes são definidos para facilitar o acesso aos dados de cada planilha.
'5. Os dados são copiados das planilhas "FUP", "MB51" e "ZSTOK" para as respectivas planilhas "BASE FUP", "BASE MB51" e "BASE ZSTOK" na pasta de trabalho "PSI BR GROUP".
'6. É aplicada a função `TextToColumns` para processar algumas colunas específicas em cada planilha.
'7. A pasta de trabalho "PSI BR GROUP" é atualizada usando o método `RefreshAll` para garantir que os dados consolidados estejam atualizados.
'8. Finalmente, todas as pastas de trabalho abertas são fechadas, salvando as alterações feitas.

Sub Epson()

Dim PSI As Workbook
Dim ZSTOK As Workbook
Dim MB51 As Workbook
Dim FUP As Workbook
Dim ZSTOKBase As Workbook
Dim MB51Base As Workbook
Dim FUPBase As Workbook
Dim wsFUP As Worksheet
Dim wsZSTOK As Worksheet
Dim wsPsiFUP As Worksheet
Dim wsPsiMB51 As Worksheet
Dim wsPsiZSTOK As Worksheet
Dim filePaths As Variant
Dim i As Long
    
'Abrir Relatórios
Application.DisplayAlerts = False


filePaths = Array("ZSTOK Epson BASE.xlsx", "MB51 Epson BASE.xlsx", "FUP Epson BASE.xlsx")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "https://konicaminoltaglobal.sharepoint.com/teams/BBRPSI/Shared%20Documents/MACRO/RELATORIOS/" & filePaths(i)
Next i

filePaths = Array("PSI Epson.xlsm")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "https://konicaminoltaglobal.sharepoint.com/teams/BBRPSI/Shared%20Documents/MACRO/PSI/" & filePaths(i)
Next i

Set wbFUP = Workbooks("FUP Epson BASE")
Set wbMB51 = Workbooks("MB51 Epson BASE")
Set wbZSTOK = Workbooks("ZSTOK Epson BASE")
Set PSI = Workbooks("PSI Epson")

Set wsFUP = wbFUP.Sheets("Sheet1")
Set wsMB51 = wbMB51.Sheets("Sheet1")
Set wsZSTOK = wbZSTOK.Sheets("Sheet1")

Set wsPsiFUP = Workbooks("PSI Epson").Sheets("BASE FUP")
Set wsPsiMB51 = Workbooks("PSI Epson").Sheets("BASE MB51")
Set wsPsiZSTOK = Workbooks("PSI Epson").Sheets("BASE ZSTOK")

Application.DisplayAlerts = True


With Sheets("PSI")
If .FilterMode Then .ShowAllData

'ajustando Data
 
 .Range("c1") = .Range("d1").Value

End With


'Limpando bases no PSI

With PSI.Sheets("BASE ZSTOK")
If .FilterMode Then .ShowAllDat
.Range("A1").CurrentRegion.ClearContents
End With

With PSI.Sheets("BASE MB51")
If .FilterMode Then .ShowAllDat

.Range("A1").CurrentRegion.ClearContents
End With


With PSI.Sheets("BASE FUP")
If .FilterMode Then .ShowAllDat
.Range("A1").CurrentRegion.ClearContents
End With

 ' COPIAR E COLAR FUP
 
With wsFUP
    .Range("A1").CurrentRegion.Copy
End With

With wsPsiFUP

    .Range("A1").PasteSpecial xlPasteValues
    
    With .Range("O2")
        .Resize(.End(xlDown).Row - .Row + 1).TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, 1)), _
            TrailingMinusNumbers:=True
    End With
    
End With
    
    
 ' COPIAR E COLAR MB51
 
With wsMB51
    .Range("A1").CurrentRegion.Copy
End With

With wsPsiMB51

    .Range("A1").PasteSpecial xlPasteValues
        
        With .Range("A2")
        .Resize(.End(xlDown).Row - .Row + 1).TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, 1)), _
            TrailingMinusNumbers:=True
    End With
    


End With
    
   
 ' COPIAR E COLAR ZSTOK

With wsZSTOK
    .Range("A1").CurrentRegion.Copy
End With

With PSI
    Set wsPsi = .Worksheets("BASE ZSTOK")
    wsPsiZSTOK.Range("A1").PasteSpecial xlPasteValues
    
 Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("O2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
        
End With
    
ActiveWorkbook.RefreshAll

 'FecharTodasAsPlanilhas
Application.DisplayAlerts = False
 With ZSTOK
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With MB51
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With FUP
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With PSI
    ActiveWindow.Close saveChanges:=True
End With


End Sub

Sub BrGroup()

Dim PSI As Workbook
Dim ZSTOK As Workbook
Dim MB51 As Workbook
Dim FUP As Workbook
Dim ZSTOKBase As Workbook
Dim MB51Base As Workbook
Dim FUPBase As Workbook
Dim wsFUP As Worksheet
Dim wsZSTOK As Worksheet
Dim wsPsiFUP As Worksheet
Dim wsPsiMB51 As Worksheet
Dim wsPsiZSTOK As Worksheet
    
'Abrir Relatórios
Application.DisplayAlerts = False

Set ZSTOK = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\ZSTOK BRGroup BASE")
Set MB51 = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\MB51 BRGroup BASE")
Set FUP = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\FUP BRGroup BASE")
Set PSI = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\PSI\PSI BR GROUP")

Set wbFUP = Workbooks("FUP BRGroup BASE")
Set wbMB51 = Workbooks("MB51 BRGroup BASE")
Set wbZSTOK = Workbooks("ZSTOK BRGroup BASE")

Set wsFUP = wbFUP.Sheets("Sheet1")
Set wsMB51 = wbMB51.Sheets("Sheet1")
Set wsZSTOK = wbZSTOK.Sheets("Sheet1")

Set wsPsiFUP = Workbooks("PSI BR GROUP").Sheets("BASE FUP")
Set wsPsiMB51 = Workbooks("PSI BR GROUP").Sheets("BASE MB51")
Set wsPsiZSTOK = Workbooks("PSI BR GROUP").Sheets("BASE ZSTOK")

Application.DisplayAlerts = True


With Sheets("PSI")
If .FilterMode Then .ShowAllData

'ajustando data
.Range("f1").Copy Psiination:=.Range("e1")
.Range("e1").TextToColumns Psiination:=.Range("e1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True

End With

'Limpando bases no PSI

Set PSI = Workbooks("PSI BR GROUP")

With PSI.Sheets("BASE ZSTOK")
If .FilterMode Then .ShowAllData
.Range("A1").CurrentRegion.ClearContents
End With

With PSI.Sheets("BASE MB51")
If .FilterMode Then .ShowAllData

.Range("A1").CurrentRegion.ClearContents
End With


With PSI.Sheets("BASE FUP")
If .FilterMode Then .ShowAllData
.Range("A1").CurrentRegion.ClearContents
End With


 ' COPIAR E COLAR FUP
 
With wsFUP
    .Range("A1").CurrentRegion.Copy
End With

With wsPsiFUP
.Range("A1").PasteSpecial xlPasteValues
End With

' transformar em numero
With Range("B2")
        .Resize(.End(xlDown).Row - .Row + 1).TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, 1)), _
            TrailingMinusNumbers:=True
    End With
    
 ' COPIAR E COLAR MB51
 
With wsMB51
    .Range("A1").CurrentRegion.Copy
End With

With wsPsiMB51
    .Range("A1").PasteSpecial xlPasteValues
    
    With .Range("A2")
        .Resize(.End(xlDown).Row - .Row + 1).TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, 1)), _
            TrailingMinusNumbers:=True
    End With
End With
   
 ' COPIAR E COLAR ZSTOK

With wsZSTOK
    .Range("A1").CurrentRegion.Copy
End With

With wsPsiZSTOK

    .Range("A1").PasteSpecial xlPasteValues
    
 Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("B2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
        
End With
    
ActiveWorkbook.RefreshAll

 'FecharTodasAsPlanilhas
Application.DisplayAlerts = False
 With ZSTOK
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With MB51
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With FUP
    ActiveWindow.Close saveChanges:=True
End With
Application.DisplayAlerts = False
 With PSI
    ActiveWindow.Close saveChanges:=True
End With

    
End Sub
'.- -.. --- .-.. .--. .... ---
