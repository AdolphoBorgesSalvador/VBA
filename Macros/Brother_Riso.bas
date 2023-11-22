Attribute VB_Name = "Brother_Riso"
'**1. Sub Riso():**
'Essa sub-rotina realiza a consolidação de dados da planilha "GERAL.xlsx" e da planilha "PSI RISO.xlsm".
'
'Passo a passo:
'1. `Dim filePaths As Variant`: Declaração da variável `filePaths`, que será usada para armazenar os nomes dos arquivos que serão abertos.
'2. `filePaths = Array("GERAL.xlsx")`: Define o nome do arquivo "GERAL.xlsx" no array `filePaths`.
'3. Um loop `For` é usado para abrir o arquivo "GERAL.xlsx" do local especificado na pasta "MACRO\RELATORIOS".
'4. A variável `filePaths` é atualizada para conter o nome do arquivo "PSI RISO.xlsm".
'5. O loop `For` é executado novamente para abrir o arquivo "PSI RISO.xlsm" do local especificado na pasta "PSI".
'6. Os objetos de pasta de trabalho `wbGERAL` e `wbPSI` são criados para armazenar as pastas de trabalho "GERAL.xlsx" e "PSI RISO.xlsm", respectivamente.
'7. Os objetos de planilha `geralWs` e `psiGeralWs` são criados para referenciar as planilhas "Sheet1" em "GERAL.xlsx" e "GERAL" em "PSI RISO.xlsm", respectivamente.
'8. A planilha "Sheet1" é usada para copiar dados da coluna C e D para a planilha "PSI RISO" nas colunas A3 e C3, respectivamente.
'9. Fórmulas são aplicadas nas colunas AI2 a AT3 e AW2 a AW3 para buscar dados da planilha "GERAL.xlsx" usando a função `INDEX` e `MATCH` e copiá-los para a planilha "PSI RISO".
'10. A coluna BV2 a CA3 é copiada para a planilha "PSI RISO".
'11. As planilhas "GERAL.xlsx" e "PSI RISO.xlsm" são fechadas, salvando as alterações feitas.
'':):
'**2. Sub Brother():**
'Essa sub-rotina realiza a consolidação de dados da planilha "GERAL.xlsx" e da planilha "PSI BROTHER.xlsm".
'
'Passo a passo:
'1. `Dim filePaths As Variant`: Declaração da variável `filePaths`, que será usada para armazenar os nomes dos arquivos que serão abertos.
'2. `filePaths = Array("GERAL.xlsx")`: Define o nome do arquivo "GERAL.xlsx" no array `filePaths`.
'3. Um loop `For` é usado para abrir o arquivo "GERAL.xlsx" do local especificado na pasta "MACRO\RELATORIOS".
'4. A variável `filePaths` é atualizada para conter o nome do arquivo "PSI BROTHER.xlsm".
'5. O loop `For` é executado novamente para abrir o arquivo "PSI BROTHER.xlsm" do local especificado na pasta "PSI".
'6. Os objetos de pasta de trabalho `wbGERAL` e `wbBrother` são criados para armazenar as pastas de trabalho "GERAL.xlsx" e "PSI BROTHER.xlsm", respectivamente.
'7. Os objetos de planilha `geralWs`, `wsSummary_A` e `wsGERAL_A` são criados para referenciar as planilhas "Sheet1", "Summary" e "GERAL" em "PSI BROTHER.xlsm", respectivamente.
'8. A planilha "Summary" é usada para ajustar a data nas células B2 e B3.
'9. A planilha "GERAL" é usada para copiar dados da coluna C e D para a planilha "PSI BROTHER" nas colunas A3 e C3, respectivamente.
'10. Fórmulas são aplicadas nas colunas C3 a O3 e P3 a AA3 para buscar dados da planilha "GERAL.xlsx" usando a função `INDEX` e `MATCH` e calcular valores com base nos dados copiados.
'11. A coluna V a Z é limpa para remover dados não utilizados.
'12. As planilhas "GERAL.xlsx" e "PSI BROTHER.xlsm" são fechadas, salvando as alterações feitas.


Sub Riso()

Dim filePaths As Variant
filePaths = Array("GERAL.xlsx")

Dim i As Long
For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\MACRO\RELATORIOS\" & filePaths(i)
Next i


filePaths = Array("PSI RISO.xlsm")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\MACRO\PSI\" & filePaths(i)
Next i

Set geralWb = Workbooks("GERAL")
Set psiRisoWb = Workbooks("PSI RISO")

Set geralWs = geralWb.Sheets("Sheet1")

Set psiGeralWs = psiRisoWb.Sheets("GERAL")

    
    'Geral
With geralWs
    If .FilterMode Then .ShowAllData
    .Range("C2:D" & .Range("D" & .Rows.Count).End(xlUp).Row).Copy _
        Destination:=psiGeralWs.Range("A3")
        
.Range("AI2:AT" & .Range("AT" & .Rows.Count).End(xlUp).Row).Copy _
        Destination:=psiGeralWs.Range("C3")
        
.Range("AW2:aw" & .Range("AW" & .Rows.Count).End(xlUp).Row).Copy _
        Destination:=psiGeralWs.Range("O3")
        
.Range("BV2:CA" & .Range("CA" & .Rows.Count).End(xlUp).Row).Copy _
        Destination:=psiGeralWs.Range("P3")
End With


    
     'FecharTodasAsPlanilhas
Application.DisplayAlerts = False

 With geralWb
    ActiveWindow.Close saveChanges:=True
End With

 With PSI
    ActiveWindow.Close saveChanges:=True
End With


Application.DisplayAlerts = True
'.- -.. --- .-.. .--. .... ---

End Sub

Sub Brother()

'Abrir relatórios
Dim i As Long
Dim filePaths As Variant
filePaths = Array("GERAL")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\" & filePaths(i)
Next i

filePaths = Array("PSI BROTHER.xlsm")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
Next i

Set wbGeral = Workbooks("GERAL.XLSX")
Set wsGERAL = wbGeral.Sheets("Sheet1")

Set wsSummary_A = Workbooks("PSI BROTHER.XLSM").Sheets("Summary")
Set wsGERAL_A = Workbooks("PSI BROTHER.XLSM").Sheets("GERAL")
Set wbBrother = Workbooks("PSI BROTHER.XLSM")
'PLAN A
wbBrother.Activate

'Ajustando Data
With wsSummary_A
    .Range("B2") = .Range("C2").Value
    .Range("B3") = .Range("C3").Value
End With

With wsGERAL
    .Range("c:c").TextToColumns _
        Destination:=Range("c1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True
End With

With wsGERAL_A
    If .FilterMode Then .ShowAllData
    Dim mySelection As Range
    Set mySelection = .Range("A3:AJ3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlUp)).ClearContents 'limpar base
    
    'trazer informações da base
    
    lastRow = wsGERAL.Cells(.Rows.Count, "C").End(xlUp).Row
    wsGERAL.Range("C3:D" & lastRow).Copy Destination:=wsGERAL_A.Range("A3")

    .Range("C3:O" & lastRow).FormulaR1C1 = _
        "=INDEX([GERAL.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[GERAL.XLSX]Sheet1!C3,0),MATCH(R2C,[GERAL.XLSX]Sheet1!R1,0))"
        
    .Range("P3:AA" & lastRow).FormulaR1C1 = _
        "=INDEX([GERAL.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[GERAL.XLSX]Sheet1!C3,0),MATCH(R2C,[GERAL.XLSX]Sheet1!R1,0))-SUMIFS(RC29:RC37,R1C29:R1C37,R1C)"

    .Range("V3:Z" & lastRow).ClearContents
End With

    wbGeral.Close saveChanges:=True
    wbBrother.Close saveChanges:=True

End Sub

'.- -.. --- .-.. .--. .... ---

