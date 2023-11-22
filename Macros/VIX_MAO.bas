Attribute VB_Name = "VIX_MAO"
'## Sub VIX():
'
'1. Definição do array `filePaths`: Nesta etapa, a macro cria um array chamado `filePaths` contendo os nomes de três arquivos: "VIX.xlsx", "Estoque" e "FUP". Esse array será utilizado para abrir os arquivos posteriormente.
'
'2. Abertura dos arquivos: A macro utiliza um loop `For` para percorrer cada elemento do array `filePaths`. Para cada elemento (nome de arquivo), ela abre o arquivo correspondente que está localizado na pasta "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\".
'
'3. Definição das variáveis: Em seguida, a macro define diversas variáveis que apontam para as planilhas dos arquivos que foram abertos. Por exemplo, `wbVIX`, `wbEstoque` e `wbFUP` são definidas para apontar para as planilhas "VIX.xlsx", "Estoque" e "FUP", respectivamente.
'':):
'4. Definição das variáveis para as planilhas "PSI": A macro também define variáveis para apontar para as planilhas "PSI_A.xlsm", "PSI_B.xlsm" e "PSI_C.xlsm", bem como as planilhas "Summary" e "VIX" dentro desses arquivos.
'
'5. Operações na planilha "PSI_A.xlsm": A macro ativa a planilha "PSI_A.xlsm" e realiza algumas operações específicas. Por exemplo, ela copia os dados da coluna "C" e "D" da planilha "VIX" (do arquivo "VIX.xlsx") e cola esses dados na planilha "VIX" da planilha "PSI_A.xlsm". Além disso, ela aplica fórmulas em algumas colunas usando a função `INDEX` e `MATCH` para fazer uma espécie de busca e correspondência de valores entre as planilhas.
'
'6. Operações na planilha "PSI_B.xlsm" e "PSI_C.xlsm": A macro repete as operações realizadas na planilha "PSI_A.xlsm", mas agora para as planilhas "PSI_B.xlsm" e "PSI_C.xlsm". Ela copia e cola os dados da planilha "VIX" em cada uma dessas planilhas, e aplica as fórmulas de busca e correspondência específicas para cada uma delas.
'
'7. Fechamento das planilhas: Ao final, a macro fecha todas as planilhas "PSI_A.xlsm", "PSI_B.xlsm" e "PSI_C.xlsm", mas mantém abertas as planilhas "VIX.xlsx", "Estoque" e "FUP".
'
'## Sub MAO():
'
'1. Definição do array `filePaths`: Nesta macro, é realizado o mesmo processo da macro anterior. Um array chamado `filePaths` é criado contendo os nomes de três arquivos: "MAO.xlsx", "Estoque" e "FUP".
'
'2. Abertura dos arquivos: A macro utiliza um loop `For` para percorrer cada elemento do array `filePaths`. Para cada elemento (nome de arquivo), ela abre o arquivo correspondente que está localizado na pasta "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\".
'
'3. Definição das variáveis: Em seguida, a macro define diversas variáveis que apontam para as planilhas dos arquivos que foram abertos. Por exemplo, `wbMAO`, `wbEstoque` e `wbFUP` são definidas para apontar para as planilhas "MAO.xlsx", "Estoque" e "FUP", respectivamente.
'
'4. Definição das variáveis para as planilhas "PSI MAO": A macro também define variáveis para apontar para as planilhas "PSI MAO A.xlsm" e "PSI MAO B.xlsm", bem como as planilhas "Summary" e "MAO" dentro desses arquivos.
'
'5. Operações na planilha "PSI MAO A.xlsm": A macro ativa a planilha "PSI MAO A.xlsm" e realiza operações específicas. Ela copia os dados das colunas "C" e "D" da planilha "MAO" e cola esses dados na planilha "MAO" da planilha "PSI MAO A.xlsm". Além disso, ela aplica fórmulas em algumas colunas usando a função `INDEX` e `MATCH` para fazer uma busca e correspondência de valores entre as planilhas.
'
'6. Operações na planilha "PSI MAO B.xlsm": A macro repete as operações realizadas na planilha "PSI MAO A.xlsm", mas agora para a planilha "PSI MAO B.xlsm".
''.- -.. --- .-.. .--. .... ---
'7. Fechamento das planilhas: Ao final, a macro fecha todas as planilhas "PSI MAO A.xlsm" e "PSI MAO B.xlsm", bem como as planilhas "MAO.xlsx", "Estoque" e "FUP".

Sub VIX()


Application.DisplayAlerts = True

'Declarando variaveis
Dim wbPSI_A As Workbook
Dim wbPSI_B As Workbook
Dim wbPSI_C As Workbook
Dim wbVIX As Workbook
Dim wbEstoque As Workbook
Dim wbFUP As Workbook
Dim wsSummary_A As Worksheet
Dim wsVIX_A As Worksheet
Dim wsSummary_B As Worksheet
Dim wsVIX_B As Worksheet
Dim wsSummary_C As Worksheet
Dim wsVIX_C As Worksheet
Dim wsVIX As Worksheet
Dim wsEstoque As Worksheet
Dim wsFUP As Worksheet
Dim filePaths As Variant
Dim i As Long
  'Abrir Relatórios
filePaths = Array("VIX.xlsx", "Estoque", "FUP")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\" & filePaths(i)
Next i

    'Abrir PSI
filePaths = Array("PSI_A.xlsm", "PSI_B.xlsm", "PSI_C.xlsm")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
Next i

'Definir objeto
    'PSI
Set wbPSI_A = Workbooks("PSI_A.xlsm")
Set wbPSI_B = Workbooks("PSI_B.xlsm")
Set wbPSI_C = Workbooks("PSI_C.xlsm")

    Set wsSummary_A = Workbooks("PSI_A.xlsm").Sheets("Summary")
    Set wsVIX_A = Workbooks("PSI_A.xlsm").Sheets("VIX")

    Set wsSummary_B = Workbooks("PSI_B.xlsm").Sheets("Summary")
    Set wsVIX_B = Workbooks("PSI_B.xlsm").Sheets("VIX")

    Set wsSummary_C = Workbooks("PSI_C.xlsm").Sheets("Summary")
    Set wsVIX_C = Workbooks("PSI_C.xlsm").Sheets("VIX")

    Set wsFUP_A = wbPSI_A.Sheets("Base FUP")
    Set wsEstoque_A = wbPSI_A.Sheets("Estoque Hist")

    Set wsFUP_B = wbPSI_B.Sheets("Base FUP")
    Set wsEstoque_B = wbPSI_B.Sheets("Estoque Hist")

    Set wsFUP_C = wbPSI_C.Sheets("Base FUP")
    Set wsEstoque_C = wbPSI_C.Sheets("Estoque Hist")
    
    'Relatorios
Set wbVIX = Workbooks("VIX.xlsx")
Set wbEstoque = Workbooks("Estoque.xlsx")
Set wbFUP = Workbooks("FUP.xlsx")

    Set wsVIX = wbVIX.Sheets("Sheet1")
    Set wsEstoque = wbEstoque.Sheets("Estoque")
    Set wsFUP = wbFUP.Sheets("Sheet1")
    
'PLAN A
wbPSI_A.Activate

'Ajustando Data
With wsSummary_A
    .Range("B2") = .Range("C2").Value
    .Range("B3") = .Range("C3").Value
End With

With wsVIX
    .Range("c:c").TextToColumns _
        Destination:=Range("c1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True
End With


'VIX
With wsVIX_A
    If .FilterMode _
        Then .ShowAllData
    Set mySelection = .Range("A3:Aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents 'limpar base
    Set mySelection = .Range("c3:aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents
    
    
' Trazer informações VIX
    lastRow = wsVIX.Cells(wsVIX.Rows.Count, "C").End(xlUp).Row
    wsVIX.Range("C3:D" & lastRow).Copy Destination:=wsVIX_A.Range("A3") ' Colar códigos
    
    .Range("C3:O" & lastRow).FormulaR1C1 = _
        "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))" ' Formula Índice-Corresp
        
    .Range("P3:AA" & lastRow).FormulaR1C1 = _
        "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))"
    
    .Range("V3:Z" & lastRow).ClearContents ' Limpar valores #ND
End With

'mc.9
With wsEstoque_A
    lastRow = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row 'conta linhas da planilha
    .Range("A4:m4").End(xlDown).ClearContents 'apaga base antiga
    wsEstoque.Range("A2:M" & lastRow).Copy 'copia a nova base
    .Range("A4:m" & lastRow).PasteSpecial xlPasteAll 'cola na planilha
End With

'Fup
With wsFUP_A
    .Range("A1").CurrentRegion.ClearContents 'apaga base antiga
    wsFUP.Range("A1").CurrentRegion.Copy 'copia a nova base
    .Range("A1").PasteSpecial xlPasteAll 'colar
End With

'PLAN B
wbPSI_B.Activate

'Ajustando Data
With wsSummary_B
    .Range("B2") = .Range("C2").Value
    .Range("B3") = .Range("C3").Value
End With

With wsVIX_B
    If .FilterMode _
        Then .ShowAllData
    Set mySelection = .Range("A3:AJ3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents 'limpar base
    Set mySelection = .Range("c3:aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents
    
    'trazer informações da base
    
    With wsVIX
        Set mySelection = .Range("C2:D2")
        .Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlUp)).Copy
    End With
    
    .Range("a3").PasteSpecial xlPasteValues

    .Range("C3:O" & lastRow).FormulaR1C1 = _
    "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))"
        
    .Range("P3:AA" & lastRow).FormulaR1C1 = _
    "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))"

    .Range("V3:Z" & lastRow).ClearContents
End With

'mc.9
With wsEstoque_B
    lastRow = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row 'conta linhas da planilha
    .Range("A4:m4").End(xlDown).ClearContents 'apaga base antiga
    wsEstoque.Range("A2:M" & lastRow).Copy 'copia a nova base
    .Range("A4:m" & lastRow).PasteSpecial xlPasteAll 'cola na planilha
End With

'Fup
With wsFUP_B
    .Range("A1").CurrentRegion.ClearContents 'apaga base antiga
    wsFUP.Range("A1").CurrentRegion.Copy 'copia a nova base
    .Range("A1").PasteSpecial xlPasteAll 'colar
End With


With wsVIX_C
    If .FilterMode _
        Then .ShowAllData
    Set mySelection = .Range("A3:AJ3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents 'limpar base
    Set mySelection = .Range("c3:aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents
    
    'trazer informações da base
    
    With wsVIX
        Set mySelection = .Range("C2:D2")
        .Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlUp)).Copy
    End With
    
    .Range("a3").PasteSpecial xlPasteValues

    .Range("C3:O" & lastRow).FormulaR1C1 = _
          "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))"
        
    .Range("P3:AA" & lastRow).FormulaR1C1 = _
         "=INDEX([VIX.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[VIX.XLSX]Sheet1!C3,0),MATCH(R2C,[VIX.XLSX]Sheet1!R1,0))"

    .Range("V3:Z" & lastRow).ClearContents
End With

'mc.9
With wsEstoque_C
    lastRow = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row 'conta linhas da planilha
    .Range("A4:m4").End(xlDown).ClearContents 'apaga base antiga
    wsEstoque.Range("A2:M" & lastRow).Copy 'copia a nova base
    .Range("A4:m" & lastRow).PasteSpecial xlPasteAll 'cola na planilha
End With

'Fup
With wsFUP_C
    .Range("A1").CurrentRegion.ClearContents 'apaga base antiga
    wsFUP.Range("A1").CurrentRegion.Copy 'copia a nova base
    .Range("A1").PasteSpecial xlPasteAll 'colar
End With

 'FecharTodasAsPlanilhas

wbPSI_A.Close saveChanges:=True
wbPSI_B.Close saveChanges:=True
wbPSI_C.Close saveChanges:=True

wbVIX.Close saveChanges:=True
wbEstoque.Close saveChanges:=False
wbFUP.Close saveChanges:=False

'.- -.. --- .-.. .--. .... ---

End Sub

Sub MAO()
Application.DisplayAlerts = False

'Abrir Relatórios
Dim wbPSI_A As Workbook
Dim wbPSI_B As Workbook
Dim wbMAO As Workbook
Dim wbEstoque As Workbook
Dim wbFUP As Workbook
Dim wsMAO As Worksheet
Dim wsEstoque As Worksheet
Dim wsFUP As Worksheet
Dim wsSummary_A As Worksheet
Dim wsMAO_A As Worksheet
Dim wsSummary_B As Worksheet
Dim wsMAO_B As Worksheet
Dim filePaths As Variant
Dim i As Long
Dim mySelection As Range

filePaths = Array("MAO.xlsx", "Estoque MAO.xlsx")
For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\" & filePaths(i)
Next i

filePaths = Array("PSI MAO A.xlsm", "PSI MAO B.xlsm")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
Next i

'Definir variaveis

    'Relatório
Set wbMAO = Workbooks("MAO.xlsx")
Set wbEstoque = Workbooks("Estoque MAO.xlsx")
'Set wbFUP = Workbooks("FUP")

    Set wsMAO = wbMAO.Sheets("Sheet1")
    Set wsEstoque = wbEstoque.Sheets("MAO MC.9")
'    Set wsFUP = wbFUP.Sheets("Sheet1")
    
    'PSI
Set wbPSI_A = Workbooks("PSI MAO A.xlsm")
Set wbPSI_B = Workbooks("PSI MAO B.xlsm")

    Set wsSummary_A = wbPSI_A.Sheets("Summary")
    Set wsMAO_A = wbPSI_A.Sheets("MAO")

    Set wsSummary_B = wbPSI_B.Sheets("Summary")
    Set wsMAO_B = wbPSI_B.Sheets("MAO")
'PLAN A
wbPSI_A.Activate

'Ajustando Data
With wsSummary_A
    .Range("B2") = .Range("C2").Value
    .Range("B3") = .Range("C3").Value
End With

With wsMAO
    .Range("c:c").TextToColumns _
        Destination:=Range("c1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=True, _
        Tab:=True
End With

With wsMAO_A
    If .FilterMode _
        Then .ShowAllData
        Set mySelection = .Range("A3:AJ3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents 'limpar base
    Set mySelection = .Range("c3:aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents
    
    
    'trazer informações da base
    
    lastRow = wsMAO.Cells(.Rows.Count, "C").End(xlUp).Row
wsMAO.Range("C3:D" & lastRow).Copy Destination:=wsMAO_A.Range("A3")

 .Range("C3:O" & lastRow).FormulaR1C1 = _
        "=INDEX([MAO.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[MAO.XLSX]Sheet1!C3,0),MATCH(R2C,[MAO.XLSX]Sheet1!R1,0))"
        
.Range("P3:AA" & lastRow).FormulaR1C1 = _
        "=INDEX([MAO.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[MAO.XLSX]Sheet1!C3,0),MATCH(R2C,[MAO.XLSX]Sheet1!R1,0))-SUMIFS(RC29:RC37,R1C29:R1C37,R1C)"

.Range("V3:Z" & lastRow).ClearContents

End With


'PLAN B
wbPSI_B.Activate

'Ajustando Data
With wsSummary_B
    .Range("B2") = .Range("C2").Value
    .Range("B3") = .Range("C3").Value
End With


With wsMAO_B
    If .FilterMode _
        Then .ShowAllData
    Set mySelection = .Range("A3:AJ3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents 'limpar base
    Set mySelection = .Range("c3:aa3")
    Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlEnd)).ClearContents
    
    'trazer informações da base
    
    lastRow = wsMAO.Cells(.Rows.Count, "C").End(xlUp).Row
wsMAO.Range("C3:D" & lastRow).Copy Destination:=wsMAO_A.Range("A3")

 .Range("C3:O" & lastRow).FormulaR1C1 = _
        "=INDEX([MAO.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[MAO.XLSX]Sheet1!C3,0),MATCH(R2C,[MAO.XLSX]Sheet1!R1,0))"
        
.Range("P3:AA" & lastRow).FormulaR1C1 = _
        "=INDEX([MAO.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[MAO.XLSX]Sheet1!C3,0),MATCH(R2C,[MAO.XLSX]Sheet1!R1,0))-SUMIFS(RC29:RC37,R1C29:R1C37,R1C)"

.Range("V3:Z" & lastRow).ClearContents

End With


 'FecharTodasAsPlanilha

wbPSI_A.Close saveChanges:=True
wbPSI_B.Close saveChanges:=True

wbMAO.Close saveChanges:=True
wbEstoque.Close saveChanges:=False
'wbFUP.Close saveChanges:=False


Application.DisplayAlerts = True

End Sub
'.- -.. --- .-.. .--. .... ---

