Attribute VB_Name = "Muratec"
'1. `Application.DisplayAlerts = False`: Esta linha desativa os alertas do Excel, permitindo que os arquivos sejam abertos sem exibir mensagens de aviso ao usu�rio.
'
'2. `Dim filePaths As Variant`: Declara��o da vari�vel `filePaths`, que ser� usada para armazenar os nomes dos arquivos que ser�o abertos.
'
'3. `filePaths = Array("GERAL.xlsx")`: Define o nome do arquivo "GERAL.xlsx" no array `filePaths`.
'':):
'4. Um loop `For` � usado para abrir os arquivos listados em `filePaths`. Neste caso, ele abre "GERAL.xlsx" do local especificado no SharePoint.
'
'5. Em seguida, a vari�vel `filePaths` � atualizada para conter o nome do arquivo "PSI Muratec.xlsm".
'
'6. O loop `For` � executado novamente para abrir o arquivo "PSI Muratec.xlsm" do local especificado no SharePoint.
'
'7. As vari�veis `wbGERAL` e `wbPSI_A` s�o usadas para armazenar os objetos referentes �s pastas de trabalho abertas "GERAL.xlsx" e "PSI Muratec.xlsm", respectivamente.
'
'8. Os objetos `wsGERAL` e `wsSummary_A` s�o usados para referenciar as planilhas "Sheet1" e "Summary" na pasta de trabalho "PSI Muratec.xlsm", respectivamente.
'
'9. A planilha "Summary" � usada para ajustar a data nas c�lulas B2 e B3.
'
'10. A planilha "GERAL" � usada para copiar informa��es da coluna C para a coluna D a partir da linha 3 e colar os dados na planilha "GERAL" na coluna A a partir da linha 3. Em seguida, s�o aplicadas f�rmulas nas colunas C3 a O3 para buscar dados da planilha "GERAL.XLSX" usando a fun��o `INDEX` e `MATCH`.
'
'11. Tamb�m s�o aplicadas f�rmulas nas colunas P3 a AA3 para calcular os dados subtraindo valores espec�ficos das colunas P a AA usando a fun��o `SUMIFS`.
'
'12. A coluna V a Z � limpa para remover dados n�o utilizados.
'
'13. As planilhas "GERAL" e "PSI Muratec" s�o fechadas com `ActiveWindow.Close SaveChanges:=True` para salvar as altera��es feitas nas planilhas antes de serem fechadas.
'
'14. `Application.DisplayAlerts = True`: Reativa os alertas do Excel.
    
    Sub Muratec()
    
    Application.DisplayAlerts = False
    
    Dim filePaths As Variant
    filePaths = Array("GERAL.xlsx")
    
    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\" & filePaths(i)
    Next i
    
    
    filePaths = Array("PSI Muratec.xlsm")
    
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
    
    'Abrir Relat�rios
    Set wbGeral = Workbooks("GERAL")
    Set wbPSI_A = Workbooks("PSI Muratec")
    
    Set wsGERAL = wbGeral.Sheets("Sheet1")
    
    Set wsSummary_A = Workbooks("PSI Muratec").Sheets("Summary")
    Set wsGERAL_A = Workbooks("PSI Muratec").Sheets("GERAL")
    
    'PLAN A
    wbPSI_A.Activate
    
    'Ajustando Data
    With wsSummary_A
        .Range("B2") = .Range("C2").Value
        .Range("B3") = .Range("C3").Value
    End With
    
    
    With wsGERAL_A
        Dim mySelection As Range
        Set mySelection = .Range("A3:AJ3")
        Range(mySelection, .Cells(.Rows.Count, mySelection.Column).End(xlUp)).ClearContents 'limpar base
        
        'trazer informa��es da base
        
        lastRow = wsGERAL.Cells(.Rows.Count, "C").End(xlUp).Row
    wsGERAL.Range("C3:D" & lastRow).Copy Destination:=wsGERAL_A.Range("A3")
    
     .Range("C3:O" & lastRow).FormulaR1C1 = _
            "=INDEX([GERAL.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[GERAL.XLSX]Sheet1!C3,0),MATCH(R2C,[GERAL.XLSX]Sheet1!R1,0))"
            
    .Range("P3:AA" & lastRow).FormulaR1C1 = _
            "=INDEX([GERAL.XLSX]Sheet1!R1:R1048576,MATCH(RC1,[GERAL.XLSX]Sheet1!C3,0),MATCH(R2C,[GERAL.XLSX]Sheet1!R1,0))-SUMIFS(RC29:RC37,R1C29:R1C37,R1C)"
    
    .Range("V3:Z" & lastRow).ClearContents
    
    End With
    
     'FecharTodasAsPlanilhas
    
     With wbGeral
        ActiveWindow.Close saveChanges:=True
    End With
    
     With wbPSI_A
        ActiveWindow.Close saveChanges:=True
    End With
    
    Application.DisplayAlerts = True
    
    
End Sub
'.- -.. --- .-.. .--. .... ---

