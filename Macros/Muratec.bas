Attribute VB_Name = "Muratec"
'1. `Application.DisplayAlerts = False`: Esta linha desativa os alertas do Excel, permitindo que os arquivos sejam abertos sem exibir mensagens de aviso ao usuário.
'
'2. `Dim filePaths As Variant`: Declaração da variável `filePaths`, que será usada para armazenar os nomes dos arquivos que serão abertos.
'
'3. `filePaths = Array("GERAL.xlsx")`: Define o nome do arquivo "GERAL.xlsx" no array `filePaths`.
'':):
'4. Um loop `For` é usado para abrir os arquivos listados em `filePaths`. Neste caso, ele abre "GERAL.xlsx" do local especificado no SharePoint.
'
'5. Em seguida, a variável `filePaths` é atualizada para conter o nome do arquivo "PSI Muratec.xlsm".
'
'6. O loop `For` é executado novamente para abrir o arquivo "PSI Muratec.xlsm" do local especificado no SharePoint.
'
'7. As variáveis `wbGERAL` e `wbPSI_A` são usadas para armazenar os objetos referentes às pastas de trabalho abertas "GERAL.xlsx" e "PSI Muratec.xlsm", respectivamente.
'
'8. Os objetos `wsGERAL` e `wsSummary_A` são usados para referenciar as planilhas "Sheet1" e "Summary" na pasta de trabalho "PSI Muratec.xlsm", respectivamente.
'
'9. A planilha "Summary" é usada para ajustar a data nas células B2 e B3.
'
'10. A planilha "GERAL" é usada para copiar informações da coluna C para a coluna D a partir da linha 3 e colar os dados na planilha "GERAL" na coluna A a partir da linha 3. Em seguida, são aplicadas fórmulas nas colunas C3 a O3 para buscar dados da planilha "GERAL.XLSX" usando a função `INDEX` e `MATCH`.
'
'11. Também são aplicadas fórmulas nas colunas P3 a AA3 para calcular os dados subtraindo valores específicos das colunas P a AA usando a função `SUMIFS`.
'
'12. A coluna V a Z é limpa para remover dados não utilizados.
'
'13. As planilhas "GERAL" e "PSI Muratec" são fechadas com `ActiveWindow.Close SaveChanges:=True` para salvar as alterações feitas nas planilhas antes de serem fechadas.
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
    
    
    'Abrir Relatórios
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
        
        'trazer informações da base
        
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

