Attribute VB_Name = "ConsumosTerceiro"
'1. Primeira parte: Abrir relatórios
'Nesta parte, o código abre dois arquivos de Excel localizados em "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\". Os nomes dos arquivos são "MAO.xlsx" e "VIX.xlsx". Ele utiliza um loop para abrir ambos os arquivos usando a função `Workbooks.Open`.
'
'2. Segunda parte: Definir variáveis
'Nesta parte, o código define várias variáveis que serão utilizadas posteriormente. Ele define variáveis para os objetos das planilhas e para os dois workbooks ("VIX" e "MAO"). As planilhas são obtidas a partir dos nomes dos objetos `wbTerceiros`, `wsBrgroup`, `wsBrother`, `wsEpson`, `wsVIX` e `wsMAO`. Os workbooks "VIX" e "MAO" são definidos apenas como strings contendo seus nomes.
'':):
'3. Terceira parte: Cópia de dados
'Nesta parte, o código copia os dados das planilhas "Sheet1" dos workbooks "VIX" e "MAO" para as planilhas "VIX" e "MAO" do workbook "Consumos Terceiros" respectivamente. Primeiro, ele ativa o workbook "Consumos Terceiros" (`wbTerceiros.Activate`) e, em seguida, copia os dados das planilhas "Sheet1" de "VIX" e "MAO" para as planilhas "VIX" e "MAO" de "Consumos Terceiros", respectivamente.
'
'4. Quarta parte: Fechar planilhas
'Após copiar os dados, o código fecha as três planilhas, mas antes disso, o workbook "Consumos Terceiros" é salvo, enquanto os workbooks "VIX" e "MAO" não são salvos.
'
'5. Quinta parte: Fechar todas as planilhas
'Finalmente, o código fecha os workbooks "VIX" e "MAO", sem salvá-los.
'

Sub ConsumosTerceiros()


Dim filePaths As Variant
Dim i As Long


'Abrir Relatórios

filePaths = Array("MAO.xlsx", "VIX.xlsx")
For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\" & filePaths(i)
Next i


'abrindo PSI
filePaths = Array("Consumos Terceiros")

For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
Next i

' Definir variavel

    'Psi
Set wbTerceiros = Workbooks("Consumos Terceiros.xlsx")
    
    Set wsBrgroup = wbTerceiros.Sheets("BRGROUP")
    Set wsBrother = wbTerceiros.Sheets("BROTHER")
    Set wsEpson = wbTerceiros.Sheets("EPSON")
    Set wsVIX = wbTerceiros.Sheets("VIX")
    Set wsMAO = wbTerceiros.Sheets("MAO")
    
    'relatorios

Set wbVIX = Workbooks("VIX.xlsx")
Set wbMAO = Workbooks("MAO.xlsx")

    Set wsVIXbase = wbVIX.Sheets("Sheet1")
    Set wsMAObase = wbMAO.Sheets("Sheet1")

wbTerceiros.Activate

With wsVIX
    If .FilterMode _
        Then .ShowAllData   'verificando filtro
    .Range("a1").CurrentRegion.ClearContents 'apagando base antiga
    lastRow = wsVIXbase.Cells(wsVIXbase.Rows.Count, "A").End(xlUp).Row 'conta linhas da planilha
    wsVIXbase.Range("A1:ca" & lastRow).Copy
    .Range("A1").PasteSpecial xlPasteAll 'colando novos dados
End With

With wsMAO
    If .FilterMode _
        Then .ShowAllData   'verificando filtro
    .Range("a1").CurrentRegion.ClearContents 'apagando base antiga
    lastRow = wsMAObase.Cells(wsMAObase.Rows.Count, "A").End(xlUp).Row 'conta linhas da planilha
    wsVIXbase.Range("A1:ca" & lastRow).Copy
    .Range("A1").PasteSpecial xlPasteAll 'colando novos dados
End With

'Fechar planilha
With wbTerceiros
    ActiveWindow.Close saveChanges:=True
End With

 'FecharTodasAsPlanilhas

    wbVIX.Close saveChanges:=True
    wbMAO.Close saveChanges:=True
'    wsVIXbase.Close saveChanges:=True
'    wsMAObase.Close saveChanges:=True
    
End Sub
'.- -.. --- .-.. .--. .... ---

