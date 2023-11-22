Attribute VB_Name = "ConsumosTerceiro"
'1. Primeira parte: Abrir relat�rios
'Nesta parte, o c�digo abre dois arquivos de Excel localizados em "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIOS\". Os nomes dos arquivos s�o "MAO.xlsx" e "VIX.xlsx". Ele utiliza um loop para abrir ambos os arquivos usando a fun��o `Workbooks.Open`.
'
'2. Segunda parte: Definir vari�veis
'Nesta parte, o c�digo define v�rias vari�veis que ser�o utilizadas posteriormente. Ele define vari�veis para os objetos das planilhas e para os dois workbooks ("VIX" e "MAO"). As planilhas s�o obtidas a partir dos nomes dos objetos `wbTerceiros`, `wsBrgroup`, `wsBrother`, `wsEpson`, `wsVIX` e `wsMAO`. Os workbooks "VIX" e "MAO" s�o definidos apenas como strings contendo seus nomes.
'':):
'3. Terceira parte: C�pia de dados
'Nesta parte, o c�digo copia os dados das planilhas "Sheet1" dos workbooks "VIX" e "MAO" para as planilhas "VIX" e "MAO" do workbook "Consumos Terceiros" respectivamente. Primeiro, ele ativa o workbook "Consumos Terceiros" (`wbTerceiros.Activate`) e, em seguida, copia os dados das planilhas "Sheet1" de "VIX" e "MAO" para as planilhas "VIX" e "MAO" de "Consumos Terceiros", respectivamente.
'
'4. Quarta parte: Fechar planilhas
'Ap�s copiar os dados, o c�digo fecha as tr�s planilhas, mas antes disso, o workbook "Consumos Terceiros" � salvo, enquanto os workbooks "VIX" e "MAO" n�o s�o salvos.
'
'5. Quinta parte: Fechar todas as planilhas
'Finalmente, o c�digo fecha os workbooks "VIX" e "MAO", sem salv�-los.
'

Sub ConsumosTerceiros()


Dim filePaths As Variant
Dim i As Long


'Abrir Relat�rios

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

