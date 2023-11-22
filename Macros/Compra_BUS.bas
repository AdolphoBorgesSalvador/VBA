Attribute VB_Name = "Compra_BUS"
'1. Declara��o de vari�veis:
'   - O c�digo come�a declarando v�rias vari�veis para armazenar informa��es, como pastas de trabalho (workbooks), planilhas (worksheets), valores de c�lulas, caminhos de arquivos, etc.
'2. Solicitar escolha de diret�rio:
'   - O usu�rio � solicitado a inserir "X" ou "Z" em uma caixa de di�logo de entrada (InputBox). O valor inserido � armazenado na vari�vel `folderChoice`.
'3. Verifica��o da escolha do usu�rio:
'   - O c�digo verifica o valor inserido pelo usu�rio e, com base nessa escolha, constr�i o caminho do arquivo Excel a ser aberto. Se o usu�rio cancelar a entrada ou deixar em branco, o c�digo sai (Exit Sub). Se o usu�rio inserir algo diferente de "X" ou "Z", exibir� uma mensagem de erro e tamb�m sair�.
'4. Abrir arquivo Excel:
'   - O c�digo usa a vari�vel `filePath` para abrir o arquivo Excel especificado.
'5. Definir planilhas:
'   - Duas planilhas do arquivo Excel aberto s�o atribu�das �s vari�veis `MacroWs` e `BaseWs`.
'6. Copiar dados:
'   - Os dados da coluna A at� a coluna E na planilha `BaseWs` s�o copiados para a planilha `MacroWs`.
'7. Aplicar f�rmulas:
'   - O c�digo utiliza um loop (For) para aplicar f�rmulas �s colunas G, H e I na planilha `MacroWs`. Essas f�rmulas s�o condicionais e dependem dos valores na coluna A.
'8. Substituir caracteres:
'   - O c�digo encontra o �ltimo valor na coluna D na planilha `MacroWs` e substitui todos os pontos (.) por v�rgulas (,) nas colunas D da planilha `MacroWs` e C da planilha `BaseWs`.
'9. Filtrar e copiar dados:
'   - � aplicado um filtro na coluna G da planilha `MacroWs` para copiar apenas as linhas onde o valor na coluna G n�o seja vazio. Os dados filtrados s�o copiados para uma nova planilha chamada "Arquivo".
'10. Dividir texto em colunas:
'    - O c�digo percorre cada coluna na planilha "Arquivo" e divide o texto em colunas com base em um delimitador de tabula��o.
'
'11. Limpar filtros:
'    - Os filtros aplicados na planilha `MacroWs` s�o removidos.
'
'.- -.. --- .-.. .--. .... ---

Sub CompraBus()

    Dim wbBUS As Workbook
    Dim MacroWs As Worksheet
    Dim BaseWs As Worksheet
    Dim ArquivoWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim folderChoice As String ' Vari�vel para armazenar a escolha do usu�rio
    Dim filePath As String
    ' Solicitar ao usu�rio que selecione O diret�rio
    folderChoice = InputBox("Escolha o diret�rio (Digite D, X, Y ou Z):", "Escolha de Diret�rio")
    
    ' Verificar a escolha do usu�rio
    If folderChoice = "X" Then
        ' Diret�rio escolhido � "X"
        filePath = "\\10.230.32.8\IMPORTA��O\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "Z" Then
        ' Diret�rio escolhido � "Z"
        filePath = "Z:\IMPORTA��O\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "D" Then
        ' Diret�rio escolhido � "Z"
        filePath = "D:\IMPORTA��O\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "Y" Then
        ' Diret�rio escolhido � "Y"
        filePath = "Y:\IMPORTA��O\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "" Then
        ' O usu�rio clicou em Cancelar ou deixou em branco
        Exit Sub
    Else
        ' A escolha do usu�rio n�o � "X" nem "Z"
        MsgBox "Escolha inv�lida. Use X ou Z.", vbExclamation, "Escolha Inv�lida"
        Exit Sub
    End If

    ' Abrir o arquivo
    Set wbBUS = Workbooks.Open(filePath)
    
    ' Set worksheets
    Set MacroWs = wbBUS.Sheets("Macro")
    Set BaseWs = wbBUS.Sheets("Base")
    
    ' Copy data from BaseWs to MacroWs
    BaseWs.Range("A:E").Copy Destination:=MacroWs.Range("A1")
    
    ' Find the number of rows in column A
    Dim numLinhas As Long
    numLinhas = MacroWs.Cells(MacroWs.Rows.Count, "A").End(xlUp).Row - 6
    
    ' Apply formulas to rows starting from column G, H, and I
    For i = 2 To numLinhas
        MacroWs.Cells(i, "G").Formula = "=IF(LEFT(A" & (i + 1) & ", 12)=""Material No:"", RIGHT(A" & (i + 1) & ", LEN(A" & (i + 1) & ")-12), """")"
        MacroWs.Cells(i, "H").Formula = "=IF(LEFT(A" & (i + 1) & ", 12)=""Material No:"", LEFT(C" & i & ", LEN(C" & i & ")-3), """")"
        MacroWs.Cells(i, "I").Formula = "=IF(LEFT(A" & (i + 1) & ", 12)=""Material No:"", RIGHT(D" & i & ", LEN(D" & i & ")-1), """")"
    Next i
    
    ' Find the last row in column D
    lastRow = MacroWs.Cells(MacroWs.Rows.Count, "D").End(xlUp).Row
    
    ' Replace dots with commas in column D of MacroWs and BaseWs
    MacroWs.Range("D1:D" & lastRow).Replace What:=".", Replacement:=",", LookAt:=xlPart
    BaseWs.Range("C1:C" & lastRow).Replace What:=".", Replacement:=",", LookAt:=xlPart
    
    ' Filter and copy data from MacroWs to ArquivoWs
    MacroWs.Range("G1:I1").AutoFilter
    MacroWs.Range("G1:I1").AutoFilter Field:=1, Criteria1:="<>"
    Sheets.Add.Name = "Arquivo"
    Set ArquivoWs = wbBUS.Sheets("Arquivo")
    
    MacroWs.Range("G:I").Copy Destination:=ArquivoWs.Range("A1:C1")
    
    ' Split text to columns in ArquivoWs
    For Each col In ArquivoWs.Range("A:C").Columns
        col.TextToColumns Destination:=col.Cells(1, 1), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    Next col
    
    ' Clear filters in MacroWs
    MacroWs.AutoFilterMode = False
    
End Sub

'.- -.. --- .-.. .--. .... ---
