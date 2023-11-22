Attribute VB_Name = "Compra_BUS"
'1. Declaração de variáveis:
'   - O código começa declarando várias variáveis para armazenar informações, como pastas de trabalho (workbooks), planilhas (worksheets), valores de células, caminhos de arquivos, etc.
'2. Solicitar escolha de diretório:
'   - O usuário é solicitado a inserir "X" ou "Z" em uma caixa de diálogo de entrada (InputBox). O valor inserido é armazenado na variável `folderChoice`.
'3. Verificação da escolha do usuário:
'   - O código verifica o valor inserido pelo usuário e, com base nessa escolha, constrói o caminho do arquivo Excel a ser aberto. Se o usuário cancelar a entrada ou deixar em branco, o código sai (Exit Sub). Se o usuário inserir algo diferente de "X" ou "Z", exibirá uma mensagem de erro e também sairá.
'4. Abrir arquivo Excel:
'   - O código usa a variável `filePath` para abrir o arquivo Excel especificado.
'5. Definir planilhas:
'   - Duas planilhas do arquivo Excel aberto são atribuídas às variáveis `MacroWs` e `BaseWs`.
'6. Copiar dados:
'   - Os dados da coluna A até a coluna E na planilha `BaseWs` são copiados para a planilha `MacroWs`.
'7. Aplicar fórmulas:
'   - O código utiliza um loop (For) para aplicar fórmulas às colunas G, H e I na planilha `MacroWs`. Essas fórmulas são condicionais e dependem dos valores na coluna A.
'8. Substituir caracteres:
'   - O código encontra o último valor na coluna D na planilha `MacroWs` e substitui todos os pontos (.) por vírgulas (,) nas colunas D da planilha `MacroWs` e C da planilha `BaseWs`.
'9. Filtrar e copiar dados:
'   - É aplicado um filtro na coluna G da planilha `MacroWs` para copiar apenas as linhas onde o valor na coluna G não seja vazio. Os dados filtrados são copiados para uma nova planilha chamada "Arquivo".
'10. Dividir texto em colunas:
'    - O código percorre cada coluna na planilha "Arquivo" e divide o texto em colunas com base em um delimitador de tabulação.
'
'11. Limpar filtros:
'    - Os filtros aplicados na planilha `MacroWs` são removidos.
'
'.- -.. --- .-.. .--. .... ---

Sub CompraBus()

    Dim wbBUS As Workbook
    Dim MacroWs As Worksheet
    Dim BaseWs As Worksheet
    Dim ArquivoWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim folderChoice As String ' Variável para armazenar a escolha do usuário
    Dim filePath As String
    ' Solicitar ao usuário que selecione O diretório
    folderChoice = InputBox("Escolha o diretório (Digite D, X, Y ou Z):", "Escolha de Diretório")
    
    ' Verificar a escolha do usuário
    If folderChoice = "X" Then
        ' Diretório escolhido é "X"
        filePath = "\\10.230.32.8\IMPORTAÇÃO\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "Z" Then
        ' Diretório escolhido é "Z"
        filePath = "Z:\IMPORTAÇÃO\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "D" Then
        ' Diretório escolhido é "Z"
        filePath = "D:\IMPORTAÇÃO\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "Y" Then
        ' Diretório escolhido é "Y"
        filePath = "Y:\IMPORTAÇÃO\Pedidos\" & "Dados BUS.xlsm"
        
    ElseIf folderChoice = "" Then
        ' O usuário clicou em Cancelar ou deixou em branco
        Exit Sub
    Else
        ' A escolha do usuário não é "X" nem "Z"
        MsgBox "Escolha inválida. Use X ou Z.", vbExclamation, "Escolha Inválida"
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
