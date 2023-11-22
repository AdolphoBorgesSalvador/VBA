Attribute VB_Name = "Validação"
Sub Validação()

Dim wb As Workbook
Dim ws As Worksheet
Dim contadorVerdadeiros As Long
Dim contadorFalsos As Long
Dim totalVerdadeiros As Long
Dim totalFalsos As Long
Dim celula As Range

Dim CaminhoVix As String
    Dim CaminhoMao As String
    Dim wbPSI_A As Workbook
    Dim wbPSI_B As Workbook
    Dim wbPSI_A_MAO As Workbook
    Dim wbPSI_B_MAO As Workbook
    Dim filePaths As Variant
    Dim i As Long

    mes = Month(Date)
    mesExtenso = UCase(MonthName(mes))
    CaminhoVix = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\"
    CaminhoMao = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI MAO\"

'' Abrir PSI
' Set wb = Workbooks.Open(CaminhoVix & "PSI_A" & "_" & mesExtenso & ".xlsm")
'
'' Inicializar contadores globais
'totalVerdadeiros = 0
'totalFalsos = 0
'
'
'    ' Loop através de todas as planilhas na pasta de trabalho
'    For Each ws In wb.Sheets
'        ' Definir a célula para contar
'        Set celula = ws.Range("P:P")
'
'        ' Inicializar os contadores locais
'        contadorVerdadeiros = 0
'        contadorFalsos = 0
'
'        ' Loop através das células na coluna P
'        For Each cel In celula
'            If VarType(cel.Value) = vbBoolean Then
'                If cel.Value Then ' Verifica se a célula é verdadeira
'                    contadorVerdadeiros = contadorVerdadeiros + 1
'                Else ' Caso contrário, a célula é falsa
'                    contadorFalsos = contadorFalsos + 1
'                End If
'            End If
'        Next cel
'
'        ' Adicionar aos contadores globais
'        totalVerdadeiros = totalVerdadeiros + contadorVerdadeiros
'        totalFalsos = totalFalsos + contadorFalsos
'    Next ws
'
'' Exibir os resultados gerais
'MsgBox "Resultados Gerais:" & vbCrLf & _
'       "Total de 'Verdadeiros': " & totalVerdadeiros & vbCrLf & _
'       "Total de 'Falsos': " & totalFalsos
'
'
''plan b
'Set wb = Workbooks.Open(CaminhoVix & "PSI_B" & "_" & mesExtenso & ".xlsm")
'
'' Inicializar contadores globais
'totalVerdadeiros = 0
'totalFalsos = 0
'
'    ' Loop através de todas as planilhas na pasta de trabalho
'    For Each ws In wb.Sheets
'        ' Definir a célula para contar
'        Set celula = ws.Range("P:P")
'
'        ' Inicializar os contadores locais
'        contadorVerdadeiros = 0
'        contadorFalsos = 0
'
'        ' Loop através das células na coluna P
'        For Each cel In celula
'            If VarType(cel.Value) = vbBoolean Then
'                If cel.Value Then ' Verifica se a célula é verdadeira
'                    contadorVerdadeiros = contadorVerdadeiros + 1
'                Else ' Caso contrário, a célula é falsa
'                    contadorFalsos = contadorFalsos + 1
'                End If
'            End If
'        Next cel
'
'        ' Adicionar aos contadores globais
'        totalVerdadeiros = totalVerdadeiros + contadorVerdadeiros
'        totalFalsos = totalFalsos + contadorFalsos
'    Next ws
'
'' Exibir os resultados gerais
'MsgBox "Resultados Gerais:" & vbCrLf & _
'       "Total de 'Verdadeiros': " & totalVerdadeiros & vbCrLf & _
'       "Total de 'Falsos': " & totalFalsos

' Abrir PSI
 Set wb = Workbooks.Open(CaminhoMao & "PSI MAO A" & "_" & mesExtenso & ".xlsm")
 ' Inicializar contadores globais
totalVerdadeiros = 0
totalFalsos = 0


    ' Loop através de todas as planilhas na pasta de trabalho
    For Each ws In wb.Sheets
        ' Definir a célula para contar
        Set celula = ws.Range("P:P")

        ' Inicializar os contadores locais
        contadorVerdadeiros = 0
        contadorFalsos = 0

        ' Loop através das células na coluna P
        For Each cel In celula
            If VarType(cel.Value) = vbBoolean Then
                If cel.Value Then ' Verifica se a célula é verdadeira
                    contadorVerdadeiros = contadorVerdadeiros + 1
                Else ' Caso contrário, a célula é falsa
                    contadorFalsos = contadorFalsos + 1
                End If
            End If
        Next cel

        ' Adicionar aos contadores globais
        totalVerdadeiros = totalVerdadeiros + contadorVerdadeiros
        totalFalsos = totalFalsos + contadorFalsos
    Next ws

' Exibir os resultados gerais
MsgBox "Resultados Gerais:" & vbCrLf & _
       "Total de 'Verdadeiros': " & totalVerdadeiros & vbCrLf & _
       "Total de 'Falsos': " & totalFalsos


'plan mao b

Set wb = Workbooks.Open(CaminhoMao & "PSI MAO B" & "_" & mesExtenso & ".xlsm")

' Loop através de todas as planilhas na pasta de trabalho
    For Each ws In wb.Sheets
        ' Definir a célula para contar
        Set celula = ws.Range("P:P")

        ' Inicializar os contadores locais
        contadorVerdadeiros = 0
        contadorFalsos = 0

        ' Loop através das células na coluna P
        For Each cel In celula
            If VarType(cel.Value) = vbBoolean Then
                If cel.Value Then ' Verifica se a célula é verdadeira
                    contadorVerdadeiros = contadorVerdadeiros + 1
                Else ' Caso contrário, a célula é falsa
                    contadorFalsos = contadorFalsos + 1
                End If
            End If
        Next cel

        ' Adicionar aos contadores globais
        totalVerdadeiros = totalVerdadeiros + contadorVerdadeiros
        totalFalsos = totalFalsos + contadorFalsos
    Next ws

' Exibir os resultados gerais
MsgBox "Resultados Gerais:" & vbCrLf & _
       "Total de 'Verdadeiros': " & totalVerdadeiros & vbCrLf & _
       "Total de 'Falsos': " & totalFalsos

End Sub

'.- -.. --- .-.. .--. .... ---
