Attribute VB_Name = "SalvaNoServidor"
'Este código que tem como objetivo salvar arquivos específicos em diferentes localizações, tanto no servidor (X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS) quanto localmente (C:\Users\fsp_adolpho.salvador\Desktop\PSI). A macro é dividida em duas partes: SalvarNoServidor() e SalvarNaLocal().
'
'Parte 1 - SalvarNoServidor():
'
'A macro começa definindo algumas variáveis, como o caminho do servidor, o mês atual e o mês por extenso (em letras maiúsculas).
'Em seguida, a macro abre os arquivos "PSI_A.xlsm" e "PSI_B.xlsm" localizados na pasta "C:\Users\fsp_adolpho.salvador\Desktop\PSI" usando um loop e a função Workbooks.Open.
'O array filePaths é definido para conter os nomes dos arquivos que serão salvos no servidor com base no mês atual (por exemplo, "PSI_A_JULHO.xlsm").
'Outro loop é utilizado para salvar os arquivos com nomes modificados na pasta do servidor "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS" usando a função ActiveWorkbook.SaveAs.
'':):
'Parte 2 - SalvarNaLocal():
'
'A segunda macro segue uma lógica semelhante à primeira, mas inverte o caminho de origem e destino. Nesta parte, os arquivos são abertos a partir do servidor e salvos na pasta local "C:\Users\fsp_adolpho.salvador\Desktop\PSI".
'Novamente, é usado o array filePaths para conter os nomes dos arquivos que serão salvos localmente com base no mês atual.
'O loop é utilizado para salvar os arquivos com os nomes modificados na pasta local.

Sub SalvarNoServidor()

    Dim caminhoServidor As String
    Dim mes As Integer
    Dim mesExtenso As String
    Dim CaminhoVix As String
    Dim CaminhoMao As String
    Dim wbPSI_A As Workbook
    Dim wbPSI_B As Workbook
    Dim wbPSI_C As Workbook
    Dim wbPSI_A_MAO As Workbook
    Dim wbPSI_B_MAO As Workbook
    Dim filePaths As Variant
    Dim i As Long

    mes = Month(Date)
    mesExtenso = UCase(MonthName(mes))
    CaminhoVix = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\"
    CaminhoMao = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI MAO\"
    CaminhoMuratec = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI MURATEC\"
    CaminhoTerceiros = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI TERCEIROS\"
    CaminhoPasta = "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\RELATORIO\"
    
    'VIX
    
    filePaths = Array("PSI_A.xlsm", "PSI_B.xlsm", "PSI_C.xlsm")
    
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
    Set wbPSI_A = Workbooks("PSI_A.xlsm")
    Set wbPSI_B = Workbooks("PSI_B.xlsm")
    Set wbPSI_C = Workbooks("PSI_C.xlsm")

    
    With wbPSI_A
        .SaveAs CaminhoVix & "PSI_A" & "_" & mesExtenso & ".xlsm"
        .Close saveChanges:=True
    End With

    With wbPSI_B
        .SaveAs CaminhoVix & "PSI_B" & "_" & mesExtenso & ".xlsm"
        .Close saveChanges:=True
    End With

    With wbPSI_C
        .SaveAs CaminhoVix & "PSI_C" & "_" & mesExtenso & ".xlsm"
        .Close saveChanges:=True
    End With

    ' MAO
    
    filePaths = Array("PSI MAO A.xlsm", "PSI MAO B.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
    Set wbPSI_A_MAO = Workbooks("PSI MAO A.xlsm")
    Set wbPSI_B_MAO = Workbooks("PSI MAO B.xlsm")
    
    With wbPSI_A_MAO
        .SaveAs CaminhoMao & "PSI MAO A" & "_" & mesExtenso & ".xlsm"
        .Close saveChanges:=True
    End With

    With wbPSI_B_MAO
        .SaveAs CaminhoMao & "PSI MAO B" & "_" & mesExtenso & ".xlsm"
        .Close saveChanges:=True
    End With
    
    'Muratec
    filePaths = Array("PSI Muratec.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
    Set wbMuratec = Workbooks("PSI Muratec.xlsm")
    
        With wbMuratec
            .SaveAs CaminhoMuratec & "PSI Muratec" & "_" & mesExtenso & ".xlsm"
            .Close saveChanges:=True
        End With
        
    'Terceiros
    filePaths = Array("Consumos Terceiros.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
    Set wbTerceiros = Workbooks("Consumos Terceiros.xlsm")
    
        With wbMuratec
            .SaveAs CaminhoTerceiros & "Consumos Terceiros" & "_" & mesExtenso & ".xlsm"
            .Close saveChanges:=True
        End With
        
    'PASTA

 filePaths = Array("VIX.xlsx", "MAO.XLSX", "GERAL.XLSX", "Estoque.xls", "FUP.XLSX", "MB51 Epson Base.XLSX", "Zstok Brgoup Base.XLSX", "MB51 Brgoup Bse.XLSX", "ZSTOK Epson BASE.XLSX", "Mapa Maq_Acess.XLSX")
    For i = LBound(filePaths) To UBound(filePaths)
        Workbooks.Open "C:\Users\fsp_adolpho.salvador\Desktop\RELATORIO\" & filePaths(i)
    Next i
    
    Set wbVIX = Workbooks("VIX.xlsx")
    Set wbMAO = Workbooks("MAO.XLSX")
    Set wbGeral = Workbooks("GERAL.XLSX")
    Set wbEstoque = Workbooks("Estoque.XLSX")
    Set wbFUP = Workbooks("FUP.XLSX")
    Set wbMb51Epson = Workbooks("MB51 Epson Base.XLSX")
    Set wbZstokBrgroup = Workbooks("Zstok Brgoup Base.XLSX")
    Set wbMb51Brgroup = Workbooks("MB51 Brgoup Bse.XLSX")
    Set wbZstokEpson = Workbooks("ZSTOK Epson BASE.XLSX")
    Set wbMapaMaq = Workbooks("Mapa Maq_Acess.XLSX")
    
    
    With wbVIX
        .SaveAs CaminhoPasta & "VIX"
        .Close saveChanges:=True
    End With

    With wbMAO
        .SaveAs CaminhoPasta & "MAO"
        .Close saveChanges:=True
    End With
    
     With wbGeral
        .SaveAs CaminhoPasta & "GERAL"
        .Close saveChanges:=True
    End With


    With wbEstoque
        .SaveAs CaminhoPasta & "Estoque"
        .Close saveChanges:=True
    End With
    
    With wbFUP
        .SaveAs CaminhoPasta & "FUP"
        .Close saveChanges:=True
    End With

    With wbMb51Epson
        .SaveAs CaminhoPasta & "MB51 Epson Base"
        .Close saveChanges:=True
    End With

    With wbZstokBrgroup
        .SaveAs CaminhoPasta & "Zstok Brgroup Base"
        .Close saveChanges:=True
    End With

    With wbMb51Brgroup
        .SaveAs CaminhoPasta & "MB51 Brgroup Base"
        .Close saveChanges:=True
    End With

    With wbMapaMaq
        .SaveAs CaminhoPasta & "Mapa Maq_Acess"
        .Close saveChanges:=True
    End With


'.- -.. --- .-.. .--. .... ---

End Sub

Sub SalvarNaLocal()

 Dim caminhoServidor As String
    Dim mes As Integer
    Dim mesExtenso As String

    mes = Month(Date)
    mesExtenso = UCase(MonthName(mes))
    
    'VIX
    
    Dim filePaths As Variant
    filePaths = Array("PSI_A" & "_" & mesExtenso & ".xlsm", "PSI_B" & "_" & mesExtenso & ".xlsm", "PSI_C" & "_" & mesExtenso & ".xlsm")
    
    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\" & filePaths(i)

 Next i

    filePaths = Array("PSI_A1.xlsm", "PSI_B1.xlsm", "PSI_C1.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
    ActiveWorkbook.SaveAs "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    Next i
    
          
      'MAO
    
    filePaths = Array("PSI MAO A" & " " & mesExtenso & ".xlsm", "PSI MAO B" & " " & mesExtenso & ".xlsm", "PSI MAO C" & " " & mesExtenso & ".xlsm")
  
    For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\" & filePaths(i)

    Next i

    filePaths = Array("PSI MAO A.xlsm", "PSI MAO B.xlsm", "PSI MAO C.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
    ActiveWorkbook.SaveAs "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    
    Next i

      'Muratec
    
    filePaths = Array("PSI Muratec" & " " & mesExtenso & ".xlsm")
  
    For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\" & filePaths(i)

    Next i

    filePaths = Array("PSI Muratec.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
    ActiveWorkbook.SaveAs "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    
    Next i

      'Muratec
    
    filePaths = Array("PSI Muratec" & " " & mesExtenso & ".xlsm")
  
    For i = LBound(filePaths) To UBound(filePaths)
    Workbooks.Open "X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\" & mes & ". " & mesExtenso & "\PSI KMI\" & filePaths(i)

    Next i

    filePaths = Array("PSI BR GROUP.xlsm")
    For i = LBound(filePaths) To UBound(filePaths)
    ActiveWorkbook.SaveAs "C:\Users\fsp_adolpho.salvador\Desktop\PSI\" & filePaths(i)
    
    Next i

End Sub
'.- -.. --- .-.. .--. .... ---

