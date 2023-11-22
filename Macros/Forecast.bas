Attribute VB_Name = "Forecast"
Sub ForecastMaquinas()

Dim MaqViWb As Workbook
Dim MaqVdWb As Workbook
Dim MaqSulWb As Workbook
Dim MaqMaoWb As Workbook
Dim DeliveryInfoWb As Workbook
Dim ZstokBaseWb As Workbook
Dim ZstokBaseWb As Workbook

Dim MaqViWs As Worksheet
Dim MaqViFupWs As Worksheet
Dim MaqViZstokWs As Worksheet

Dim MaqVdWs As Worksheet
Dim MaqVdFupWs As Worksheet
Dim MaqVdZstokWs As Worksheet

Dim MaqSulPoaWs As Worksheet
Dim MaqSulFlnWs As Worksheet
Dim MaqSulFupWs As Worksheet
Dim MaqSulZstokWs As Worksheet

Dim MaqMaoWs As Worksheet
Dim ZstokBaseWs As Worksheet
Dim MaqSulFupWs As Worksheet
Dim ZstokMaoWs As Worksheet
Dim ZstokVIXWs As Worksheet
Dim DeliveryInfoFUP As Worksheet
Dim ZstokSheet1Ws As Worksheet

Set MaqVi = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Máquinas - INDIRETO")
Set MaqVd = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Máquinas - DIRETO")
Set MaqSul = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Máquinas - SUL")
Set MaqMao = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Máquinas - MAO")
Set ZSTOK = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\zstok")
Set ZSTOKBase = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\zstok BASE")
Set DeliveryInfo = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Delivery Info - ETA Machines  Acessories")

Set MaqViWb = Workbooks("_Forecast_Máquinas - INDIRETO")
Set MaqVdWb = Workbooks("_Forecast_Máquinas - DIRETO")
Set MaqSulWb = Workbooks("_Forecast_Máquinas - SUL")
Set MaqMaoWb = Workbooks("_Forecast_Máquinas - MAO")
Set DeliveryInfoWb = Workbooks("Delivery Info - ETA Machines  Acessories")
Set ZstokWb = Workbooks("zstok")
Set ZstokBaseWb = Workbooks("zstok BASE")

Set MaqViWs = MaqViWb.Sheets("Indireto")
Set MaqViFupWs = MaqViWb.Sheets("FUP")
Set MaqViZstokWs = MaqViWb.Sheets("Estoque")

Set MaqVdWs = MaqVdWb.Sheets("Direto")
Set MaqVdFupWs = MaqVdWb.Sheets("FUP")
Set MaqVdZstokWs = MaqVdWb.Sheets("Estoque")

Set MaqSulPoaWs = MaqSulWb.Sheets("POA")
Set MaqSulFlnWs = MaqSulWb.Sheets("FLN")
Set MaqSulFupWs = MaqSulWb.Sheets("FUP")
Set MaqSulZstokWs = MaqSulWb.Sheets("Estoque")

Set MaqMaoWs = MaqMaoWb.Sheets("Direto")
Set MaqMaoFupWs = MaqMaoWb.Sheets("FUP")

Set DeliveryInfoFUP = DeliveryInfoWb.Sheets("PIVOT FUP")
Set ZstokBaseWs = ZstokWb.Sheets("BASE")
Set ZstokMaoWs = ZstokWb.Sheets("MAO")
Set ZstokSulWs = ZstokWb.Sheets("SUL")
Set ZstokVIXWs = ZstokWb.Sheets("VIX")
Set ZstokSheet1Ws = ZstokBaseWb.Sheets("Sheet1")

'Zstok

With ZstokSheet1Ws

    Dim lastCell As Range
    Set lastCell = .Range("A1").End(xlToRight).End(xlDown)
    
    .Range("A1", lastCell).Copy _
        Destination:=ZstokBaseWs.Range("A1")
        
End With

ZstokBaseWs.Parent.RefreshAll


'VI SP

With MaqViWs

If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("d1:d2").Copy
.Range("b1").PasteSpecial xlPasteValues

.Range("d7:k21").Copy
.Range("c7").PasteSpecial xlPasteValues

    Range("f7:k21").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
   Range("c7:e21").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
        
        
    With DeliveryInfoFUP
    
    Set lastCell = .Range("A4").End(xlToRight).End(xlDown)
    
    .Range("A4", lastCell).Copy _
         Destination:=MaqViFupWs.Range("A1")
    End With
    
    With ZstokVIXWs
    
    Set lastCell = .Range("A1").End(xlToRight).End(xlDown)
    
    .Range("A1", lastCell).Copy _
        Destination:=MaqViZstokWs.Range("A1")
    End With


MaqViWb.Save
MaqViWb.Close


'VD SP

With MaqVdWs

If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("d1:d2").Copy
.Range("b1").PasteSpecial xlPasteValues

.Range("d7:k21").Copy
.Range("c7").PasteSpecial xlPasteValues

    Range("f7:k21").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
   Range("c7:e21").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
        
      With MaqVdZstokWs
      .Range("a1:p100").ClearContents
    End With
    
      With MaqVdFupWs
      .Range("a1:p100").ClearContents
    End With
        
    With DeliveryInfoFUP
    
    Set lastCell = .Range("A4").End(xlToRight).End(xlDown)
    
    .Range("A4", lastCell).Copy _
         Destination:=MaqVdFupWs.Range("A1")
    End With
    
    With ZstokVIXWs
    
    Set lastCell = .Range("A1").End(xlToRight).End(xlDown)
    
    .Range("A1", lastCell).Copy _
        Destination:=MaqVdZstokWs.Range("A1")
    End With


MaqVdWb.Save
MaqVdWb.Close



'SUL

    
With MaqSulFlnWs

If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("d1:d2").Copy
.Range("b1").PasteSpecial xlPasteValues

.Range("d7:k21").Copy
.Range("c7").PasteSpecial xlPasteValues

    Range("f7:k20").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
   Range("c7:e20").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
    
   With MaqSulPoaWs

If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("d1:d2").Copy
.Range("b1").PasteSpecial xlPasteValues

.Range("d7:k21").Copy
.Range("c7").PasteSpecial xlPasteValues

    Range("f7:k20").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
   Range("c7:e20").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
     
    
    With DeliveryInfoFUP
    Set lastCell = .Range("A4").End(xlToRight).End(xlDown)
    
    .Range("A4", lastCell).Copy _
         Destination:=MaqSulFupWs.Range("A1")
         
      With MaqSulZstokWs
      .Range("a1:p100").ClearContents
      
    End With
    
    With ZstokSulWs
    
    Set lastCell = .Range("A1").End(xlToRight).End(xlDown)
    
    .Range("A1", lastCell).Copy _
         Destination:=MaqSulZstokWs.Range("A1")
    End With
    
End With


.Protect Password:="km2023"

'MaqSulWb.Save
'MaqSulWb.Close



'MAO

With MaqMaoWs

If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("d1:d2").Copy
.Range("b1").PasteSpecial xlPasteValues

.Range("d7:k21").Copy
.Range("c7").PasteSpecial xlPasteValues

    Range("f7:k20").Select
    
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Selection.Locked = False
    Selection.FormulaHidden = False
   Range("c7:e20").Select
    
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    
    Selection.Locked = True
    Selection.FormulaHidden = False

        
    With DeliveryInfoFUP
    
    Set lastCell = .Range("A4").End(xlToRight).End(xlDown)
    
    .Range("A4", lastCell).Copy _
         Destination:=MaqMaoFupWs.Range("A1")
    End With

.Protect Password:="km2023"

MaqMaoWb.Save
MaqMaoWb.Close

With DeliveryInfoWb
.Save
.Close
End With

With ZstokWb
.Save
.Close
End With

With ZstokBaseWb
.Save
.Close
End With

End With
End With
End With
End With
End With

End Sub

Sub ForecastConsumo()
'
'

Set PE = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_PE")
Set SUL = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_SUL")
Set VI_SP = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_VI SP")
Set MAO = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_MAO")
Set VD_SP = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_VD")
    Set Giro = Workbooks.Open("X:\PLANEJAMENTO\2. PSI\2023\3. CONSUMOS\Dispon Abr-23\3. Novo MB51 base_Abr-23.xlsx")

Set PeWb = Workbooks("Forecast Vendas_PE")
Set SulWb = Workbooks("Forecast Vendas_SUL")
Set ViWb = Workbooks("Forecast Vendas_VI SP")
Set MaoWb = Workbooks("Forecast Vendas_MAO")
Set VdWb = Workbooks("Forecast Vendas_VD")

Set GiroWb = Workbooks("3. Novo MB51 base_Abr-23.xlsx")

Set PeWs = PeWb.Sheets("PE")
Set PoaWs = SulWb.Sheets("POA")
Set FlnWs = SulWb.Sheets("FLN")
Set ViWs = ViWb.Sheets("VI SP")
Set MaoWs = MaoWb.Sheets("MAO")
Set VdWs = VdWb.Sheets("VD")

Set GiroPeWs = GiroWb.Sheets("DIRETO PE")
Set GiroSc952Ws = GiroWb.Sheets("DIRETO SC (952)")
Set GiroRs952Ws = GiroWb.Sheets("DIRETO RS (952)")
Set GiroSp952Ws = GiroWb.Sheets("DIRETO SP (952)")
Set GiroSc601Ws = GiroWb.Sheets("DIRETO SC (601)")
Set GiroRs601Ws = GiroWb.Sheets("DIRETO RS (601)")
Set GiroSp601Ws = GiroWb.Sheets("DIRETO SP (601)")
Set GiroIndiretoSpWs = GiroWb.Sheets("INDIRETO POA")
Set GiroIndiretoFloWs = GiroWb.Sheets("INDIRETO FLO")
Set GiroIndiretoPoaWs = GiroWb.Sheets("INDIRETO SP")
Set GiroUsoConsumoWs = GiroWb.Sheets("Uso e Consumo")

Dim lookupValue As Variant
Dim lookupRange As Range
Dim resultRange As Range

With PeWs
If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:k304").Copy _
        Destination:=PeWs.Range("e6")
Set lookupRange = GiroPeWs.Range("Q:Q")
Set resultRange = GiroPeWs.Range("AG:AG")

PeWs.Range("D6").Value = "=XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO PE'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO PE'!C33,0)"

.Range("D6").Copy _
       Destination:=PeWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

.Protect Password:="km2023"

PeWb.Save
PeWb.Close

End With



With PoaWs
If PoaWs.ProtectContents Then PoaWs.Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:k304").Copy _
        Destination:=PoaWs.Range("e6")

PoaWs.Range("D6").Value = "=SUM(XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO RS (601)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO RS (601)'!C33,0),XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO POA'!C17,'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO POA'!C33,0),XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO RS (952)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO RS (952)'!C33,0))"
.Range("D6").Copy _
       Destination:=PoaWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

.Protect Password:="km2023"

End With

With FlnWs
 If FlnWs.ProtectContents Then FlnWs.Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:k304").Copy _
        Destination:=FlnWs.Range("e6")

FlnWs.Range("D6").Value = "=SUM(XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SC (601)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SC (601)'!C33,0),XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO FLO'!C17,'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO FLO'!C33,0),XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SC (952)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SC (952)'!C33,0))"

.Range("D6").Copy _
       Destination:=FlnWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

.Protect Password:="km2023"

SulWb.Save
SulWb.Close

End With

With ViWs
 If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:k304").Copy _
        Destination:=ViWs.Range("e6")

ViWs.Range("D6").Value = "=XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO SP'!C17,'[3. Novo MB51 base_Abr-23.xlsx]INDIRETO SP'!C33,0)"
.Range("D6").Copy _
       Destination:=ViWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

.Protect Password:="km2023"

ViWb.Save
ViWb.Close

End With

With MaoWs
 If .ProtectContents Then .Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:i304").Copy _
        Destination:=MaoWs.Range("e6")

MaoWs.Range("D6").Value = "=XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]mao'!C1,'[3. Novo MB51 base_Abr-23.xlsx]mao'!C33,0)"
.Range("D6").Copy _
       Destination:=MaoWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

MaoWs.Protect Password:="km2023"

MaoWb.Save
MaoWb.Close

End With

With VdWs
 If VdWs.ProtectContents Then VdWs.Unprotect Password:="km2023"
    If .FilterMode Then .ShowAllData
.Range("f6:i304").Copy _
        Destination:=VdWs.Range("e6")

VdWs.Range("D6").Value = "=SUM(XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SP (601)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SP (601)'!C33,0),XLOOKUP(RC[-3],'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SP (952)'!C17,'[3. Novo MB51 base_Abr-23.xlsx]DIRETO SP (952)'!C33,0))"
.Range("D6").Copy _
       Destination:=VdWs.Range("D6:D304")
          
.Range("D6:D304").Copy
.Range("D6:D304").PasteSpecial xlPasteValues

.Protect Password:="km2023"

VdWb.Save
VdWb.Close

End With

With GiroWb
.Save
.Close
End With

End Sub



