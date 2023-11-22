Attribute VB_Name = "Forecast_Consolidados"
Sub consolidar_Forecast_email()

Dim respostas As VbMsgBoxResult

resposta = MsgBox("Deseja inicar o forecast de máquinas?", vbYesNo, "FORECAST DE MÁQUINAS")

If resposta = vbYes Then

Application.ScreenUpdating = False

'Rodar Forecast
'Call ForecastMaquinas
'Call Email_Maq_Direto
'Call Email_Maq_Indireto
'Call Email_Maq_SUL
'Call Email_Maq_MAO


End If

resposta = MsgBox("Deseja inicar o forecast de consumo?", vbYesNo, "FORECAST DE CONSUMOS")
If resposta = vbYes Then


'Rodar Email
Call ForecastConsumo
Call Email_Cons_VI
Call Email_Cons_PE
'Call Email_Cons_SUL

Else
MsgBox "Processo Concluído!!! :D "

End If
MsgBox "Processo Concluído!!! :D "


End Sub

Sub consolidadoConsumos()

Set ConsVi = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_VI SP")
Set ConsVd = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_VD")
Set ConsSul = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_SUL")
Set ConsMao = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_MAO")
Set ConsPE = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_PE")
Set consolidado = Workbooks.Open("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Consolidado Consumo")

Set ConsViWb = Workbooks("Forecast Vendas_VI SP")
Set ConsVdWb = Workbooks("Forecast Vendas_VD")
Set ConsSulWb = Workbooks("Forecast Vendas_SUL")
Set ConsMaoWb = Workbooks("Forecast Vendas_MAO")
Set ConsPeWb = Workbooks("Forecast Vendas_PE")
Set ConsConsWb = Workbooks("Consolidado Consumo")

Set ConsViWs = ConsViWb.Sheets("VI SP")
Set ConsVdWs = ConsVdWb.Sheets("VD")
Set ConsSulPoaWs = ConsSulWb.Sheets("POA")
Set ConsSulFlnWs = ConsSulWb.Sheets("FLN")
Set ConsMaoWs = ConsMaoWb.Sheets("MAO")
Set ConsPeWs = ConsPeWb.Sheets("PE")

Set ConsConsViWs = ConsConsWb.Sheets("VI")
Set ConsConsVdWs = ConsConsWb.Sheets("VD")
Set ConsConsPoaWs = ConsConsWb.Sheets("POA")
Set ConsConsFlnWs = ConsConsWb.Sheets("FLN")
Set ConsConsMaoWs = ConsConsWb.Sheets("MAO")
Set ConsConsPeWs = ConsConsWb.Sheets("PE")
Set ConsConsWs = ConsConsWb.Sheets("TOTAL")

With ConsConsViWs
If ConsViWs.ProtectContents Then ConsViWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsViWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsViWs.Protect Password:="km2023"
End With
    
With ConsConsVdWs
If ConsVdWs.ProtectContents Then ConsVdWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsVdWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsVdWs.Protect Password:="km2023"
End With

With ConsConsPoaWs
If ConsSulPoaWs.ProtectContents Then ConsSulPoaWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsSulPoaWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsSulPoaWs.Protect Password:="km2023"
End With
    
With ConsConsFlnWs
If ConsSulFlnWs.ProtectContents Then ConsSulFlnWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsSulFlnWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsSulFlnWs.Protect Password:="km2023"
End With

With ConsConsMaoWs
If ConsMaoWs.ProtectContents Then ConsMaoWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsMaoWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsMaoWs.Protect Password:="km2023"
End With
    
With ConsConsPeWs
If ConsPeWs.ProtectContents Then ConsPeWs.Unprotect Password:="km2023"
    .Range("a2").CurrentRegion.ClearContents
ConsPeWs.Range("a5").CurrentRegion.Copy
    .Range("a1").PasteSpecial xlPasteValues
ConsPeWs.Protect Password:="km2023"
End With

filePaths = Array("Forecast Vendas_VI SP", "Forecast Vendas_VD", "Forecast Vendas_SUL", "Forecast Vendas_MAO", "Forecast Vendas_PE", "Consolidado Consumo")

For i = LBound(filePaths) To UBound(filePaths)
     With filePaths(i)
    ActiveWindow.Close saveChanges:=True
     End With
     
Next i

End Sub

