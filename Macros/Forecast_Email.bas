Attribute VB_Name = "Forecast_Email"
Sub Email_Maq_Direto()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail
'Forecast Máquinas direto

.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "william.espinelli@konicaminolta.com" ' para
.Subject = "Forecast de máquinas Diretas" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de máquina dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Diretas.xlsx") ' anexar email

'.Send 'enviar
End With




End Sub

Public Sub Email_Maq_Indireto()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail
'Forecast Máquinas Indirera

.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "eduardo.nieto@konicaminolta.com" ' para
.Subject = "Forecast de máquinas Indiretas" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de máquina dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast_Indirect.xlsx") ' anexar email

'.Send 'enviar

End With

End Sub

Public Sub Email_Maq_SUL()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail
'Forecast Máquinas Indirera

.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "giovani.perini@konicaminolta.com" & ";" & " " & "willian.luz@konicaminolta.com" ' para
.Subject = "Forecast de Maquinas SUL" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de máquina dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast - SUL.xlsx") ' anexar email

'.Send 'enviar
End With

End Sub

Public Sub Email_Maq_MAO()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail
'Forecast Máquinas Indirera

.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "kaythy.duarte@konicaminolta.com" ' para
.Subject = "Forecast de Maquinas MAO" ' Assunto

' textos do email
texto1 = "Prezada,"
texto2 = "Segue planilha para preenchimento do forecast de máquina dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\_Forecast - MAO.xlsx") ' anexar email

'.Send 'enviar
End With

End Sub

Public Sub Email_Cons_VI()


Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail
'Forecast Máquinas Indirera

.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "eduardo.nieto@konicaminolta.com" ' para
.Subject = "Forecast de Consumos VI" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de consumo dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_VI SP.xlsx") ' anexar email

'.Send 'enviar

End With

End Sub

Public Sub Email_Cons_PE()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail


.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "raphael.carvalho@konicaminolta.com" ' para
.Subject = "Forecast de Consumos PE" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de consumo dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_PE.xlsx") ' anexar email

'.Send 'enviar

End With

End Sub

Public Sub Email_Cons_SUL()

Dim objOutlook As Outlook.Application
Dim objEmail As Outlook.MailItem

Set objOutlook = GetObject(, "Outlook.Application")

Set objEmail = objOutlook.CreateItem(olMailItem) ' RODAR COM O OUTLOOK ABERTO

Dim dtRetorno As Date
dtRetorno = WorksheetFunction.WorkDay(Date, 2)

With objEmail


.Display ' abrir email
SendUsingAccount = objOutlook.Session.Accounts(1) 'escilher email
.CC = "sabrina.sakagawa@konicaminolta.com" & ";" & " " & "rosana.odani@konicaminolta.com" & ";" & " " & "carlos.silva@konicaminolta.com" ' copia
.To = "giovani.perini@konicaminolta.com" & ";" & " " & "willian.luz@konicaminolta.com" ' para
.Subject = "Forecast de Consumos SUL" ' Assunto

' textos do email
texto1 = "Prezado,"
texto2 = "Segue planilha para preenchimento do forecast de consumo dos meses. Por gentileza retornar até o dia " & Format(dtRetorno, "dd/mmm") & "."
texto3 = "Qualquer dúvida estamos a disposição"

.HTMLBody = texto1 & "<br><br>" & texto2 & "<br><br><br>" & texto3 & .HTMLBody 'formatação do email
.Attachments.Add ("C:\Users\fsp_adolpho.salvador\Desktop\FORECAST\Forecast Vendas_SUL.xlsx") ' anexar email

'.Send 'enviar

End With

End Sub

