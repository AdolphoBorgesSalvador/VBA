Attribute VB_Name = "Consolidado_Mapas"
'1. A macro `consolidarMapas()` inicia com a declara��o de uma vari�vel `resposta` do tipo `VbMsgBoxResult`, que ser� utilizada para armazenar a resposta do usu�rio � mensagem exibida.
'
'2. A linha `ThisWorkbook.UpdateLinks = xlUpdateLinksNever` define que os links para outros arquivos n�o ser�o atualizados automaticamente quando a pasta de trabalho for aberta. Isso garante que os dados sejam consolidados a partir dos arquivos existentes no momento da execu��o da macro.
'
'3. S�o declaradas as vari�veis `tempoInicial`, `tempoFinal` e `tempoDecorrido` do tipo `Double`, que ser�o utilizadas para medir o tempo de execu��o da macro.
'':):
'4. Uma mensagem � exibida ao usu�rio com a fun��o `MsgBox`, perguntando se deseja iniciar a macro. O resultado da escolha � armazenado na vari�vel `resposta`.
'
'5. Em seguida, a macro verifica o valor da vari�vel `resposta` para determinar se o processo deve ser executado. Se o usu�rio escolher "Sim" (`vbYes`), o c�digo dentro do bloco `If resposta = vbYes Then` ser� executado.
'
'6. O tempo inicial � registrado usando a fun��o `Timer`.
'
'7. A anima��o da aplica��o e a atualiza��o da tela s�o desativadas com `Application.EnableAnimations = False` e `Application.ScreenUpdating = False`, respectivamente. Isso melhora o desempenho da macro, tornando a execu��o mais r�pida.
'
'8. S�o chamadas v�rias sub-rotinas (macros menores) para consolidar os mapas de diferentes marcas: `Brother`, `MAO`, `Muratec`, `VIX`, `BrGroup` e `Epson`. Essas sub-rotinas devem estar definidas em algum lugar do mesmo arquivo ou em um arquivo externo.
'
'9. Ap�s a consolida��o dos dados, a atualiza��o da tela � ativada novamente com `Application.ScreenUpdating = True` e as anima��es s�o reativadas com `Application.EnableAnimations = True`.
'
'10. O tempo final � registrado usando a fun��o `Timer`.
'
'11. O tempo decorrido � calculado subtraindo o tempo final pelo tempo inicial e convertido em minutos.
'
'12. Uma mensagem � exibida ao usu�rio informando que o processo foi conclu�do e mostrando o tempo decorrido.
'
'13. Caso o usu�rio escolha "N�o" (`vbNo`) na mensagem inicial, a macro exibir� uma mensagem informando que o processo foi cancelado.
'
'14. Finalmente, a linha `ThisWorkbook.UpdateLinks = xlUpdateLinksAlways` � utilizada para restaurar o comportamento padr�o de atualiza��o de links, garantindo que os links voltem a ser atualizados automaticamente quando a pasta de trabalho for aberta em futuras ocasi�es

Sub consolidarMapas()

Dim resposta As VbMsgBoxResult
Dim tempoInicial As Double
Dim tempoFinal As Double
Dim tempoDecorrido As Double
    
ThisWorkbook.UpdateLinks = xlUpdateLinksNever

resposta = MsgBox("Deseja inicar a macro?", vbYesNo, "MAPAS CONSOLIDADOS")

If resposta = vbYes Then

tempoInicial = Timer ' armazena o tempo inicial

    Application.EnableAnimations = False
    Application.ScreenUpdating = False
    
    Call MAO
    Call VIX
    Call ConsumosTerceiros
    Application.Wait Now + TimeValue("00:00:06")
    Call SalvarNoServidor
    
    Application.ScreenUpdating = True
    Application.EnableAnimations = True

    tempoFinal = Timer ' armazena o tempo final
    tempoDecorrido = tempoFinal - tempoInicial ' calcula o tempo decorrido
    tempoDecorrido = tempoDecorrido / 60
    
    MsgBox "Processo Conclu�do!!! :) "
    
    MsgBox "Tempo decorrido: " & tempoDecorrido & " Minutos."

Else
   MsgBox "Processo cancelado ", vbInformation, "MAPAS CONSOLIDADOS"
   
End If

ThisWorkbook.UpdateLinks = xlUpdateLinksAlways

End Sub
'.- -.. --- .-.. .--. .... ---

