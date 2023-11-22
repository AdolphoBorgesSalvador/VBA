Attribute VB_Name = "Consolidado_Mapas"
'1. A macro `consolidarMapas()` inicia com a declaração de uma variável `resposta` do tipo `VbMsgBoxResult`, que será utilizada para armazenar a resposta do usuário à mensagem exibida.
'
'2. A linha `ThisWorkbook.UpdateLinks = xlUpdateLinksNever` define que os links para outros arquivos não serão atualizados automaticamente quando a pasta de trabalho for aberta. Isso garante que os dados sejam consolidados a partir dos arquivos existentes no momento da execução da macro.
'
'3. São declaradas as variáveis `tempoInicial`, `tempoFinal` e `tempoDecorrido` do tipo `Double`, que serão utilizadas para medir o tempo de execução da macro.
'':):
'4. Uma mensagem é exibida ao usuário com a função `MsgBox`, perguntando se deseja iniciar a macro. O resultado da escolha é armazenado na variável `resposta`.
'
'5. Em seguida, a macro verifica o valor da variável `resposta` para determinar se o processo deve ser executado. Se o usuário escolher "Sim" (`vbYes`), o código dentro do bloco `If resposta = vbYes Then` será executado.
'
'6. O tempo inicial é registrado usando a função `Timer`.
'
'7. A animação da aplicação e a atualização da tela são desativadas com `Application.EnableAnimations = False` e `Application.ScreenUpdating = False`, respectivamente. Isso melhora o desempenho da macro, tornando a execução mais rápida.
'
'8. São chamadas várias sub-rotinas (macros menores) para consolidar os mapas de diferentes marcas: `Brother`, `MAO`, `Muratec`, `VIX`, `BrGroup` e `Epson`. Essas sub-rotinas devem estar definidas em algum lugar do mesmo arquivo ou em um arquivo externo.
'
'9. Após a consolidação dos dados, a atualização da tela é ativada novamente com `Application.ScreenUpdating = True` e as animações são reativadas com `Application.EnableAnimations = True`.
'
'10. O tempo final é registrado usando a função `Timer`.
'
'11. O tempo decorrido é calculado subtraindo o tempo final pelo tempo inicial e convertido em minutos.
'
'12. Uma mensagem é exibida ao usuário informando que o processo foi concluído e mostrando o tempo decorrido.
'
'13. Caso o usuário escolha "Não" (`vbNo`) na mensagem inicial, a macro exibirá uma mensagem informando que o processo foi cancelado.
'
'14. Finalmente, a linha `ThisWorkbook.UpdateLinks = xlUpdateLinksAlways` é utilizada para restaurar o comportamento padrão de atualização de links, garantindo que os links voltem a ser atualizados automaticamente quando a pasta de trabalho for aberta em futuras ocasiões

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
    
    MsgBox "Processo Concluído!!! :) "
    
    MsgBox "Tempo decorrido: " & tempoDecorrido & " Minutos."

Else
   MsgBox "Processo cancelado ", vbInformation, "MAPAS CONSOLIDADOS"
   
End If

ThisWorkbook.UpdateLinks = xlUpdateLinksAlways

End Sub
'.- -.. --- .-.. .--. .... ---

