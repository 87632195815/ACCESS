# ACCESS
Código VBA para ACCESS

If SENHA = CONFIRMA_SENHA Then
MsgBox (" Cadastro efetuado com sucesso ")
DoCmd.GoToRecord , , acNewRec
Else
MsgBox (" Senha confirmada não confere, digite a senha novamente")
End If
