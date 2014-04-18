<%
lbProtectLoaded = True

Function isLogado()
	If (Session("clienteLogado") = False) or (isEmpty(Session("clienteLogado"))) Then
		isLogado = False
	Else
		isLogado = True
	End If
End Function

Function fazerLogin(fazerLogin_email, fazerLogin_senha)
	fazerLogin_sql = "Select " & vbNewLine & _
	"	id_user, nome, profissao, exibir " & vbNewLine & _
	"From tusers " & vbNewLine & _
	"Where (email = '" & fazerLogin_email & "') " & vbNewLine & _
	"	And (senha = '" & fazerLogin_senha & "') " & vbNewLine & _
	"	And (email <> '') And (senha <> '')"
	
	connOpen	
	Set fazerLogin_cliente = db.Execute(fazerLogin_sql)			
	If fazerLogin_cliente.eof Then
		fazerLogin = 1
	Else
		If fazerLogin_cliente("exibir") <> "S" Then
			fazerLogin = 2
		Else
			Session("clienteLogado") = "true"
			Session("nome") = fazerLogin_cliente("nome")
			Session("id") = fazerLogin_cliente("id_user")
			Session("profissao") = fazerLogin_cliente("profissao")
			
			fazerLogin = 0
		End if
	End If
	Set fazerLogin_cliente = Nothing
End Function

Function protectGetVoltar
	protectGetVoltar = Session("voltar")
End Function

Sub protectLogout
	voltar = protectGetVoltar
	Session.Contents.RemoveAll()
	Session("id") = 0
	protectSetVoltar voltar
End Sub

Sub protectMakeVoltar
	voltar = Request.ServerVariables("URL")
	If voltar = "/Default.asp" Then voltar = "/"
	If Request.QueryString <> "" Then voltar = voltar & "?" & Request.QueryString
	
	protectSetVoltar voltar
End Sub

Sub protectSetVoltar( protectSetVoltar_voltar )
	Session("voltar") = protectSetVoltar_voltar
End Sub
%>