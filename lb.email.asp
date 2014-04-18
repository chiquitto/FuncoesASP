<%
Function lbEmailObj()
	Dim MailAut
	Set MailAut	= Server.CreateObject("Persits.MailSender")
	
	MailAut.Host		= ""
	MailAut.Timestamp	= Now()
	MailAut.Username	= ""
	MailAut.Password	= "" 
	Set lbEmailObj = MailAut
End Function

Function enviaMail(remetenteNome, remetenteEmail, destinatario, subject, body, isHtml)
	Set MailAut	= lbEmailObj
	
	MailAut.FromName	= remetenteNome
	MailAut.From		= remetenteEmail
	
	MailAut.AddAddress destinatario
	
	MailAut.Subject		= MailAut.EncodeHeader(subject)
	MailAut.Body		= body
	MailAut.IsHTML		= isHtml
	
	'MailAut.SendToQueue
	MailAut.Send
	
	Set MailAut = Nothing
	
	enviaMail = True
End Function

Function enviaMailcomAnexo(remetenteNome, remetenteEmail, destinatario, subject, body, isHtml, anexos)
	Set MailAut	= lbEmailObj
	
	MailAut.FromName	= remetenteNome
	MailAut.From		= remetenteEmail
	
	MailAut.AddAddress destinatario
	
	MailAut.Subject		= subject
	MailAut.Body		= body
	MailAut.IsHTML		= isHtml
	
	For iAnexo=0 To UBound(anexos)
		MailAut.AddAttachment anexos(iAnexo)
	Next
	
	'MailAut.SendToQueue
	MailAut.Send
	
	Set MailAut = Nothing
	
	enviaMailcomAnexo = True
End Function

Function enviaMail2(remetenteNome, remetenteEmail, destinatario, subject, body, isHtml)
	enviaMail2 = enviaMail(remetenteNome, remetenteEmail, destinatario, subject, body, isHtml)
	Exit Function
	
	If isHtml=True Then
		isHtml="1"
	Else
		isHtml="0"
	End If

	strFile = privateSite & "\email\" & destinatario & ".e"
	rt = False
	
	SEP = "<<"&"%*%"&">>"
	bodyFile = ""
	bodyFile = bodyFile & remetenteNome & SEP
	bodyFile = bodyFile & remetenteEmail & SEP
	bodyFile = bodyFile & destinatario & SEP
	bodyFile = bodyFile & destinatario & SEP
	bodyFile = bodyFile & subject & SEP
	bodyFile = bodyFile & isHtml & SEP
	bodyFile = bodyFile & body & SEP
	
	Set fso = Server.CreateObject("scripting.FileSystemObject")
	If fso.FileExists(strFile) Then fso.DeleTeFile strFile
	t = saveInFile(strFile, bodyFile, True)
	
	If fso.FileExists(strFile) Then
		url = virtualSite & "library/email.php?tipo=envia&file=" & Server.URLEncode(destinatario) & ".e"
		Set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
		objHttp.Open "GET", url, false
		objHttp.Send
		
		If (objHttp.status = 200) Then
			If objHttp.responseText="0" Then
				rt = True
			End If
			rt = True
		End If
		fso.DeleTeFile strFile
		Set objHttp = Nothing
	End If
	enviaMail2 = rt
End Function
%>