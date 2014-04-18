<%
Function getObjHttp()
	Set getObjHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
End Function

Function simpleGetHttp(simpleGetHttp_file, simpleGetHttp_param)
	Set simpleGetHttpObj = getObjHttp()
	
	If isArray(simpleGetHttp_param) Then
		simpleGetHttp_param = Join(simpleGetHttp_param, "&")
	End If
	If simpleGetHttp_param <> "" Then simpleGetHttp_file = simpleGetHttp_file & "?" & simpleGetHttp_param
	simpleGetHttpObj.Open "GET", simpleGetHttp_file, false
	simpleGetHttpObj.Send
	
	If simpleGetHttpObj.Status <> 200 Then
		simpleGetHttp = ""
	Else
		simpleGetHttp = simpleGetHttpObj.responseText
	End If
	Set simpleGetHttpObj = Nothing
End Function

Sub logHttpRequest
	str = "=============================================================================" & vbNewLine
	str = str & "* DATE: " & Now & vbNewLine
	str = str & "* PATH_TRANSLATED: " & Request.ServerVariables("PATH_TRANSLATED") & vbNewLine
	str = str & "* REQUEST_METHOD: " & Request.ServerVariables("REQUEST_METHOD") & vbNewLine
	str = str & "* ALL_HTTP: " & vbNewLine & Request.ServerVariables("ALL_RAW")
	str = str & "* ALL_RAW: " & vbNewLine & Request.ServerVariables("ALL_RAW")
	
	str = str & Request.ServerVariables("REQUEST_METHOD") & ":" & vbNewLine
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		Set itens = Request.Form
	Else
		Set itens = Request.QueryString
	End If
	
	For each it in itens
		str = str & "	" & it & "=" & itens(it) & vbNewLine
	Next
	logHttpRequest_pathFile = pathSite & "tmp\requesthttp\" & date2String(Date, "%Y%m%d") & ".txt"
	str = loadFile(logHttpRequest_pathFile, False) & str & vbNewLine & vbNewLine
	saveInFile logHttpRequest_pathFile, str, True
End Sub
%>