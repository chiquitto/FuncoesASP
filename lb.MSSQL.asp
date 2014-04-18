<%
Function replaceCaracInvMSSQL(str, length)
	replaceCaracInvMSSQL = str
	'replaceCaracInvMSSQL = replace(replaceCaracInvMSSQL, chr(39), "")
	replaceCaracInvMSSQL = replace(replaceCaracInvMSSQL, "'", "''")
	If length>0 Then replaceCaracInvMSSQL = Left(replaceCaracInvMSSQL, length)
End Function

Function getOne(sqlOne)
	getOne = ""
	Set rsGetOne = db.Execute(sqlOne)
	If not rsGetOne.Eof Then
		getOne = rsGetOne(0)
	End If
	Set rsGetOne = Nothing
End Function
%>