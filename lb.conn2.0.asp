<%
conn20 = True

Dim db

Sub connOpen()
	If not isObject(db) Then
		%><!--#include file="conn.asp"--><%
	End If
End Sub

Sub connClose()
	If isObject(db) Then
		db.Close
	End If
	Set db = Nothing
End Sub
%>