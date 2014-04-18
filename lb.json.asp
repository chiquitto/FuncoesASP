<%
'Function recordset2Json(recordset2Json_rs)
'	Dim r, r2
'	r = Array()
'	While not recordset2Json_rs.eof
'		r2 = Array()
'		For Each it In recordset2Json_rs.Fields
'			name = it.name
'			value = it.value
'			If not (it.type = 3) Then value = """" & value & """"
'		
'			r2 = arrayPush(r2, "" & name & ":" & value & "")
'		Next
'		r = arrayPush(r, "{" & Join(r2, ",") & "}")
'		recordset2Json_rs.moveNext
'	Wend
'	recordset2Json = "[" & Join(r, ",") & "]"
'End Function

Function recordset2Json(recordset2Json_rs)
	Dim r
	r = Array()
	While not recordset2Json_rs.eof
		r = arrayPush(r, "{" & Join(recordsetRow2Json(recordset2Json_rs), ",") & "}")
		recordset2Json_rs.moveNext
	Wend
	recordset2Json = "[" & Join(r, ",") & "]"
End Function

Function recordset2JsonLimit(recordset2JsonLimit_rs, recordset2JsonLimit_limit)
	Dim r
	r = Array()
	While ( not recordset2JsonLimit_rs.Eof ) And ( recordset2JsonLimit_limit > 0 )
		r = arrayPush(r, "{" & Join(recordsetRow2Json(recordset2JsonLimit_rs), ",") & "}")
		recordset2JsonLimit_limit = recordset2JsonLimit_limit - 1
		recordset2JsonLimit_rs.moveNext
	Wend
	recordset2JsonLimit = "[" & Join(r, ",") & "]"
End Function

recordsetRow2Json_ServerHTMLEncode = False
Function recordsetRow2Json( recordsetRow2Json_rs )
	recordsetRow2Json = Array()
	For Each it In recordsetRow2Json_rs.Fields
		name = it.name
		value = it.value
		If not (it.type = 3) Then
			value = Replace(value, """", "\""")
			If recordsetRow2Json_ServerHTMLEncode Then value = Server.HTMLEncode(value)
			value = """" & value & """"
		End If
	
		recordsetRow2Json = arrayPush(recordsetRow2Json, """" & name & """:" & value & "")
	Next
End Function

Function jsonVazio()
	jsonVazio = "[]"
End Function

Function jsonErro(jsonErro_number, jsonErro_desc)
	jsonErro = "{""erro"":" & jsonErro_number & ",""data"":[], ""msg"":""" & jsonErro_desc & """}"
End Function

Function jsonData(jsonData_erro, jsonData_data, jsonData_desc)
	jsonData = "{""erro"":" & jsonData_erro & ",""data"":" & jsonData_data & ",""msg"":""" & jsonData_desc & """}"
End Function
%>