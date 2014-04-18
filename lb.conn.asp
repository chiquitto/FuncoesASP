<%

Function getValueBD(getValueBDTable, getValueBDCol, getValueBDWhere, getValueBDOrder)
	getValueBD = ""
	
	getValueBDSql = configQids & "SELECT TOP 1 " & getValueBDCol & " FROM " & userToBd & getValueBDTable
	If getValueBDWhere<>"" Then getValueBDSql = getValueBDSql & " WHERE " & getValueBDWhere
	If getValueBDOrder<>"" Then getValueBDSql = getValueBDSql & " ORDER BY " & getValueBDOrder
	Set getValueBDQid = db.Execute(getValueBDSql)
	If not getValueBDQid.Eof Then
		getValueBD = getValueBDQid(0)
	End If
	Set getValueBDQid = Nothing
End Function

Function recordset_getValue(rs, campo)
	recordset_getValue = ""
	For each it in rs.Fields
		If it.Name = campo Then recordset_getValue = it.Value
	Next
End Function
%>