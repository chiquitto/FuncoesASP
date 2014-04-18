<%
Sub headerLocation(url)
	Response.AddHeader "Location", url
End Sub

Sub headerMovedPermanently
	Response.Status = "301 Moved Permanently"
End Sub

Sub headerMovedTemporarily
	Response.Status = "302 Moved Temporarily"
End Sub

Sub headerNotFound
	Response.Status = "404 Not Found"
End Sub

Sub persolanizado404
	headerNotFound
	%><!--#include file="../error_docs/404.html"--><%
	Response.End()
End Sub
%>