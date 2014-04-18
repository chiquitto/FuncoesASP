<%
lbZipLoaded = True

Set lbZip_obj = Nothing
Sub lbZip_loadObj
	If lbZip_obj is Nothing Then Set lbZip_obj = Server.CreateObject("XStandard.Zip")
End Sub

' Adiciona um arquivo a um arquivo Zip
Sub lbZip_simplePack(lbZip_simplePack_file, lbZip_simplePack_filezip)
	lbZip_loadObj
	lbZip_obj.Pack lbZip_simplePack_file, lbZip_simplePack_filezip
End Sub
%>