<%
'PAGINAÇÃO
'CRIADO POR CHIQUITTO <GASPARCHIQUITTO@YAHOO.COM.BR>
'22/07/2005

Function insere_pagXli(pag, xli, nrows, varPag, vars)
	'DEFINIÇÕES DE VARS
	'	pag				= pagina atual
	'	per_pag			= quantidade de rows por pagina
	'	nrows			= rows encontrados no banco de dados
	'	varPag			= variavel que será passado por meio QueryString
	'	vars			= variaveis adicionais que serao inseridas para o metodo GET
	
	pag			= CInt(pag)
	xli			= CInt(xli)
	'nrows		= CInt(nrows)
	varPag		= CStr(varPag)
	vars		= CStr(vars)
	ultimo		= (nrows+(xli-1))\xli
	
	echo = ""
	If ultimo>1 Then
		echo = echo & "<scr"&"ipt language=""javascript"">"& _
			"function alterPag(pag, query){ window.location = '"& Request.ServerVariables("URL") &"?"
		If vars<>"" Then
			echo = echo & vars & "&"
		End If
		echo = echo & varPag &"='+pag+query; }"&vbNewLine&"</scr"&"ipt>"&vbNewLine
		echo = echo & "<input type=""button"" name=""Button"" value=""Anterior"" class=""buttonAnt"""
		If pag>1 Then
			echo = echo & " onClick=""alterPag('"& CStr(pag-1) &"','')"""
		Else
			echo = echo & " disabled"
		End If
		echo = echo & ">"&vbNewLine&"<input type=""button"" name=""Button"" value=""Pr&oacute;ximo"" class=""buttonPro"""
		If pag<ultimo Then
			echo = echo & " onClick=""alterPag('"& CStr(pag+1) &"','')"""
		Else
			echo = echo & " disabled"
		End If
		echo = echo & ">"&vbNewLine&"&nbsp;&nbsp;&nbsp;<select name=""select"" onChange=""alterPag(this.value,'')"">"&vbNewLine
	
		y=1
		For x=1 to nrows step xli
			xu = x+xli-1
			If xu>nrows Then xu=nrows
				echo = echo & "<option value="""& y &""""
				If pag=y Then
					echo = echo & " selected"
				End If
				echo = echo & ">"& Right("00000"&x, Len(nrows)) &" - "& Right("00000"&xu, Len(nrows)) & "</option>"&vbNewLine
			y = y+1
		Next
		echo = echo & "</select>"
	End If
	insere_pagXli = echo
End Function
%>