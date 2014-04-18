<%
Function permalinkClassificados( permalinkClassificados_codanuncio, permalinkClassificados_titulo )
	permalinkClassificados = "classificado.asp?c=" & permalinkClassificados_codanuncio & "&anuncio=" & url2siteSearch(permalinkClassificados_titulo)
End Function

Function permalinkCurriculos( permalinkCurriculos_id, permalinkCurriculos_profissao )
	permalinkCurriculos = "profdetalhes.asp?id=" & permalinkCurriculos_id & "&profissao=" & url2siteSearch(permalinkCurriculos_profissao)
End Function

Function permalinkEventos( permalinkEventos_codevento, permalinkEventos_titulo )
	permalinkEventos = "eventost.asp?codevento=" & permalinkEventos_codevento & "&n=" & url2siteSearch(permalinkEventos_titulo)
End Function

Function permalinkGaleria(permalinkGaleria_id, permalinkGaleria_codfoto, permalinkGaleria_nome)
	permalinkGaleria_url = "noticiaInt_detalhes.asp?id=" & permalinkGaleria_id
	permalinkGaleria_url = permalinkGaleria_url & "&i=" & permalinkGaleria_codfoto
	permalinkGaleria_url = permalinkGaleria_url & "&n=" & url2siteSearch(Trim(permalinkGaleria_nome))
	
	permalinkGaleria = permalinkGaleria_url
End Function

Function permalinkImagens(linkImagens_id, linkImagens_fot, linkImagens_nome)
	url = "noticiaInt_detalhes.asp?id=" & linkImagens_id
	url = url & "&fot=" & linkImagens_fot
	
	If linkImagens_nome <> "" Then
		linkImagens_nomeSplit = Split(linkImagens_nome, " ")
		linkImagens_maxWords = 10
		linkImagens_countWords = 0
		linkImagens_i = 0
		linkImagens_nome = ""
		While (linkImagens_i <= UBound(linkImagens_nomeSplit)) And (linkImagens_countWords < linkImagens_maxWords)
			linkImagens_word = linkImagens_nomeSplit(linkImagens_i)
			If Len(linkImagens_word) > 2 Then linkImagens_countWords = linkImagens_countWords + 1
			linkImagens_nome = linkImagens_nome & " " & linkImagens_word
			linkImagens_i = linkImagens_i + 1
		Wend
		url = url & "&n=" & url2siteSearch(Trim(linkImagens_nome))
	End If
	
	permalinkImagens = url
End Function

Function permalinkImagensModa( permalinkImagensModa_cod, permalinkImagensModa_q )
	permalinkImagensModa = "imagens-moda.asp?q=" & url2siteSearch(permalinkImagensModa_q)
End Function

Function permalinkVideosModa( permalinkVideosModa_cod, permalinkVideosModa_q )
	permalinkVideosModa = "videos-moda.asp?q=" & url2siteSearch(permalinkVideosModa_q)
End Function

Function permalinkNews(permalinkNews_idnews, permalinkNews_nome)
	permalinkNews = "noticiaInt.asp?id=" & permalinkNews_idnews & "&n=" & url2siteSearch(permalinkNews_nome)
End Function

Function permalinkBuscaClassificados(permalinkBuscaClassificados_tipo, permalinkBuscaClassificados_q)
	permalinkBuscaClassificados = "classificados.asp?q="
	If permalinkBuscaClassificados_tipo <> "" Then
		permalinkBuscaClassificados = permalinkBuscaClassificados & permalinkBuscaClassificados_tipo & "-" & url2siteSearch(permalinkBuscaClassificados_q)
	Else
		permalinkBuscaClassificados = permalinkBuscaClassificados & url2siteSearch(permalinkBuscaClassificados_q)
	End If
End Function

Function permalinkBuscanews(permalinkBuscanews_tipo, permalinkBuscanews_q)
	permalinkBuscanews = "buscanews.asp?q="
	If permalinkBuscanews_tipo <> "" Then
		permalinkBuscanews = permalinkBuscanews & permalinkBuscanews_tipo & "-" & url2siteSearch(permalinkBuscanews_q)
	Else
		permalinkBuscanews = permalinkBuscanews & url2siteSearch(permalinkBuscanews_q)
	End If
End Function

Function permalinkBuscavid(permalinkBuscavid_tipo, permalinkBuscavid_q)
	permalinkBuscavid = "buscavid.asp?q="
	If permalinkBuscavid_tipo <> "" Then
		permalinkBuscavid = permalinkBuscavid & permalinkBuscavid_tipo & "-" & url2siteSearch(permalinkBuscavid_q)
	Else
		permalinkBuscavid = permalinkBuscavid & url2siteSearch(permalinkBuscavid_q)
	End If
End Function

Function permalinkVideos( permalinkVideos_idvid, permalinkVideos_nomevid )
	permalinkVideos = "notVideos.asp?id=" & permalinkVideos_idvid & "&n=" & url2siteSearch(permalinkVideos_nomevid)
End Function

Function permalink2LinkFreeze( permalink2LinkFreeze_link )
	permalink2LinkFreeze_link = Replace(permalink2LinkFreeze_link, ".asp?", "~")
	permalink2LinkFreeze_link = Replace(permalink2LinkFreeze_link, "&", "~")
	permalink2LinkFreeze_link = Replace(permalink2LinkFreeze_link, "=", "~")
	permalink2LinkFreeze_link = permalink2LinkFreeze_link & ".htm"
	permalink2LinkFreeze = permalink2LinkFreeze_link
End Function

Function permalinkCatalogoCategoria(permalinkCatalogoCategoria_idcat, permalinkCatalogoCategoria_nome)
	permalinkCatalogoCategoria = "catalogo.asp?cat=" & url2siteSearch(permalinkCatalogoCategoria_nome) & "&idcat=" & permalinkCatalogoCategoria_idcat
End Function
%>