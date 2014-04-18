<%
'Array(
'	QTD_DESTAQUE,
'	QTD_LISTA_ULTIMAS
')
lbDefaultTemplateConfig = Array( _
	Array( _
		2, _
		3 _
	), _
	Array( _
		3, _
		4 _
	), _
	Array( _
		3, _
		1 _
	), _
	Array( _
		3, _
		3 _
	), _
	Array( _
		4, _
		3 _
	) _
)

lbDefaultMaxCod = 3

Function lbDefault_codArray2String(lbDefault_codArray2String_array)
	lbDefault_codArray2String = Join(lbDefault_codArray2String_array, ",")
End Function

Function lbDefault_codString2Array(lbDefault_codString2Array_string)
	lbDefault_codString2Array_string_tmp = Split(lbDefault_codString2Array_string, ",")
	lbDefault_codString2Array_string_tmpUBound = UBound(lbDefault_codString2Array_string_tmp)
	
	lbDefault_codString2Array_r = Array()
	For lbDefault_codString2Array_i=0 To lbDefaultMaxCod-1
		lbDefault_codString2Array_i_padrao = 0
		If lbDefault_codString2Array_i <= lbDefault_codString2Array_string_tmpUBound Then
			If isNumeric(lbDefault_codString2Array_string_tmp(lbDefault_codString2Array_i)) Then
				lbDefault_codString2Array_i_padrao = Int(lbDefault_codString2Array_string_tmp(lbDefault_codString2Array_i))
			End If
		End If
		
		lbDefault_codString2Array_r = arrayPush(lbDefault_codString2Array_r, lbDefault_codString2Array_i_padrao)
	Next
	
	lbDefault_codString2Array = lbDefault_codString2Array_r
End Function

Function lbDefault_destaqueArray2String(lbDefault_destaqueArray2String_template, lbDefault_destaqueArray2String_array)
	lbDefault_destaqueArray2String = Join(lbDefault_destaqueArray2String_array, ",")
End Function

Function lbDefault_destaqueString2Array(lbDefault_destaqueString2Array_template, lbDefault_destaqueString2Array_string)
	lbDefault_destaqueString2Array_tmp = Split(lbDefault_destaqueString2Array_string, ",")
	lbDefault_destaqueString2Array_tmpUBound = UBound(lbDefault_destaqueString2Array_tmp)
	
	lbDefault_destaqueString2Array_max = lbDefault_getDestaqueMax(lbDefault_destaqueString2Array_template) - 1
	
	lbDefault_destaqueString2Array_r = Array()
	For lbDefault_destaqueString2Array_i=0 To lbDefault_destaqueString2Array_max
		lbDefault_destaqueString2Array_i_padrao = 0
		If lbDefault_destaqueString2Array_i <= lbDefault_destaqueString2Array_tmpUBound Then
			If isNumeric(lbDefault_destaqueString2Array_tmp(lbDefault_destaqueString2Array_i)) Then
				lbDefault_destaqueString2Array_i_padrao = Int(lbDefault_destaqueString2Array_tmp(lbDefault_destaqueString2Array_i))
			End If
		End If
		
		lbDefault_destaqueString2Array_r = arrayPush(lbDefault_destaqueString2Array_r, lbDefault_destaqueString2Array_i_padrao)
	Next
	
	lbDefault_destaqueString2Array = lbDefault_destaqueString2Array_r
End Function

Function lbDefault_getDestaqueMax(lbDefault_getDestaqueMax_template)
	lbDefault_getDestaqueMax = lbDefaultTemplateConfig(lbDefault_getDestaqueMax_template)(0)
End Function

Function lbDefault_getListaMax(lbDefault_getListaMax_template)
	lbDefault_getListaMax = lbDefaultTemplateConfig(lbDefault_getListaMax_template)(1)
End Function

' MODULOS
lbDefault_jaFoi = Array(0)

Sub lbDefault_image(lbDefault_image_idnews, lbDefault_image_alt)
	%><img src="<%= virtualStaticVersion %>/spacer.gif" class="img" alt="<%= Server.HTMLEncode(lbDefault_image_alt) %>" style="background-position:<%= lbString_cods_foto_calcpos %>;" width="200" height="130" /><%
	lbString_cods_foto_addcod lbDefault_image_idnews
End Sub

Sub lbDefault_modRs( lbDefault_modRs_rs )
	Select Case lbDefault_modRs_rs("codtemplate")
		Case 0
			lbDefault_modTemplate0 lbDefault_modRs_rs("boxn"), lbDefault_modRs_rs("nome"), lbDefault_modRs_rs("destaque"), lbDefault_modRs_rs("tipo"), lbDefault_modRs_rs("cod")
		Case 1
			lbDefault_modTemplate1 lbDefault_modRs_rs("boxn"), lbDefault_modRs_rs("nome"), lbDefault_modRs_rs("destaque"), lbDefault_modRs_rs("tipo"), lbDefault_modRs_rs("cod")
		Case 2
			lbDefault_modTemplate2 lbDefault_modRs_rs("boxn"), lbDefault_modRs_rs("nome"), lbDefault_modRs_rs("destaque"), lbDefault_modRs_rs("tipo"), lbDefault_modRs_rs("cod")
		Case 3
			lbDefault_modTemplate3 lbDefault_modRs_rs("boxn"), lbDefault_modRs_rs("nome"), lbDefault_modRs_rs("destaque"), lbDefault_modRs_rs("tipo"), lbDefault_modRs_rs("cod")
		Case 4
			lbDefault_modTemplate4 lbDefault_modRs_rs("boxn"), lbDefault_modRs_rs("nome"), lbDefault_modRs_rs("destaque"), lbDefault_modRs_rs("tipo"), lbDefault_modRs_rs("cod")
	End Select
End Sub

Sub lbDefault_modTemplate0(lbDefault_modTemplate0_boxn, lbDefault_modTemplate0_nome, lbDefault_modTemplate0_destaque, lbDefault_modTemplate0_tipo, lbDefault_modTemplate0_cod)
%>
<div id="boxn<%= lbDefault_modTemplate0_boxn %>" class="boxTemplate boxTemplate0">
  <h2><%= Server.HTMLEncode(lbDefault_modTemplate0_nome) %></h2>
  <ul class="lista0"><%
	Set lbDefault_modTemplate0_rsLista = lbDefault_modTemplateLista(lbDefault_modTemplate0_boxn, 0, lbDefault_modTemplate0_tipo, lbDefault_modTemplate0_cod)
	lbDefault_modTemplate0_i=0
	While not lbDefault_modTemplate0_rsLista.Eof
%>
    <li class="pos<%= lbDefault_modTemplate0_i %>"><a href="/<%= permalinkNews(lbDefault_modTemplate0_rsLista("idnews"), lbDefault_modTemplate0_rsLista("nome")) %>">
      <% lbDefault_image lbDefault_modTemplate0_rsLista("idnews"), lbDefault_modTemplate0_rsLista("nome") %>
      <span><%= Server.HTMLEncode(lbDefault_modTemplate0_rsLista("nome")) %></span>
    </a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate0_rsLista("idnews"))
		lbDefault_modTemplate0_i = lbDefault_modTemplate0_i + 1
		
		lbString_cods_foto_addcod lbDefault_modTemplate0_rsLista("idnews")
		lbDefault_modTemplate0_rsLista.moveNext
	Wend
	lbDefault_modTemplate0_rsLista.Close
	Set lbDefault_modTemplate0_rsLista = Nothing
%>
  </ul>
  <ul class="destaque0"><%
	Set lbDefault_modTemplate0_rsDestaque = lbDefault_modTemplateDestaques(lbDefault_modTemplate0_boxn, 0, lbDefault_modTemplate0_destaque)
	While not lbDefault_modTemplate0_rsDestaque.Eof
%>
    <li><a href="/<%= permalinkNews(lbDefault_modTemplate0_rsDestaque("idnews"), lbDefault_modTemplate0_rsDestaque("nome")) %>"><%= Server.HTMLEncode(lbDefault_modTemplate0_rsDestaque("nome")) %></a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate0_rsDestaque("idnews"))
		lbDefault_modTemplate0_rsDestaque.moveNext
	Wend
	lbDefault_modTemplate0_rsDestaque.Close
	Set lbDefault_modTemplate0_rsDestaque = Nothing
%>
  </ul>
</div><%
End Sub

Sub lbDefault_modTemplate1(lbDefault_modTemplate1_boxn, lbDefault_modTemplate1_nome, lbDefault_modTemplate1_destaque, lbDefault_modTemplate1_tipo, lbDefault_modTemplate1_cod)
%>
<div id="boxn<%= lbDefault_modTemplate1_boxn %>" class="boxTemplate boxTemplate1">
  <h2><%= Server.HTMLEncode(lbDefault_modTemplate1_nome) %></h2><%
	Set lbDefault_modTemplate1_rsLista = lbDefault_modTemplateLista(lbDefault_modTemplate1_boxn, 1, lbDefault_modTemplate1_tipo, lbDefault_modTemplate1_cod)
	If not lbDefault_modTemplate1_rsLista.Eof Then
		lbDefault_modTemplate1_link = permalinkNews(lbDefault_modTemplate1_rsLista("idnews"), lbDefault_modTemplate1_rsLista("nome"))
		lbDefault_modTemplate1_foto = fotoMateria(2, 5, lbDefault_modTemplate1_rsLista("idnews"))
%>
  <a href="/<%= lbDefault_modTemplate1_link %>"><img src="<%= lbDefault_modTemplate1_foto(2)(1) %>" alt="<%= Server.HTMLEncode(lbDefault_modTemplate1_rsLista("nome")) %>" class="listaImg" /></a>
  <p class="listaTitle"><a href="/<%= lbDefault_modTemplate1_link %>"><strong><%= Server.HTMLEncode(lbDefault_modTemplate1_rsLista("nome")) %></strong>
    <%= Server.HTMLEncode(lbDefault_modTemplate1_rsLista("nome2")) %></a></p><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate1_rsLista("idnews"))
		lbDefault_modTemplate1_rsLista.moveNext
	End If
  %>
  <ul class="destaque0"><%
	Set lbDefault_modTemplate1_rsDestaque = lbDefault_modTemplateDestaques(lbDefault_modTemplate1_boxn, 1, lbDefault_modTemplate1_destaque)
	While not lbDefault_modTemplate1_rsDestaque.Eof
%>
    <li><a href="/<%= permalinkNews(lbDefault_modTemplate1_rsDestaque("idnews"), lbDefault_modTemplate1_rsDestaque("nome")) %>"><%= Server.HTMLEncode(lbDefault_modTemplate1_rsDestaque("nome")) %></a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate1_rsDestaque("idnews"))
		lbDefault_modTemplate1_rsDestaque.moveNext
	Wend
	lbDefault_modTemplate1_rsDestaque.Close
	Set lbDefault_modTemplate1_rsDestaque = Nothing
%>
  </ul>
  <ul class="lista0"><%
	lbDefault_modTemplate1_i=0
	While not lbDefault_modTemplate1_rsLista.Eof
%>
    <li class="pos<%= lbDefault_modTemplate1_i %>"><a href="/<%= permalinkNews(lbDefault_modTemplate1_rsLista("idnews"), lbDefault_modTemplate1_rsLista("nome")) %>">
      <% lbDefault_image lbDefault_modTemplate1_rsLista("idnews"), lbDefault_modTemplate1_rsLista("nome") %>
      <span><%= Server.HTMLEncode(lbDefault_modTemplate1_rsLista("nome")) %></span>
    </a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate1_rsLista("idnews"))
		lbDefault_modTemplate1_i = lbDefault_modTemplate1_i + 1
		
		lbString_cods_foto_addcod lbDefault_modTemplate1_rsLista("idnews")
		lbDefault_modTemplate1_rsLista.moveNext
	Wend
	lbDefault_modTemplate1_rsLista.Close
	Set lbDefault_modTemplate1_rsLista = Nothing
%>
  </ul>
</div><%
End Sub

Sub lbDefault_modTemplate2(lbDefault_modTemplate2_boxn, lbDefault_modTemplate2_nome, lbDefault_modTemplate2_destaque, lbDefault_modTemplate2_tipo, lbDefault_modTemplate2_cod)
%>
<div id="boxn<%= lbDefault_modTemplate2_boxn %>" class="boxTemplate boxTemplate2">
  <h2><%= Server.HTMLEncode(lbDefault_modTemplate2_nome) %></h2><%
	Set lbDefault_modTemplate2_rsLista = lbDefault_modTemplateLista(lbDefault_modTemplate2_boxn, 2, lbDefault_modTemplate2_tipo, lbDefault_modTemplate2_cod)
	If not lbDefault_modTemplate2_rsLista.Eof Then
		lbDefault_modTemplate2_link = permalinkNews(lbDefault_modTemplate2_rsLista("idnews"), lbDefault_modTemplate2_rsLista("nome"))
		lbDefault_modTemplate2_foto = fotoMateria(2, 5, lbDefault_modTemplate2_rsLista("idnews"))
%>
  <a href="/<%= lbDefault_modTemplate2_link %>"><img src="<%= lbDefault_modTemplate2_foto(2)(1) %>" alt="<%= Server.HTMLEncode(lbDefault_modTemplate2_rsLista("nome")) %>" class="listaImg" /></a>
  <p class="listaTitle"><a href="/<%= lbDefault_modTemplate2_link %>"><strong><%= Server.HTMLEncode(lbDefault_modTemplate2_rsLista("nome")) %></strong>
    <%= Server.HTMLEncode(lbDefault_modTemplate2_rsLista("nome2")) %></a></p><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate2_rsLista("idnews"))
		lbDefault_modTemplate2_rsLista.moveNext
	End If
	lbDefault_modTemplate2_rsLista.Close
	Set lbDefault_modTemplate2_rsLista = Nothing
  %>
  <ul class="destaque0"><%
	Set lbDefault_modTemplate2_rsDestaque = lbDefault_modTemplateDestaques(lbDefault_modTemplate2_boxn, 2, lbDefault_modTemplate2_destaque)
	While not lbDefault_modTemplate2_rsDestaque.Eof
%>
    <li><a href="/<%= permalinkNews(lbDefault_modTemplate2_rsDestaque("idnews"), lbDefault_modTemplate2_rsDestaque("nome")) %>"><%= Server.HTMLEncode(lbDefault_modTemplate2_rsDestaque("nome")) %></a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate2_rsDestaque("idnews"))
		lbDefault_modTemplate2_rsDestaque.moveNext
	Wend
	lbDefault_modTemplate2_rsDestaque.Close
	Set lbDefault_modTemplate2_rsDestaque = Nothing
%>
  </ul>
</div><%
End Sub

Sub lbDefault_modTemplate3(lbDefault_modTemplate3_boxn, lbDefault_modTemplate3_nome, lbDefault_modTemplate3_destaque, lbDefault_modTemplate3_tipo, lbDefault_modTemplate3_cod)
%>
<div id="boxn<%= lbDefault_modTemplate3_boxn %>" class="boxTemplate boxTemplate3">
  <h2><%= Server.HTMLEncode(lbDefault_modTemplate3_nome) %></h2><%
	Set lbDefault_modTemplate3_rsLista = lbDefault_modTemplateLista(lbDefault_modTemplate3_boxn, 3, lbDefault_modTemplate3_tipo, lbDefault_modTemplate3_cod)
	While not lbDefault_modTemplate3_rsLista.Eof
%>
  <p class="listaTitle"><a href="/<%= permalinkNews(lbDefault_modTemplate3_rsLista("idnews"), lbDefault_modTemplate3_rsLista("nome")) %>"><strong><%= Server.HTMLEncode(lbDefault_modTemplate3_rsLista("nome")) %></strong>
    <%= Server.HTMLEncode(lbDefault_modTemplate3_rsLista("nome2")) %></a></p><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate3_rsLista("idnews"))
		lbDefault_modTemplate3_rsLista.moveNext
	Wend
	lbDefault_modTemplate3_rsLista.Close
	Set lbDefault_modTemplate3_rsLista = Nothing
  %>
  <ul class="destaque0"><%
	Set lbDefault_modTemplate3_rsDestaque = lbDefault_modTemplateDestaques(lbDefault_modTemplate3_boxn, 3, lbDefault_modTemplate3_destaque)
	While not lbDefault_modTemplate3_rsDestaque.Eof
%>
    <li><a href="/<%= permalinkNews(lbDefault_modTemplate3_rsDestaque("idnews"), lbDefault_modTemplate3_rsDestaque("nome")) %>"><%= Server.HTMLEncode(lbDefault_modTemplate3_rsDestaque("nome")) %></a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate3_rsDestaque("idnews"))
		lbDefault_modTemplate3_rsDestaque.moveNext
	Wend
	lbDefault_modTemplate3_rsDestaque.Close
	Set lbDefault_modTemplate3_rsDestaque = Nothing
%>
  </ul>
</div><%
End Sub

Sub lbDefault_modTemplate4(lbDefault_modTemplate4_boxn, lbDefault_modTemplate4_nome, lbDefault_modTemplate4_destaque, lbDefault_modTemplate4_tipo, lbDefault_modTemplate4_cod)
%>
<div id="boxn<%= lbDefault_modTemplate4_boxn %>" class="boxTemplate boxTemplate4">
  <h2><%= Server.HTMLEncode(lbDefault_modTemplate4_nome) %></h2><%
	Set lbDefault_modTemplate4_rsDestaque = lbDefault_modTemplateDestaques(lbDefault_modTemplate4_boxn, 4, lbDefault_modTemplate4_destaque)
	If not lbDefault_modTemplate4_rsDestaque.Eof Then
		lbDefault_modTemplate4_link = permalinkNews(lbDefault_modTemplate4_rsDestaque("idnews"), lbDefault_modTemplate4_rsDestaque("nome"))
%>
  <p class="listaTitle"><a href="/<%= lbDefault_modTemplate4_link %>"><strong><%= Server.HTMLEncode(lbDefault_modTemplate4_rsDestaque("nome")) %></strong>
    <%= Server.HTMLEncode(lbDefault_modTemplate4_rsDestaque("nome2")) %>. <%= Server.HTMLEncode(Left(clearBB(lbDefault_modTemplate4_rsDestaque("materia")),200)) %>...</a></p><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate4_rsDestaque("idnews"))
		lbDefault_modTemplate4_rsDestaque.moveNext
	End If
  %>
  <ul class="destaque1"><%
	While not lbDefault_modTemplate4_rsDestaque.Eof
%>
    <li><a href="/<%= permalinkNews(lbDefault_modTemplate4_rsDestaque("idnews"), lbDefault_modTemplate4_rsDestaque("nome")) %>"><%= Server.HTMLEncode(lbDefault_modTemplate4_rsDestaque("nome")) %></a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate4_rsDestaque("idnews"))
		lbDefault_modTemplate4_rsDestaque.moveNext
	Wend
	lbDefault_modTemplate4_rsDestaque.Close
	Set lbDefault_modTemplate4_rsDestaque = Nothing
%>
  </ul>
  <ul class="lista0"><%
	lbDefault_modTemplate4_i=0
	Set lbDefault_modTemplate4_rsLista = lbDefault_modTemplateLista(lbDefault_modTemplate4_boxn, 4, lbDefault_modTemplate4_tipo, lbDefault_modTemplate4_cod)
	While not lbDefault_modTemplate4_rsLista.Eof
%>
    <li class="pos<%= lbDefault_modTemplate4_i %>"><a href="/<%= permalinkNews(lbDefault_modTemplate4_rsLista("idnews"), lbDefault_modTemplate4_rsLista("nome")) %>">
      <% lbDefault_image lbDefault_modTemplate4_rsLista("idnews"), lbDefault_modTemplate4_rsLista("nome") %>
      <span><%= Server.HTMLEncode(lbDefault_modTemplate4_rsLista("nome")) %></span>
    </a></li><%
		lbDefault_modTemplateJafoi(lbDefault_modTemplate4_rsLista("idnews"))
		lbDefault_modTemplate4_i = lbDefault_modTemplate4_i + 1
		
		lbDefault_modTemplate4_rsLista.moveNext
	Wend
	lbDefault_modTemplate4_rsLista.Close
	Set lbDefault_modTemplate4_rsLista = Nothing
%>
  </ul>
</div><%
End Sub

Function lbDefault_modTemplateDestaques( lbDefault_modTemplateDestaques_boxn, lbDefault_modTemplateDestaques_template, lbDefault_modTemplateDestaques_cods )
	lbDefault_modTemplateDestaques_sql = "Select Distinct Top " & lbDefault_getDestaqueMax(lbDefault_modTemplateDestaques_template) & vbNewLine & _
		"	vn.idnews, vn.nome, vn.nome2, vn.materia " & vbNewLine & _
		"From vNews vn " & vbNewLine & _
		"Where (vn.idnews In (" & lbDefault_modTemplateDestaques_cods & ")) " & vbNewLine & _
		"Order By vn.idnews Desc"
	Set lbDefault_modTemplateDestaques = getRecordset(lbDefault_modTemplateDestaques_sql, "lbDefault/2011091901/boxn" & lbDefault_modTemplateDestaques_boxn & "/destaque", 3600)
End Function

Sub lbDefault_modTemplateJafoi(lbDefault_modTemplateJafoi_idnews)
	lbDefault_jaFoi = arrayPush(lbDefault_jaFoi, lbDefault_modTemplateJafoi_idnews)
End Sub

Function lbDefault_modTemplateLista( lbDefault_modTemplateLista_boxn, lbDefault_modTemplateLista_template, lbDefault_modTemplateLista_tipo, lbDefault_modTemplateLista_cods )

	Select Case lbDefault_modTemplateLista_tipo
		Case "C"
			lbDefault_modTemplateLista_sql = "Select Top " & lbDefault_getListaMax(lbDefault_modTemplateLista_template) & " " & vbNewLine & _
			"	vn.idnews, vn.nome, vn.nome2 " & vbNewLine & _
			"From vNews vn " & vbNewLine & _
			"Inner Join trel_news_cat tr " & vbNewLine & _
			"	On tr.idnews = vn.idnews " & vbNewLine & _
			"Where (tr.idcat In (" & lbDefault_modTemplateLista_cods & ")) And (vn.idnews Not In (" & Join(lbDefault_jaFoi, ",") & ")) " & vbNewLine & _
			"Order by vn.idnews Desc"
		Case Else
			lbDefault_modTemplateLista_tipo = "T"
			
			lbDefault_modTemplateLista_sql = "Select Distinct Top " & lbDefault_getListaMax(lbDefault_modTemplateLista_template) & " " & vbNewLine & _
			"	vn.idnews, vn.nome, vn.nome2 " & vbNewLine & _
			"From vNews vn " & vbNewLine & _
			"Inner Join tnews_tags tt " & vbNewLine & _
			"	On tt.idnews = vn.idnews " & vbNewLine & _
			"Where (tt.idtag In (" & lbDefault_modTemplateLista_cods & ")) And (vn.idnews Not In (" & Join(lbDefault_jaFoi, ",") & ")) " & vbNewLine & _
			"Order By vn.idnews Desc"
	End Select
	
	Set lbDefault_modTemplateLista = getRecordset(lbDefault_modTemplateLista_sql, "lbDefault/2011091901/boxn" & lbDefault_modTemplateLista_boxn & "/lista" & lbDefault_modTemplateLista_tipo, 3600)
End Function
%>