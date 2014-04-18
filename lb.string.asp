<%
Function alert(alert_msg, alert_url)
	If isArray(alert_msg) Then alert_msg = "* " & Join(alert_msg, "\n* ")

	If alert_msg <> "" Then
		Response.Write "<script type=""text/javascript"">"& vbNewLine
		Response.Write "	alert("""& alert_msg &""");"& vbNewLine
		If alert_url<>"" Then Response.Write "	window.location = """& alert_url &""";"& vbNewLine End If
		Response.Write "</script>"
	End If
End Function

Function serverTranslate(str) 
	serverTranslate = str&""
	serverTranslate = Server.HTMLEncode(serverTranslate)
	serverTranslate = nl2br(serverTranslate)
End Function

'essa aqui permite usar o <b> no default
Function serverTranslateB(str)
	serverTranslateB = nl2br(serverTranslateB)
End Function

'-----------------------------------------------------
'Funcao:	ereg_replace - www.php.net/ereg_replace
'Sinopse:	Substituicao atraves de expressoes regulares
'Parametro:	pattern: Expressao regular
'			replacement: String a substituir
'			str: String de entrada
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function ereg_replace(pattern, replacement, str)
	ereg_replace = replaceRegex(str, pattern, replacement, False, True)
End Function

'-----------------------------------------------------
'Funcao:	eregi_replace - www.php.net/eregi_replace
'Sinopse:	Substituicao atraves de expressoes regulares, insensivel a maiusculas e minusculas
'Parametro:	pattern: Expressao regular
'			replacement: String a substituir
'			str: String de entrada
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function eregi_replace(pattern, replacement, str)
	eregi_replace = replaceRegex(str, pattern, replacement, True, True)
End Function

'-----------------------------------------------------
'Funcao:	replaceRegex
'Sinopse:	Replace com Expressao Regular
'Parametro:	str: String principal
'			regEx: Expressao Regular
'			str2: String para substituicao
'			ignoreCase (Boolean): Ignorar caixa alta/baixa
'			global (Boolean): Substituir todos/somente o primeiro casamento
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function replaceRegex(str, regEx, str2, ignoreCase, global)
	replaceRegex = ""
	str = CStr(str & "")
	If str="" Then Exit Function
	
	replaceRegex = str
	Set objRegExp = New RegExp
	With objRegExp
		.ignoreCase = ignoreCase
		.pattern = regEx
		.global = global
		replaceRegex = .replace(replaceRegex, str2)
	End With
	Set objRegExp = Nothing
End Function

'-----------------------------------------------------
'Funcao:	ereg - www.php.net/ereg
'Sinopse:	Casando expressıes regulares
'Parametro:	pattern: Expressao regular
'			str: String de entrada
'Retorno:	Array
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function ereg(pattern, str)
	ret = Array()
	Set Matches = RegExpTest(pattern, str, False)
	For Each Match In Matches
		ret = InsertInArray(ret, Match.Value)
	Next
	Set Matches = Nothing
	ereg = ret
End Function

'-----------------------------------------------------
'Funcao:	eregi - www.php.net/eregi
'Sinopse:	Casando expressıes regulares, insensivel a maiusculas e minusculas
'Parametro:	pattern: Expressao regular
'			str: String de entrada
'Retorno:	Array
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function eregi(pattern, str)
	ret = Array()
	Set Matches = RegExpTest(pattern, str, True)
	For Each Match In Matches
		ret = InsertInArray(ret, Match.Value)
	Next
	Set Matches = Nothing
	eregi = ret
End Function

'-----------------------------------------------------
'Funcao:	testRegex
'Sinopse:	Testa uma String usando Regex
'Parametro:	str: String principal
'			regEx: Expressao Regular
'			ignoreCase (Boolean): Ignorar caixa alta/baixa
'			global (Boolean): Substituir todos/somente o primeiro casamento
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function testRegex(str, regEx, ignoreCase, global)
	testRegex = False
	str = CStr(str & "")
	If str="" Then Exit Function
	
	Set objRegExp = New RegExp
	With objRegExp
		.ignoreCase = ignoreCase
		.pattern = regEx
		.global = global
		testRegex = .test(str)
	End With
	Set objRegExp = Nothing
End Function

' http://msdn.microsoft.com/en-us/library/yab2dx62(VS.85).aspx
Function RegExpTest(patrn, strng, ignoreCase)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = patrn   ' Set pattern.
   regEx.IgnoreCase = ignoreCase   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set RegExpTest = regEx.Execute(strng)   ' Execute search.
   'RegExpTest = Matches
End Function

'-----------------------------------------------------
'Funcao:	bbToText
'Sinopse:	Converte bbCode para HTML
'Parametro:	str: String para ser convertida
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function bbToText(str)
	str = replaceRegex(str, "\[url\=([^ \]]+)\](.*?)\[\/url\]", "<a href=""$1"">$2</a>", True, True)
	str = replaceRegex(str, "\[url\=(.*?) target=(.*?)\](.*?)\[\/url\]", "<a href=""$1"" rel=""nofollow"" target=""$2"">$3</a>", True, True)
	'str = replaceRegex(str, "\[url\=(.*?) modelo=(.*?) rel=(.*?)\](.*?)\[\/url\]", "<a href=""$1"" class=""$2"" rel=""$3"">$4</a>", True, True)
	str = replaceRegex(str, "\[url\=(.*?) rel=(.*?)\](.*?)\[\/url\]", "<a href=""$1"" rel=""$2"">$3</a>", True, True)

	str = replaceRegex(str, "\[img( alt\=(.*?))?( align\=(.*?))?( modelo\=(.*?))?](.*?)\[\/img\]", "<img src=""$7"" alt=""$2"" align=""$4"" class=""$6"" />", True, True)
	str = replaceRegex(str, "\[swf tam=([0-9]+)x([0-9]+)\](.*?)\[\/swf\]", "<center><object type=""application/x-shockwave-flash"" data=""$3"" width=""$1"" height=""$2""><param name=""movie"" value=""quadro.swf"" /><param name=""wmode"" value=""transparent"" /></object></center>", True, True)
	
	' Necessario tirar a linha abaixo
	str = replaceRegex(str, "\[glossario\=(.*?)\](.*?)\[\/glossario\]", "<a href=""glossario-moda.asp?$1"">$2</a>", True, True)
	
	str = replaceRegex(str, "\[b\](.*?)\[\/b\]", "<strong>$1</strong>", True, True)
	str = replaceRegex(str, "\[clear\]", "<br clear=""all"">", True, True)
	
	' Titulos
	str = replaceRegex(str, "\[h1\](.*?)\[\/h1\]", "<h1>$1</h1>", True, True)
	str = replaceRegex(str, "\[h2\](.*?)\[\/h2\]", "<h2>$1</h2>", True, True)
	str = replaceRegex(str, "\[h3\](.*?)\[\/h3\]", "<h3>$1</h3>", True, True)
	str = replaceRegex(str, "\[h4\](.*?)\[\/h4\]", "<h4>$1</h4>", True, True)
	str = replaceRegex(str, "\[h5\](.*?)\[\/h5\]", "<h5>$1</h5>", True, True)
	str = replaceRegex(str, "\[h6\](.*?)\[\/h6\]", "<h6>$1</h6>", True, True)
	
	' tabelas
	str = replaceRegex(str, "\[table( id=(.*?))?( class=(.*?))?\](.*?)\[\/table\]", "<table id=""$2"" class=""$4"">$5</table>", True, True)
	str = replaceRegex(str, "\[tr\](.*?)\[\/tr\]", "<tr>$1</tr>", True, True)
	str = replaceRegex(str, "\[td\](.*?)\[\/td\]", "<td>$1</td>", True, True)
	
	str = replaceRegex(str, "\[p\]", "<p>", True, True) :: str = replaceRegex(str, "\[\/p\]", "</p>", True, True)
	str = replaceRegex(str, "\[i\]", "<em>", True, True) :: str = replaceRegex(str, "\[\/i\]", "</em>", True, True)
	str = replaceRegex(str, "\[center\]", "<center>", True, True) :: str = replaceRegex(str, "\[\/center\]", "</center>", True, True)
	
	str = replaceRegex(str, "\[span=(.*?)\]", "<span class=""$1"">", True, True)
	str = replaceRegex(str, "\[span\]", "<span>", True, True)
	str = replaceRegex(str, "\[\/span\]", "</span>", True, True)
	
	str = replaceRegex(str, "\[br\]", "<br />", True, True)
	
	str = replaceRegex(str, "\[div=(.*?)\]", "<div class=""$1"">", True, True) :: str = replaceRegex(str, "\[\/div\]", "</div>", True, True)
	
	bbToText = str
End Function

'-----------------------------------------------------
'Funcao:	clearBB
'Sinopse:	Remove as tags BBCode do texto
'Parametro:	str: String para ser convertida
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function clearBB(str)
	str = replaceRegex(str, "\[clear\]", " ", True, True)
	str = replaceRegex(str, "\[\/?url(.*?)\]", " ", True, True)
	str = replaceRegex(str, "\[foto_(direita|centro|esquerda)([0-9]{1,2})\]", " ", True, True)
	str = replaceRegex(replaceRegex(str, "\[noticia=([0-9]+)\]", " ", True, True), "\[/noticia\]", " ", True, True)
	str = replaceRegex(str, "\[tabela_foto\][0-9,]+\[/tabela_foto\]", " ", True, True)
	
	str = replaceRegex(str, "["&vbNewLine&"]+", vbNewLine, True, True)
	str = replaceRegex(str, "[ ]+", " ", True, True)
	str = replaceRegex(str, "\[video[0-9]+\]", " ", True, True)
	str = replaceRegex(str, "\[\/?[bi]\]", "", True, True)
	
	str = replaceRegex(str, "\[video_link=[0-9]+\](.*?)\[/video_link\]", "$1", True, True)
	str = replaceRegex(str, "\[h[1-6]\](.*?)\[\/h[1-6]\]", "$1", True, True)
	
	clearBB = str
End Function

Function URLEncode2(str)
	URLEncode2 = ""
	For x=1 to Len(str)
		URLEncode2 = URLEncode2 & "%" & hex(asc(mid(str, x, 1)))
	Next
End Function

'-----------------------------------------------------
'Funcao:	strRot13
'Sinopse:	Aplica Rot13 para uma String
'Parametro:	xVar: String a ser formatada
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function strRot13(xVar)
	alf1 = "abcdefghijklmnopqrstuvwxyz"
	alf2 = UCase(alf1)
	
	If (Len(alf1) mod 2)<>0 Then
		strRot13 = ""
		Exit Function
	End If
	
	return = ""
	
	For x=1 to Len(xVar)
		char = mid(xVar, x, 1)
		If UCase(char)=char Then
			useChars = alf2
		Else
			useChars = alf1
		End If
		pos = inStr(useChars, char)
		If pos>0 Then		
			pos = pos + (Len(useChars)/2)
			If pos>Len(useChars) Then
				pos = pos - Len(useChars)
			End If			
			return = return & Mid(useChars, pos, 1)
		Else
			return = return & char
		End If
	Next
	
	strRot13 = return
End Function

'-----------------------------------------------------
'Funcao:	geraSenha
'Sinopse:	Alias para geraSenha2, com a diferenca que o retorno e em Caixa Alta
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function geraSenha(max)
	geraSenha = UCase(geraSenha2(max))
End Function

'-----------------------------------------------------
'Funcao:	geraSenha2
'Sinopse:	Gera uma senha
'Parametro:	max: Numero de caracteres da senha gerada
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function geraSenha2(max)
	vars = "abcdefghijklmnopqrstuvwxyz0123456789"
	lenVars = Len(vars)	
	geraSenha2 = ""
	
	For x=1 to max
		y = CInt(Rnd() * (lenVars-1)) + 1
		geraSenha2 = geraSenha2 & Mid(vars, y, 1)
	Next
End Function

'-----------------------------------------------------
'Funcao:	mt_rand
'Sinopse:	Retorna um numero aleatorio
'Parametro:	min: Menor numero
'			max: Maior numero
'Retorno:	Integer
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function mt_rand(min, max)
	mt_rand = min
	If min>max Then Exit Function
	
	mt_rand = Int( ( max - min + 1 ) * Rnd + min )
	
	'mt_rand = Fix(Rnd()*(max-min))
	'mt_rand = mt_rand + min
End Function

'-----------------------------------------------------
'Funcao:	date2String
'Sinopse:	Formata uma data para uma String
'Parametro:	dt: Valor no formato Date
'			format: Formata da saida da String
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function date2String(dt, format)
	Dim mesesNome, mesesName, namesWeekDay
	mesesNome = Array("", "Janeiro", "Fevereiro", "MarÁo", "Abril", "Maio", "Junho", "Junho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
	mesesName = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	namesWeekDay = Array("", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

	' Dias
	format = Replace(format, "%d", Right("0"&Day(dt), 2))
	format = Replace(format, "%D", Left(namesWeekDay(WeekDay(dt)), 3))
	
	' Mes
	format = Replace(format, "%m", Right("0"&Month(dt), 2))
	format = Replace(format, "%M", Left(MesesName(Month(dt)), 3))
	format = Replace(format, "%F", mesesNome(Month(dt)))
	
	' Ano
	format = Replace(format, "%Y", Year(dt))
	format = Replace(format, "%y", Right("0"&Year(dt), 2))
	
	' Tempo
	format = Replace(format, "%H", Right("0"&Hour(dt), 2))
	format = Replace(format, "%i", Right("0"&Minute(dt), 2))
	format = Replace(format, "%s", Right("0"&Second(dt), 2))
	
	date2String = format
End Function

'-----------------------------------------------------
'Funcao:	nl2br
'Sinopse:	Traduz quebra de linha para o HTML
'Parametro:	txt: String a ser convertida
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function nl2br(txt)
	nl2br = Replace(txt, vbNewLine, "<br />" & vbNewLine)
End Function

'-----------------------------------------------------
'Funcao:	SplitX
'Sinopse:	Divide uma string e retorna uma parte
'Parametro:	str: String a ser dividida
'			delimiter: Delimitador para a divisao
'			key: Parte da String a ser retornada
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function SplitX(str, delimiter, key)
	SplitX = ""
	
	SplitX2 = Split(str, delimiter)
	If UBound(SplitX2)>=key Then
		SplitX = SplitX2(key)
	End If
End Function

'-----------------------------------------------------
'Funcao:	UCFirst
'Sinopse:	Primeira letra da String em Caixa Alta
'Parametro:	str: String a ser convertida
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function UCFirst(str)
	If Len(str)>1 Then
		UCFirst = UCase(Left(str, 1)) & Right(str, Len(str)-1)
	Else
		UCFirst = UCase(str)
	End If
End Function

'-----------------------------------------------------
'Funcao:	UCWords
'Sinopse:	Deixa todas as palavras com a primeira em Caixa Alta
'Parametro:	str: String a ser convertida
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function UCWords(str)
	rt = Split(str, " ")
	
	For uncr=0 to UBound(rt)
		rt(uncr) = UCFirst(rt(uncr))
	Next
	UCWords = Join(rt, " ")
End Function

'-----------------------------------------------------
'Funcao:	getFromArray
'Sinopse:	Pega o valor de uma posicao de um array
'Parametro:	arr: Array
'			key: Key a ser recuperado o valor
'			padrao: Valor retornado caso a key nao exista
'Retorno:	String()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function getFromArray(arr, key, padrao)
	getFromArray = padrao
	For i = 0 to UBound(arr)
		If i = Cint(key) Then
			getFromArray = arr(i)
			Exit Function
		End If
	Next
End Function

Sub ShowMsgToUser(ShowMsgToUserMsg, ShowMsgToUserType)
	If ShowMsgToUserType = "ok" Then
		%><div class="msgOK"><%= serverTranslate(ShowMsgToUserMsg) %></div><%
	ElseIf ShowMsgToUserType = "aviso" Then
		%><div class="msgAviso"><%= serverTranslate(ShowMsgToUserMsg) %></div><%
	ElseIf ShowMsgToUserType = "erro" Then
		%><div class="msgError"><%= serverTranslate(ShowMsgToUserMsg) %></div><%
	End If
End Sub

Sub showSQL(showSQL_sql, showSQL_stop)
	If (isMarknet) Or (isChiquitto) Or (Request.QueryString("chiquittodebug") = "1") Then
		If isNull(showSQL_sql) Then
			showSQL_sql_escreve = "NULL"
		Else
			showSQL_sql_escreve = """" & showSQL_sql & """"
		End If
		%><pre style="background:#000000; color:#00FF00; padding:15px;"><%= Server.HTMLEncode(showSQL_sql_escreve) %></pre><%
		Response.Flush()
		If showSQL_stop=1 Then Response.End()
	End If
End Sub

Function isMarknet()
	isMarknet = (Request.ServerVariables("REMOTE_ADDR") = "201.55.145.6") Or (InStr(Request.ServerVariables("HTTP_USER_AGENT"), "PM.isMarknet=On") > 0)
End Function

Function isChiquitto()
	isChiquitto = _
		( _
			(Request.ServerVariables("HTTP_HOST") = "chiquitto:8888") _
			And (Request.ServerVariables("REMOTE_ADDR") = "192.168.192.6") _
		) _
		Or ( InStr(Request.ServerVariables("HTTP_USER_AGENT"), "PM.isChiquitto=On") > 0 )
End Function

'-----------------------------------------------------
'Funcao:	formatNumber2
'Sinopse:	Formata um numero
'Parametro:	nb: Numero a ser formatado
'			dec: Quantidade de algarismos depois da virgula
'Retorno:	String
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function formatNumber2(nb, dec)
	nb1 = Fix(nb)
	lenx = Len(nb)-Len(nb1)-1
	If lenx<0 Then lenx=0
	nb2 = Left(Right(CStr(nb), lenx) & Right(10^dec, dec), dec)
	formatNumber2 = nb1
	If Len(nb2)>0 Then formatNumber2 = formatNumber2 & "." & nb2
End Function

' d = Direcao
' f = tamanho da foto - olhar variavel fGaleria
' fotoGaleria_codgal = codigo da galeria
' fotoGaleria_codfot = codigo da foto
Function fotoGaleria(d, f, fotoGaleria_codgal, fotoGaleria_codfot)
	If f = "I" Then ' Foto destaque
		fotoGaleria_foto1 = Replace(notPrincipalNP, "<idnews>", fotoGaleria_codfot)
		fotoGaleria_foto2 = Replace(notPrincipalN, "<idnews>", fotoGaleria_codfot)
	
		fotoGaleria = Array( _
			Array(fotoGaleria_foto1, "http://img02.portaisdamoda.com.br/" & fotoGaleria_foto1, notPrincipalWP, notPrincipalHP), _
			Array(fotoGaleria_foto2, "http://img02.portaisdamoda.com.br/" & fotoGaleria_foto2, notPrincipalW, notPrincipalH) _
		)
	Else
		fotoGaleria_divGal = 72
	
		If isArray(fotoGaleria_codgal) Then
			fotoGaleria_codgal_ubound = Ubound(fotoGaleria_codgal)
			If fotoGaleria_codgal_ubound > -1 Then
				fotoGaleria_gal = fotoGaleria_codgal(0)(0)
				tipoArr = 0
				fotoGaleria_gal_replace = "g"
				
				For fotoGaleria_i=0 To fotoGaleria_codgal_ubound
					fotoGaleria_codgal(fotoGaleria_i) = Join(fotoGaleria_codgal(fotoGaleria_i), "/")
				Next
				
				fotoGaleria_key = Replace(Replace(fGaleria(tipoArr)(f)(0), "gal<codgaleria>/<codfoto>", ""), ".jpg", "")
				
				fotoGaleria_content = fotoGaleria_key & vbNewLine & _
					Join(fotoGaleria_codgal, ",")
				
				fotoGaleria_fot_replace = MD5(fotoGaleria_content)
				
				lbString_md5_save fotoGaleria_fot_replace, fotoGaleria_content, False
				
				FOT__ = "g/" & fotoGaleria_key & "/" & fotoGaleria_fot_replace & ".jpg"
			Else
				fotoGaleria = Null
				Exit Function
			End If
		Else
			fotoGaleria_gal = Int(fotoGaleria_codgal)
			fotoGaleria_gal_replace = fotoGaleria_codgal
			fotoGaleria_fot_replace = fotoGaleria_codfot
			
			If d="V" Then
				tipoArr = 1
			Else
				tipoArr = 0
			End If
			
			FOT__ = fGaleria(tipoArr)(f)(0)
		End If
		
		FOT__ = Replace(FOT__, "<codgaleria>", fotoGaleria_gal_replace)
		FOT__ = Replace(FOT__, "<codfoto>", fotoGaleria_fot_replace)
		
		If fotoGaleria_gal > fotoGaleria_divGal Then
			fotoGaleria_urlSite = "http://im.portaisdamoda.com.br/"
		Else
			fotoGaleria_urlSite = "http://img02.portaisdamoda.com.br/"
		End If
		
		FOT__ = fotoGaleria_urlSite & FOT__
		'FOT__ = "http://www.portaisdamoda.com.br/_fotnews/" & FOT__
		
		fotoGaleria = Array(FOT__, fGaleria(tipoArr)(f)(1), fGaleria(tipoArr)(f)(2))
	End If
End Function

' Carrega a foto principal das materias
' @param fotoMateria_full
'		Se 0 retorna somente o diretorio
'		Se 1, retorna o PATH completo
'		Se 2, retorna a URL completa
' @param fotoMateria_formato Formato da imagem
'		1 = Formato original
'		2 = Foto G para nFlash2
'		3 = Foto P para nFlash2
' @param fotoMateria_idnews Id na materia
Function fotoMateria(fotoMateria_full, fotoMateria_formato, fotoMateria_idnews)
	If fotoMateria_formato > 0 Then
		fotoMateria_idnews_isArray = False
		If isArray(fotoMateria_idnews) Then
			fotoMateria_idnews_isArray = True
			fotoMateria_idnews = Join(fotoMateria_idnews, ",")
		End If
	
		fotoMateria_tmp = Array( _
			config_fotoMateria(0), _
			Array( Replace(Replace(diretorionot, "\", "/"), "//", "/") ), _
			config_fotoMateria(fotoMateria_formato) _
		)
		
		fotoMateria_tmp(2)(0) = Replace(fotoMateria_tmp(2)(0), "<idnews>", fotoMateria_idnews)
		fotoMateria_tmp(2)(1) = Replace(fotoMateria_tmp(2)(1), "<idnews>", fotoMateria_idnews)
		
		Select Case fotoMateria_full
			Case 1
				fotoMateria_tmp(2)(0) = fotoMateria_tmp(1)(0) & fotoMateria_tmp(2)(0)
				'fotoMateria_tmp(2)(0) = Replace(fotoMateria_tmp(2)(0), "\", "/")
				'fotoMateria_tmp(2)(0) = Replace(fotoMateria_tmp(2)(0), "//", "/")
			Case 2
				fotoMateria_tmp(2)(1) = "http://" & fotoMateria_tmp(0)(0) & "/" & fotoMateria_tmp(2)(1)
				
				If fotoMateria_idnews_isArray Then
					fotoMateria_md5 = md5(fotoMateria_idnews)
					lbString_md5_save fotoMateria_md5, fotoMateria_idnews, False
				
					fotoMateria_tmp(2)(2) = "http://" & fotoMateria_tmp(0)(0) & "/" & _
						Replace(fotoMateria_tmp(2)(2), "<md5>", fotoMateria_md5)
				End If
		End Select
		
		fotoMateria = fotoMateria_tmp
	End If
End Function

lbString_cods_foto = Array()
lbString_cods_foto_addpos = 0
lbString_cods_foto_count = 0
lbString_cods_foto_direction = "H"
Sub lbString_cods_foto_addcod(lbString_cods_foto_addcod_cod)
	lbString_cods_foto = arrayPush(lbString_cods_foto, lbString_cods_foto_addcod_cod)
	lbString_cods_foto_count = lbString_cods_foto_count + 1
End Sub
Function lbString_cods_foto_calcpos()
	If lbString_cods_foto_direction = "V" Then
		lbString_cods_foto_calcpos = "0 -" & (lbString_cods_foto_count * lbString_cods_foto_addpos) & "px"
	Else
		lbString_cods_foto_calcpos = "-" & (lbString_cods_foto_count * lbString_cods_foto_addpos) & "px 0"
	End If
End Function
Sub lbString_cods_foto_reset(lbString_cods_foto_reset_addpos, lbString_cods_foto_reset_direction)
	lbString_cods_foto = Array()
	lbString_cods_foto_count = 0
	lbString_cods_foto_pos = 0
	lbString_cods_foto_addpos = lbString_cods_foto_reset_addpos
	lbString_cods_foto_direction = lbString_cods_foto_reset_direction
End Sub

' Retorna o script para gerar o CSS da imagem
' @param string lbString_cods_foto_writescript_selector		- Seletor CSS
' @param lbString_cods_foto_writescript_f					- Codigo da foto
' @param lbString_cods_foto_writescript_type				- Tipo da imagem
'																= 1 - Imagem da galeria
'																= 2 - Imagem principal da materia
Sub lbString_cods_foto_writescript(lbString_cods_foto_writescript_selector, lbString_cods_foto_writescript_f, lbString_cods_foto_writescript_type)
	Select Case lbString_cods_foto_writescript_type
		Case 1
			lbString_cods_foto_writescript_foto = fotoGaleria("V", lbString_cods_foto_writescript_f, lbString_cods_foto, Null)
			lbString_cods_foto_writescript_img = lbString_cods_foto_writescript_foto(0)
		Case 2
			lbString_cods_foto_writescript_foto = fotoMateria(2, lbString_cods_foto_writescript_f, lbString_cods_foto)
			lbString_cods_foto_writescript_img = lbString_cods_foto_writescript_foto(2)(2)
	End Select
	%><script type="text/javascript">document.write('<style type="text/css"><%= lbString_cods_foto_writescript_selector %> {background-image:url("<%= lbString_cods_foto_writescript_img %>");}</style>');</script><%
End Sub

Function fotoCatalogo(fotoCatalogo_img, fotoCatalogo_id, fotoCatalogo_ext, fotoCatalogo_typert)
	FOT__ = Replace(fotoCatalogo_img, "<id>", fotoCatalogo_id) & "." & fotoCatalogo_ext
	If fotoCatalogo_typert=1 Then FOT__ = "http://" & virtualServerCatalogo & "/" & FOT__
	'If fotoCatalogo_typert=1 Then FOT__ = "http://www.portaisdamoda.com.br/_fotnews.new/" & FOT__
	If fotoCatalogo_typert=2 Then FOT__ = diretorionot & FOT__
	fotoCatalogo = FOT__
End Function

Function fotoVideo( fotoVideo_pos, fotoVideo_id )
	fotoVideo_tmp = fVideo(fotoVideo_pos)
	
	fotoVideo_tmp(0) = Replace(fotoVideo_tmp(0), "<id>", fotoVideo_id)
	fotoVideo_tmp(0) = Array( _
		fotoVideo_tmp(0), _
		"http://img02.portaisdamoda.com.br/" & fotoVideo_tmp(0) _
	)
	fotoVideo = fotoVideo_tmp
End Function

' Carrega a foto do Video
' @param fotoVideo2_full
'		Se 0 retorna somente o diretorio
'		Se 1, retorna o PATH completo
'		Se 2, retorna a URL completa
' @param fotoVideo2_formato Formato da imagem
' @param fotoVideo2_idvid Id do video
Function fotoVideo2(fotoVideo2_full, fotoVideo2_formato, fotoVideo2_idvid)
	If fotoVideo2_formato > 0 Then
		fotoVideo2_tmp = Array( _
			config_fotoVideo(0), _
			Array( Replace(Replace(diretorionot, "\", "/"), "//", "/") ), _
			config_fotoVideo(fotoVideo2_formato) _
		)
		
		fotoVideo2_tmp(2)(0) = Replace(fotoVideo2_tmp(2)(0), "<idvid>", fotoVideo2_idvid)
		fotoVideo2_tmp(2)(1) = Replace(fotoVideo2_tmp(2)(1), "<idvid>", fotoVideo2_idvid)
		
		Select Case fotoVideo2_full
			Case 1
				fotoVideo2_tmp(2)(0) = fotoVideo2_tmp(1)(0) & fotoVideo2_tmp(2)(0)
			Case 2
				fotoVideo2_tmp(2)(1) = "http://" & fotoVideo2_tmp(0)(0) & "/" & fotoVideo2_tmp(2)(1)
		End Select
		
		fotoVideo2 = fotoVideo2_tmp
	End If
End Function

' Carrega a foto do Evento
' @param fotoEvento_full
'		Se 0 retorna somente o diretorio
'		Se 1, retorna o PATH completo
'		Se 2, retorna a URL completa
' @param fotoEvento_formato Formato da imagem
' @param fotoEvento_codevento Id na materia
Function fotoEvento(fotoEvento_full, fotoEvento_formato, fotoEvento_codevento)
	If fotoEvento_formato > 0 Then
		fotoEvento_tmp = Array( _
			config_fotoEvento(0), _
			Array( Replace(Replace(diretorionot, "\", "/"), "//", "/") ), _
			config_fotoEvento(fotoEvento_formato) _
		)
		
		fotoEvento_tmp(2)(0) = Replace(fotoEvento_tmp(2)(0), "[codevento]", fotoEvento_codevento)
		fotoEvento_tmp(2)(1) = Replace(fotoEvento_tmp(2)(1), "[codevento]", fotoEvento_codevento)
		
		fotoEvento_tmp(2)(0) = Replace(fotoEvento_tmp(2)(0), "[codevento_per10]", Fix(fotoEvento_codevento/10))
		fotoEvento_tmp(2)(1) = Replace(fotoEvento_tmp(2)(1), "[codevento_per10]", Fix(fotoEvento_codevento/10))
		
		Select Case fotoEvento_full
			Case 1
				fotoEvento_tmp(2)(0) = fotoEvento_tmp(1)(0) & fotoEvento_tmp(2)(0)
			Case 2
				fotoEvento_tmp(2)(1) = "http://" & fotoEvento_tmp(0)(0) & "/" & fotoEvento_tmp(2)(1)
		End Select
		
		fotoEvento = fotoEvento_tmp
	End If
End Function

' Carrega a foto do Autor
' @param fotoAutor_full
'		Se 0 retorna somente o diretorio
'		Se 1, retorna o PATH completo
'		Se 2, retorna a URL completa
' @param fotoAutor_formato Formato da imagem
' @param fotoAutor_codadmin Id na materia
Function fotoAutor(fotoAutor_full, fotoAutor_formato, fotoAutor_codadmin)
	If fotoAutor_formato > 0 Then
		fotoAutor_tmp = Array( _
			config_fotoAutor(0), _
			Array( Replace(Replace(diretorionot, "\", "/"), "//", "/") ), _
			config_fotoAutor(fotoAutor_formato) _
		)
		
		fotoAutor_tmp(2)(0) = Replace(fotoAutor_tmp(2)(0), "[codadmin]", fotoAutor_codadmin)
		fotoAutor_tmp(2)(1) = Replace(fotoAutor_tmp(2)(1), "[codadmin]", fotoAutor_codadmin)
		
		fotoAutor_tmp(2)(0) = Replace(fotoAutor_tmp(2)(0), "[codadmin_per10]", Fix(fotoAutor_codadmin/10))
		fotoAutor_tmp(2)(1) = Replace(fotoAutor_tmp(2)(1), "[codadmin_per10]", Fix(fotoAutor_codadmin/10))
		
		Select Case fotoAutor_full
			Case 1
				fotoAutor_tmp(2)(0) = fotoAutor_tmp(1)(0) & fotoAutor_tmp(2)(0)
			Case 2
				fotoAutor_tmp(2)(1) = "http://" & fotoAutor_tmp(0)(0) & "/" & fotoAutor_tmp(2)(1)
		End Select
		
		fotoAutor = fotoAutor_tmp
	End If
End Function

Function fotoCla(fotoCla_pos, fotoCla_codanuncio, fotoCla_codfoto, fotoCla_typert)
	fCla_tmp = fCla(fotoCla_pos)
	fCla_tmp(0) = Replace(fCla_tmp(0), "<codanuncio>", fotoCla_codanuncio)
	fCla_tmp(0) = Replace(fCla_tmp(0), "<codfoto>", fotoCla_codfoto)
	fCla_tmp(0) = fCla_tmp(0)
	
	If fotoCla_typert=1 Then fCla_tmp(0) = "http://img02.portaisdamoda.com.br/" & fCla_tmp(0)
	'If fotoCla_typert=1 Then fCla_tmp(0) = "_fotnews/" & fCla_tmp(0)
	fotoCla = fCla_tmp
End Function

Function geraBusca(str, campos)
	Dim a
	a = Array()

	s = Split(str, " ")
	For i=0 To UBound(s)
		If LCase(Right(s(i), 1))="s" Then s(i) = Left(s(i), Len(s(i))-1)
		If Len(s(i))>2 Then
			For i2=0 To UBound(campos)
				a = insertInArray(a, "(" & campos(i2) & " like '%" & s(i) & "%')")
			Next
		End If
	Next
	geraBusca = Join(a, " Or ")
End Function

'-----------------------------------------------------
'Funcao:    RemoveAcentos(Texto)
'Sinopse:    Remove todos os acentos do texto
'Parametro: Texto: Texto a ser transformado
'Retorno: String
'Autor: Gabriel FrÛes - www.codigofonte.com.br
'Adaptado por Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function RemoveAcentos(Texto)
    Dim ComAcentos
    Dim SemAcentos
    Dim Resultado
	Dim Cont
    'Conjunto de Caracteres com acentos
    ComAcentos = "¡Õ”⁄…ƒœ÷‹À¿Ã“Ÿ»√’¬Œ‘€ ·ÌÛ˙È‰Ôˆ¸Î‡ÏÚ˘Ë„ı‚ÓÙ˚Í«Á—Ò"
    'Conjunto de Caracteres sem acentos
    SemAcentos = "AIOUEAIOUEAIOUEAOAIOUEaioueaioueaioueaoaioueCcnn"
    Cont = 0
    Resultado = Texto
    Do While Cont < Len(ComAcentos)
		Cont = Cont + 1
		Resultado = Replace(Resultado, Mid(ComAcentos, Cont, 1), Mid(SemAcentos, Cont, 1))
    Loop
    RemoveAcentos = Resultado
End Function

'-----------------------------------------------------
'Nome: URLDecode(ByVal Texto)
'Tipo: Funcao
'Sinopse: Faz um URL Decode em uma String
'Parametros:
'   Texto: Texto com a URL
'Retorno: String
'Autor: http://www.aspnut.com/reference/encoding.asp
'Adaptado: www.chiquitto.com.br
'-----------------------------------------------------
Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
	
    If IsArray(aSplit) Then
	  If UBound(aSplit) > -1 Then
        sOutput = aSplit(0)
        For URLDecode_I = 0 to UBound(aSplit) - 1
          sOutput = sOutput & _
            Chr("&H" & Left(aSplit(URLDecode_i + 1), 2)) &_
            Right(aSplit(URLDecode_i + 1), Len(aSplit(URLDecode_i + 1)) - 2)
        Next
	  End If
    End If
	
    URLDecode = sOutput
End Function

'-----------------------------------------------------
'Funcao:    text2textSimples
'Sinopse:    Retira todos caracteres especiais, e espacos a mais do texto
'Parametro: Texto: Texto a ser traduzido
'Retorno: String
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function text2textSimples(Texto)
	text2textSimples = RemoveAcentos(Texto)
	text2textSimples = LCase(text2textSimples)
	text2textSimples = eregi_replace("[^0-9a-z/ ]", "", text2textSimples) ' Tudo que n„o seja Letras,Numeros,Barra e EspaÁo
	text2textSimples = eregi_replace("\s+", " ", text2textSimples)
	'text2textSimples = Trimx(text2textSimples)
End Function

'-----------------------------------------------------
'Funcao:    url2siteSearch(Texto)
'Sinopse:    Traduz palavras para as URLs, num formato amigavel a sites de busca
'Parametro: Texto: Texto a ser traduzido
'Retorno: String
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function url2siteSearch(Texto)
	url2siteSearch = Texto
	url2siteSearch = text2textSimples(url2siteSearch)
	url2siteSearch = Server.URLEncode(url2siteSearch)
End Function

'-----------------------------------------------------
'Funcao:    Trimx(Texto)
'Sinopse:    Melhor que Trim, porque retira tambem tabs e newlines
'Parametro: Texto: Texto a ser tratado
'Retorno: String
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function Trimx(Texto)
	Trimx = replaceRegex(Texto, "^\s+|\s+$", "", True, True)
End Function

'-----------------------------------------------------
'Funcao:    UserAgent(Texto)
'Sinopse:    Pega um UserAgent
'Parametro: v: Se for -1 retorna o UserAgent atual, senao retorna um aleatorio
'Retorno: String
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function UserAgent(v)
	If v = -1 Then
		UserAgent = Request.ServerVariables("HTTP_USER_AGENT")
	Else v = 0
		Dim arUA(3)
		arUA(0) = "Mozilla/5.0 (Windows; U; Windows NT 5.1; pt-BR; rv:1.9.0.3) Gecko/2008092417 Firefox/3.0.3" 'FF 3.0.3
		arUA(1) = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; InfoPath.2; .NET CLR 2.0.50727)" 'IE 6.0
		arUA(2) = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US) AppleWebKit/525.13 (KHTML, like Gecko) Chrome/0.2.149.30 Safari/525.13" '0.2.149.30
		arUA(3) = "Opera/9.52 (Windows NT 5.1; U; pt-BR)" ' Opera 9.52
	
		UserAgent = arUA( mt_rand(LBound(arUA), UBound(arUA)) )
	End If
End Function

Function isMail( isMail_email )
	isMail = UBound(eregi("^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$", isMail_email)) = 0
End Function

'-----------------------------------------------------
'Funcao:    die(s)
'Sinopse:    Escreve e mata o programa
'Parametro: s: Texto a ser escrito
'Retorno: Void()
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Sub die(s)
	Response.Write(s)
	Response.End()
End Sub

'-----------------------------------------------------
'Funcao:    writel(s)
'Sinopse:    Escreve e pula linha
'Parametro: s: Texto a ser escrito
'Retorno: Void()
'Autor: Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Sub writel(s)
	Response.Write(s & vbNewLine)
End Sub

'-----------------------------------------------------
'Funcao:    print_r(ar)
'Sinopse:    Escreve um array/string na tela
'Parametro: ar: Variavel a ser escrita
'Retorno: Void()
'Autor: Chiquitto - www.chiquitto.com.br
' http://w3schools.com/vbscript/func_vartype.asp
'-----------------------------------------------------
Sub print_r(ar)
	%><pre style="background:#000000; color:#00FF00; padding:15px;"><% print_r2 ar, 0 %></pre><%
End Sub
Sub print_r2(ar, cx)
	Response.Flush()
	c = ""
	For ic=0 To (cx*4)
		c = c & " "
	Next

	If isArray(ar) Then
		If UBound(ar) = -1 Then
			writel("Array[-1]()")
		Else
			writel("Array[" & UBound(ar) & "](")
			For a=0 To UBound(ar)
				Response.Write(c & "    [" & a & "] = ")
				print_r2 ar(a), cx+1
			Next
			writel(c & ")")
		End If
	ElseIf (varType(ar) = 2) Or (varType(ar) = 3) Then
		writel(ar)
	Else
		writel("""" & ar & """")
	End If
End Sub

Function lbString_md5_path(lbString_md5_path_md5)
	lbString_md5_path = pathTmpMd5 & eregi_replace("([0-9a-z]{8})([0-9a-z]{8})([0-9a-z]{8})([0-9a-z]{8})", "$1/$2/$3/$4", lbString_md5_path_md5) & ".txt"
End Function

Sub lbString_md5_save(lbString_md5_save_md5, lbString_md5_save_content, lbString_md5_save_forceSave)
	lbString_md5_save_path = lbString_md5_path(lbString_md5_save_md5)
	If (lbString_md5_save_forceSave) Or (not lbFile_fso.fileExists(lbString_md5_save_path)) Then
		mkDirRecursive(basename(lbString_md5_save_path))
		saveInFile lbString_md5_save_path, lbString_md5_save_content, True
	End If
End Sub
%>