<%
connCacheLoaded = True
connCacheCanUseCache = True

getRecordsetDebugMode = False
If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "PM.DebugRecordsetCache=On") Then getRecordsetDebugMode = True
If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "PM.ShowSQL=On") Then getRecordsetShowSQL = True

If Session("clearCache") = True Then
	If getRecordsetDebugMode Then showSQL "# connCacheCanUseCache = False",0
	connCacheCanUseCache = False
	Session("clearCache") = False
End If

getRecordset_CacheSize = 0
getRecordset_CreateCacheIfRecordCountIsZero = True
getRecordset_CursorLocation = 0
getRecordset_CursorType = 0
getRecordset_LockType = 0
getRecordset_PageSize = 0
getRecordset_TypeReset = "0"

Sub getRecordset_resetVars(typeReset)
	typeReset = typeReset & ""
	If typeReset = "0" Then ' somente leitura
		getRecordset_CacheSize = 0
		getRecordset_CreateCacheIfRecordCountIsZero = True
		getRecordset_CursorLocation = 3
		getRecordset_CursorType = 0
		getRecordset_LockType = 1
		getRecordset_PageSize = 0
	ElseIf typeReset = "1" Then ' inclusao/edicao
		getRecordset_CursorLocation = 3
		getRecordset_CursorType = 1
		getRecordset_LockType = 2
	ElseIf typeReset = "2.FALSE" Then ' Mudar getRecordset_CreateCacheIfRecordCountIsZero para false
		getRecordset_CreateCacheIfRecordCountIsZero = False
	End If
End Sub
getRecordset_resetVars(getRecordset_TypeReset)

Function connCacheId2FileName( connCacheId2FileName_id )
	connCacheId2FileName_id = ereg_replace("[^a-zA-Z0-9_\-\/\.]", "", connCacheId2FileName_id)
	connCacheId2FileName_id = ereg_replace("[\/_]+", "/", connCacheId2FileName_id)
	If InStr(connCacheId2FileName_id, "/")=0 Then
		' Se nao foi definido diretorio, entao o diretorio sera gerado automaticamente
		If Len(connCacheId2FileName_id) > 4 Then
			connCacheId2FileName_newId = "auto"
			For connCacheId2FileName_cont=1 To Len(connCacheId2FileName_id) Step 4
				'If connCacheId2FileName_newId <> "" Then connCacheId2FileName_newId = connCacheId2FileName_newId & "/"
				connCacheId2FileName_newId = connCacheId2FileName_newId & "/" & Mid(connCacheId2FileName_id, connCacheId2FileName_cont, 4)
			Next
			connCacheId2FileName_id = connCacheId2FileName_newId
		End If
	End If
	connCacheId2FileName = connCacheId2FileName_id
End Function

connCacheCounterIsValid_valid = Null
connCacheCounterIsValid_file = pathTmpRecordesets & "counter.dat"
Function connCacheCounterIsValid()
	connCacheCounterIsValid_valid = True
	Exit Function
	'If isNull(connCacheCounterIsValid_checked) Then
	'	If validFile2(connCacheCounterIsValid_file, 1) <> 1 Then
	'		connCacheCounterIsValid_valid = True
	'	Else
	'		connCacheCounterIsValid_valid = False
	'	End If
	'End If
	'connCacheCounterIsValid = connCacheCounterIsValid_valid
End Function

'Dim recordsetsId, recordsetsIdFileName
'recordsetsId = Array()
'recordsetsIdFileName = Array()
Function getRecordset(getRecordset_sql, getRecordset_id, getRecordset_valid)
	getRecordset_saida = "#CACHE " & getRecordset_id & vbNewLine
	getRecordset_saida = getRecordset_saida & "#SQL " & getRecordset_sql & vbNewLine
	getRecordset_saida = getRecordset_saida & "#VALIDADE " & getRecordset_valid & vbNewLine
	
	Set getRecordset_rs = Server.CreateObject("ADODB.Recordset")
	
	getRecordset_rs.CursorLocation = getRecordset_CursorLocation
	getRecordset_rs.CursorType = getRecordset_CursorType
	getRecordset_rs.LockType = getRecordset_LockType
	
	If getRecordset_PageSize > 0 Then getRecordset_rs.PageSize = getRecordset_PageSize
	If getRecordset_CacheSize > 0 Then getRecordset_rs.CacheSize = getRecordset_CacheSize
	
	If getRecordsetShowSQL Then showSQL getRecordset_sql, 0
	
	If (getRecordset_id = "") Or (getRecordset_valid = 0) Then
		connOpen
		getRecordset_saida = getRecordset_saida & "#FORCAR CARREGAMENTO DO BD" & vbNewLine
		getRecordset_rs.Open getRecordset_sql, db, 3, 3
	Else
		getRecordset_idFileName = connCacheId2FileName( getRecordset_id )
		getRecordset_file = pathTmpRecordesets & getRecordset_idFileName & ".dat"
		
		If getRecordset_sql = "" Then
			getRecordset_loadFrom = "cache"
			getRecordset_saida = getRecordset_saida & "#FORCAR CARREGAMENTO DO CACHE" & vbNewLine
		Else
			If connCacheCanUseCache Then
				getRecordset_loadFrom = "cache"
				getRecordset_saida = getRecordset_saida & "#VERIFICAR VALIDADADE DO CACHE" & vbNewLine
				
				validCache = validFile2(getRecordset_file, getRecordset_valid)
				If validCache <> 1 Then ' Cache com problemas
					If validCache = 0 Then ' Cache expirado
						getRecordset_saida = getRecordset_saida & "#CACHE EXPIRADO, CRIADO EM " & validFile2_dateLastModified & vbNewLine
						
						If not connCacheCounterIsValid() Then
							getRecordset_saida = getRecordset_saida & "#CARREGAR DO BD (COUNTER EXPIRADO)" & vbNewLine
							getRecordset_loadFrom = "db"
							'saveInFile connCacheCounterIsValid_file, getRecordset_id, True
						Else
							getRecordset_saida = getRecordset_saida & "#CARREGAR DO CACHE (BLOQUEIO DE CRIAÇÃO DE CACHE)" & vbNewLine
						End If
					ElseIf validCache = -1 Then ' Cache não existe
						getRecordset_saida = getRecordset_saida & "#CARREGAR DO BD (CACHE NÃO EXISTE)" & vbNewLine
						getRecordset_loadFrom = "db"
						'saveInFile connCacheCounterIsValid_file, getRecordset_id, True
					End If
				Else
					getRecordset_saida = getRecordset_saida & "#CACHE VALIDO, CRIADO EM " & validFile2_dateLastModified & vbNewLine
				End If
			Else
				getRecordset_saida = getRecordset_saida & "#FORCAR CARREGAMENTO DO BD" & vbNewLine
				getRecordset_loadFrom = "db"
			End If
		End If
		
		If getRecordset_loadFrom = "db" Then
			connOpen
			'connCacheCounterIsValid_valid = True
			
			'Response.Write getRecordset_sql
			getRecordset_rs.Open getRecordset_sql, db, 1, 3
			getRecordset_saida = getRecordset_saida & "#CARREGANDO DO BD" & vbNewLine
			
			If (not getRecordset_rs.Eof) Or (getRecordset_CreateCacheIfRecordCountIsZero) Then
				validCache = validFile2(getRecordset_file, getRecordset_valid)
				If validCache <> -1 Then
					lbFile_fso.deleteFile(getRecordset_file)
				End If
				mkDirRecursive basename(getRecordset_file)
				getRecordset_rs.Save getRecordset_file, adPersistXML
				
				getRecordset_saida = getRecordset_saida & "#CACHE CRIADO (" & getRecordset_file & ")" & vbNewLine
			Else
				getRecordset_saida = getRecordset_saida & "#CACHE NÃO CRIADO. Não existem dados suficientes." & vbNewLine
			End If
		Else
			getRecordset_saida = getRecordset_saida & "#CARREGANDO DO CACHE (" & getRecordset_file & ")" & vbNewLine
			getRecordset_rs.Open getRecordset_file
		End If
	End If
	getRecordset_resetVars(getRecordset_TypeReset)
	
	getRecordset_saida = getRecordset_saida & "#ENCONTRADOS " & getRecordset_rs.RecordCount & " REGISTROS NA CONSULTA" & vbNewLine
	If getRecordsetDebugMode Then showSQL getRecordset_saida, 0
		
	Set getRecordset = getRecordset_rs
End Function

Sub clearRecordsetCached(clearRecordsetCached_id)
	lbFile_loadFso
	clearRecordsetCached_file = pathTmpRecordesets & connCacheId2FileName(clearRecordsetCached_id) & ".dat"
	If getRecordsetDebugMode Then showSQL "#LIMPAR CACHE " & clearRecordsetCached_id, 0
	If lbFile_fso.fileExists(clearRecordsetCached_file) Then
		lbFile_fso.deleteFile clearRecordsetCached_file
		If getRecordsetDebugMode Then showSQL "#LIMPOU CACHE " & clearRecordsetCached_id, 0
	End If
End Sub
%>