<%
lbFileLoaded = True

Set lbFile_fso = Nothing
Sub lbFile_loadFso
	If lbFile_fso is Nothing Then Set lbFile_fso = Server.CreateObject("scripting.fileSystemObject")
End Sub

'On Error Resume Next
Function loadFile(loadFile_pathFile, loadFile_createIfNotExists)
	loadFile = ""
	
	lbFile_loadFso
	If lbFile_fso.fileExists(loadFile_pathFile) Then
		Set loadFile_file = lbFile_fso.getFile(loadFile_pathFile)
		If loadFile_file.Size > 0 Then
			Set loadFile_txt = lbFile_fso.OpenTextFile(loadFile_pathFile, 1)
			loadFile = loadFile_txt.readAll
			loadFile_txt.Close
			Set loadFile_txt = Nothing
		Else
			loadFile = ""
		End If
		Set loadFile_file = Nothing
	ElseIf loadFile_createIfNotExists Then
		loadFile_t = saveInFile(loadFile_pathFile, " ", True)
	End If
End Function

Function saveInFile(saveInFile_pathFile, saveInFile_txt, saveInFile_createIfNotExists)
	saveInFile = False
	
	lbFile_loadFso
	On Error Resume Next
	Set saveInFile_gravar = lbFile_fso.createTextFile(saveInFile_pathFile, saveInFile_createIfNotExists)
	If Err Then
		showSQL "saveInFile - Permissão negada para " & saveInFile_pathFile, 0
		showSQL Err.Description, 1
	End If
	On Error Goto 0
	If lbFile_fso.fileExists(saveInFile_pathFile) Then
		saveInFile_gravar.Write(saveInFile_txt)
		saveInFile_gravar.Close()
		
		If Err=0 Then saveInFile = True
	End If

	Set saveInFile_gravar = Nothing
End Function

Function addInFile(addInFile_pathFile, addInFile_txt, addInFile_createIfNotExists, addInFile_pos)
	addInFile = False
	On Error Goto 0
	addInFile_loaded = loadFile(addInFile_pathFile, False)
	If addInFile_pos = "BOF" Then
		addInFile_loaded = addInFile_txt & addInFile_loaded
	Else
		addInFile_loaded = addInFile_loaded & addInFile_txt
	End If
	addInFile_saved = saveInFile(addInFile_pathFile, addInFile_loaded, addInFile_createIfNotExists)
	If Err=0 Then addInFile = True
End Function

' Retorna True se a data de modificacao do validFile_arquivo estiver dentro do prazo de validFile_validade
validFile_dateLastModified = 0
Function validFile(validFile_arquivo, validFile_validade)
	'showSQL "lb.file.validFile(""" & validFile_arquivo & """, " & validFile_validade & ")", 0
	validFile = False
	
	lbFile_loadFso
	If not lbFile_fso.fileExists(validFile_arquivo) Then Exit Function
	If validFile_validade = -1 Then
		validFile = True
		Exit Function
	End If
	
	Set validFile_file = lbFile_fso.getFile(validFile_arquivo)
	validFile_dateLastModified = validFile_file.dateLastModified
	If dateDiff("s", validFile_file.dateLastModified, Now()) <= validFile_validade Then
		validFile = True
	End If
	Set validFile_file = Nothing
End Function

' Retorna 1 se a data de modificacao do validFile2_arquivo estiver dentro do prazo de validFile2_validade
' Retorna 0 se a data de modificacao do validFile2_arquivo estiver fora do prazo de validFile2_validade
' Retorna -1 se validFile2_arquivo não existe
validFile2_checked = Array()
validFile2_dateLastModified = 0
Function validFile2(validFile2_arquivo, validFile2_validade)
	For validFile2_i = 0 To UBound(validFile2_checked)
		If validFile2_checked(validFile2_i)(0) = validFile2_arquivo Then
			validFile2 = validFile2_checked(validFile2_i)(1)
			Exit Function
		End If
	Next
	
	validFile2 = 0
	lbFile_loadFso
	If not lbFile_fso.fileExists(validFile2_arquivo) Then
		validFile2 = -1
	Else
		If validFile2_validade = -1 Then
			validFile2 = 1
		Else
			Set validFile2_file = lbFile_fso.getFile(validFile2_arquivo)
			validFile2_dateLastModified = validFile2_file.dateLastModified
			If dateDiff("s", validFile2_file.dateLastModified, Now()) <= validFile2_validade Then
				validFile2 = 1
			End If
			Set validFile2_file = Nothing
		End If
	End If
	validFile2_checked = arrayPush(validFile2_checked, Array(validFile2_arquivo, validFile2))
End Function

Function mkDir( mkdir_pathname )
	mkDir = False
	lbFile_loadFso
	If mkdir_pathname = "C:" Then
		Exit Function
	End If
	If lbFile_fso.folderExists(mkdir_pathname) Then
		Exit Function
	End If
	
	showSQL mkdir_pathname, 0
	If not lbFile_fso.folderExists(mkdir_pathname) Then
		'Response.write(mkdir_pathname)
		'Response.end
		lbFile_fso.CreateFolder mkdir_pathname
		mkDir = True
	End If
End Function

Function mkDirRecursive( mkDirRecursive_pathname )
	mkDirRecursive = False
	lbFile_loadFso
	If lbFile_fso.folderExists(mkDirRecursive_pathname) Then
		Exit Function
	End If
	
	If Right(mkDirRecursive_pathname, 1) = "/" Then mkDirRecursive_pathname = Left(mkDirRecursive_pathname, Len(mkDirRecursive_pathname)-1)
	mkDirRecursive_split = Split(Replace(mkDirRecursive_pathname, "\", "/"), "/")
	mkDirRecursive_pathname2 = mkDirRecursive_split(0)
	showSQL mkDirRecursive_pathname2, 0
	mkDirRecursive_iniciar = UBound(split(pathPortais,"\"))+1
	showSQL mkDirRecursive_iniciar,0
	For mkDirRecursive_count=1 To UBound(mkDirRecursive_split)
		mkDirRecursive_pathname2 = mkDirRecursive_pathname2 & "/" & mkDirRecursive_split(mkDirRecursive_count) '& "/"
		If mkDirRecursive_count > mkDirRecursive_iniciar Then
			mkDir mkDirRecursive_pathname2
		End If
	Next
	mkDirRecursive = True
End Function

Function GeneratePath(pFolderPath)
	GeneratePath = False
	If Not objFSO.FolderExists(pFolderPath) Then
		//If GeneratePath(objFSO.GetParentFolderName(pFolderPath)) Then
		//	GeneratePath = True
		//	Call objFSO.CreateFolder(pFolderPath)
		//End If
	Else
		GeneratePath = True
	End If
End Function

Function basename(basename_path)
	basename = ""
	basename_path = Replace(basename_path, "\", "/")
	
	basename_2 = Split(basename_path, "/")
	For basename_i = 0 To UBound(basename_2)-1
		basename = basename & basename_2(basename_i) & "/"
	Next
End Function

Sub lbFile_deleteFile(lbFile_deleteFile_file, lbFile_deleteFile_check)
	lbFile_loadFso
	If lbFile_deleteFile_check Then
		If not lbFile_fso.fileExists(lbFile_deleteFile_file) Then
			Exit Sub
		End If
	End If
	
	lbFile_fso.DeleteFile lbFile_deleteFile_file
End Sub

Sub lbFile_deleteFolder(lbFile_deleteFolder_folder, lbFile_deleteFolder_check)
	lbFile_loadFso
	If lbFile_deleteFolder_check Then
		If not lbFile_fso.FolderExists(lbFile_deleteFolder_folder) Then
			Exit Sub
		End If
	End If
	
	If Right(lbFile_deleteFolder_folder, 1) = "/" Then
		lbFile_deleteFolder_folder = Left(lbFile_deleteFolder_folder, Len(lbFile_deleteFolder_folder)-1)
	End If
	
	lbFile_fso.DeleteFolder lbFile_deleteFolder_folder
End Sub
%>
