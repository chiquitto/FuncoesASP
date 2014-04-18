<%
Function addMask(mask, str)
	addMask = ""
	
	If Len(str) > 0 Then
		j=1
		For addMask_i=1 to Len(mask)
			If Mid(mask, addMask_i, 1)="#" Then
				addMask = addMask & Mid(str, j, 1)
				j = j+1
			Else
				addMask = addMask & Mid(mask, addMask_i, 1)
			End If
			
			If j > Len(str) Then Exit Function
		Next
	End If
End Function

Function addMaskCEP(str)
	str	= replaceRegex(CStr(str&""), "[^0-9,]", "", True, True)
	addMaskCEP = addMask("#####-###", str)
End Function

Function addMaskDDDFone(str)
	str	= replaceRegex(CStr(str&""), "[^0-9,]", "", True, True)
	If Left(str,4)="0800" Then
		addMaskDDDFone = addMask("####-########", str)
	Else
		addMaskDDDFone = addMask("(##) ####-####", str)
	End If
End Function

Function addMaskFone(str)
	str	= replaceRegex(CStr(str&""), "[^0-9,]", "", True, True)
	If Left(str,4)="0800" Then
		addMaskFone = addMask("####-########", str)
	Else
		addMaskFone = addMask("####-####", str)
	End If
End Function

Function addMaskCPF(str)
	str	= replaceRegex(CStr(str&""), "[^0-9,]", "", True, True)
	addMaskCPF = addMask("###.###.###-##", str)
End Function

Function addMaskCNPJ(str)
	str	= replaceRegex(CStr(str&""), "[^0-9,]", "", True, True)
	addMaskCNPJ = addMask("##.###.###/####-##", str)
End Function
%>