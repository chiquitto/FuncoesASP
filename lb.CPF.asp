<%
'/*
' * Valida o CPF
' *
' * @name ValidaCpf
' * @param IdCnd String
' * @return Boolean
' *
'*/
Function ValidaCpf(IdCnd)
	Dim Cpfnumerico
	Dim strCPF 
	Dim strCPFTemp 
	Dim intSoma 
	Dim intResto 
	Dim strDigito 
	Dim intMultiplicador
	intMultiplicador = 10
	Const constIntMultiplicador = 11
	Dim i
	
	ValidaCpf = False
	Cpfnumerico = Trim(IdCnd)
	
	Cpfnumerico = replaceRegex(Cpfnumerico, "[^0-9,]", "", True, True)
	
	If Cpfnumerico="" Then Exit Function
	
	strCPF = Mid(Cpfnumerico, 1, 9)
	For i = 1 To 9 
		intSoma = intSoma +( Cint( Mid(strCPF,i,1)) * intMultiplicador )
		intMultiplicador = intMultiplicador - 1
	Next
	If (intSoma Mod constIntMultiplicador) < 2 Then
		intResto = 0
	Else
		intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
	End If
	strDigito = intResto
	intSoma = 0
	strCPFTemp = strCPF & strDigito
	intMultiplicador = 11
	For i = 1 To 10
		intSoma = intSoma + (CInt (Mid(strCPFTemp,i,1) ) * intMultiplicador)
		intMultiplicador = intMultiplicador - 1
	Next
	If (intSoma Mod constIntMultiplicador) < 2 Then
		intResto = 0
	Else
		intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
	End If
	strDigito = strDigito & intResto
	If strDigito = right(Cpfnumerico,2) Then ValidaCpf = True
End Function
%>