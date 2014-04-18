<%
'/*
' * Valida o CPF
' *
' * @name validCNPJ
' * @param Numero_CNPJ String
' * @return Boolean
' *
'*/

Function validCNPJ(Numero_CNPJ)
    Dim RecebeCNPJ, Numero(14), Soma, Resultado1, Resultado2
    Dim Vb_Valido, Vs_String, xCont, Ch
	
    Vb_Valido = False ' o CNPJ é falso até que prove o contrario
	RecebeCNPJ = Trim(Numero_CNPJ)
	
	If Len(RecebeCNPJ) = 0 Then
		validCNPJ = False
		Exit Function
	End If
	
	'retira todos os caracteres diferentes de numero
	RecebeCNPJ = replaceRegex(RecebeCNPJ, "[^0-9,]", "", True, True)
	
	RecebeCNPJ = Right(RecebeCNPJ, 14)
    If (Len(RecebeCNPJ) = 14) Then
		For xCont=1 to 14
			Numero(xCont) = Cint(Mid(RecebeCNPJ,xCont,1))
		Next

		Soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
        Soma = Soma -(11 * (Int(Soma / 11)))
		
        If (Soma = 0) or (Soma = 1) Then
			Resultado1 = 0
        Else
			Resultado1 = 11 - Soma
		End If
		
        If Resultado1 = Numero(13) Then
            Soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
            Soma = Soma - (11 * (Int(Soma/11)))
            
            If Soma = 0 or Soma = 1 Then
                Resultado2 = 0
            Else
                Resultado2 = 11 - Soma
            End If

            If Resultado2 = Numero(14) Then Vb_Valido = True
        End If
    End If    
	validCNPJ = Vb_Valido
End Function
%>