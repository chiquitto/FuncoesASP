<%
'Funcao para o divisao arredondar sempre pra mais
Function Ceil(nb1, nb2)
	if (nb1/nb2) = (nb1\nb2) Then
		Ceil = nb1\nb2
	else
		Ceil = (nb1\nb2) + 1	
	end if
End function
%>