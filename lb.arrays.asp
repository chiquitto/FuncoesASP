<%
lbArraysLoaded = True

'alias de arrayPush
Function insertInArray(varArray, row)
	insertInArray = arrayPush(varArray, row)
End Function

'-----------------------------------------------------
'Funcao:	inArray
'Sinopse:	Procura dentro do array, e retorna a posição
'Parametro:	inArray_arr : Array a ser filtrado
'			inArray_str : String a ser pesquisada
'Retorno:	Int
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function inArray(inArray_arr, inArray_str)
	inArray = -1
	For inArray_i=0 to UBound(inArray_arr)
		If inArray_arr(inArray_i)=inArray_str Then
			inArray = inArray_i
			inArray_i = UBound(inArray_arr)+1
		End If
	Next
End Function

'-----------------------------------------------------
'Funcao:	arrayUnique - www.php.net/array_unique
'Sinopse:	Remove os valores duplicados de um array e retorna o novo array
'Parametro:	arrayUnique_array : Array a ser filtrado
'Retorno:	Array()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function arrayUnique ( arrayUnique_array )
	Dim arrayUnique_new
	arrayUnique_new = Array()
	
	For arrayUnique_i = 0 To uBound(arrayUnique_array)
		If inArray(arrayUnique_new, arrayUnique_array(arrayUnique_i)) = -1 Then
			arrayUnique_new = arrayPush(arrayUnique_new, arrayUnique_array(arrayUnique_i))
		End If
	Next
	arrayUnique = arrayUnique_new
End Function

'-----------------------------------------------------
'Funcao:	arrayPush - www.php.net/array_push
'Sinopse:	Adiciona um ou mais elementos no final de um array e retorna o novo array
'Parametro:	arrayPush_array : Array de entrada
'			arrayPush_add : Valor a adicionar ao final do arrayPush_array
'Retorno:	Array()
'Autor:		Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function arrayPush ( arrayPush_array, arrayPush_add )
	ReDim preserve arrayPush_array(UBound(arrayPush_array) +1)
	arrayPush_array(UBound(arrayPush_array)) = arrayPush_add
	arrayPush = arrayPush_array
End Function

'-----------------------------------------------------
'Funcao:	arrayUnshift - www.php.net/array_unshift
'Sinopse:	Adiciona um ou mais elementos no inicio de um array e retorna o novo array
'Parametro:	arrayUnshift_array : Array de entrada
'			arrayUnshift_add : Valor a adicionar ao inicio do arrayUnshift_array
'Retorno:	Array()
'Autor:		http://www.aspfree.com/c/a/Code-Examples/Creating-Useful-Array-Functions/1/
'Alterado:	Chiquitto - www.chiquitto.com.br
'-----------------------------------------------------
Function arrayUnshift ( arrayUnshift_array, arrayUnshift_add )
	ReDim Preserve arrayUnshift_array(UBound(arrayUnshift_array) + 1)
	For arrayUnshift_i = UBound(arrayUnshift_array) To 1 Step -1
		arrayUnshift_array(arrayUnshift_i) = arrayUnshift_array(arrayUnshift_i - 1)
	Next
	arrayUnshift_array(0) = arrayUnshift_add
	arrayUnshift = arrayUnshift_array
End Function

'-----------------------------------------------------
'Funcao:	BubbleSort
'Sinopse:	Ordena um array usando o algoritmo BubbleSort
'Parametro:	matriz : Array de entrada
'Retorno:	Array()
'Autor:		Rubens Farias
'-----------------------------------------------------
Function BubbleSort( matriz )
	dim i, j, aux
	For i = 0 To UBound(matriz)
		For j = 0 To UBound(matriz)
			If( matriz(i) < matriz(j) ) Then
				aux = matriz(j)
				matriz(j) = matriz(i)
				matriz(i) = aux
			End If
		Next
	Next
	BubbleSort = matriz
End Function
%>