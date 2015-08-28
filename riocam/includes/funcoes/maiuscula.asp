<%
'************************************************************************************
'FUNÇÃO QUE DEIXA SOMENTE AS INICIAIS DA STRING EM MAIÚSCULO.
'************************************************************************************
Function maiuscula(str)
	MeuArray = Split(str," ")
	for i = LBound(MeuArray) to UBound(MeuArray)
		resultado = resultado & UCase(LEFT(MeuArray(i),1)) & LCase(RIGHT(MeuArray(i),Len(MeuArray(i))-1)) & " "
	next
	maiuscula = resultado
End Function
%>