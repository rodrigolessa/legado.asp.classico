<%
'************************************************************************************
'VALIDAR EMAIL COM EXPRESSÕES REGULARES.
'************************************************************************************
Function validarMail(myEmail)
DIM isValidE
DIM regEx

isValidE = True
SET regEx = New RegExp

regEx.IgnoreCase = False

regEx.Pattern = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
isValidE = regEx.Test(myEmail)

validarMail = isValidE
End Function
%>