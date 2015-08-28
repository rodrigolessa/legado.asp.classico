<%
'************************************************************************************
'FUNวรO PARA PREVINIR SQL INJECTION
'************************************************************************************
function limpaLogin(strL)
	strL = trim(strL)
	strL = lcase(strL)
	strL = replace(strL,"=","")
	strL = replace(strL,"'","")
	strL = replace(strL,"""""","")
	strL = replace(strL," or ","")
	strL = replace(strL," and ","")
	strL = replace(strL,"(","")
	strL = replace(strL,")","")
	strL = replace(strL,"<","[")
	strL = replace(strL,">","]")
	strL = replace(strL,"update","")
	strL = replace(strL,"-shutdown","")
	strL = replace(strL,"--","")
	strL = replace(strL,"'","")
	strL = replace(strL,"#","")
	strL = replace(strL,"$","")
	strL = replace(strL,"%","")
	strL = replace(strL,"จ","")
	strL = replace(strL,"&","")
	strL = replace(strL,"'or'1'='1'","")
	strL = replace(strL,"--","")
	strL = replace(strL,"insert","")
	strL = replace(strL,"drop","")
	strL = replace(strL,"delet","")
	strL = replace(strL,"xp_","")
	strL = replace(strL,"select","")
	strL = replace(strL,"*","")
	limpaLogin = strL
end function
%>