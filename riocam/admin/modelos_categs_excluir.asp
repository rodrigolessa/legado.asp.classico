<% Option Explicit %>
<!--#include file='includes/incVerifLogado.asp'-->
<!--#include file='includes/incConnectDB.asp'-->
<!--#include file='includes/incVars.asp'-->
<%
Dim query
Dim amodelo_categ, i
Dim strRedir, msgErro, excluidos, naoExcluidos
Dim erro

Call conectar()

Response.Buffer = TRUE

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType")

If Request("amodelo_categ") <> "" Then 

	msgErro = ""
	excluidos = 0
	naoExcluidos = 0
	amodelo_categ = Split(Request("amodelo_categ"), ", ")
	
	For i=0 To UBound(amodelo_categ)
	
		ON ERROR RESUME NEXT
		
		conn.BeginTrans
		
		query = "DELETE FROM XCA_modelos_categs WHERE modelo_categ=" & Replace(amodelo_categ(i), """", "")
		conn.Execute query
		
		If conn.Errors.Count = 0 Then
			conn.CommitTrans
			excluidos = excluidos + 1
		Else
			conn.RollbackTrans
			erro = "Ocorreu erro na exclusão de um ou mais registro, a exclusão não foi efetuada!"
		End If
		
		ON ERROR GOTO 0
		
	Next
	
	If erro <> "" Then
		Response.Redirect "msg.asp?m=" & Server.URLEncode(erro)
		Response.End
	End If
	
End If

Response.Redirect "modelos_categs.asp?" & strRedir
%>
