<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
Dim query
Dim usuario, i
Dim erro
Dim strRedir, excluidos, naoExcluidos

Call conectar()

Response.Buffer = True
strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
If Request("ausuario")<>"" Then 
	erro = ""
	excluidos = 0
	naoExcluidos = 0
	usuario = Split(Request("ausuario"), ", ")
	ON ERROR RESUME NEXT
	conn.BeginTrans
	For i=0 To UBound(usuario)	
			'FINALMENTE EXCLUINDO O USUARIO
			query = "DELETE FROM XCA_usuarios WHERE usuario=" & Replace(usuario(i), """", "") & ""
			conn.Execute query
			excluidos = excluidos + 1
	Next
	If conn.Errors.Count = 0 Then
		conn.CommitTrans
		erro = ""
	Else
		conn.RollbackTrans
		erro="Ocorreu erro na exclusão de um ou mais registro, a exclusão não foi efetuada!"
	End If
	ON ERROR GOTO 0
	If erro <> "" Then
		Response.Redirect "msg.asp?m=" & Server.URLEncode(erro)
		Response.End
	End If
End If		
Response.Redirect "usuarios.asp?" & strRedir
%>