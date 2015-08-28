<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
Dim query
Dim usuario, i
Dim msgErro
Dim strRedir, excluidos, naoExcluidos

Call conectar()

Response.Buffer = True

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType")

If Request("ausuario") <> "" Then 
	msgErro = ""
	excluidos = 0
	naoExcluidos = 0 
	usuario = Split(Request("ausuario"), ", ")
		For i=0 To UBound(usuario)	
			ON ERROR RESUME NEXT
			conn.BeginTrans 	'INICIO DA TRANSACAO
			query = "DELETE FROM XCA_usuariosADM WHERE usuario=" & Replace(usuario(i), """", "") & ""
			conn.Execute query
			If conn.Errors.Count = 0 Then		'VERIFICA SE HOUVE ERRO NA TRANSACAO
				conn.CommitTrans				'NAO DEU ERRO ACEITA TODAS AS MODIFICACOES
			Else
				'DEU ERRO, VOLTA TODAS AS MODIFICACOES
				Dim sErr
				For Each sErr In conn.Errors 
					Response.Write sErr.number & ": " & sErr.Description & "<br>"
				Next
				conn.RollbackTrans
				Response.End
			End If
			ON ERROR GOTO 0 'DESLIGA TRATAMENTO DE ERRO
			excluidos = excluidos + 1
		Next
	If msgErro <> "" Then
		Response.Redirect "msg.asp?m=" & Server.URLEncode(msgErro)
		Response.End
	End If
End If		
Response.Redirect "usuariosadm.asp?" & strRedir
%>