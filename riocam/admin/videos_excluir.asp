<% Option Explicit %>
<!--#include file='includes/incVerifLogado.asp'-->
<!--#include file='includes/incConnectDB.asp'-->
<!--#include file='includes/incVars.asp'-->
<%
Dim query
Dim avideo, i
Dim strRedir, msgErro, excluidos, naoExcluidos
Dim erro, exVStatus

Call conectar()

Response.Buffer = TRUE

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType")

If Request("avideo") <> "" Then 

	msgErro			= ""
	excluidos		= 0
	naoExcluidos	= 0
	avideo			= Split(Request("avideo"), ", ")
	
	For i=0 To UBound(avideo)
	
		ON ERROR RESUME NEXT
		
		conn.BeginTrans
		
		exVStatus = exVideos(Replace(avideo(i), """", ""))
		
		if exVStatus then
			query = "DELETE FROM XCA_modelos_videos WHERE modelo_video = " & Replace(avideo(i), """", "")
			conn.Execute query
		else
			erro = "O arquivo não pode ser excluido, a exclusão não foi efetuada!"
		end if
		
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

Response.Redirect "videos.asp?" & strRedir
%>
