<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
Dim query, erro
Dim acalendario, i, nomeArquivo
Dim strRedir, msgErro, excluidos, naoExcluidos
Dim mainDir, programaArquivo, fs, imgP, imgI
Dim nomeTabela

nomeTabela = "XCA_calendarios"

Call conectar()

'Response.Buffer = TRUE
strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
If Request("acalendario") <> "" Then 
	msgErro = ""
	excluidos = 0
	naoExcluidos = 0
	mainDir = Server.MapPath("imgs\fotos") & "\" & nomeTabela & "\"
 	Set fs = CreateObject("Scripting.FileSystemObject")
	acalendario = split(request("acalendario"), ", ")
	ON ERROR RESUME NEXT
	For i=0 To UBound(acalendario)
		'INICIA TRANSAÇÃO PARA EXCLUSÃO
		conn.BeginTrans
		query = "DELETE FROM " & nomeTabela & " WHERE calendario=" & Replace(acalendario(i), """", "")
		conn.Execute query
		If conn.Errors.Count = 0 Then
			conn.CommitTrans
			excluidos = excluidos + 1
			erro = ""
		Else
			sErros = "Descrição do erro: <br>"
			For Each objErro In conn.Errors
				sErros = sErros & "<b>Erro : </b>" & objErro.number & " - " & objErro.Description & "<br>"
			Next
			conn.RollbackTrans
			erro="Ocorreu um ou mais erros na exclusão de um ou mais registro, a exclusão não foi efetuada!<br><br>" & sErros & "<br>"
		End If
	Next
	ON ERROR GOTO 0
	Set fs = Nothing
End If
If erro="" Then
	Response.Redirect "calendarios.asp?" & strRedir
Else
	Response.Redirect "msg.asp?t=Erro na exclusão&m=" & erro & "&v=calendarios.asp?" & Server.URLEncode(strRedir)
End If
%>
