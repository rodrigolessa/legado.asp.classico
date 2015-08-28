<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
Dim query
Dim anoticia, i, nomeArquivo, erro
Dim strRedir, msgErro, excluidos, naoExcluidos
Dim mainDir, movimentoArquivo, fs, imgP, imgI
Dim nomeTabela

nomeTabela = "XCA_noticias"

Call conectar()

'Response.Buffer = TRUE
strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
If Request("anoticia")<>"" Then 
	msgErro=""
	excluidos = 0
	naoExcluidos = 0
'	mainDir = Server.MapPath("imgs\fotos") & "\" & nomeTabela & "\"
' 	Set fs = CreateObject("Scripting.FileSystemObject")
	anoticia = Split(Request("anoticia"), ", ")
	On Error Resume Next
	For i=0 To UBound(anoticia)
		'inicia transação para exclusão
		conn.BeginTrans
		query = "DELETE FROM " & nomeTabela & " WHERE noticia=" & Replace(anoticia(i), """", "")
		conn.Execute query
		If conn.Errors.Count = 0 Then
			conn.CommitTrans
			'exclui arquivos
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
	On Error Goto 0
	Set fs = Nothing
End If
If erro="" Then
	Response.Redirect "noticias.asp?" & strRedir
Else
	Response.Redirect "msg.asp?t=Erro na exclusão&m=" & erro & "&v=noticias.asp?" & Server.URLEncode(strRedir)
End If
%>