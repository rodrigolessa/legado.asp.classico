<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/geral_gerencia.asp"-->
<!--#include file="includes/incVars.asp"-->
<!--#include file="includes/incUpload.asp"-->
<% 
Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046

DIM nome, nomeArquivo, tamanho, modelo, modelo_video
DIM tipo, extArquivo, aModeloVideo
DIM oper, cmd, erro, msg_erro, msg_conf
DIM s, i, countFiles
DIM upFile
DIM sErro
DIM objUpload
DIM mainDir, fs, nomeArquivoAtual, tamanhoAtual
DIM aExcFoto, nmExcArq

CONST maxNome = 100
CONST maxNomeArquivo = 255

Call Conectar()

upFile = Request.QueryString("upfile")
If upFile Then
	SET objUpload	= New clsUpload
	If objUpload.Form.Count > 0 then
		countFiles		= objUpload.Files.Count
		oper			= objUpload.Form.itembykey("oper")
		cmd				= objUpload.Form.itembykey("cmd")
		modelo			= objUpload.Form.itembykey("modelo")
		modelo_video	= objUpload.Form.itembykey("modelo_video")
		tipo			= objUpload.Form.itembykey("tipo")
  End if
Else
	oper			= trim(request("oper"))
	modelo			= trim(request("modelo"))
	modelo_video	= trim(request("modelo_video"))
	tipo			= trim(uCase(request("tipo")))
End If


if len(trim(modelo_video)) > 0 then
	sSQL = "SELECT	modelo " & _
			"FROM	XCA_modelos_videos " & _
			"WHERE	modelo_video = " & modelo_video & " "
	aModeloVideo = getArray(sSQL)
	if isArray(aModeloVideo) then
		modelo	= trim(aModeloVideo(0,0))
		'DIRETORIO RAIZ DAS MODELOS
		mainDir	= Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & modelo & "\video\"
	else
		erro = "Código da Modelo ou do Vídeo incorreto!"
	end if
end if


'VERIFICA SE É A PRIMEIRA VEZ
If cmd = "S" Then

	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	nome				= objUpload.Form.itembykey("nomeArquivo")
	nomeArquivoAtual	= objUpload.Form.itembykey("nomeArquivoAtual")
	If objUpload.Files.Count > 0 Then
		nomeArquivo	= lCase(objUpload.Files.Item(0).FileName)
		tamanho		= objUpload.Files.Item(0).Size
		extArquivo	= getExt(nomeArquivo)
	End If

	if extArquivo <> "JPG" and extArquivo <> "GIF" then
		erro = "Arquivo informado (" & extArquivo & ") não é do tipo esperado, deve ser 'jpg' ou 'gif'!"
	end if
  	    
	nomeArquivo = lCase("vf_mod" & modelo & "_" & modelo_video & "." & extArquivo)
	objUpload.Files.Item(0).FileName = nomeArquivo

	'PREPARA CAMPO PARA SALVAR
	If erro = "" Then

		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		
		If oper = "I" Then
		
			If countFiles > 0 Then
			
					fileSave()
					
					If fs.FileExists(mainDir & nomeArquivo) Then
						
						sSQL = "UPDATE XCA_modelos_videos SET nomeArquivo2 = '" & nomeArquivo & "', tamanhoArquivo2 = '" & tamanho & "' WHERE modelo_video = " & modelo_video & " "
						conn.execute(sSQL)
						
						msg_conf = " - Arquivo gravado com sucesso: " & nome
						
					%>
						<script>
							alert("Arquivo gravado com sucesso!");
							close();
						</script>
					<%
						
					Else
					
						erro = erro & " - Erro na cópia do arquivo, favor tentar novamente!<br>"
						
					End If
					
			Else
			
				erro = erro & " - É preciso selecionar um arquivo!<br>"
				
			End If
			
		End If
		
	End If
	
End If

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
		<title><%=Application("XCA_APP_TITLE")%></title>
		<link href="includes/box.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/geral.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/formulario.css" type="text/css" rel="stylesheet" media="screen">
		<script src="includes/geral_gerencia.js" language="Javascript"></script>
		<script>
			function salvar() {
				document.forms[0].oper.value="I";
				document.forms[0].cmd.value="S";
				document.forms[0].submit();
			}
		</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="280" border="0" cellpadding="0" cellspacing="10" class="texto2">
	<form action="dlgVideoFoto_incluir.asp?upFile=true" method="post" enctype="multipart/form-data">
	<input type="hidden" name="oper" value="<%=oper%>">
	<input type="hidden" name="cmd">
	<input type="hidden" name="modelo" value="<%=modelo%>">
	<input type="hidden" name="modelo_video" value="<%=modelo_video%>">
	<input type="hidden" name="tipo" value="<%=tipo%>">

	<% If len(trim(erro)) > 0 Then %>
	<tr>
		<td class="erro" valign="top">ERRO:</td>
		<td class="erro"><%=erro%></td>
	</tr>
	<% End If %>
	
	<% If len(trim(msg_conf)) > 0 Then %>
	<tr>
		<td class="msg_conf" colspan=2><%=msg_conf%></td>
	</tr>
	<% End If %>
	
	<tr>
		<td colspan=2>Cadastro da imagem <b>pequena</b> (62 x 72 pixels)</td>
	</tr>
	
	<tr>
		<td><b>arquivo:</b></td>
		<td align="right">
			<input type="file" name="nomeArquivo" class="cx" style="width:250px;">
			<input type="hidden" name="nomeArquivoAtual" value="<%=nomeArquivoAtual%>">
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right"><input type="button" name="btnSalvar" value="Salvar arquivo" onClick="JavaScript:salvar();" class="bt"></td>
	</tr>
	</form>
</table>

</body>

</html>

<%
Sub fileSave()

	DIM erro
	
	ON ERROR RESUME NEXT
	
 	objUpload.Files.Item(0).Save mainDir
 	
	If err.number <> 0 Then
	
		erro =	"Ocorreu o seguinte erro ao tentar salvar o arquivo: <br><br>" & _
				"<b>Erro : </b>" & Err.number & "<br>" & _
				"Descrição do erro: " & Err.Description & "<br>"
		
	End If
	
	ON ERROR GOTO 0
	
End Sub
%>