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

DIM nome, nomeArquivo, tamanho, modelo, modelo_galeria, modelo_galeria_foto
DIM tipo, extArquivo, ultRegistro
DIM oper, cmd, erro, msg_erro, msg_conf
DIM s, i, countFiles
DIM upFile
DIM sErro
DIM objUpload
DIM mainDir, fs, nomeArquivoAtual, tamanhoAtual
DIM aExcFoto, nmExcArq, nmExcArq2, strTipoImg

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
		modelo_galeria	= objUpload.Form.itembykey("modelo_galeria")
		modelo_galeria_foto	= objUpload.Form.itembykey("modelo_galeria_foto")
		tipo			= objUpload.Form.itembykey("tipo")
  End if
Else
	oper			= trim(request("oper"))
	modelo			= trim(request("modelo"))
	modelo_galeria	= trim(request("modelo_galeria"))
	modelo_galeria_foto	= trim(request("modelo_galeria_foto"))
	tipo			= trim(uCase(request("tipo")))
End If

'DIRETORIO RAIZ DAS MODELOS
if len(modelo_galeria) > 0 then
	mainDir = Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & modelo & "\galeria\"
else
	erro = "Código da Modelo ou da Galeria incorreto!"
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
	
	if len(trim(modelo_galeria_foto)) < 1 then
		sSQL = "INSERT INTO XCA_modelos_galerias_fotos (modelo_galeria, status) VALUES (" & modelo_galeria & ", 'I')"
		conn.execute(sSQL)

		modelo_galeria_foto = retUltRegFoto()
	end if
  	    
	If tipo = "G" Then
		nomeArquivo = lCase("gg_mod" & modelo & "_" & modelo_galeria & "_" & modelo_galeria_foto & "." & extArquivo)
		objUpload.Files.Item(0).FileName = nomeArquivo
	Else
		nomeArquivo = lCase("gp_mod" & modelo & "_" & modelo_galeria & "_" & modelo_galeria_foto & "." & extArquivo)
		objUpload.Files.Item(0).FileName = nomeArquivo
	End If

	'PREPARA CAMPO PARA SALVAR
	If erro = "" Then

		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		
		If oper = "I" Then
		
			If countFiles > 0 Then
			
					fileSave()
					
					If fs.FileExists(mainDir & nomeArquivo) Then
						
						if tipo = "G" then
							sSQL = "UPDATE XCA_modelos_galerias_fotos SET nomeArquivo = '" & nomeArquivo & "', tamanhoArquivo = '" & tamanho & "' WHERE modelo_galeria_foto = " & modelo_galeria_foto & " "
						else
							sSQL = "UPDATE XCA_modelos_galerias_fotos SET nomeArquivo2 = '" & nomeArquivo & "', tamanhoArquivo2 = '" & tamanho & "', status = 'A' WHERE modelo_galeria_foto = " & modelo_galeria_foto & " "
						end if
						conn.execute(sSQL)
						
						'msg_conf = " - Arquivo gravado com sucesso: " & nome
						
						if tipo = "G" then
					%>
						<script>
							alert("Arquivo gravado com sucesso!");
						</script>
					<%
							tipo = "P"
						else
					%>
						<script>
							window.opener.location.reload();
							alert("Arquivo gravado com sucesso!");
							close();
						</script>
					<%
						end if
						
					Else
					
						erro = erro & " - Erro na cópia do arquivo, favor tentar novamente!<br>"
						
					End If
					
			Else
			
				erro = erro & " - É preciso selecionar um arquivo!<br>"
				
			End If
			
		End If
		
	End If
	
End If


'DESCRIÇÃO DO TIPO DE IMAGEM, GRANDE OU PEQUENA
if tipo = "G" then
	strTipoImg = "Cadastro da imagem <b>grande</b> (555 x 289 pixels)"
else
	strTipoImg = "Cadastro da imagem <b>pequena</b> (62 x 72 pixels)"
end if


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
	<form action="dlgFotoGaleria_incluir.asp?upFile=true" method="post" enctype="multipart/form-data">
	<input type="hidden" name="oper" value="<%=oper%>">
	<input type="hidden" name="cmd">
	<input type="hidden" name="modelo" value="<%=modelo%>">
	<input type="hidden" name="modelo_galeria" value="<%=modelo_galeria%>">
	<input type="hidden" name="modelo_galeria_foto" value="<%=modelo_galeria_foto%>">
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
		<td colspan=2><%=strTipoImg%></td>
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

Function retUltRegFoto()

	DIM aNFotos
	
	sSQL = "SELECT MAX(modelo_galeria_foto) FROM XCA_modelos_galerias_fotos"
	aNFotos	= getArray(sSQL)
	if isArray(aNFotos) then
		retUltRegFoto = cDbl(aNFotos(0,0))
	end if
			
end function
%>