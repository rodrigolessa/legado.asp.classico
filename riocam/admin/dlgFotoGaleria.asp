<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/geral_gerencia.asp"-->
<!--#include file="includes/incVars.asp"-->
<% 
Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046

DIM query
DIM modelo, modelo_galeria, modelo_galeria_foto, nomeArquivo, tamanhoArquivo
DIM nomeArquivo2, tamanhoArquivo2, tipo, aModeloGaleria
DIM oper, cmd, erro, erro2, fs
DIM msg_erro
DIM s, i, codRemocao
DIM mainDir
DIM aExFoto, exModelo, exMainDir

Call Conectar()

'RECUPERA DADOS
if len(trim(request("oper"))) > 0 then
	oper = uCase(trim(request("oper")))
else
	oper = "L"
end if

if len(trim(request("cmd"))) > 0 then
	cmd = uCase(trim(request("cmd")))
else
	cmd = ""
end if


if len(trim(request("modelo_galeria"))) > 0 then
	modelo_galeria	= trim(request("modelo_galeria"))
	
	sSQL = "SELECT	modelo " & _
			"FROM	XCA_modelos_galerias " & _
			"WHERE	modelo_galeria = " & modelo_galeria & " "
	aModeloGaleria = getArray(sSQL)
	if isArray(aModeloGaleria) then
		modelo	= trim(aModeloGaleria(0,0))
		mainDir	= Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & modelo & "\galeria\"
	end if
else
	modelo	= "0"
end if

if len(trim(request("tipo"))) > 0 then
	tipo = uCase(trim(request("tipo")))
else
	tipo = "0"
end if

'---------------------------------------------------
'SCRIP PARA EXCLUIR UMA FOTO, PEQUENA E GRANDE
if cmd = "X" then
	modelo_galeria_foto	= trim(request("modelo_galeria_foto"))
	
	if exFotosGaleria(modelo_galeria_foto) then
		msg_erro = "Foto excluida com sucesso!"
	else
		msg_erro = "Não foi possível excluír essa foto. Tente novamente!"
	end if
end if
'---------------------------------------------------
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
		<title><%=Application("XCA_APP_TITLE")%></title>
		<link href="includes/box.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/geral.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/formulario.css" type="text/css" rel="stylesheet" media="screen">
		<script src="includes/geral_gerencia.js" language="Javascript"></script>
		<script language="javaScript">
			
			function incFotos(ifModelo, ifGaleria, ifTipo) {
				var w			= 350;
				var h			= 100;
				var parametros	= "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,copyhistory=no,width="+w+",height="+h+",top="+((screen.height / 2) - (h / 2))+",left="+((screen.width / 2) - (w / 2));
				var sURL		= "dlgFotoGaleria_incluir.asp?tipo="+ifTipo+"&modelo="+ifModelo+"&modelo_galeria="+ifGaleria;
				var	incAltFoto	= window.open(sURL, "incFotos", parametros);
					incAltFoto.focus();
			}
			
			function excFotos(nFoto) {
				document.forms[0].oper.value="L";
				document.forms[0].cmd.value="X";
				document.forms[0].modelo_galeria_foto.value=nFoto;
				document.forms[0].submit();
			}

		</script>
</head>


	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">
	
	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="hidden" name="oper">
	<input type="hidden" name="cmd">
	<input type="hidden" name="modelo" value="<%=modelo%>">
	<input type="hidden" name="modelo_galeria" value="<%=modelo_galeria%>">
	<input type="hidden" name="modelo_galeria_foto">
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="dlgPrincipal">

			<!-- DIV PARA ENGLOBAR TODO O CONTEÚDO -->
			<div id="janela">

			<%	if len(trim(msg_erro)) > 0 then	%>
				<div id="msg_erro"><%=msg_erro%></div>
			<%	end if	%>
			
				<input type="button" name="bt_incFotos" value="incluir fotos" class="bt" onClick="incFotos('<%=modelo%>', '<%=modelo_galeria%>', 'G');">
			
				<div id="centro_espaco"></div>

				<div id="conteudo">
					<%=impDlgFoto(modelo, modelo_galeria)%>
				</div>

				<div id="centro_espaco"></div>

			</div>


		</div>
		
	</form>

	</body>
</html>
<%
Function impDlgFoto(idfModelo, idfModeloGaleria)

	DIM aFotos, temp, nFoto, nmArq2, tamArq2, i
	
	temp			= ""
	
	'===== FOTO GRANDE E PEQUENA DE MODELOS PARA EXIBIR NO POPUP DA CADASTRO =====
	sSQL = "SELECT	modelo_galeria_foto, nomeArquivo2, tamanhoArquivo2 " & _ 
			"FROM	XCA_modelos_galerias_fotos " & _ 
			"WHERE	modelo_galeria = " & idfModeloGaleria & " "
			
	aFotos	= getArray(sSQL)
	if isArray(aFotos) then
	
		for i = 0 to uBound(aFotos, 2)
	
			nFoto	= trim(aFotos(0,i))
			nmArq2	= trim(aFotos(1,i))
			tamArq2	= trim(aFotos(2,i))
	
			if len(nmArq2) > 0 then
				temp =	temp &	"<div id=""dlgImgThumb"">" & _
								"<img src="""&Application("XCA_APP_HOME")&"modelos/" & idfModelo & "/galeria/" & nmArq2 & """ width=62 height=72 border=0>" & _
								"<br><a href=""#"" onClick=""javaScript: excFotos('" & nFoto & "');"">excluir</a>" & _
								"</div>"
			else
				temp =	temp &	"<div id=""dlgMensagem"">" & _
								"Nenhuma imagem grande cadastrada!" & _
								"</div>"
			end if
			
		next
		
	else
	
		temp =	temp &	"<div id=""msg_erro"">Arquivos não encontrados!</div>"
		
	end if
	
	impDlgFoto = temp

End Function
%>