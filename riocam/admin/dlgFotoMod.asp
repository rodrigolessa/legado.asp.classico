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
DIM modelo, nomeArquivo, tamanhoArquivo
DIM nomeArquivo2, tamanhoArquivo2, tipo
DIM oper, cmd, erro, erro2, fs
DIM msg_erro
DIM s, i, codRemocao
DIM mainDir

Const nomeTabela = "XCA_modelos"

Call Conectar()

'RECUPERA DADOS
if len(trim(request("oper"))) > 0 then
	oper = uCase(trim(request("oper")))
else
	oper = "L"
end if

if len(trim(request("modelo"))) > 0 then
	modelo	= trim(request("modelo"))
	mainDir	= Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & modelo & "\"
else
	modelo	= "0"
end if

if len(trim(request("tipo"))) > 0 then
	tipo = uCase(trim(request("tipo")))
else
	tipo = "0"
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
		<script language="javaScript">
			function impImgGrande(numMod, strFoto, strTipo) {
				if (strTipo=="G") {
					var strTag = "<img src='<%=Application("XCA_APP_HOME")%>modelos/"+numMod+"/"+strFoto+"' width=275 height=260 border=0>";
				} else {
					var strTag = "<img src='<%=Application("XCA_APP_HOME")%>modelos/"+numMod+"/"+strFoto+"' width=60 height=70 border=0>";
				}
				document.getElementById("imgGrande").innerHTML = strTag;
				return false;
			}
			
			function incAltFoto(tipo) {
				var w			= 350;
				var h			= 100;
				var parametros	= "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width="+w+",height="+h+",top="+((screen.height / 2) - (h / 2))+",left="+((screen.width / 2) - (w / 2));
				var f 			= document.forms[0];
				if (f.modelo.value>0) {
					var sURL		= "dlgFotoMod_incluir.asp?tipo="+tipo+"&modelo="+f.modelo.value;
					var	incAltFoto	= window.open(sURL, "incAltFoto", parametros);
						incAltFoto.focus();
				} else {
					alert("Modelo não encontrada!");
				}
			}
			
			function excFoto(tipo) {
				var w			= 350;
				var h			= 100;
				var parametros	= "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width="+w+",height="+h+",top="+((screen.height / 2) - (h / 2))+",left="+((screen.width / 2) - (w / 2));
				var f 			= document.forms[0];
				if (f.modelo.value>0) {
					var sURL		= "dlgFotoMod_incluir.asp?oper=X&tipo="+tipo+"&modelo="+f.modelo.value;
					var	incAltFoto	= window.open(sURL, "incAltFoto", parametros);
						incAltFoto.focus();
				} else {
					alert("Modelo não encontrada!");
				}
			}

		</script>
</head>


	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">
	
	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="hidden" name="oper">
	<input type="hidden" name="cmd">
	<input type="hidden" name="modelo" value="<%=modelo%>">
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- DIV PARA ENGLOBAR TODO O CONTEÚDO -->
			<div id="janela">

			<%	if len(trim(msg_erro)) > 0 then	%>
				<div id="msg_erro"><%=msg_erro%></div>
			<%	end if	%>

				<div id="conteudo">
					<%=impDlgFoto(modelo)%>
				</div>


			</div>


		</div>
		
	</form>

	</body>
</html>
<%
Function impDlgFoto(idfModelo)

	DIM aModelos, temp, nModelo, nmArq, nmArq2
	
	temp = ""
	
	'===== FOTO GRANDE E PEQUENA DE MODELOS PARA EXIBIR NO POPUP DA CADASTRO =====
	sSQL = "SELECT	modelo, nomeArquivo, tamanhoArquivo, nomeArquivo2, tamanhoArquivo2 " & _ 
			"FROM	XCA_modelos " & _ 
			"WHERE	modelo = " & idfModelo & " "
			
	aModelos	= getArray(sSQL)
	if isArray(aModelos) then
	
			nModelo	= trim(aModelos(0,0))
			nmArq	= trim(aModelos(1,0))
			nmArq2	= trim(aModelos(3,0))
	
			if len(nmArq) > 0 then
				temp =	temp &	"<p>" & _
								"<span id=""imgGrande"" style=""height:260px;"">" & _
								"<img src="""&Application("XCA_APP_HOME")&"modelos/" & nModelo & "/" & nmArq & """ width=275 height=260 border=0>" & _
								"</span>" & _
								"</p>"
								
				temp =	temp &	"<a href=""#"" onClick=""javaScript: impImgGrande('" & nModelo & "', '" & nmArq & "', 'G');"">[ Foto grande &nbsp;&nbsp;]</a> " & _
								"<input type=""button"" name=""excFotoG"" value=""excluir"" class=""bt"" onClick=""javaScript: excFoto('G');""><br><br>"
			else
				temp =	temp &	"<p>" & _
								"<span id=""imgGrande"">" & _
								"Nenhuma imagem grande cadastrada!" & _
								"</span>" & _
								"</p>"
				temp =	temp & "[ Foto grande &nbsp;&nbsp;] <input type=""button"" name=""incFotoG"" value=""incluir"" class=""bt"" onClick=""javaScript: incAltFoto('G');""><br><br>"
			end if

			if len(nmArq2) > 0 then
				temp =	temp &	"<a href=""#"" onClick=""javaScript: impImgGrande('" & nModelo & "', '" & nmArq2 & "', 'P');"">[ Foto pequena ]</a> " & _
								"<input type=""button"" name=""excFotoP"" value=""excluir"" class=""bt"" onClick=""javaScript: excFoto('P');""><br>"
			else
				temp =	temp & "[ Foto pequena ] <input type=""button"" name=""incFotoP"" value=""incluir"" class=""bt"" onClick=""javaScript: incAltFoto('P');""><br>"
			end if

		
	else
	
		temp =	temp &	"<div id=""msg_erro"">Modelo não encontrada!</div>"
		
	end if
	
	impDlgFoto = temp

End Function
%>