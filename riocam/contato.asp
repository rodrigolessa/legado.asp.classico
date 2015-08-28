<% option explicit %>
<!--#include file="includes/incConnectDB.asp" -->
<!--#include file="includes/geral.asp" -->
<!--#Include file="includes/incEnviaEmail.asp"-->
<%
'Response.Buffer = False
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir
	DIM login, senha
	DIM i
	
	'***** VARIAVEIS E-MAIL ******
	DIM	usuario, nome, email, bairro, cidade
	DIM	estado, pais, telefone
	DIM nomeFrom, emailFrom, nomeTo, emailTo
	DIM	strSubject, strBody, strMsg
	
	Const maxNome		= 60
	Const maxEmail		= 60
	Const maxBairro		= 30
	Const maxCidade		= 30
	Const maxTelefone	= 15

	
	'***** ABRE A CONEXÃO *****
	Call conectar()

	
	'***** RECUPERA AS VARIAVEIS ******
	if len(trim(request("oper"))) > 0 then
		oper = cStr(uCase(trim(request("oper"))))
	else
		oper = "L"
	end if
	
	if len(trim(request("cmd"))) > 0 then
		cmd = cStr(uCase(trim(request("cmd"))))
	else
		cmd = ""
	end if
	
	msg_erro	= cStr(trim(request("msg_erro")))


	nome		= trim(request("nome"))
	email		= trim(lCase(request("email")))
	cidade		= trim(request("cidade"))
	estado		= trim(Ucase(request("estado")))
	telefone	= trim(request("telefone"))
	strMsg		= trim(request("strMsg"))
	
	strSubject	= trim(request("strSubject"))
	


	If cmd = "E" Then
		'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
		If Len(nome) < 1 or Len(nome) > maxNome Then
			erro = erro & " - " & id_getText("msg_erro_02") & " <br>"
		end if
		If Len(email) < 1 or Len(email) > maxEmail Then
			erro = erro & " - " & id_getText("msg_erro_03") & " <br>"
		end if
		If Len(strMsg) < 1 Then
			erro = erro & " - " & id_getText("msg_erro_04") & " <br>"
		end if

		'PREPARA CAMPO PARA SALVAR
		If erro = "" Then
			nome	= replace(nome, "'", "''")
			email	= replace(email, "'", "")
			strMsg	= replace(strMsg, VbCrLf, "<br>")
			
			strBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & VbCrLf & _
						"<html>" & VbCrLf & _
						"<head>" & VbCrLf & _
						"<title>" & Application("XCA_APP_TITLE") & "</title>" & VbCrLf & _
						"</head>" & VbCrLf & _
						"<body bgcolor=""#FFFFFF"">" & VbCrLf & _
						"	Administrador,<br><br>" & vbCrLf & _
						"	<b><a href='mailto:"&email&"'>"&nome&"</a></b>, " & _
						"	enviou a mensagem abaixo em "&date&".<br>" & _
						"	E-mail: "&email&"<br>Telefone: "&telefone&"<br>Cidade: "&cidade&"<br>Estado: "&estado&"<br>" & _
						"	<blockquote>"&strMsg&"</blockquote><br>" & _
						"</body>" & vbCrLf & _
						"</html>"
			nomeTo		= Application("XCA_APP_OWNER")
			emailTo		= Application("XCA_APP_EMAIL")
			nomeFrom	= nome 
			emailFrom	= email 
			
			enviaEmail nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, "CDO"
			
			nome = "" : email = "" : msg = "" : oper = "V"
		End If
	End If
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html;charset=iso-8859-1">
		<title><%=Application("XCA_APP_TITLE")%></title>
		<link href="includes/box.css"			type="text/css" rel="stylesheet" media="screen">
		<link href="includes/geral.css"			type="text/css" rel="stylesheet" media="screen">
		<link href="includes/formulario.css"	type="text/css" rel="stylesheet" media="screen">
		<script src="includes/funcoes.js" language="Javascript"></script>
		
		<script language="JavaScript">
		//Variáveis Globais

		//Funções gerais
			function abreJanela(a,w,h) {
				var b
				b = b + "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars="+s+",resizable=no,copyhistory=no";
				b = b + ",width="+w+",height="+h+",top="+((screen.height/2)-(h/2))+",left="+((screen.width/2)-(w/2));
				window.open(a,"_blank",b);
			}

			function validarEmail(valor) {
				if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(valor)){
					return true
				} else {
					return false;
				}
			}

			function enviaFaleConosco() {
				var f0 = document.forms[0];
				var fok = false;
				if (validarEmail(f0.email.value) == true) fok = true;
				if (fok == true) {	
		//			f0.oper.value="F";	
					f0.cmd.value="E";
					f0.action="contato.asp";
					f0.submit();
				} else {
					alert ("<%=id_getText("msg_erro_05")%>");
				}
			}
		</script>
	</head>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">

	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="Hidden" name="oper"		value="">
	<input type="Hidden" name="cmd"			value="">
	<input type="Hidden" name="strSubject"	value="Contato de visitante">


		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- BOX PÁGINA PARA CONTER TODO O LAYOUT -->
			<div id="pagina">

				<!-- CABECALHO DA PÁGINA -->
				<!--#include file="cabecalho.asp" -->
				
<%
				if len(msg_erro) > 0 then
%>
					<div id="msg_erro"><%=msg_erro%></div>
<%
				end if
%>

				<!-- BOX DE CONTEÚDO DA PÁGINA -->
				<div id="hm_conteudo">

					<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
					<div id="hm_conteudo_centro" style="width:400px;margin-left:30px;margin-top:10px;">
						
						<p class="inter_titulo"><%=id_getText("contato_txt_01")%></p>
<%
			if oper = "V" then
%>
						<p class="inter_texto"><%=id_getText("contato_txt_02")%></p>
<%
			else
%>
						<p class="inter_texto">
							<%=id_getText("contato_txt_03")%><br>
							<input type="text" name="nome" size=60 maxlength="<%=maxNome%>" value="<%=nome%>" class=int_cx>
						</p>

						<p class="inter_texto">
							<%=id_getText("contato_txt_04")%><br>
							<input type="text" name="email" size=60 maxlength="<%=maxEmail%>" value="<%=email%>" class=int_cx>
						</p>

						<p class="inter_texto">
							<%=id_getText("contato_txt_05")%><br>
							<input type="text" name="cidade" size=60 maxlength="<%=maxCidade%>" value="<%=cidade%>" class=int_cx>
						</p>

						<p class="inter_texto">
							<%=id_getText("contato_txt_06")%><br>
							<select name="estado" class=int_sel>
								<option value="" <% If estado="" Then Response.Write " SELECTED "%>>...</option>
								<option value="AC" <% If estado="AC" Then Response.Write " SELECTED "%>>AC</option>
								<option value="AL" <% If estado="AL" Then Response.Write " SELECTED "%>>AL</option>
								<option value="AM" <% If estado="AM" Then Response.Write " SELECTED "%>>AM</option>
								<option value="AP" <% If estado="AP" Then Response.Write " SELECTED "%>>AP</option>
								<option value="BA" <% If estado="BA" Then Response.Write " SELECTED "%>>BA</option>
								<option value="CE" <% If estado="CE" Then Response.Write " SELECTED "%>>CE</option>
								<option value="DF" <% If estado="DF" Then Response.Write " SELECTED "%>>DF</option>
								<option value="ES" <% If estado="ES" Then Response.Write " SELECTED "%>>ES</option>
								<option value="GO" <% If estado="GO" Then Response.Write " SELECTED "%>>GO</option>
								<option value="MA" <% If estado="MA" Then Response.Write " SELECTED "%>>MA</option>
								<option value="MG" <% If estado="MG" Then Response.Write " SELECTED "%>>MG</option>
								<option value="MS" <% If estado="MS" Then Response.Write " SELECTED "%>>MS</option>
								<option value="MT" <% If estado="MT" Then Response.Write " SELECTED "%>>MT</option>
								<option value="PA" <% If estado="PA" Then Response.Write " SELECTED "%>>PA</option>
								<option value="PB" <% If estado="PB" Then Response.Write " SELECTED "%>>PB</option>
								<option value="PE" <% If estado="PE" Then Response.Write " SELECTED "%>>PE</option>
								<option value="PI" <% If estado="PI" Then Response.Write " SELECTED "%>>PI</option>
								<option value="PR" <% If estado="PR" Then Response.Write " SELECTED "%>>PR</option>
								<option value="RJ" <% If estado="RJ" Then Response.Write " SELECTED "%>>RJ</option>
								<option value="RN" <% If estado="RN" Then Response.Write " SELECTED "%>>RN</option>
								<option value="RO" <% If estado="RO" Then Response.Write " SELECTED "%>>RO</option>
								<option value="RR" <% If estado="RR" Then Response.Write " SELECTED "%>>RR</option>
								<option value="RS" <% If estado="RS" Then Response.Write " SELECTED "%>>RS</option>
								<option value="SC" <% If estado="SC" Then Response.Write " SELECTED "%>>SC</option>
								<option value="SE" <% If estado="SE" Then Response.Write " SELECTED "%>>SE</option>
								<option value="SP" <% If estado="SP" Then Response.Write " SELECTED "%>>SP</option>
								<option value="TO" <% If estado="TO" Then Response.Write " SELECTED "%>>TO</option>
							</select><%=id_getText("contato_txt_07")%><br>
						</p>

						<p class="inter_texto">
							<%=id_getText("contato_txt_08")%><br>
							<textarea name="strMsg" rows=8 cols=60 class="int_tx"><%=strMsg%></textarea>
						</p>

						<p class="inter_texto">
							<nobr>
								<input type="Reset" name="btnLimpar" value="<%=id_getText("contato_txt_09")%>" class="int_bt">&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="Button" name="btnEnviar" value="<%=id_getText("contato_txt_10")%>" onClick="JavaScript: enviaFaleConosco();" class="int_bt">
							</nobr>
						</p>							
<%
			end if
%>
						<div id="centro_espaco"></div>

					</div>
					<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->

				</div>
				<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->
				
				<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
				<!--#include file="menu.asp" -->

				<!-- BOX DE RODAPÉ DA PÁGINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX PÁGINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>