<% option explicit %>
<!--#include file="includes/geral.asp" -->
<%
'Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir 
	DIM login, senha
	DIM i

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<meta http-equiv="Content-Language" content="pt-br">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title><%=Application("XCA_APP_TITLE")%></title>
	<META NAME="keywords" CONTENT="modelos, sexo, fotos, erotismo, sexual, pornô, foto, sacanagem, x-rated, chat, adult, adulto, fotolog, sexlog, sexy, baby, babies, teens, gay, gays, lesbian, lésbica, lésbicas, lesbians, girls, xxx, pornografia, amadoras, voyeur, peladas, amadores, amador, sexo amador, caseiras, câmeras, câmeras ao vivo, cam">
	<META NAME="description" CONTENT="Fotos e vídeos de Sexo.">
	<META NAME="rating" CONTENT="Sex">
	<META NAME="revisit-after" CONTENT="7 days">
	<META NAME="objecttype" CONTENT="Homepage">
	<META HTTP-EQUIV="reply-to" CONTENT="contato@3xcam.com.br">
	<META HTTP-EQUIV="copyright" CONTENT="3XCam © 2007">
	<link href="includes/box.css"			type="text/css" rel="stylesheet" media="screen">
	<link href="includes/geral.css"			type="text/css" rel="stylesheet" media="screen">
	<link href="includes/formulario.css"	type="text/css" rel="stylesheet" media="screen">
	<script src="includes/funcoes.js" language="Javascript"></script>
</head>

<body link="#ED4141" vlink="#ED4141" alink="#ED4141" bgcolor="#FFFFFF" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">

<div align="center">

<table width="600" cellspacing="0" cellpadding="0" style="padding: 1px">
		<tr>
			<td align="center">&nbsp;</td>
			<td align="center" colspan="3">&nbsp;</td>
			<td align="center">&nbsp;</td>

		</tr>
		<tr>
			<td align="center">&nbsp;</td>
			<td align="center" colspan="3">
			<img border="0" src="imgs/logo.gif" width="229" height="120"><br>&nbsp;</td>
			<td align="center">&nbsp;</td>
		</tr>
		<tr>

			<td bgcolor="#333333">&nbsp;</td>
			<td colspan="3" bgcolor="#333333" align="center">&nbsp;</td>
			<td bgcolor="#333333">&nbsp;</td>
		</tr>
		<tr>
			<td bgcolor="#333333">&nbsp;</td>
			<td colspan="3" bgcolor="#333333" align="center"><font face="Verdana" color="#E31E16" size="2"><b><font size="3"><%=id_getText("inicio_txt_01")%></font></b><%=id_getText("inicio_txt_02")%></font></td>

			<td bgcolor="#333333">&nbsp;</td>
		</tr>
		<tr>
			<td bgcolor="#333333">&nbsp;</td>
			<td colspan="3" bgcolor="#333333">
			<p align="justify"><font face="Verdana" size="1"><%=id_getText("inicio_txt_03")%></font>
			</td>
			<td bgcolor="#333333">&nbsp;</td>
		</tr>
		<tr>
			<td bgcolor="#333333">&nbsp;</td>
			<td colspan="3" bgcolor="#333333">&nbsp;</td>
			<td bgcolor="#333333">&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td colspan="3">&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td align="center">
			<b>
			<font face="Arial" size="3" color="#666666"><%=id_getText("inicio_txt_06")%></font><br>
			<font color="#ED4141" face="Arial" size="5">(<a href="home.asp"><%=id_getText("inicio_txt_07")%></a>)</font>
			</b>
			</td>
			<td>&nbsp;</td>
			<td align="center">
			<b>
			<font face="Arial" size="3" color="#666666"><%=id_getText("inicio_txt_04")%></font><br>
			<font color="#ED4141" face="Arial" size="4">(<a href="about:blank"><%=id_getText("inicio_txt_05")%></a>)</font>
			</b>
			</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td colspan="3">&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td align="center"><a href="<%=request.ServerVariables("URL")%>?id_lang=1" class="lMenuHorizontal">português <img src="imgs/bandeiras/flag_br.gif" width=15 height=10 border=0></a></td>
			<td>&nbsp;</td>
			<td align="center"><a href="<%=request.ServerVariables("URL")%>?id_lang=2" class="lMenuHorizontal">english <img src="imgs/bandeiras/flag_us.gif" width=15 height=10 border=0></a></td>
			<td>&nbsp;</td>
		</tr>
	</table>
	
</div>

</body>

</html>