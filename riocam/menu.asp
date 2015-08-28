<%
	DIM	aMenuCamisas, aMenuEstampas
	DIM strMenuCamisas, strMenuEstampas


%>
	<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
	<div id="hm_menu">

		<!-- BOX DO FORMULÁRIO DE LOGIN -->
		<!--#include file="form_login.asp" -->

		<div id="mn_espaco"></div>
<%
	if verLogado then
%>
		<!-- BOX SLIDE DE FOTOS DAS MODELOS -->
		<!--#include file="slide_foto.asp" -->
		
		<div id="mn_espaco"></div>
<%
	else
%>
		<script language="javaScript">
			criaFlash("banners/banner_assine.swf", "180", "300");
		</script>
		
		<div id="mn_espaco"></div>
<%
	end if
%>
		<%=impUltNoticias("H")%>

	</div>