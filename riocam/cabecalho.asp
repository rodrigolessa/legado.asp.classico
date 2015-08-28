		<!-- CABECALHO DA PÁGINA -->
		<div id="hm_cabecalho">
<%
	if inStr(request.ServerVariables("URL"), "home.asp") then
%>
			<div id="mn_espaco"><img src="imgs/shim.gif" width=1 height=85 border=0></div>
			<div id="mn_cab_sep"><img src="imgs/shim.gif" width=1 height=1 border=0></div>
			<div id="mn_cab_item"><a href="<%=request.ServerVariables("URL")%>?id_lang=2" class="lMenuHorizontal">english <img src="imgs/bandeiras/flag_us.gif" width=15 height=10 border=0></a></div>
			<div id="mn_cab_sep">-</div>
			<div id="mn_cab_item"><a href="<%=request.ServerVariables("URL")%>?id_lang=1" class="lMenuHorizontal">português <img src="imgs/bandeiras/flag_br.gif" width=15 height=10 border=0></a></div>
<%
	end if
%>
		</div>
				
		<!-- BOX DE MENU HORIZONTAL DO CABECALHO DA PÁGINA -->
		<div id="hm_menu_cab">
<%
	if inStr(request.ServerVariables("URL"), "tour.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_06")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="tour.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_06")%></a></div>
<%
	end if
%>
			<div id="mn_cab_sep">|</div>
<%
	if inStr(request.ServerVariables("URL"), "contato.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_04")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="contato.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_04")%></a></div>
<%
	end if
%>
			<div id="mn_cab_sep">|</div>
<%
	if inStr(request.ServerVariables("URL"), "cadastro_modelo.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_05")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="cadastro_modelo.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_05")%></a></div>
<%
	end if
%>
			<div id="mn_cab_sep">|</div>
<%
	if inStr(request.ServerVariables("URL"), "modelos_todas.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_03")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="modelos_todas.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_03")%></a></div>
<%
	end if
%>
			<div id="mn_cab_sep">|</div>
<%
	if inStr(request.ServerVariables("URL"), "assine.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_02")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="assine.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_02")%></a></div>
<%
	end if
%>
			<div id="mn_cab_sep">|</div>
<%
	if inStr(request.ServerVariables("URL"), "home.asp") then
%>
			<div id="mn_cab_item"><%=id_getText("menu_cab_01")%></div>
<%
	else
%>
			<div id="mn_cab_item"><a href="home.asp" class="lMenuHorizontal"><%=id_getText("menu_cab_01")%></a></div>
<%
	end if
%>
		</div>

		<!-- BOX DE FAIXA ABAIXO DO MENU HORIZONTAL -->
		<div id="hm_menu_faixa"></div>