<% Option Explicit %>
<%
Dim dia, mes, ano, primeira
Dim diaCal, mesCal, anoCal
Dim diaHoje, mesHoje, anoHoje
Dim anoIni, anoFim, nSemanas
Dim i
Dim data
Dim fData

If Request("oper") <> "" Then
	dia = Cint(Request("dia"))
	mes = Cint(Request("mes"))
	ano = Cint(Request("ano"))
Else
	If Session("data") = "" Then
		dia = Day(Now()) : mes = Month(Now()) : ano = Year(now())
	Else
		If Not IsDate(Session("data")) Then Session("data") = Date()
		dia = Day(Session("data")) : mes = Month(Session("data")) : ano = Year(Session("data"))
	End If
End If
data = dia & "/" & mes & "/" & ano
Session("data") = data
fData = Request("fData")
'Caso nao reconheca o campo para o qual devese colocar a data o campo fica vazio
diaCal=1
mesCal=mes
anoCal=ano
diaHoje = Cint(Day(Now))
mesHoje = Cint(Month(Now))
anoHoje = Cint(Year(Now))
primeira = True

%>
<html>
<head>
	<title>Calendário</title>
	<link rel="STYLESHEET" type="text/css" href="includes/calendario.css">
	<link rel="STYLESHEET" type="text/css" href="includes/geral.css">
<script>
	function putDate() {
		var f = document.forms[0]
		if ('<%=fData%>'!='') {
			var d = f.dia.value; var m = f.mes.value; var a = f.ano.value;
			if (d.length==1) { d = "0" + d }; if (m.length==1) {m = "0" + m};
			var sData = d + "/" + m + "/" + a;
			if (!opener.closed) { 
				opener.document.forms[0].<%=fData%>.value=sData;
				opener.document.forms[0].<%=fData%>.focus();
			}; 
			close();
		}
	}

	function nav(intDia) {
		document.forms[0].dia.value=intDia;
//		document.forms[0].submit();
		putDate();
	}

	function navMes(intMes, intAno) {
		document.forms[0].mes.value=intMes;
		document.forms[0].ano.value=intAno;
		document.forms[0].submit();
	}

	function trocaMes(intMes) {
		document.forms[0].mes.value=intMes;
		//document.forms[0].submit();
	}

</script>

</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
<form method="post" name="formData" action="<%= Request.ServerVariables("SCRIPT_NAME") %>">
	<input type="hidden" name="fData" value="<%=fData%>">
	<tr> 
		<td> 
			<table width="100%" border="0" cellspacing="1" cellpadding="0">
					<tr> 
						<td> 
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="bg2">
								<tr> 
									<td>&nbsp;<a class="setas" href="javascript: void(0)" onClick="navMes('<%=Month(DateAdd("m", -1, data))%>', '<%=Year(DateAdd("m", -1, data))%>'); return false;">«</a></td>
									<td align="center" class="texto2"><b><%= MonthName(mes) & " " & Year(data) %></b></td>
									<td align="right"><a class="setas" href="javascript: void(0)" onClick="navMes('<%=Month(DateAdd("m", 1, data))%>', '<%=Year(DateAdd("m", 1, data))%>'); return false;" >»</a>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr> 
						<td> 
							<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff">
								<tr> 
									<td class="textoP" align="center"> D </td>
									<td class="textoP" align="center"> S </td>
									<td class="textoP" align="center"> T </td>
									<td class="textoP" align="center"> Q </td>
									<td class="textoP" align="center"> Q </td>
									<td class="textoP" align="center"> S </td>
									<td class="textoP" align="center"> S </td>
								</tr>
<%	nSemanas = 0
		While IsDate(diaCal & "/" & mesCal & "/" & anoCal) 
			nSemanas = nSemanas + 1
			%>
								<tr bgcolor=#FFFFFF>
<% 		For i=1 to 7 %>
									<td align=center <% If diaCal=dia Then Response.Write "bgcolor='#c0c0c0'"%>>
<%			If IsDate(diaCal & "/" & mesCal & "/" & anoCal) Then 
					If primeira Then 
						If WeekDay(1 & "/" & mesCal & "/" & anoCal) = i Then %>
										<a href="javascript: void(0)" onClick="nav('<%=diaCal%>'); return false;">
										<span class='
<% 										If WeekDay(diaCal & "/" & mesCal & "/" & anoCal)=1 Then 
												Response.Write "diaF" 
											Else
												Response.Write "diaN"
											End If %>
										'><%=diaCal%></span>
										</a>
<%						primeira = False
							diaCal = diaCal + 1
						Else %>
							&nbsp;
<%					End If
					Else %>
							<a href="javascript: void(0)" onClick="nav('<%=diaCal%>'); return false;">
							<span class='
<% 						If WeekDay(diaCal & "/" & mesCal & "/" & anoCal)=1 Then 
								Response.Write "diaF" 
							Else 
								Response.Write "diaN"
							End If
							%>
							'><%=diaCal%></span></a>
<%						diaCal = diaCal + 1
					End If
				Else %>
										&nbsp;
<% 			End If%>
									</td>
<%		Next %>
								</tr>

<%	Wend %>
							</table>
						</td>
					</tr>
				<tr> 
					<td class="bg1"> 
						<table width="100%" cellspacing="0" cellpadding="4" border="0">
							<input type="Hidden" name="oper" value="R">
							<input type="Hidden" name="dia" value="<%=dia%>">
							<tr> 
								<td align="center">
									<select name="mes" class="textoP" onChange="document.forms[0].submit()">
										<option value=1 <% If mes=1 Then Response.Write " SELECTED " %>>Janeiro
										<option value=2 <% If mes=2 Then Response.Write " SELECTED " %>>Fevereiro
										<option value=3 <% If mes=3 Then Response.Write " SELECTED " %>>Março
										<option value=4 <% If mes=4 Then Response.Write " SELECTED " %>>Abril
										<option value=5 <% If mes=5 Then Response.Write " SELECTED " %>>Maio
										<option value=6 <% If mes=6 Then Response.Write " SELECTED " %>>Junho
										<option value=7 <% If mes=7 Then Response.Write " SELECTED " %>>Julho
										<option value=8 <% If mes=8 Then Response.Write " SELECTED " %>>Agosto
										<option value=9 <% If mes=9 Then Response.Write " SELECTED " %>>Setembro
										<option value=10 <% If mes=10 Then Response.Write " SELECTED " %>>Outubro
										<option value=11 <% If mes=11 Then Response.Write " SELECTED " %>>Novembro
										<option value=12 <% If mes=12 Then Response.Write " SELECTED " %>>Dezembro
									</select>
								</td>
								<td valign="middle" align="middle"> 
<%								If Cint(ano) <= anoHoje Then
											anoIni = ano - 4
											anoFim = anoHoje + 2
									Else
											anoIni = anoHoje - 1
											anoFim = ano + 1
									End If %>
									<select name="ano" class="textoP" onChange="document.forms[0].submit()">
<%									For i = anoIni To anoFim
											Response.Write "<option value='" & i & "' "
											If i=ano Then Response.Write " SELECTED "
											Response.Write ">" & i 
										Next %>
									</select>
								</td>
								<td><!--<input type="Submit" value=" ir " class="textoP">-->&nbsp;
								</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
</form>
</table>
</body>
</html>
<script>
<% 
If nSemanas = 5 Then Response.Write "resizeTo(200, 182)"
If nSemanas = 6 Then Response.Write "resizeTo(200, 202)"
%>
</script>
