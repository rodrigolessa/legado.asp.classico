<%@codepage = 850 %>
<%
'Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Response.CacheControl = "no-cache"
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#3366FF"><tr><td>
</td></tr></table>

<%
'*****************************************************************************'
'*                                                                           *'
'* DADOS DAS VARIÁVEIS DE AMBIENTE                                           *'
'*                                                                           *'
'*****************************************************************************'
%>
<br>
<table border="1" width=75% align="center" cellspacing=0><small>

<tr>
<th colspan=2 bgcolor="#3366FF"><font face=Verdana size=2 color=#FFCC99>Status do Site</font></th>
</tr>
<tr>
<td colspan=2><font face="Verdana" size="2">Sessão: (<%=Session.SessionID%>)</font></td>
</tr>
<tr>
<td colspan=2><font face="Verdana" size="2">Data e Hora: <%=now() & " - " & date%></font></td>
</tr>
<tr>
<td colspan=2><font face="Verdana" size="2">A página já foi visitada por <B><%=Application("visitas")%></B> usuários desde 10/03/2002.</font></td>
</tr>
<tr>
<td colspan=2><font face="Verdana" size="2">Existem <B><%=Application("visitamomento")%></B> usuários online.</font></td>
</tr>

<tr>
<th valign=top bgcolor="#3366FF" colspan=2><font face=Verdana size=2 color=#FFCC99><b>Variável de Aplicação</b></font></th>
<%
For Each Item in Application.Contents %>
	<tr>
	<td><font face="Verdana" size="2"><%= Item %></font></td>
	<td><font face="Verdana" size="2">
	<%
		If IsArray(Application.Contents(Item)) Then
			For i = 0 to ubound(Application.Contents(Item))%>
				<%= Item %>(<%= i %>)= <%= Application.Contents(Item)(i) %><br>
			<%Next%>
		<%Else
			If IsObject(Application.Contents(Item)) Then
				Response.Write "Object"
			Else
				Response.Write Application.Contents(Item) & "&nbsp;"
			End If
		End If
	%></font>
	</td></tr>
<% Next %>

<tr>
<th valign=top bgcolor="#3366FF" colspan=2><font face=Verdana size=2 color=#FFCC99><b>Variável de Seção</b></font></th>
<%
For Each Item in Session.Contents %>
	<tr>
	<td><font face="Verdana" size="2"><%= Item %></font></td>
	<td><font face="Verdana" size="2">
	<%
		If IsArray(Session.Contents(Item)) Then
			For i = 0 to ubound(Session.Contents(Item))%>
				<%= Item %>(<%= i %>)= <%= Session.Contents(Item)(i) %><br>
	<%
			Next
		Else
			If IsObject(Session.Contents(Item)) Then
				Response.Write "Object"
			Else
				Response.Write Session.Contents(Item) & "&nbsp;"
			End If
		End If
	%></font>
	</td></tr>
<% Next %>

<tr><th colspan=2 bgcolor="#3366FF"><font face=Verdana size=2 color=#FFCC99>Listando todas as Variáveis de ambiente</font></th></tr>
<%
  for each x in request.servervariables
    response.write("<tr><td><font face=Verdana size=2>" & x & "</font></td><td><font face=Verdana size=2>")
    response.write(request.servervariables(x) & "&nbsp;")
    response.write("</font></td></tr>")
  next
%>
</small>

</table>

<%
'*****************************************************************************'
'*                                                                           *'
'* FIM DA DADOS DAS VARIÁVEIS DE AMBIENTE                                    *'
'*                                                                           *'
'*****************************************************************************'
%>
