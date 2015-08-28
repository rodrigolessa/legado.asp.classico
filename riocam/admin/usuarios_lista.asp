<%

rCount = 0
ordemField = Request("ordemField")
ordemType = Request("ordemType")

If Request("pagina")="" Then pagina=1 Else pagina = Request("pagina")

Select Case ordemField
	Case "dataCadastro"
		If ordemType="DESC" Then ordem="dataCadastro DESC" Else ordem="dataCadastro"
	Case "nome"
		If ordemType="DESC" Then ordem="nome DESC" Else ordem="nome"
	Case "email"
		If ordemType="DESC" Then ordem="email DESC" Else ordem="email"
	Case "status"
		If ordemType="DESC" Then ordem="status DESC" Else ordem="status"
	Case Else
		ordemField = "dataCadastro"
		ordemType = "DESC"
		ordem="dataCadastro DESC"
End Select


filtro = ""
If fTexto<> "" Then
	If fCampo<> "" Then filtro = fCampo & " like " & FormatLikeSearch(Replace(fTexto, "'", "''")) & " "
End If
If fMalaDir<>"" Then
	If filtro<>"" Then filtro = filtro & " AND "
	filtro = filtro & " malaDir = " & fMalaDir
End If
If fStatus<>"" Then
	If filtro<>"" Then filtro = filtro & " AND "
	filtro = filtro & " status='" & fStatus & "' "
End If
If filtro <> "" Then filtro = " WHERE " & filtro

Conn.CursorLocation = 3 ' adUseClient, coloca o cursor no cliente para agilizar a paginação

query = "SELECT usuario, nome, email, dataCadastro, status" & _
		"	FROM " & nomeTabela & _
			filtro & _
		" ORDER BY " & ordem
		
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = conn
rs.CursorType = 1 'adOpenKeyset
rs.LockType = 1 'adLockReadOnly
rs.PageSize = linhasDetalhe
rs.CacheSize = linhasDetalhe 'define o cache no mesmo tamanho da página por questoes de performance
rs.Open query

filtroUsuario = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">

	<input type='Hidden' name='oper' value='L'>
	<input type='Hidden' name='pagina' value='<%=pagina%>'>
	<input type='Hidden' name='nPaginas' value='<%=rs.PageCount%>'>
	<input type='Hidden' name='ordemField' value='<%=ordemField%>'>
	<input type='Hidden' name='ordemType' value='<%=ordemType%>'>
	<input type='hidden' name='filtroUsuario' value="<%=filtroUsuario%>">
	<input type='hidden' name='usuario' value="<%=Request("usuario")%>">
<tr>
<td align="center" valign="top">
<%
'1a. linha do cabecalho
%>
	<table width="100%" cellpadding="1" cellspacing="0" border="0" bordercolor="#004754">
		<tr>
			<td class="texto4">
				<table width="100%" border="0">
				<tr>
					<td class="texto4">
				<%		If NOT rs.EOF Then 
							If Cint(rs.PageCount) < Cint(pagina) Then pagina = Cint(rs.PageCount)
							rs.AbsolutePage = pagina
							rCount = rs.RecordCount
							If rs.RecordCount > 1 Then
								s = rs.RecordCount & " clientes cadastrados"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 cliente cadastrado"
							End If
						Else
							s = "Nenhum cliente cadastrado"
						End If
						Response.Write s 
						s = ""
				%>
					</td>
					<td align="right" class="texto4">
				<% If rs.PageCount > 1 Then %>
						<a href='javascript:avoid(0)' onclick='navegar("<%= pagina - 1%>"); return false;' class='link2'><<&nbsp;</a>
						<span class='texto'>
				<%		'processa intervalo de páginas para mostrar apenas uma qtd de página
						Dim nIni, nFim
						If pagina-intervaloPaginas < 1 Then 
							nIni=1
							nFim = (intervaloPaginas * 2) + 1
							If nFim > rs.PageCount Then nFim = rs.PageCount
						Else 
							nIni = pagina-intervaloPaginas
							If pagina+intervaloPaginas > rs.PageCount Then nFim = rs.PageCount : nIni = nIni - ((pagina+intervaloPaginas) - rs.PageCount) Else nFim = pagina+intervaloPaginas
							If nIni < 1 Then nIni = 1
						End If
						For i=nIni To nFim
							If Cstr(i) <> Cstr(pagina) Then
								Response.Write "<a href='javascript:avoid(0)' onclick='navegar(" & i & "); return false;' class='link2'>" & i & "</a>&nbsp;"
							Else
								Response.Write "<font color='#FF0000'>" & i & "</font>&nbsp;"
							End If
						Next %>
						</span>
						<a href='javascript:avoid(0)' onclick='navegar("<%= pagina + 1 %>"); return false;' class='link2'>&nbsp;>></a>
				<%	End If %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
		<td>		
<%
s = ""
branco = True
If NOT rs.EOF Then %>
	<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
		<tr class='tituloLista'>
			<td width=10 height=20>&nbsp;</td>
			<td width=130 height=20>
			<%	If ordemField<>"dataCadastro" Then
					Response.Write	"<nobr>&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataCadastro', 'ASC'); return false;" & Chr(34) & "' class=link1>Data Cadastro</a>&nbsp;</nobr>"
				Else
					Response.Write	"<nobr>&nbsp;Data Cadastro&nbsp;</nobr>"
				End If
				If ordemField="dataCadastro" Then
					If ordemType="ASC" Then
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataCadastro', 'DESC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ograve;</font></a>"
					Else
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataCadastro', 'ASC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ntilde;</font></a>"
					End If
				End If %>
			</td>
			<td width=280 height=20>
			<%	If ordemField<>"nome" Then
					Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'ASC'); return false;" & Chr(34) & "' class=link1>Nome</a>"
				Else
					Response.Write	"&nbsp;&nbsp;Nome&nbsp;"
				End If
				If ordemField="nome" Then
					If ordemType="ASC" Then
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'DESC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ograve;</font></a>"
					Else
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'ASC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ntilde;</font></a>"
					End If
				End If %>
			</td>
			<td width=200 height=20>
			<%	If ordemField<>"email" Then
					Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('email', 'ASC'); return false;" & Chr(34) & "' class=link1>E-mail</a>"
				Else
					Response.Write	"&nbsp;&nbsp;E-mail&nbsp;"
				End If
				If ordemField="email" Then
					If ordemType="ASC" Then
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('login', 'DESC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ograve;</font></a>"
					Else
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('login', 'ASC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ntilde;</font></a>"
					End If
				End If %>
			</td>
			<td width=70 height=20>
			<%	If ordemField<>"status" Then
					Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'ASC'); return false;" & Chr(34) & "' class=link1>Status</a>"
				Else
					Response.Write	"&nbsp;&nbsp;Status&nbsp;"
				End If
				If ordemField="status" Then
					If ordemType="ASC" Then
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'DESC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ograve;</font></a>"
					Else
						Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'ASC'); return false;" & Chr(34) & "' class=link1><font face='Wingdings'>&ntilde;</font></a>"
					End If
				End If %>
			</td>
			<td width=10 height=20>&nbsp;</td>
		</tr>
<%
	'Detalhes
	For i=1 To linhasDetalhe
	
		If branco Then nBG = "1" Else nBG = "2"
		branco = NOT branco
		
		If Not rs.EOF Then
		
			usuario = Trim(rs("usuario"))
			nome = Trim(rs("nome"))
			email = Trim(rs("email"))
			email = "<a href='mailto:" & email & "' class='linkLista'>" & email & "</a>"
			dataCadastro = FormatDate(rs("dataCadastro"), "DD/MM/AA hh:mm")
			status = uCase(trim(rs("status")))
			Select Case status 
				Case "A"
					status = "Ativo" 
				Case "I"
					status = "<font color='#FF0000'>Inativo</font>"
				Case "P"
					status = "Pendente"
			End Select
%>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20 align="center"><input type="checkbox" name="ausuario" value="<%=usuario%>"></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=dataCadastro%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=nome%>&nbsp;</span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=email%>&nbsp;</span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%>&nbsp;</span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=usuario%>" <% If CStr(usuario) = CStr(Request("usuario")) Then Response.Write " checked " %>></td>
			</tr> 
<%
			rs.MoveNext
		Else
%>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
			</tr>
<%
		End If
	Next
%>
	</table>
<%
End If
rs.Close
conn.Close 
%>
</td>
</tr>
</table>
</td>
</tr>
</table>