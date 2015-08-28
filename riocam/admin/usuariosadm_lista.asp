<%
	rCount		= 0


	Select Case ordemField
		Case "login"
			If ordemType="DESC" Then ordem="login DESC" Else ordem="login"
		Case "email"
			If ordemType="DESC" Then ordem="email DESC" Else ordem="email"
		Case Else
			ordemField = "login"
			ordemType = "ASC"
			ordem="login"
	End Select


	query =	"SELECT usuario, login, email, status" & _
			"	FROM " & nomeTabela & _
			"	ORDER BY " & ordem

	conn.CursorLocation		= 3   ' adUseClient, coloca o cursor no cliente para agilizar a paginação
	rs.Source			= query
	rs.ActiveConnection	= conn
	rs.CursorType		= 1 'adOpenKeyset
	rs.LockType			= 1 'adLockReadOnly
	rs.PageSize			= linhasDetalhe
	rs.CacheSize		= linhasDetalhe 'define o cache no mesmo tamanho da página por questoes de performance
	rs.Open


%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">

	<input type="hidden" name="nPaginas" value="<%=rs.PageCount%>">
	<input type="hidden" name="filtroUsuario" value="<%=filtroUsuario%>">
	<input type="hidden" name="usuario">
<tr>
<td align="center" valign="top">
<%
'1a. linha do cabecalho
%>
	<table width="100%" cellpadding="1" cellspacing="0" border="0">
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
								s = rs.RecordCount & " usuários cadastrados"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 usuário cadastrado"
							End If
						Else
							s = "Nenhum usuário cadastrado"
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
		<tr class="tituloLista">
			<td width=10 height=20>&nbsp;</td>
			<td width=150 height=20>
<%If ordemField<>"login" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('login', 'ASC'); return false;" & Chr(34) & "'>Login</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Login&nbsp;"
	End If
	If ordemField="login" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('login', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('login', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=500 height=20>
<%If ordemField<>"email" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('email', 'ASC'); return false;" & Chr(34) & "'>E-mail</a>"
	Else
		Response.Write	"&nbsp;&nbsp;E-mail&nbsp;"
	End If
	If ordemField="email" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('email', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('email', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=70 height=20>&nbsp;Status&nbsp;</td>
			<td width=10 height=20>&nbsp;</td>
		</tr>
<%
	'Detalhes
	For i=1 To linhasDetalhe
	
		If branco Then nBG = "1" Else nBG = "2"
		branco = NOT branco
		If Not rs.EOF Then
		usuario = Trim(rs("usuario"))
		login = Trim(rs("login"))
		email = Trim(rs("email"))
		status = Trim(rs("status"))
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
				<td height=20 align="center"><input type='checkbox' name='ausuario' value='"<%=usuario%>"'></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=login%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=email%>&nbsp;</span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("usuario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%>&nbsp;</span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=usuario%>" <% If CStr(usuario) = CStr(Request("usuario")) Then Response.Write " checked " %>></td>
			</tr> 
				
<%		rs.MoveNext

		Else %>
		
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
				<td height=20>&nbsp;</td>
			</tr>
			
<%		End If

	Next %>
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