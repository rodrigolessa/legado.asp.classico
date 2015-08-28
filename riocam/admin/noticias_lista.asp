<%
Dim filtronoticia

rCount = 0
ordemField = Request("ordemField")
ordemType = Request("ordemType")
If Request("pagina")="" Then pagina=1 Else pagina = Request("pagina")
Select Case ordemField
	Case "Titulo"
		If ordemType="DESC" Then ordem="titulo DESC" Else ordem="titulo"
	Case "data"
		If ordemType="DESC" Then ordem="data DESC" Else ordem="data"
	Case "dataExpiracao"
		If ordemType="DESC" Then ordem="dataExpiracao DESC" Else ordem="dataExpiracao"
	Case "status"
		If ordemType="DESC" Then ordem="status DESC, titulo" Else ordem="status, titulo"
	Case Else
		ordemField = "data"
		ordemType = "DESC"
		ordem="data Desc"
End Select

filtro = ""
If fTexto<> "" Then
	If fCampo<> "" Then filtro = fCampo & " like " & FormatLikeSearch(Replace(fTexto, "'", "''")) & " "
End If
If fStatus<>"" Then
	If filtro<>"" Then filtro = filtro & " AND "
	filtro = filtro & " status='" & fStatus & "' "
End If
If filtro <> "" Then filtro = " WHERE " & filtro
Conn.CursorLocation = 3 ' adUseClient, coloca o cursor no cliente para agilizar a paginação
query = "SELECT noticia, titulo, lead, data, dataExpiracao, status " & _
				"	FROM " & nomeTabela & " " & _
				filtro & _
				" ORDER BY " & ordem
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = conn
rs.CursorType = 1 'adOpenKeyset
rs.LockType = 1 'adLockReadOnly
rs.PageSize = linhasDetalhe
rs.CacheSize = linhasDetalhe 'define o cache no mesmo tamanho da página por questoes de performance
rs.Open query

filtronoticia = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
%>
<table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0'>

	<input type='Hidden' name='oper' value='L'>
	<input type='Hidden' name='pagina' value='<%=pagina%>'>
	<input type='Hidden' name='nPaginas' value='<%=rs.PageCount%>'>
	<input type='Hidden' name='ordemField' value='<%=ordemField%>'>
	<input type='Hidden' name='ordemType' value='<%=ordemType%>'>
	<input type='hidden' name='noticia' value="<%=Request("noticia")%>">
	<input type='hidden' name='filtronoticia' value="<%=filtronoticia%>">
<tr>
<td valign="top">
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
								s = rs.RecordCount & " notícias cadastradas"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 notícia cadastrada"
							End If
						Else
							s = "Nenhum notícia cadastrada"
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
			<td width=100 height=20>
<%If ordemField<>"data" Then
		Response.Write	"&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('data', 'ASC'); return false;" & Chr(34) & "'>Data</a>"
	Else
		Response.Write	"&nbsp;Data&nbsp;"
	End If
	If ordemField="data" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('data', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('data', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=120 height=20>
<%If ordemField<>"dataExpiracao" Then
		Response.Write	"&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataExpiracao', 'ASC'); return false;" & Chr(34) & "'>Data Expiração</a>"
	Else
		Response.Write	"&nbsp;Data Expiração&nbsp;"
	End If
	If ordemField="dataExpiracao" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataExpiracao', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dataExpiracao', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=440 height=20>
<%If ordemField<>"titulo" Then
		Response.Write	"&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'ASC'); return false;" & Chr(34) & "'>Título</a>"
	Else
		Response.Write	"&nbsp;Título&nbsp;"
	End If
	If ordemField="titulo" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=60 height=20>
<%If ordemField<>"status" Then
		Response.Write	"&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'ASC'); return false;" & Chr(34) & "'>Status</a>"
	Else
		Response.Write	"&nbsp;Status&nbsp;"
	End If
	If ordemField="status" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
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
		noticia = Trim(rs("noticia"))
		titulo = Trim(rs("titulo"))
		data = FormatDate(rs("data"), "DD/MM/AA")
		dataExpiracao = rs("dataExpiracao")
		If Not IsNull(dataExpiracao) Then
			If dataExpiracao < Now() Then 
				dataExpiracao = "<font color='#FF0000'>" & FormatDate(dataExpiracao, "DD/MM/AA") & "</font>"
			Else
				dataExpiracao = FormatDate(dataExpiracao, "DD/MM/AA")
			End If
		Else
			dataExpiracao = "não expira"
		End If
		status = UCase(rs("status"))
		If status = "A" Then status = "ativo" Else If status = "I" Then status = "<font color='#FF0000'>inativo</font>"
%>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20 align="center"><input type='checkbox' name='anoticia' value='"<%=noticia%>"'></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("noticia")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=data%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("noticia")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=dataExpiracao%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("noticia")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=titulo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("noticia")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%></span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=noticia%>" <% If CStr(noticia) = CStr(Request("noticia")) Then Response.Write " checked " %>></td>
			</tr> 
<%		rs.MoveNext
		Else %>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20>&nbsp;</td>
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