<%

rCount = 0
ordemField = Request("ordemField")
ordemType = Request("ordemType")

If Request("pagina")="" Then pagina=1 Else pagina = Request("pagina")

Select Case ordemField
	Case "descricao"
		If ordemType="DESC" Then ordem="descricao DESC" Else ordem="descricao"
	Case "status"
		If ordemType="DESC" Then ordem="status DESC, descricao" Else ordem="status, descricao"
	Case Else
		ordemField = "descricao"
		ordemType = "ASC"
		ordem="descricao"
End Select

Conn.CursorLocation = 3 ' adUseClient, coloca o cursor no cliente para agilizar a paginação

query = "SELECT modelo_categ, descricao, status " & _
		"	FROM " & nomeTabela & " " & _
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
<table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0'>

	<input type='Hidden' name='oper' value='L'>
	<input type='Hidden' name='pagina' value='<%=pagina%>'>
	<input type='Hidden' name='nPaginas' value='<%=rs.PageCount%>'>
	<input type='Hidden' name='ordemField' value='<%=ordemField%>'>
	<input type='Hidden' name='ordemType' value='<%=ordemType%>'>
	<input type='hidden' name='modelo_categ'>
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
								s = rs.RecordCount & " categorias cadastradas"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 categoria cadastrada"
							End If
						Else
							s = "Nenhuma categoria cadastrada"
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
			<td width=670 height=20>
<%If ordemField<>"descricao" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('descricao', 'ASC'); return false;" & Chr(34) & "'>Descrição</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Descrição&nbsp;"
	End If
	If ordemField="descricao" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('descricao', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('descricao', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=60 height=20>
<%If ordemField<>"status" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('status', 'ASC'); return false;" & Chr(34) & "'>Status</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Status&nbsp;"
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
		
			modelo_categ = Trim(rs("modelo_categ"))
			descricao = Trim(rs("descricao"))
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
				<td height=20 align="center"><input type='checkbox' name='amodelo_categ' value='"<%=modelo_categ%>"'></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_categ")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=descricao%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_categ")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%></span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=modelo_categ%>" <% If CStr(modelo_categ) = CStr(Request("modelo_categ")) Then Response.Write " checked " %>></td>
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
</form>
</table>
