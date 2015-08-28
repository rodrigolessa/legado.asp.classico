<%

rCount = 0
ordemField = Request("ordemField")
ordemType = Request("ordemType")

If Request("pagina") = "" Then pagina = 1 Else pagina = Request("pagina")

Select Case ordemField
	Case "nome"
		If ordemType="DESC" Then ordem="nome DESC" Else ordem="nome"
	Case "v.titulo"
		If ordemType="DESC" Then ordem="v.titulo DESC" Else ordem="v.titulo"
	Case "v.data"
		If ordemType="DESC" Then ordem="v.data DESC" Else ordem="v.data"
	Case "v.status"
		If ordemType="DESC" Then ordem="v.status DESC, v.titulo" Else ordem="v.status, v.titulo"
	Case Else
		ordemField	= "v.titulo"
		ordemType	= "ASC"
		ordem		= "v.titulo"
End Select

Conn.CursorLocation = 3 'adUseClient, coloca o cursor no cliente para agilizar a paginação

query = "SELECT	v.modelo_video, v.modelo, v.titulo, v.nomeArquivo, v.nomeArquivo2, v.data, v.status, " & _
		"		m.nome " & _
		"FROM			XCA_modelos m " & _
		"	INNER JOIN	XCA_modelos_videos v	ON m.modelo = v.modelo " & _
		"ORDER BY	" & ordem
				
SET rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = conn
rs.CursorType = 1 'adOpenKeyset
rs.LockType = 1 'adLockReadOnly
rs.PageSize = linhasDetalhe
rs.CacheSize = linhasDetalhe 'define o cache no mesmo tamanho da página por questoes de performance
rs.Open query

filtroUsuario = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") 
%>

<table width="100%" height="100%" border='0' cellpadding='0' cellspacing='0'>

	<input type="Hidden" name="oper" value="L">
	<input type="Hidden" name="pagina" value="<%=pagina%>">
	<input type="Hidden" name="nPaginas" value="<%=rs.PageCount%>">
	<input type="Hidden" name="ordemField" value="<%=ordemField%>">
	<input type="Hidden" name="ordemType" value="<%=ordemType%>">
	<input type="hidden" name="modelo_video" value="<%=modelo_video%>">
	
<tr>
	<td align="center" valign="top">

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
								s = rs.RecordCount & " vídeos cadastrados"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 vídeo cadastrado"
							End If
						Else
							s = "Nenhum vídeo cadastrado"
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
			<td width=85 height=20>
<%
	If ordemField <> "v.data" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.data', 'ASC'); return false;" & Chr(34) & "'>Data</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Data&nbsp;"
	End If
	If ordemField="v.data" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.data', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.data', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If
%>
			</td>
			<td width=210 height=20>
<%
	If ordemField<>"nome" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'ASC'); return false;" & Chr(34) & "'>Modelo</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Modelo&nbsp;"
	End If
	If ordemField="nome" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('nome', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If 
%>
			</td>
			<td width=250 height=20>
<%
	If ordemField<>"titulo" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'ASC'); return false;" & Chr(34) & "'>Título</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Título&nbsp;"
	End If
	If ordemField="titulo" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('titulo', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If
%>
			</td>
			<td width=60 height=20>&nbsp;&nbsp;Vídeo&nbsp;</td>
			<td width=60 height=20>&nbsp;&nbsp;Foto&nbsp;</td>
			<td width=60 height=20>
<%
	If ordemField <> "v.status" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.status', 'ASC'); return false;" & Chr(34) & "'>Status</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Status&nbsp;"
	End If
	If ordemField="v.status" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.status', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('v.status', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If
%>
			</td>
			<td width=10 height=20>&nbsp;</td>
		</tr>
<%
	'DETALHES
	For i=1 To linhasDetalhe
	
		If branco Then nBG = "1" Else nBG = "2"
		branco = NOT branco
		
		If Not rs.EOF Then
		
			modelo_video			= Trim(rs("modelo_video"))
			modelo			= Trim(rs("modelo"))
			nomeModelo		= Trim(rs("nome"))
			titulo			= Trim(rs("titulo"))
			nomeArquivo		= Trim(rs("nomeArquivo"))
			nomeArquivo2	= Trim(rs("nomeArquivo2"))
			dataCad			= verSQL(rs("data"), "D", "F")
			status			= uCase(trim(rs("status")))
			Select Case status
				Case "A"
					status = "Ativo" 
				Case "I"
					status = "<font color='#FF0000'>Inativo</font>"
				Case "P"
					status = "Pendente"
			End Select
			if existArq(mainDir & "\" & modelo & "\video\" & nomeArquivo) then
				bolArquivo	= "sim"
			else
				bolArquivo	= "<font color='#FF0000'>não</font>"
			end if
			if existArq(mainDir & "\" & modelo & "\video\" & nomeArquivo2) then
				bolArquivo2	= "sim"
			else
				bolArquivo2	= "<font color='#FF0000'>não</font>"
			end if
%>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20 align="center"><input type='checkbox' name='avideo' value='"<%=modelo_video%>"'></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=dataCad%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=nomeModelo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=titulo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=bolArquivo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=bolArquivo2%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("modelo_video")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%></span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=modelo_video%>" <% If CStr(modelo_video) = CStr(Request("modelo_video")) Then Response.Write " checked " %>></td>
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
