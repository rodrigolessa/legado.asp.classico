<%
Dim filtrocalendario

rCount = 0
ordemField = Request("ordemField")
ordemType = Request("ordemType")

If Request("pagina")="" Then pagina=1 Else pagina = Request("pagina")

Select Case ordemField
	Case "titulo"
		If ordemType="DESC" Then ordem="titulo DESC" Else ordem="titulo"
	Case "dia"
		If ordemType="DESC" Then ordem="dia DESC" Else ordem="dia"
	Case "mes"
		If ordemType="DESC" Then ordem="mes DESC" Else ordem="mes"
	Case "ano"
		If ordemType="DESC" Then ordem="ano DESC" Else ordem="ano"
	Case "tipo"
		If ordemType="DESC" Then ordem="tipo DESC" Else ordem="tipo"
	Case "status"
		If ordemType="DESC" Then ordem="status DESC, titulo" Else ordem="status, titulo"
	Case Else
		ordemField = "ano"
		ordemType = "DESC"
		ordem="ano Desc"
End Select

'Define o filtro da página
fTexto = Request("fTexto") 
fCampo = Request("fCampo")
fTipo = Request("fTipo")
fStatus = Request("fStatus")

filtro = ""
If fTexto<> "" Then
	If fCampo<> "" Then filtro = fCampo & " like " & FormatLikeSearch(Replace(fTexto, "'", "''")) & " "
End If
If fStatus<>"" Then
	If filtro<>"" Then filtro = filtro & " AND "
	filtro = filtro & " status='" & fStatus & "' "
End If
If fTipo<>"" Then
	If filtro<>"" Then filtro = filtro & " AND "
	filtro = filtro & " tipo='" & fTipo & "' "
End If

If filtro <> "" Then filtro = " WHERE " & filtro

Conn.CursorLocation = 3 ' adUseClient, coloca o cursor no cliente para agilizar a paginação

query = "SELECT calendario, dia, IIF(mes Is Null, '--', mes) AS sMes, IIF(ano Is Null, '--', ano)  AS sAno, titulo, descricao, tipo, status " & _
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

filtrocalendario = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus")  & "&fTipo=" & Request("fTipo") 

%>
<table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0'>

	<input type='Hidden' name='oper' value='L'>
	<input type='Hidden' name='pagina' value='<%=pagina%>'>
	<input type='Hidden' name='nPaginas' value='<%=rs.PageCount%>'>
	<input type='Hidden' name='ordemField' value='<%=ordemField%>'>
	<input type='Hidden' name='ordemType' value='<%=ordemType%>'>
	<input type='hidden' name='calendario' value="<%=Request("calendario")%>">
	<input type='hidden' name='filtrocalendario' value="<%=filtrocalendario%>">
	
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
								s = rs.RecordCount & " datas cadastradas"
								If rs.PageCount > 1 Then
									s = s & " (mostrando página " & pagina & " de " & rs.PageCount & " páginas)"
								End If
							Else
								s = "1 data cadastrada"
							End If
						Else
							s = "Nenhuma data cadastrada"
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
			<td width=20 height=20>&nbsp;</td>
			<td width=50 height=20>
<%If ordemField<>"dia" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dia', 'ASC'); return false;" & Chr(34) & "'>Dia</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Dia&nbsp;"
	End If
	If ordemField="dia" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dia', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('dia', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=50 height=20>
<%If ordemField<>"mes" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('mes', 'ASC'); return false;" & Chr(34) & "'>Mês</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Mês&nbsp;"
	End If
	If ordemField="mes" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('mes', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('mes', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=50 height=20>
<%If ordemField<>"ano" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('ano', 'ASC'); return false;" & Chr(34) & "'>Ano</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Ano&nbsp;"
	End If
	If ordemField="ano" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('ano', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('ano', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
		End If
	End If %>
			</td>
			<td width=350 height=20>
<%If ordemField<>"titulo" Then
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
	End If %>
			</td>
			<td width=140 height=20>
<%If ordemField<>"tipo" Then
		Response.Write	"&nbsp;&nbsp;<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('tipo', 'ASC'); return false;" & Chr(34) & "'>Tipo</a>"
	Else
		Response.Write	"&nbsp;&nbsp;Tipo&nbsp;"
	End If
	If ordemField="tipo" Then
		If ordemType="ASC" Then
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('tipo', 'DESC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ograve;</font></a>"
		Else
			Response.Write	"<a href='javascript:void(0)' onclick=" & Chr(34) & "ordenar('tipo', 'ASC'); return false;" & Chr(34) & "'><font face='Wingdings'>&ntilde;</font></a>"
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
		calendario = Trim(rs("calendario"))
		titulo = Trim(rs("titulo"))
		dia = Trim(rs("dia"))
		mes = Trim(rs("sMes"))
		ano = Trim(rs("sAno"))
		tipo = UCase(rs("tipo"))
		status = uCase(trim(rs("status")))
		Select Case status 
			Case "A"
				status = "Ativo" 
			Case "I"
				status = "<font color='#FF0000'>Inativo</font>"
			Case "P"
				status = "Pendente"
		End Select
		Select Case tipo
			case "N"
				tipo = "Não se repete"
				If cDate(dia&"/"&mes&"/"&ano) < date() Then dia = "<font color='#FF0000'>"&dia&"</font>" : mes = "<font color='#FF0000'>"&mes&"</font>" : ano = "<font color='#FF0000'>"&ano&"</font>"
			case "A"
				tipo = "repete anualmente"
			case "M"
				tipo = "repete mensalmente"
			case else
				tipo = "não cadastrado"
		End Select
%>
			<tr class="bg<%=nBG%>" onMouseover="onColor<%=nBG%>(this);" onMouseout="offColor<%=nBG%>(this);">
				<td height=20 align="center"><input type='checkbox' name='acalendario' value='<%=calendario%>'></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=dia%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=mes%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=ano%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=titulo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=tipo%></span></td>
				<td height=20 onClick="editarLinha('<%= Trim(rs("calendario")) %>'); return false;"><span class='textoLista'>&nbsp;&nbsp;<%=status%></span></td>
				<td height=20 align="center"><input type="radio" name="registroSel" value="<%=calendario%>" <% If CStr(calendario) = CStr(Request("calendario")) Then Response.Write " checked " %>></td>
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
</form>
</table>