<%
Dim strRedir
Const maxTitulo = 100
Const maxdescricao = 255

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto")  & "&fStatus=" & Request("fStatus") 
'response.write strRedir

'informações da foto
Dim dirArq, nomeArqP, nomeArqI
'verifica se é a primeira vez
If Request("cmd")="" Then
	'verifica se é a primeira vez da edição ou da inclusão
	If oper = "E" Then
		'editar
		calendario = Request("calendario")
		query = "SELECT calendario, dia, mes, ano, titulo, descricao, tipo, status FROM " & nomeTabela & " WHERE calendario=" & calendario 
		rs.Open query, conn
		If Not rs.EOF Then
			dia = Trim(rs("dia"))
			mes = Trim(rs("mes"))
			ano = Trim(rs("ano"))
			titulo = Trim(rs("titulo"))
			descricao = Trim(rs("descricao"))
			tipo = UCase(rs("tipo"))
			status = UCase(rs("status"))
		Else
 			erro = "É preciso informar um registro para poder edita-lo"
		End If
		rs.close
	Else
		'incluir
	End If
Else
	'segunda vez no formulário, valida dados para salvar
	calendario = Request("calendario")
	dia = Request("dia")
	mes = Request("mes")
	ano = Request("ano")
	titulo = Request("titulo")
	descricao = Request("descricao")
	tipo = UCase(Request("tipo"))
	status = UCase(Request("status"))
	If Len(titulo) < 1 Or Len(titulo) > maxTitulo Then erro = erro & "-O campo titulo deve ser informado e deve conter no máximo " & maxTitulo & " caracteres<br>"
	If Len(descricao) > maxDescricao Then erro = erro & "-O campo descricao deve conter no máximo " & maxDescricao & " caracteres<br>"
	If tipo = "N" Then	If Not IsDate(dia&"/"&mes&"/"&ano) Then erro = erro & "-Data inválida<br>"
	'prepara campo para salvar
	If erro="" Then
		'na funcao prepara campo eu prepara alguns campos para nulos caso ele nao seja informado
		'pois no modo cliente o teste para mostrar ou nao o campo é ver se ele esta vazio
		Call preparaCampos()
		If oper="I" Then
				Application.Lock
				query = "INSERT INTO " & nomeTabela & " (dia, mes, ano, titulo, descricao, tipo, status) VALUES " & _
								"	(" & dia & ", " & mes & ", " & ano & ", '" & sTitulo & "', '" & sDescricao & "', '" & tipo & "', '" & status & "')"
				conn.Execute query  
				Application.UnLock
				If Request("cmd") = "S" Then
					dia = ""
					mes = ""
					ano = ""
					titulo = ""
					descricao = ""
					tipo = "N"
					status = "A" 
					Response.Write "<script language='JavaScript'> alert('registro cadastrado com sucesso!') </script>"
				Else
					query = "Select Max(calendario) From " & nomeTabela
					rs.open query, conn
					If not rs.Eof Then calendario = rs(0)
					rs.close
					oper = "E"
				End If
		Else
			'grava edicao
				query = "UPDATE " & nomeTabela & " SET " & _
								"		dia=" & dia & ", " & _
								"		mes=" & mes & ", " & _
								"		ano=" & ano & ", " & _
								"		titulo='" & sTitulo & "', " & _
								"		descricao='" & sDescricao & "', " & _
								"		texto='" & sTexto & "', " & _
								"		tipo='" & tipo & "', " & _
								"		status='" & status & "' " & _
								" WHERE calendario=" & calendario & ""
				conn.Execute query
				Response.Redirect "calendarios.asp?calendario=" & calendario & "&" & strRedir
				Response.End
		End If
	End If
End If
%>
<table width="100%" height="100%" border="0">
	<tr>
		<td valign="top">
			<table width="100%" cellpadding="3" cellspacing="0" border="0">
				<tr>
					<td>
					<div class="erro" align="right">* preenchimento obrigatório</div>
							<table border="0" cellpadding="0" cellspacing="0" class="texto2">
								
								<input type="hidden" name="oper" value="<%=oper%>">
								<input type="hidden" name="cmd">
								<input type='hidden' name='calendario' value='<%= calendario %>'>
								<input type='Hidden' name='pagina' value='<%=Request("pagina")%>'>
								<input type='Hidden' name='ordemField' value='<%=Request("ordemField")%>'>
								<input type='Hidden' name='ordemType' value='<%=Request("ordemType")%>'>
								<input type='Hidden' name='fCampo' value='<%=Request("fCampo")%>'>
								<input type='Hidden' name='fTexto' value='<%=Request("fTexto")%>'>
								<input type='Hidden' name='fTipo' value='<%=Request("fTipo")%>'>
								<input type='Hidden' name='fStatus' value='<%=Request("fStatus")%>'>
								<% If erro <> "" Then %>
									<tr>
										<td class="erro">&nbsp;ERRO:</td>
										<td >&nbsp;&nbsp;</td>
										<td class="erro"><%= erro %></td>
									</tr>
								<% End If %>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Tipo:</td>
									<td>&nbsp;&nbsp;</td>
									<td>
									<select name="tipo" size="1" class="sel">
										<option value="N" <% If tipo = "N" Then Response.Write " SELECTED " %>>Não se repete
										<option value="A" <% If tipo = "A" Then Response.Write " SELECTED " %>>Repete anualmente
										<option value="M" <% If tipo = "M" Then Response.Write " SELECTED " %>>Repete mensalmente
									</select>
									</td>
								</tr>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Data:</td>
									<td>&nbsp;&nbsp;</td>
									<td>
										<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td>
											<select name="dia" size="1" class="sel">
												<option value="1"  <% If dia = "1"  Then Response.Write " SELECTED " %>>01
												<option value="2"  <% If dia = "2"  Then Response.Write " SELECTED " %>>02
												<option value="3"  <% If dia = "3"  Then Response.Write " SELECTED " %>>03
												<option value="4"  <% If dia = "4"  Then Response.Write " SELECTED " %>>04
												<option value="5"  <% If dia = "5"  Then Response.Write " SELECTED " %>>05
												<option value="6"  <% If dia = "6"  Then Response.Write " SELECTED " %>>06
												<option value="7"  <% If dia = "7"  Then Response.Write " SELECTED " %>>07
												<option value="8"  <% If dia = "8"  Then Response.Write " SELECTED " %>>08
												<option value="9"  <% If dia = "9"  Then Response.Write " SELECTED " %>>09
												<option value="10" <% If dia = "10" Then Response.Write " SELECTED " %>>10
												<option value="11" <% If dia = "11" Then Response.Write " SELECTED " %>>11
												<option value="12" <% If dia = "12" Then Response.Write " SELECTED " %>>12
												<option value="13" <% If dia = "13" Then Response.Write " SELECTED " %>>13
												<option value="14" <% If dia = "14" Then Response.Write " SELECTED " %>>14
												<option value="15" <% If dia = "15" Then Response.Write " SELECTED " %>>15
												<option value="16" <% If dia = "16" Then Response.Write " SELECTED " %>>16
												<option value="17" <% If dia = "17" Then Response.Write " SELECTED " %>>17
												<option value="18" <% If dia = "18" Then Response.Write " SELECTED " %>>18
												<option value="19" <% If dia = "19" Then Response.Write " SELECTED " %>>19
												<option value="20" <% If dia = "20" Then Response.Write " SELECTED " %>>20
												<option value="21" <% If dia = "21" Then Response.Write " SELECTED " %>>21
												<option value="22" <% If dia = "22" Then Response.Write " SELECTED " %>>22
												<option value="23" <% If dia = "23" Then Response.Write " SELECTED " %>>23
												<option value="24" <% If dia = "24" Then Response.Write " SELECTED " %>>24
												<option value="25" <% If dia = "25" Then Response.Write " SELECTED " %>>25
												<option value="26" <% If dia = "26" Then Response.Write " SELECTED " %>>26
												<option value="27" <% If dia = "27" Then Response.Write " SELECTED " %>>27
												<option value="28" <% If dia = "28" Then Response.Write " SELECTED " %>>28
												<option value="29" <% If dia = "29" Then Response.Write " SELECTED " %>>29
												<option value="30" <% If dia = "30" Then Response.Write " SELECTED " %>>30
												<option value="31" <% If dia = "31" Then Response.Write " SELECTED " %>>31
											</select>
											</td>
											<td class="texto2">&nbsp;/&nbsp;</td>
											<td>
											<select name="mes" size="1" class="sel">
												<option value="1"  <% If mes = "1"  Then Response.Write " SELECTED " %>>Janeiro
												<option value="2"  <% If mes = "2"  Then Response.Write " SELECTED " %>>Fevereiro
												<option value="3"  <% If mes = "3"  Then Response.Write " SELECTED " %>>Março
												<option value="4"  <% If mes = "4"  Then Response.Write " SELECTED " %>>Abril
												<option value="5"  <% If mes = "5"  Then Response.Write " SELECTED " %>>Maio
												<option value="6"  <% If mes = "6"  Then Response.Write " SELECTED " %>>Junho
												<option value="7"  <% If mes = "7"  Then Response.Write " SELECTED " %>>Julho
												<option value="8"  <% If mes = "8"  Then Response.Write " SELECTED " %>>Agosto
												<option value="9"  <% If mes = "9"  Then Response.Write " SELECTED " %>>Setembro
												<option value="10" <% If mes = "10" Then Response.Write " SELECTED " %>>Outubro
												<option value="11" <% If mes = "11" Then Response.Write " SELECTED " %>>Novembro
												<option value="12" <% If mes = "12" Then Response.Write " SELECTED " %>>Dezembro
											</select>
											</td>
											<td class="texto2">&nbsp;/&nbsp;</td>
											<td><input type="text" name="ano" maxlength="4" class="cx" size="5" value='<%=ano%>'></td>
										</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Título:</td>
									<td>&nbsp;&nbsp;</td>
									<td><input type="text" name="titulo" maxlength="<%=maxTitulo%>" class="cx" size="97" value='<%= Replace(titulo, """", "&quot;") %>'></td>
								</tr>
								<tr>
									<td valign="top">&nbsp;&nbsp;Descricao:</td>
									<td>&nbsp;&nbsp;</td>
									<td valign="top">
										<table cellpadding="0" cellspacing="0" border="0">
											<tr>
												<td><textarea cols="80" rows="12" name="descricao" class="tx"><%=descricao%></textarea></td>
												<td valign="top">&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td>&nbsp;&nbsp;Status:</td>
									<td>&nbsp;&nbsp;</td>
									<td><% Call criaComboStatus(status) %></td>
								</tr>

								</form>
								
								<script language="javascript1.2">
									document.forms[0].titulo.focus();
								</script>
								
							</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<%
Sub preparaCampos()
		'prepara campo para salvar
		tipo = UCase(tipo)
		Select case tipo
			case "N"
				dia = cInt(dia)
				mes = cInt(mes)
				ano = cInt(ano)
			case "A"
				dia = cInt(dia)
				mes = cInt(mes)
				ano = "Null"
			case "M"
				dia = cInt(dia)
				mes = "Null"
				ano = "Null"
		End Select
		sTitulo = Replace(titulo, "'", "''")
		sDescricao = Replace(descricao, "'", "''")
		status = UCase(status)
End Sub
%>