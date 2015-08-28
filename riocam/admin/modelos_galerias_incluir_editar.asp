<% 
Dim strRedir
Const maxTitulo = 200
Const maxTitulo_eng = 200
Const maxStatus = 1

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus")

'VERIFICA SE É A PRIMEIRA VEZ
If Request("cmd")="" Then
	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
		'EDITAR
		modelo_galeria	= Request("modelo_galeria")
		
		query =	"SELECT	modelo_galeria, modelo, titulo, titulo_eng, " & _
				"		descricao, descricao_eng, " & _
				"		data, status " & _
				"FROM " & nomeTabela & " WHERE modelo_galeria = " & modelo_galeria & ""
				
		rs.Open query, conn
		
		If Not rs.EOF Then
			modelo_galeria	= trim(rs("modelo_galeria"))
			modelo			= trim(rs("modelo"))
			titulo			= trim(rs("titulo"))
			titulo_eng		= trim(rs("titulo_eng"))
			descricao		= trim(rs("descricao"))
			descricao_eng	= trim(rs("descricao_eng"))
			data			= FormatDate(rs("data"), "DD/MM/AA hh:mm")
			status			= trim(rs("status"))
		Else
			rs.close
			erro = "É preciso informar uma galeria para poder edita-la"
		End If
		
	Else
		'INCLUIR
		status	= "P"
	End If
	
Else

	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	modelo_galeria	= Request("modelo_galeria")
	modelo			= Request("modelo")
	titulo			= Request("titulo")
	titulo_eng		= Request("titulo_eng")
	descricao		= trim(request("descricao"))
	descricao_eng	= trim(request("descricao_eng"))
	status			= request("status")	
	If status <> "A" And status <> "I" Then
		status = "P"
	end if
	
	If Len(modelo) = 0 Then erro = " - O campo modelo deve ser informado<br>"
	If Len(titulo) = 0 Or Len(titulo) > maxTitulo Then erro = " - O campo título deve ser informado e deve conter no máximo " & maxtitulo & " caracteres<br>"
	If Len(titulo_eng) = 0 Or Len(titulo_eng) > maxTitulo Then erro = " - O campo título(inglês) deve ser informado e deve conter no máximo " & maxtitulo & " caracteres<br>"
	
	'PREPARA CAMPO PARA SALVAR
	If erro = "" And erro2 = "" Then
		
		If oper = "I" Then
		
			'*************************************************
			'INCLUI MODELO_GALERIA
			'*************************************************
			
			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
			
			If erro = "" Then
			
				Call preparaCampos()
				
				query =	"INSERT INTO " & nomeTabela & " ( " & _
						"		modelo, titulo, titulo_eng, " & _
						"		descricao, descricao_eng, " & _
						"		data, status) " & _
						"	VALUES ( " & _
						"		" & modelo & ", '" & titulo & "', '" & titulo_eng & "', " & _
						"		'" & descricao & "', '" & descricao_eng & "', " & _
						"		getDate(), '" & status & "')"

				conn.Execute query
				
				'CRIA UM DIRETÓRIO PARA A GALERIA DA MODELO, SE NÃO EXISTIR
				dirModeloGaleria = Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & modelo & "\galeria"

				if not existeDiretorio(dirModeloGaleria) then
					dirModGal	= criaDiretorio(dirModeloGaleria)
				end if
				
				If conn.Errors.Count = 0 Then

					conn.CommitTrans

					'LIMPA TODOS OS CAMPOS DO FORMULÁRIO
					modelo			= ""
					titulo			= ""
					titulo_eng		= ""
					descricao		= ""
					descricao_eng	= ""
					data			= ""
					status			= "I"
					
					oper			= "I"
					cmd				= ""

					%>
					<script language="JavaScript">
						alert("modelo_galeria cadastrado com sucesso!!")
					</script>
					<%

				Else

					conn.RollbackTrans
					erro = "Ocorreu erro na inclusão de um registro, o cadastro não foi efetuado!"

				End If
				
			Else
			
				conn.RollbackTrans
	
			End If

			ON ERROR GOTO 0
			
		Else
		
			'************************************************************
			'GRAVA EDICAO
			'************************************************************
			
			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
			
			If erro = "" Then
			
				Call preparaCampos()
				
				query =	"UPDATE " & nomeTabela & " SET " & _
						" titulo			='" & titulo & "', " & _
						" titulo_eng		='" & titulo_eng & "', " & _
						" descricao			='" & descricao & "', " & _
						" descricao_eng		='" & descricao_eng & "', " & _
						" status			='" & status & "' " & _
						"WHERE modelo_galeria = " & modelo_galeria
						
				conn.Execute query
				
				If conn.Errors.Count = 0 Then

					conn.CommitTrans

					response.redirect "modelos_galerias.asp?modelo_galeria=" & modelo_galeria & "&" & strRedir
					response.end

				Else

					conn.RollbackTrans
					erro = "Ocorreu erro na alteração de um registro, a atualização não foi efetuada!"

				End If
				
			Else
			
				conn.RollbackTrans
				
			End If

			ON ERROR GOTO 0
	
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
						<table border="0" cellpadding="0" cellspacing="2" class="texto2">
							
							<input type="hidden" name="oper"			value="<%=oper%>">
							<input type="hidden" name="cmd">
							<input type="hidden" name="modelo_galeria"	value="<%=modelo_galeria%>">
							<input type="Hidden" name="pagina"			value="<%=Request("pagina")%>">
							<input type="Hidden" name="ordemField"		value="<%=Request("ordemField")%>">
							<input type="Hidden" name="ordemType"		value="<%=Request("ordemType")%>">
							<input type="Hidden" name="fCampo"			value="<%=Request("fCampo")%>">
							<input type="Hidden" name="fTexto"			value="<%=Request("fTexto")%>">
							<input type="Hidden" name="fStatus"			value="<%=Request("fStatus")%>">
							
							<% If erro <> "" Then %>
								<tr>
									<td class="erro" valign="top">&nbsp;ERRO:</td>
									<td >&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Título:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="titulo" class="cx" size="50" maxlength="<%=maxTitulo%>" value="<%=titulo%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Título (inglês):</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="titulo_eng" class="cx" size="50" maxlength="<%=maxTitulo%>" value="<%=titulo_eng%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Modelo:</td>
								<td>&nbsp;&nbsp;</td>
								<td><%=criaComboModelo(modelo)%></td>
							</tr>

							<tr>
								<td valign="top">&nbsp;&nbsp;Descrição:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><textarea cols="80" rows="12" name="descricao" class="tx" style="width:500; height:100"><%=descricao%></textarea></td>
											<td valign="top">&nbsp;</td>
											<script language="javascript1.2">
												editor_generate("descricao");
											</script>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="top">&nbsp;&nbsp;Descrição (inglês):</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><textarea cols="80" rows="12" name="descricao_eng" class="tx" style="width:500; height:100"><%=descricao_eng%></textarea></td>
											<td valign="top">&nbsp;</td>
											<script language="javascript1.2">
												editor_generate("descricao_eng");
											</script>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;status:</td>
								<td>&nbsp;&nbsp;</td>
								<td><% Call criaComboStatus(status) %></td>
							</tr>
						</table>
					</td>
				</tr>
						
				<%
					if oper = "I" then
				%>
						<script language="javascript1.2">
							document.forms[0].titulo.focus()
						</script>
				<%
					elseif oper = "E" then
				%>
						<script language="javascript1.2">
							document.forms[0].titulo.focus()
							document.forms[0].modelo.disabled = true;
						</script>
				<%
					end if
				%>
				
			</table>
		</td>
	</tr>
</table>
<%
Sub preparaCampos()
	titulo			= replace(titulo, "'", "´")
	titulo_eng		= replace(titulo_eng, "'", "´")
	descricao		= replace(descricao, "'", "´")
	descricao_eng	= replace(descricao_eng, "'", "´")
	status			= Ucase(status)
End Sub

Sub criaComboStatus(status)
	s = "<select name='status' class='cx'>" & vbCrLf
	s = s & "<option value='A' " 
	If status = "A" Then s = s & " SELECTED "
	s = s & ">Ativo" & vbCrLf
	s = s & "<option value='P' " 
	If status = "P" Then s = s & " SELECTED "
	s = s & ">Pendente" & vbCrLf
	s = s & "<option value='I' " 
	If status = "I" Then s = s & " SELECTED "
	s = s & ">Inativo" & vbCrLf
	s = s & "</select>" & vbCrLf
	Response.Write s
End Sub


'FUNÇÃO PARA CRIAR A COMBO DE MODELOS
function criaComboModelo(ccModelo)
	DIM rs2, sSQL2, temp
	SET rs2	= CreateObject("ADODB.Recordset")
	
	ccModelo = trim(ccModelo)
	if len(ccModelo) > 0 then
		if isNumeric(ccModelo) then
			ccModelo = cInt(ccModelo)
		else
			ccModelo = 0
		end if
	else
		ccModelo = 0
	end if
	
	temp = ""

	sSQL2 = "SELECT modelo, nome FROM XCA_modelos WHERE status = 'A' ORDER BY nome "
	rs2.open sSQL2, conn

	if not rs2.EOF then

		temp = temp &	"<SELECT name=modelo class=""sel"">"
						'"<option value=""""> </option>"
		while not rs2.EOF 
			temp = temp & "<option value="""&rs2("modelo")&""" "
			if rs2("modelo") = ccModelo then temp = temp & "SELECTED"
			temp = temp & ">"&rs2("nome")&"</option>"
			rs2.MoveNext
		wend
		temp = temp &	"</SELECT>"
		
	end if
	
		rs2.close
	set rs2 = nothing

	criaComboModelo = temp
end function
%>