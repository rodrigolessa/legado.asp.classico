<% 
Dim strRedir
Const maxtitulo = 100
CONST maxNomeArquivo = 255

erro = ""

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType")

'VERIFICA SE É A PRIMEIRA VEZ
If Request("cmd") = "" Then

	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
	
		'EDITAR
		modelo_video = request("modelo_video")
		
		query = "SELECT modelo_video, modelo, titulo, titulo_eng, status " & _
				"FROM " & nomeTabela & " WHERE modelo_video = " & modelo_video
		rs.Open query, conn
		If Not rs.EOF Then
			titulo			= Trim(rs("titulo"))
			titulo_eng		= Trim(rs("titulo_eng"))
			modelo			= Trim(rs("modelo"))
			status			= UCase(rs("status"))
		Else
			erro = "É preciso informar um registro para poder edita-lo"
		End If
		rs.close
		
	Else
		'INCLUIR
	End If
	
Else
	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	modelo_video		= Request("modelo_video")
	titulo		= Request("titulo")
	titulo_eng	= Request("titulo_eng")
	modelo		= Request("modelo")
	status		= Request("status")
	If Len(titulo) < 1 Or Len(titulo) > maxtitulo Then erro = " - O campo titulo deve ser informado e deve conter no máximo " & maxtitulo & " caracteres<br>"
	If Len(titulo_eng) < 1 Or Len(titulo_eng) > maxtitulo Then erro = " - O campo titulo (inglês) deve ser informado e deve conter no máximo " & maxtitulo & " caracteres<br>"
	
	'PREPARA CAMPO PARA SALVAR
	If erro = "" Then
	
		titulo		= replace(trim(titulo), "'", "''")
		titulo_eng	= replace(trim(titulo_eng), "'", "''")
		modelo		= replace(trim(modelo), "'", "''")
		status		= uCase(status)
		
		If oper = "I" Then
		
			'*************************************************
			'INCLUI CATEGORIAS DE VIDEOS
			'*************************************************
			
'			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
		
			query = "SELECT modelo_video FROM " & nomeTabela & " WHERE titulo = '" & titulo & "' "
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & "Já existe um vídeo com esta descrição"
			end if
			rs.close
			
			If erro = "" Then
				sSQL =	"INSERT INTO " & nomeTabela & " (titulo, titulo_eng, modelo, status) VALUES " & _
						"	('" & titulo & "', '" & titulo_eng & "', " & modelo & ", '" & status & "')"
				conn.execute(sSQL)
			End If
			
			If conn.Errors.Count = 0 Then
			
				conn.CommitTrans
				erro			= ""
				
				'LIMPA TODOS OS CAMPOS DO FORMULÁRIO
				titulo		= ""
				titulo_eng	= ""
				status		= "I"
				
				%>
				<script language="JavaScript">
					alert("Vídeo cadastrado com sucesso!!")
				</script>
				<%

			Else
			
				conn.RollbackTrans
				erro = "Ocorreu erro na inclusão de um registro, o cadastro não foi efetuado!"
				
			End If

'			ON ERROR GOTO 0
			
		Else
		
			'************************************************************
			'GRAVA EDICAO
			'************************************************************
			
'			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
			
			query = "SELECT modelo_video FROM " & nomeTabela & " WHERE titulo = '" & titulo & "' AND modelo_video <> " & modelo_video
			rs.Open query, conn
			If Not rs.EOF Then erro = "Já existe um vídeo cadastrada com este título<br>"
			rs.close
			If erro = "" Then
				query =	"UPDATE " & nomeTabela & " SET " & _
						"		titulo		= '" & titulo & "', " & _
						"		titulo_eng	= '" & titulo_eng & "', " & _
						"		status		= '" & status & "' " & _
						" WHERE modelo_video = " & modelo_video & ""
				conn.execute query
			End If
			
			If conn.Errors.Count = 0 Then
			
				conn.CommitTrans
				
				erro = ""
				
				response.redirect "videos.asp?modelo_video=" & modelo_video & "&" & strRedir
				response.end
				
			Else
			
				conn.RollbackTrans
				erro = "Ocorreu erro na alteração de um registro, a atualização não foi efetuada!"
				
			End If

'			ON ERROR GOTO 0
			
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
						<table border="0" cellpadding="0" cellspacing="5" class="texto2">
						
							<input type="hidden" name="oper"			value="<%=oper%>">
							<input type="hidden" name="cmd">
							<input type="Hidden" name="pagina"			value="<%=Request("pagina")%>">
							<input type="Hidden" name="ordemField"		value="<%=Request("ordemField")%>">
							<input type="Hidden" name="ordemType"		value="<%=Request("ordemType")%>">
							<input type="hidden" name="modelo_video"			value="<%=modelo_video%>">
							
							<% If erro <> "" Then %>
								<tr>
									<td class="erro">&nbsp;ERRO:</td>
									<td >&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Descrição:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<input type="text" name="titulo" maxlength="<%=maxtitulo%>" class="cx" size="50" value="<%=titulo%>">
								</td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Descrição (inglês):</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<input type="text" name="titulo_eng" maxlength="<%=maxtitulo%>" class="cx" size="50" value="<%=titulo_eng%>">
								</td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Modelo:</td>
								<td>&nbsp;&nbsp;</td>
								<td><%=criaComboModelo(modelo)%></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Status:</td>
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