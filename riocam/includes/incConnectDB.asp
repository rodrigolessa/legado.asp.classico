<%

	DIM conn, rs, sSQL

	Sub conectar()
		DIM strConnect
		strConnect	= Application("XCA_StringConexao")
		SET conn	= Server.CreateObject("ADODB.Connection")
		SET rs		= Server.CreateObject("ADODB.RecordSet")
	'	conn.CursorLocation		= 3   'adUseClient
	'	conn.ConnectionTimeout	= 7200
	'	conn.CommandTimeout		= 7200
		conn.open strConnect
	End Sub

	Sub desconectar()
		conn.close
		SET conn = Nothing
	End Sub

	' FUNÇÃO PARA RETORNAR A CONSULTA EM UM ARRAY
	Function getArray(strRSSource)
	DIM rs2
	
	ON ERROR RESUME NEXT
	
	SET rs2 = Server.CreateObject("ADODB.RecordSet")
		rs2.source = strRSSource
		rs2.activeconnection = conn
		rs2.CursorType = 2   ' Use 2 instead of adCmdTable
		rs2.LockType = 3     ' Use 3 instead of adLockOptimistic
		rs2.open
'		rs2.Open strRSSource, conn

		If err.number <> 0 Then
			response.write "<div id=""msg_erro"">Erro no sistema <br> Descrição do Erro: " & Err.Description & VbCrLf & "Fonte: " & Err.Source & VbCrLf & "Arquivo: " & Request.ServerVariables("SCRIPT_NAME") & VbCrLf & "ID do Usuário: " & Session("Id_Usuario") & VbCrLf & "String SQL: " & strRSSource & "</div>"
			response.flush
		End If

		If rs2.EOF Then
			getArray = False
		Else
			getArray = rs2.GetRows
		End If

			rs2.Close
		SET rs2 = Nothing
		
	ON ERROR GOTO 0

	End Function
	
%>