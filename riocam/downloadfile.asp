<% option explicit %>
<% 

	DIM strFileName, FPath, ext 
	DIM adoStream
	DIM strTipo, strCodigo

	CONST adTypeBinary = 1 

	'NOME DO ARQUIVO
	strFileName	= Trim(Request.QueryString("file"))
	strTipo		= Trim(uCase(Request.QueryString("tipo")))
	strCodigo	= Trim(Request.QueryString("codigo"))
	
	'CAMINHO COMPLETO DO ARQUIVO
	select case strTipo
	case "V"
		FPath = Application("XCA_APP_SERVERDIR_BASE") & "modelos\" & strCodigo & "\video\" & strFileName
	end select
	
	'PEGA A EXTENÇÃO DO ARQUIVO
	ext = uCase(right(strFileName, 3))

	if ext <> "ASP" and ext <> "CSS" and ext <> "MDB" then
		'response.Buffer = True
		response.clear
		response.ContentType = "application/asp-unknown" 'Content-Disposition attachment; filename=" & FPath
		'Response.ContentType = "application/x-msexcel"
		'Response.ContentType = "application/x-msexcel Content-Disposition attachment; filename=name1.xls"
		response.AddHeader "content-disposition", "attachment; filename=" & strFileName
		SET adoStream = Server.CreateObject("ADODB.Stream")
		adoStream.Open()
		adoStream.Type = adTypeBinary
		adoStream.LoadFromFile(FPath)
		response.BinaryWrite adoStream.Read()
		adoStream.Close: Set adoStream = Nothing
		response.write "arquivo: " & FPath
		response.end
	else
		response.write "Extenção proibida!"
		response.end
	end if
%>
