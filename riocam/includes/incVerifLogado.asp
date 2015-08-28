<% 
if session("XCA_LOGADO_SITE") <> "TRUE" then
	if session("XCA_MODELO_LOGADO_SITE") <> "TRUE" then
		response.redirect application("XCA_APP_HOME") & "assine.asp?msg_erro=NLOG"
	end if
end if
%>