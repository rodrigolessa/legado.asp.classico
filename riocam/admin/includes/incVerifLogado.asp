<% 
If Session("XCA_LOGADO_ADM") <> "TRUE" Then
	Response.Redirect Application("XCA_APP_HOME") & "admin/"
End If
%>