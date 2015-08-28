<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/geral_gerencia.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
DIM query
DIM modelo, i
DIM erro
DIM strRedir, excluidos, naoExcluidos
DIM dirModRaiz, dirModGal, dirModVid

CALL conectar()

response.buffer = true

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus")

if request("aModelo") <> "" then

	erro			= ""
	excluidos		= 0
	naoExcluidos	= 0
	modelo			= split(request("aModelo"), ", ")
	
'	ON ERROR RESUME NEXT
	
	conn.BeginTrans
	
	for i = 0 to uBound(modelo)
			
			'EXCLUINDO FOTOS DAS GALERIAS DAS MODELOS
			query = "DELETE FROM XCA_modelos_galerias_fotos WHERE modelo_galeria IN (SELECT mg.modelo_galeria FROM XCA_modelos_galerias mg WHERE mg.modelo = " & Replace(modelo(i), """", "") & ") "
			conn.Execute query
			
			'EXCLUINDO GALERIAS DAS MODELOS
			query = "DELETE FROM XCA_modelos_galerias WHERE modelo = " & replace(modelo(i), """", "") & " "
			conn.Execute query
			
			'EXCLUINDO FERENCIAS DE VIDEOS, NA TABELA VIDEOS, DAS MODELOS SELECIONADAS
			query = "DELETE FROM XCA_modelos_videos WHERE modelo = " & Replace(modelo(i), """", "") & " "
			conn.Execute query
			
			'FINALMENTE EXCLUINDO A MODELO
			query = "DELETE FROM XCA_modelos WHERE modelo = " & replace(modelo(i), """", "") & " "
			conn.Execute query
			
			excluidos = excluidos + 1
			
	next
	
	If conn.Errors.Count = 0 Then
	
		for i=0 to uBound(modelo)
		
			dirModRaiz = Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & replace(modelo(i), """", "")
			dirModGal = dirModRaiz & "\galeria"
			dirModVid = dirModRaiz & "\video"

			if existeDiretorio(dirModRaiz) then
				apagarDiretorio(dirModRaiz)
			end if
				
		next
		
		if ERR.number = 0 then
			conn.CommitTrans
			erro = ""
		else
			conn.RollbackTrans
			erro =	"Ocorreu o seguinte erro: <br><br>" & _
					"<b>Erro : </b>" & ERR.number & "<br>" & _
					"Descrição do erro: " & ERR.Description & "<br>"
		end if

	Else
	
		conn.RollbackTrans
		erro = "Ocorreu erro na exclusão de um ou mais registro, a exclusão não foi efetuada!"
		
	End If
	
	ON ERROR GOTO 0
	
	If erro <> "" Then
		Response.Redirect "msg.asp?m=" & Server.URLEncode(erro)
		Response.End
	End If
	
End If

Response.Redirect "modelos.asp?" & strRedir
%>