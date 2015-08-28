<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/geral_gerencia.asp"-->
<!--#include file="includes/incVars.asp"-->
<%
DIM query
DIM modelo_galeria, i, j
DIM msg_ok, msg_erro
DIM strRedir, excluidos, naoExcluidos
DIM dirModRaiz, dirModGal, dirModVid
DIM aFotos

CALL conectar()

response.buffer = true

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus")

if request("amodelo_galeria") <> "" then

	msg_erro		= ""
	excluidos		= 0
	naoExcluidos	= 0
	modelo_galeria	= split(request("amodelo_galeria"), ", ")
	
	ON ERROR RESUME NEXT
	
	conn.BeginTrans
	
	for i = 0 to uBound(modelo_galeria)
	
			'EXCLUINDO FOTOS DAS GALERIAS DAS MODELOS_GALERIAS
			sSQL = "SELECT modelo_galeria_foto FROM XCA_modelos_galerias_fotos WHERE modelo_galeria = " & replace(modelo_galeria(i), """", "") & " "
			aFotos = getArray(sSQL)
			if isArray(aFotos) then
				for j = 0 to uBound(aFotos, 2)
					'CHAMA A FUNÇÃO QUE EXCLUI OS ARQUIVO DE IMAGEM E A LINHA DA XCA_MODELOS_GALERIAS_FOTOS
					if exFotosGaleria(aFotos(0,j)) then
						msg_ok		= "Foto excluida com sucesso!"
					else
						msg_erro	= "Não foi possível excluír essa foto. Tente novamente!"
					end if

				next
			end if
			
			if len(trim(msg_erro)) < 1 then
				'EXCLUINDO GALERIAS DAS MODELOS_GALERIAS
				query = "DELETE FROM XCA_modelos_galerias WHERE modelo_galeria = " & replace(modelo_galeria(i), """", "") & " "
				conn.Execute query
			
				excluidos = excluidos + 1
			end if
			
	next
	
	If conn.Errors.Count = 0 Then
		
		if ERR.number = 0 then
			conn.CommitTrans
		else
			conn.RollbackTrans
			msg_erro =	"Ocorreu o seguinte erro ao tentar excluir o diretório: <br><br>" & _
						"<b>Erro : </b>" & ERR.number & "<br>" & _
						"Descrição do erro: " & ERR.Description & "<br>"
		end if

	Else
	
		conn.RollbackTrans
		msg_erro = "Ocorreu erro na exclusão de um ou mais registro, a exclusão não foi efetuada!"
		
	End If
	
	ON ERROR GOTO 0
	
	If msg_erro <> "" Then
		Response.Redirect "msg.asp?m=" & Server.URLEncode(msg_erro)
		Response.End
	End If
	
End If

Response.Redirect "modelos_galerias.asp?" & strRedir
%>