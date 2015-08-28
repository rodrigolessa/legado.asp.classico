<%
function quantFotos(qfGaleriaModelo)

	DIM aQuanFotos
	
	'===== MODELOS E GALERIAS =====
	sSQL =	("SELECT	COUNT(gf.modelo_galeria_foto) " & _
			"FROM	XCA_modelos_galerias_fotos gf " & _
			"WHERE	gf.modelo_galeria = " & qfGaleriaModelo & " ")
			
	aQuanFotos	= getArray(sSQL)
	if isArray(aQuanFotos) then
		quantFotos = aQuanFotos(0,0)
	else
		quantFotos = 0
	end if

end function
%>