<%
Sub wLog(str)
  Const ForAppending = 8, TristateFalse = 0
  Dim fs, f
  Dim aUsuario, pathFile
	pathFile = Server.MapPath("/logs") & "\log\log.txt"
	str = " : " & str
  If Session("LOGADO_ADM") Then str = Session("LOGIN") & str Else
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile(pathFile, ForAppending, True, TristateFalse)
  f.Write Now & " : " & str & Chr(13) & Chr(10) 
  f.Close
End Sub
%>