<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
set lConsulta = Server.CreateObject("ADODB.Recordset")
lConsulta.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Bombs</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin=0 topmargin=0><font class="text3">
<%
If Not Request.QueryString("cat") = "" Then
  'Verificar existência da categoria
  Rs.Source = "SELECT * FROM TblCategorias WHERE CodigoCategoria=" & Request.QueryString("cat")
  Rs.Open()
  If Not Rs.EOF Then
	'Se existir a categoria, procurar pelos bomburgões
	Rs.Close()
 	Rs.Source = "SELECT TblBomburgoes.*, TblCategorias.* FROM TblCategorias INNER JOIN TblBomburgoes ON TblCategorias.CodigoCategoria = TblBomburgoes.CategoriaBomburgao WHERE CategoriaBomburgao=" & Request.QueryString("cat")
	Rs.Open()
	If Rs.EOF Then
	  'Se não houverem bomburgões na categoria...
	  %><br><center>Esta categoria ainda não tem um competidor</center><br><%
	Else
	  'Se houver registros, verificar o sistema da categoria
	  If Rs("SistemaCategoria") = "0" Then
		lConsulta.Source = "SELECT TblBomburgoes.*, TblIntegrantes.* FROM TblIntegrantes INNER JOIN (TblCategorias INNER JOIN TblBomburgoes ON TblCategorias.CodigoCategoria = TblBomburgoes.CategoriaBomburgao) ON TblIntegrantes.CodigoIntegrante = TblBomburgoes.IntegranteBomburgao WHERE (((TblBomburgoes.CategoriaBomburgao)=" & Request.QueryString("cat") & ")) ORDER BY TblBomburgoes.Parametro2Bomburgao DESC"
		lConsulta.Open()
		Contador = 0
		Ultimo = 0
		Iguais = 0
		Do Until lConsulta.EOF
		  Contador = Contador + 1
		  If Not lConsulta.BOF Then
		  	If lConsulta("Parametro2Bomburgao") = Ultimo Then
			  Iguais = Iguais + 1
		  	Else
			  Iguais = 0
		  	End If
		  End If
		  'Tirar o erro do zero
		  If Rs("MostrarParametro2Categoria") = 1 Then
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%> (<%=lConsulta("Parametro2Bomburgao") & " " & lConsulta("Parametro1Bomburgao")%>)<br><%
		  Else
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%> (<%=lConsulta("Parametro1Bomburgao")%>)<br><%
		  End If
		  Ultimo = lConsulta("Parametro2Bomburgao")
		  lConsulta.MoveNext
		  If Contador = 25 Then
			Exit Do
		  End If
		Loop
		lConsulta.Close()
	  ElseIf Rs("SistemaCategoria") = "1" Then
		lConsulta.Source = "SELECT TblBomburgoes.*, TblIntegrantes.* FROM TblIntegrantes INNER JOIN (TblCategorias INNER JOIN TblBomburgoes ON TblCategorias.CodigoCategoria = TblBomburgoes.CategoriaBomburgao) ON TblIntegrantes.CodigoIntegrante = TblBomburgoes.IntegranteBomburgao WHERE (((TblBomburgoes.CategoriaBomburgao)=" & Request.QueryString("cat") & ")) ORDER BY TblBomburgoes.Parametro2Bomburgao DESC"
		lConsulta.Open()
		Contador = 0
		Ultimo = 0
		Iguais = 0
		Do Until lConsulta.EOF
		  Contador = Contador + 1
		  If Not lConsulta.BOF Then
		  	If lConsulta("Parametro2Bomburgao") = Ultimo Then
			  Iguais = Iguais + 1
		  	Else
			  Iguais = 0
		  	End If
		  End If
		  'Tirar o erro do zero
		  If Rs("MostrarParametro2Categoria") = 1 Then
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%> (<%=lConsulta("Parametro2Bomburgao") & " " & lConsulta("Parametro1Bomburgao")%>)<br><%
		  Else
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%> (<%=lConsulta("Parametro1Bomburgao")%>)<br><%
		  End If
		  Ultimo = lConsulta("Parametro2Bomburgao")
		  lConsulta.MoveNext
		  If Contador = 25 Then
			Exit Do
		  End If
		Loop
		lConsulta.Close()
	  ElseIf Rs("SistemaCategoria") = "2" Then
		lConsulta.Source = "SELECT TblBomburgoes.*, TblIntegrantes.* FROM TblIntegrantes INNER JOIN (TblCategorias INNER JOIN TblBomburgoes ON TblCategorias.CodigoCategoria = TblBomburgoes.CategoriaBomburgao) ON TblIntegrantes.CodigoIntegrante = TblBomburgoes.IntegranteBomburgao WHERE (((TblBomburgoes.CategoriaBomburgao)=" & Request.QueryString("cat") & ")) ORDER BY TblBomburgoes.Parametro2Bomburgao DESC"
		lConsulta.Open()
		Contador = 0
		Ultimo = 0
		Iguais = 0
		Do Until lConsulta.EOF
		  Contador = Contador + 1
		  If Not lConsulta.BOF Then
		  	If lConsulta("Parametro2Bomburgao") = Ultimo Then
			  Iguais = Iguais + 1
		  	Else
			  Iguais = 0
		  	End If
		  End If
		  'Tirar o erro do zero
		  If Not lConsulta("Parametro2Bomburgao") = 0 Then
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%><br><%
		  Else
		  	%><font class="text2m"><b><%=Contador - Iguais%>º - </b></font><%=lConsulta("NomeIntegrante")%><br><%
		  End If
		  Ultimo = lConsulta("Parametro2Bomburgao")
		  lConsulta.MoveNext
		  If Contador = 10 Then
			Exit Do
		  End If
		Loop
		lConsulta.Close()
	  Else
		%><br><center>Erro interno. Conte um técnico em informática antes que seu computador se autodestrua em 5 segundos</center><br><%
	  End If
	End If
	Rs.Close()
  Else
	%><br><center>A Categoria não existe</center><br><%
  End If
End If
%>
</font>
</body>
</html>
