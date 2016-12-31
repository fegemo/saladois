<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Rs.Source = "SELECT * FROM TblFracassados"
Rs.Open()
%>
<html>
<head>
<title>-=: SalaDois .:. Fracassados :=-</title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
Do Until Rs.EOF
  Response.Write("Codigo: " & Rs("CodigoFracassado") & "<br>IP: " & Rs("IPFracassado") & "<br>Data: " & Rs("DataFracassado") & "<br><br><hr><br>")
  Rs.MoveNext
Loop
Rs.Close()
%>
</body>
</html>
