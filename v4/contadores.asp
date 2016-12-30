<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Rs.Source = "SELECT * FROM TblContadores"
Rs.Open()
%>
<html>
<head>
<title>-=: SalaDois .:. Contadores :=-</title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
Página: <%=Rs("NumeroContador")%><br><% Rs.MoveNext %>Cobaias: <%=Rs("NumeroContador")%><% Rs.Close() %>
</body>
</html>
