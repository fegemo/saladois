<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<a href="teste.asp?mensagem=leonciocriança">lçsdfjka</a>
<%
if request.querystring("mensagem") <> "" then
	response.write(request.Querystring("mensagem") & "<br>")
	%><%=request.Querystring("mensagem")%><%
end if
%>
</body>
</html>
