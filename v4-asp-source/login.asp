<%@LANGUAGE="VBSCRIPT" %> 
<html>
<head>
<title>Login</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p align="center" class="text2mm"><%
If Request.QueryString("cod") = "senha" Then
  %>Fodas. Dê seu jeito de arrumar outra, patife! <b>Irí</b><%
ElseIf Request.QueryString("cod") = "logout" Then
  Session.Abandon
  %>Logged Out<br><br><br><a href="login.htm" class="text3mm" onMouseOver="this.className='text2mm';"  onMouseOut="this.className='text3mm';">Login</a><%
Else
  If Trim(Request.Form("Usuario")) = "" Then
	%>Digite o nome do usuário<%
	Response.Redirect "http://www.saladois.hpg.com.br/login.htm"
  End If
  If Trim(Request.Form("Senha")) = "" Then
	%>Digite a senha do usuário<%
	Response.Redirect "http://www.saladois.hpg.com.br/login.htm"
  End If
  Response.Buffer="true" %>
  <!--#include file="Connections/Conneccao.asp" -->
  <%
  set Rs = Server.CreateObject("ADODB.Recordset")
  Rs.ActiveConnection = MM_Conneccao_STRING
  Rs.Source = "SELECT * FROM TblUsuarios WHERE NomeUsuario='" & Trim(Request.Form("Usuario")) & "'"
  Rs.LockType = 3
  Rs.Open()
  If Rs.EOF Then
	%>Usuário inexistente no banco de dados. Tenta outra palhaço<%
  Else
	If Not Rs("SenhaUsuario") = Trim(Request.Form("Senha")) Then
	  %>Senha incorreta. Seu computador se autodestruirá em 5 segundos...<%
	Else
      %>Entrando...<%
	  Session("User") = Rs("NomeUsuario")
      Session("Logged") = "yes"
	  Rs("LogsUsuario") = Rs("LogsUsuario") + 1
	  Rs.Update
	  Session("Logs") = Rs("LogsUsuario")
      Response.Redirect "gerenc.asp"
	End If
  End If
End If
  %>
</p>
</body>
</html>
