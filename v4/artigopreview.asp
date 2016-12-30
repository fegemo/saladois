<%@LANGUAGE="VBSCRIPT"%> 
<%
If Session("Logged")="" Then
  Response.Redirect"branco.htm"
End If
%>
<html>
<head>
<title>Preview</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table align="center" width="100%" height="100%" class="text2m">
  <tr>
	<td align="center">
	  Preview:<br><br>
      <form action="salvar.asp?func=artigos&cod=<%=Request.QueryString("Codigo")%>" method="POST">         
        <table cellpadding=0 cellspacing=0 width="484">
          <tr align="center">
          <td bgcolor="#DEDEDE"><font face="Verdana, Arial, Helvetica" size="5"><b><%=Request.QueryString("Titulo")%></b></font><br></td>
		</tr>
		<tr class="text3m">
          <td><br>
            <%=Request.QueryString("Texto")%><br>
          </td>
		</tr>
		<tr>
          <td align="left" class="text3" bgcolor="#DEDEDE"><%=Request.QueryString("Autor")%> (<%=Request.QueryString("Data")%>)</td>
		</tr>
	  </table><br><img src="imgs\r484.gif" width="484" height="50"><br>
        Se o artigo estiver maior do que 484 pixels, volte e modifique alguma 
        coisa de modo a ficar ajustado ao tamanho m&aacute;ximo permitido (484 
        px) <br>
        <br><input type="hidden" name="Codigo" value="<%=Request.QueryString("codigo")%>"><input type="hidden" name="Titulo" value="<%=Request.QueryString("titulo")%>"><input type="hidden" name="Autor" value="<%=Request.QueryString("autor")%>"><input type="hidden" name="Data" value="<%=Request.QueryString("data")%>"><input type="hidden" name="Texto" value="<%=Request.QueryString("texto")%>"><input type="hidden" name="delartigo" value="<%=Request.QueryString("delartigo")%>"><input type="hidden" name="Preview" value=""><br><input type="submit" class="botao1" value="Salvar"></form>
	</td>
  </tr>
</table>

</body>
</html>
