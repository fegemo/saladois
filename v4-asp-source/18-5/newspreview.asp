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
<table align="center" width="100%" height="100%" class="text2">
  <tr>
	<td align="center">
	  Preview:<br><br>
      <form action="salvar.asp?func=newspro&cod=<%=Request.QueryString("Codigo")%>" method="POST">
        <table width="383" border="0" cellspacing="0" cellpadding="0">
      	<tr> 
      	  <td colspan="2" background="imgs/news/00.gif"> 
      		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="55">
              <tr> 
                <td width="50%">&nbsp;</td>
                <td class="text3" valign="top"><br><%=Request.QueryString("Titulo")%></td>
              </tr>
            </table>
          </td>
            <td width="55" height="55" bgcolor="#DEDEDE"><img src="<%=Request.QueryString("Avatar")%>" width="55" height="55"></td>
        </tr>
        <tr> 
          <td><img src="imgs/news/10.gif" width="45" height="47"></td>
          <td background="imgs/news/11.gif" width="283" valign="top" class="text3">Data:. 
                           <%=Request.QueryString("Data")%> <br>
          </td>
          <td bgcolor="#DEDEDE">&nbsp;</td>
        </tr>
        <tr> 
          <td width="45" height="95" bgcolor="#DEDEDE">&nbsp;</td>
          <td valign="top" class="text2"> 
            <p><%=Request.QueryString("Texto")%></p>
            <p>&nbsp;</p>
          </td>
          <td width="55" bgcolor="#DEDEDE">&nbsp;</td>
        </tr>
        <tr> 
          <td height="16" background="imgs/news/30.gif"></td>
          <td height="16" background="imgs/news/31.gif" align="right" class="text3">por&nbsp;&nbsp; 
          </td>
          <td height="16" bgcolor="#DEDEDE" class="text2"><%=Request.QueryString("Autor")%></td>
        </tr>
      </table><br><img src="imgs\r383.gif" width="383" height="50"><br>
        Se o news estiver maior do que 383 pixels, volte e modifique alguma coisa 
        de modo a ficar ajustado ao tamanho m&aacute;ximo permitido (383 px) <br>
        <br>
        <br><input type="hidden" name="Codigo" value="<%=Request.QueryString("codigo")%>"><input type="hidden" name="Titulo" value="<%=Request.QueryString("titulo")%>"><input type="hidden" name="Autor" value="<%=Request.QueryString("autor")%>"><input type="hidden" name="Avatar" value="<%=Request.QueryString("avatar")%>"><input type="hidden" name="Data" value="<%=Request.QueryString("data")%>"><input type="hidden" name="Texto" value="<%=Request.QueryString("texto")%>"><input type="hidden" name="delnews" value="<%=Request.QueryString("delnews")%>"><input type="hidden" name="Preview" value=""><br><input type="submit" class="botao1" value="Salvar"></form><br><br>
	</td>
  </tr>
</table>
</body>
</html>
