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
      <form action="salvar.asp?func=jornais&cod=<%=Request.QueryString("codigo")%>" method="POST">         
        <table cellpadding=0 cellspacing=0 width="484">
          <tr align="center">
                      <td bgcolor="#DEDEDE"><font face="Verdana, Arial, Helvetica" size="5"><b><%=Request.QueryString("edicao")%></b></font><br></td></tr><tr class="text3m">
                      <td><br>
                        <table width="100%" border="0" cellspacing="2" cellpadding="2" class="text3" bgcolor="#FFFFFF">
                          <tr valign="top">
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n1"),1,1) & "</b></font>" & Mid(Request.QueryString("n1"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n2"),1,1) & "</b></font>" & Mid(Request.QueryString("n2"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n3"),1,1) & "</b></font>" & Mid(Request.QueryString("n3"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n4"),1,1) & "</b></font>" & Mid(Request.QueryString("n4"),2)%></td>
                          </tr>
                          <tr valign="top">
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n5"),1,1) & "</b></font>" & Mid(Request.QueryString("n5"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n6"),1,1) & "</b></font>" & Mid(Request.QueryString("n6"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n7"),1,1) & "</b></font>" & Mid(Request.QueryString("n7"),2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(Request.QueryString("n8"),1,1) & "</b></font>" & Mid(Request.QueryString("n8"),2)%></td>
                          </tr>
                        </table>
                        <br>
                      </td></tr><tr>
                      <td align="left" class="text3" bgcolor="#DEDEDE"><%=Request.QueryString("autor")%></td></tr></table>
                  <br><img src="imgs\r484.gif" width="484" height="50"><br>
        Se o jornal estiver maior do que 484 pixels, volte e modifique alguma 
        coisa de modo a ficar ajustado ao tamanho m&aacute;ximo permitido (484 
        px) <br>
        <br>
        <br><input type="hidden" name="Codigo" value="<%=Request.QueryString("codigo")%>"><input type="hidden" name="Edicao" value="<%=Request.QueryString("edicao")%>"><input type="hidden" name="Autor" value="<%=Request.QueryString("autor")%>"><input type="hidden" name="N1" value="<%=Request.QueryString("n1")%>"><input type="hidden" name="N2" value="<%=Request.QueryString("n2")%>"><input type="hidden" name="N3" value="<%=Request.QueryString("n3")%>"><input type="hidden" name="N4" value="<%=Request.QueryString("n4")%>"><input type="hidden" name="N5" value="<%=Request.QueryString("n5")%>"><input type="hidden" name="N6" value="<%=Request.QueryString("n6")%>"><input type="hidden" name="N7" value="<%=Request.QueryString("n7")%>"><input type="hidden" name="N8" value="<%=Request.QueryString("n8")%>"><input type="hidden" name="deljornal" value="<%=Request.QueryString("deljornal")%>"><input type="hidden" name="Preview" value=""><br><input type="submit" class="botao1" value="Salvar"></form>
	</td>
  </tr>
</table>

</body>
</html>
