<%@LANGUAGE="VBSCRIPT"%> 
<html>
<head>
<title>.: SalaDois :: Fotos :.</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>

<body topmargin=0 leftmargin=0 bgcolor="#DEDEDE">
<%
If Request.QueryString("largura") = "" Or Request.QueryString("altura") = "" Or Request.QueryString("nome") = "" Then
	Response.Write("<font class='text2m'><center><br><br>Houve um erro ao calcular as matrizes vetoriais trigonométricas. Favor certificar-se da validade dos sprites isométricos</center></font>")
Else
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="132" background="imgs\fotos.gif"> 
      <div align="center"><font class="text1m"><%=Request.QueryString("texto")%></font></div>
    </td>
  </tr>
  <tr>
    <td height="100%" bgcolor="#DEDEDE">
      <div align="center"><img src="<%=Request.QueryString("nome")%>" width="<%=Request.QueryString("largura")%>" height="<%=Request.QueryString("altura")%>" alt="<%=Request.QueryString("texto")%>"></div>
    </td>
  </tr>
</table>
<%
End If
%>
</body>
</html>
