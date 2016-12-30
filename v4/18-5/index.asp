<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Index</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF" text="#000000">
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <!--#include file="topo.htm" -->
  <tr> 
    <td> 
	  <!--#include file="banner.asp" --></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td width="126" height="100%"> 
            <!--#include file="esquerda.asp" -->
          </td>
          <td width="509" height="100%"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
              <tr> 
                <td height="17" background="imgs/41.gif" class="text3m" align="center"> 
                  <%
				  Dim CodigoFrase
				  Rs.Source = "SELECT * FROM TblFrases"
				  Rs.Open()
				  Dim Contador
				  Do Until Rs.EOF
			  	    Contador = Contador + 1
			  	    Rs.MoveNext
				  Loop
				  Dim Temp
				  Randomize
				  Temp = Int(Rnd * Contador) 
				  Rs.MoveFirst
				  While Contador > Temp And Temp > 0
			  		Rs.MoveNext
					Temp = Temp - 1
				  Wend
				  CodigoFrase = Rs("CodigoFrase")
				  Rs.Close()
				  Rs.Source = "SELECT * FROM TblFrases WHERE CodigoFrase=" & CodigoFrase
				  Rs.Open()
				  %>
                  <b>"<%=Rs("TextoFrase")%>"<br>
                  (<%=Rs("AutorFrase")%>)</b> </td>
              </tr>
              <tr> 
                <td height="100%" valign="top"> 
                  <p>&nbsp;</p>
                  <%
				Rs.Close()
				Rs.Source = "SELECT * FROM TblNews ORDER BY CodigoNews DESC"
				Rs.Open()
				%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <%
					Do Until Rs.EOF 
					%>
                    <tr> 
                      <td width="73">&nbsp;</td>
                      <td width="383" valign="top"> 
                        <table width="383" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td colspan="2" background="imgs/news/00.gif"> 
                              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="55">
                                <tr> 
                                  <td width="50%">&nbsp;</td>
                                  <td class="text3" valign="top"><br>
                                    <%=Rs("TituloNews")%></td>
                                </tr>
                              </table>
                            </td>
                            <td width="55" height="55" bgcolor="#DEDEDE"><img src="<%=Server.MapPath(Rs("AvatarNews"))%>" width="55" height="55"></td>
                          </tr>
                          <tr> 
                            <td><img src="imgs/news/10.gif" width="45" height="47"></td>
                            <td background="imgs/news/11.gif" width="283" valign="top" class="text3">Data:. 
                              <%=Rs("DataNews")%> <br>
                            </td>
                            <td bgcolor="#DEDEDE">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td width="45" height="95" bgcolor="#DEDEDE">&nbsp;</td>
                            <td valign="top" class="text2"> 
                              <p><%=Rs("TextoNews")%></p>
                              <p>&nbsp;</p>
                            </td>
                            <td width="55" bgcolor="#DEDEDE">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="16" background="imgs/news/30.gif"></td>
                            <td height="16" background="imgs/news/31.gif" align="right" class="text3">por&nbsp;&nbsp; 
                            </td>
                            <td height="16" bgcolor="#DEDEDE" class="text2"><%=Rs("AutorNews")%></td>
                          </tr>
                        </table>
                      </td>
                      <td>&nbsp;</td>
                    </tr>
                    <%
						Rs.MoveNext
					Loop
					Rs.Close()
					%>
                    <tr>
                      <td><br>
                        <br>
                        <br>
                        <br>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td height="23" bgcolor="13AB13">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td height="100%"> 
            <!--#include file="direita.asp" -->
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
<script language="JavaScript">
function Enquete()
{
    window.open("espera.htm","enquete","toolbar=no,scrollbars=no,directories=no,status=no,menubar=no,resizable=yes,width=400,height=280");
}
</script>
</html>
