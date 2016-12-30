			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
              <tr> 
                <td height="387" background="imgs\40.gif" valign="top" align="center">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <%
					Rs.Source = "SELECT * FROM TblMenus ORDER BY NomeMenu ASC"
					Rs.Open()
					Do Until Rs.EOF
					%>
                    <tr>
                      <td align="center">
						<% If Rs("StatusMenu") = 0 Then %>
						<span class="text6mm">
						<%=Rs("NomeMenu")%>
						</span>
						<% ElseIf Rs("StatusMenu") = 1 Then %>
						<a href="<%=Rs("EnderecoMenu")%>" class="text3mm" onMouseOver="this.className='text2mm';" onMouseOut="this.className='text3mm';"><%=Rs("NomeMenu")%></a>
						<% End If %>
					  <font size="1"><br><br></font></td>
                    </tr>
					<%
					  Rs.MoveNext
					Loop
					Rs.Close()
					%>
                  </table>
                </td>
              </tr>
              <tr> 
                <td><img src="imgs/50.gif" width="126" height="24"></td>
              </tr>
              <tr> 
                <td height="100%" background="imgs\60.gif">&nbsp;</td>
              </tr>
              <tr> 
                <td><img src="imgs/70.gif" width="126" height="23"></td>
              </tr>
            </table>