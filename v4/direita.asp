            <%
			Dim CodigoEnquete
			Rs.Source = "SELECT * FROM TblEnquetes WHERE StatusEnquete=1"
			Rs.Open()
			Contador = 0
			Do Until Rs.EOF
			  Contador = Contador + 1
			  Rs.MoveNext
			Loop
			Temp = 0
		    Randomize
			Temp = Int(Rnd * Contador) 
			Rs.MoveFirst
			While Contador > Temp And Temp > 0
			  Rs.MoveNext
			  Temp = Temp - 1
			Wend
			CodigoEnquete = Rs("CodigoEnquete")
			%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
              <tr> 
                <td height="387" background="imgs/62.gif" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="text3mm" align="center" colspan="2"> -: <span class="text2mm">Enquete</span> 
                        :- </td>
                    </tr>
                    <tr> 
                      <td align="center" class="text3mm" colspan="2"> 
                        <p>&nbsp;</p>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <tr> 
                      <td class="text2m" colspan="2"> 
                        <p><img src="imgs/int.gif" width="16" height="16"><span class="text3m"><%=Rs("PerguntaEnquete")%></span></p>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <form action="enquete.asp?cod=<%=CodigoEnquete%>&action=votar" method="post" target="enquete" onSubmit="Enquete">
                      <%
					  Rs.Close()
				      Rs.Source = "SELECT TblEnquetes.*, TblOpcoes.* FROM TblEnquetes INNER JOIN TblOpcoes ON TblEnquetes.CodigoEnquete = TblOpcoes.EnqueteOpcao WHERE (((TblEnquetes.CodigoEnquete)=" & CodigoEnquete & "));"
					  Rs.Open()
					  Do Until Rs.EOF
					  %>
                      <tr> 
                        <td width="16" bgcolor="#FFFFFF"> 
                          <input type="radio" name="radiobutton" value="<%=Rs("CodigoOpcao")%>">
                        </td>
                        <td bgcolor="#FFFFFF" class="text2"><%=Rs("TextoOpcao")%></td>
                      </tr>
                      <tr> 
                        <td colspan="2"><img src="imgs/falso.gif" width="1" height="4"></td>
                      </tr>
                      <%
						Rs.MoveNext
					  Loop
					  Rs.Close()
					  %>
                      <tr> 
                        <td colspan="2" align="center"> 
                          <input type="submit" value="Votar" class="botao1" onClick="Enquete();">
                        </td>
                      </tr>
                    </form>
                    <form action="enquete.asp?cod=<%=CodigoEnquete%>&action=ver" method="post" target="enquete" onSubmit="Enquete">
                      <tr> 
                        <td colspan="2" align="center"> 
                          <input type="submit" value="Resultados Parciais" class="botao1" onClick="Enquete();">
                        </td>
                      </tr>
                    </form>
                    <tr> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td><img src="imgs/52.gif" width="145" height="24"></td>
              </tr>
              <tr> 
                <td background="imgs/62.gif" height="100%" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="text3mm" align="center">-: <span class="text2mm">Links</span> 
                        :-</td>
                    </tr>
                    <tr> 
                      <td> 
                        <p>&nbsp;</p>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <tr> 
                      <td align="center"> 
                        <%
						Rs.Source = "SELECT * FROM TblLinks"
						Rs.Open()
						Do Until Rs.EOF
						  %>
                        <a href="<%=Rs("EnderecoLink")%>" class="text3" onMouseOver="this.className='text2';" onMouseOut="this.className='text3';"><%=Rs("NomeLink")%></a><br>
                        <%
						  Rs.MoveNext
						Loop
						%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td><img src="imgs/72.gif" width="145" height="23"></td>
              </tr>
            </table>
          