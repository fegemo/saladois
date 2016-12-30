<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Dicion&aacute;rio</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF" text="#000000">
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <!--#include file="topo.htm" -->
  <tr> 
    <td> 
      <!--#include file="banner.asp" -->
    </td>
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
                <td height="17" background="imgs/41.gif" class="text3mm" align="center">.: 
                  Dicion&aacute;rio :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text3"> 
                  <p><a href="dicionario.asp?func=letra&cod=a" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">A</a> | <a href="dicionario.asp?func=letra&cod=b" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">B</a> | <a href="dicionario.asp?func=letra&cod=c" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">C</a> | <a href="dicionario.asp?func=letra&cod=d" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">D</a> | <a href="dicionario.asp?func=letra&cod=e" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">E</a> | <a href="dicionario.asp?func=letra&cod=f" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">F</a> | <a href="dicionario.asp?func=letra&cod=g" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">G</a> | <a href="dicionario.asp?func=letra&cod=h" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">H</a> | <a href="dicionario.asp?func=letra&cod=i" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">I</a> | <a href="dicionario.asp?func=letra&cod=j" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">J</a> | <a href="dicionario.asp?func=letra&cod=k" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">K</a> | <a href="dicionario.asp?func=letra&cod=l" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">L</a> | <a href="dicionario.asp?func=letra&cod=m" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">M</a> | <a href="dicionario.asp?func=letra&cod=n" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">N</a> | <a href="dicionario.asp?func=letra&cod=o" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">O</a> 
                    | <a href="dicionario.asp?func=letra&cod=p" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">P</a> | <a href="dicionario.asp?func=letra&cod=q" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">Q</a> | <a href="dicionario.asp?func=letra&cod=r" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">R</a> | <a href="dicionario.asp?func=letra&cod=s" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">S</a> | <a href="dicionario.asp?func=letra&cod=t" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">T</a> | <a href="dicionario.asp?func=letra&cod=u" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">U</a> | <a href="dicionario.asp?func=letra&cod=v" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">V</a> | <a href="dicionario.asp?func=letra&cod=w" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">W</a> | <a href="dicionario.asp?func=letra&cod=x" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">X</a> | <a href="dicionario.asp?func=letra&cod=y" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">Y</a> | <a href="dicionario.asp?func=letra&cod=z" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">Z</a></p>
                  <p>&nbsp;</p>
				  <%
					Select Case Request.QueryString("func")
					  Case "letra"
						Rs.Source = "SELECT * FROM TblPalavras ORDER BY TextoPalavra ASC"
						Rs.Open()
						Minimo = False
						%><font class="text3mm"><%=UCase(Request.QueryString("cod"))%></font><br><br><%
						Do Until Rs.EOF
						  If lCase(Mid(Rs("TextoPalavra"),1,1)) = lCase(Request.QueryString("cod")) Then
							%><a href="dicionario.asp?func=uma&cod=<%=Rs("CodigoPalavra")%>" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';"><%=Rs("TextoPalavra")%></a><br><%
							Minimo = True
						  End If
						  Rs.MoveNext
						Loop
						If Not Minimo Then
						  %><br>Ainda não existem palavras cadastradas no banco de dados com a determinada letra<br><br><%
						End If
						Rs.Close()
					  Case "todas"
						Rs.Source = "SELECT * FROM TblPalavras ORDER BY TextoPalavra ASC"
						Rs.Open()
						If Not Rs.EOF Then
						  %><br><br><%
						  Do Until Rs.EOF
							%><a href="dicionario.asp?func=uma&cod=<%=Rs("CodigoPalavra")%>" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';"><%=Rs("TextoPalavra")%></a><br><%
							Rs.MoveNext
						  Loop
						Else
						  %><br>Ainda não existem palavras cadastradas no banco de dados<br><br><%
						End If
						Rs.Close()
					  Case "uma"
						Rs.Source = "SELECT * FROM TblPalavras WHERE CodigoPalavra=" & Request.QueryString("cod")
						Rs.Open()
						%><br>
                  <table width="90%" border="0" cellspacing="2" cellpadding="0">
                    <tr valign="top" bgcolor="#DEDEDE"> 
                      <td width="20%" class="text2" bgcolor="#DEDEDE"> 
                        <div align="right"><%=Rs("TextoPalavra")%> <font class="text4">::</font> </div>
                      </td>
                      <td width="80%" class="text3">&nbsp;<%=Rs("SignificadoPalavra")%></td>
                    </tr>
                  </table><br><br><a href="dicionario.asp" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';">Voltar</a>
                  <br><%
						Rs.Close()
					  Case ""
						%>
                  &quot;Conhecer o idioma nacional e usa-lo com propriedade hoje 
                  me dia n&atilde;o &eacute; s&oacute; uma quest&atilde;o de educa&ccedil;&atilde;o. 
                  &Eacute; uma verdadeira necessidade, que ultrapassa os limites 
                  da escola e da forma&ccedil;&atilde;o acad&ecirc;mica, para 
                  atingir a qualidade de vida dos cidad&atilde;os. Aquele que 
                  encontra dificuldades em escrever, em falar est&aacute; despreparado 
                  para a luta di&aacute;ria na conquista de seu bem-estar e da 
                  aceita&ccedil;&atilde;o social. N&atilde;o consegue expressar 
                  seu pensamento, suas opini&otilde;es s&atilde;o incompreendidas, 
                  suas reivindica&ccedil;&otilde;es passam despercebidas, enfim, 
                  todo seu processo de comunica&ccedil;&atilde;o com os outros 
                  fica prejudicado. Al&eacute;m da ordena&ccedil;&atilde;o l&oacute;gica 
                  das id&eacute;ias, falta-lhe a no&ccedil;&atilde;o precisa do 
                  significado das palavras, ou n&atilde;o encontra a palavra certa 
                  para traduzir o significado que tem em mente. Quando escreve, 
                  o desconhecimento da grafia, da acentua&ccedil;&atilde;o torna 
                  rid&iacute;culas suas tentativas&quot;.<br>
                  (Dicion&aacute;rio Brasileiro GLOBO)
                  <p> Com toda a certeza, podemos aplicar essa perspectiva &agrave; 
                    nossa realidade. Toda a constru&ccedil;&atilde;o de um texto 
                    depende da prepara&ccedil;&atilde;o do autor em termos de 
                    gram&aacute;tica, produ&ccedil;&atilde;o de texto, e por que 
                    n&atilde;o a pr&oacute;pria ling&uuml;&iacute;stica carregada 
                    de aspectos etimol&oacute;gicos? O Dicion&aacute;rio Sala&ETH;ois 
                    tem por objetivo utilizar esses aspectos que formam a r&iacute;gida 
                    cobran&ccedil;a que a sociedade moderna nos faz e disponibilizar 
                    uma s&eacute;rie de voc&aacute;bulos que formam toda a prolixidade 
                    dos indiv&iacute;duos que formam esse grupo. No entanto, apresentamos 
                    um modelo diferente de dicion&aacute;rio, sendo este personalizado, 
                    de modo &agrave; dar maior explica&ccedil;&atilde;o dos termos 
                    aqui publicados. Agora lhe fazemos um convite para que acrescente 
                    tais termos em seu dicion&aacute;rio particular e desfrute 
                    da grande prestigiosidade &agrave; qual ser&aacute; submetido. 
                    Boa leitura!<br>
                  </p>
                  <br><br><%
						Rs.Source = "SELECT * FROM TblPalavras"
						Rs.Open()
						Contador = 0
						Do Until Rs.EOF
						  Contador = Contador + 1
						  Rs.MoveNext
						Loop
						%><font class="text4"><%=Contador%> palavra(s) cadastrada(s)</font><br><br><%
						Rs.Close()						
					  Case Else
						%><br><br>Houve um erro durante a operação. Volte e tente novamente<br><br><%
					End Select
				  %>
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
