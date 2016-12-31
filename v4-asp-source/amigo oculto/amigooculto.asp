<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/AmigoOculto.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_AmigoOculto_STRING
  MM_editTable = "TblMembros"
  MM_editRedirectUrl = "amigooculto.asp"
  MM_fieldsStr  = "TxbNome|value"
  MM_columnsStr = "NomeMembro|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim RsMembros
Dim RsMembros_numRows

Set RsMembros = Server.CreateObject("ADODB.Recordset")
RsMembros.ActiveConnection = MM_AmigoOculto_STRING
RsMembros.Source = "SELECT * FROM TblMembros"
RsMembros.CursorType = 0
RsMembros.CursorLocation = 2
RsMembros.LockType = 1
RsMembros.Open()

RsMembros_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Amigo Oculto S&ETH;</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-color: #dcdcdc;
	margin-right: 0px;
	margin-bottom: 0px;
}
.titulo {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 36px;
	font-style: normal;
	font-weight: bold;
	text-decoration: none;
}
.texto1 {
	font-family: Georgia, "Times New Roman", Times, serif;
	font-size: 14px;
	text-decoration: none;
	font-weight: normal;
}
.form1 {
	font-family: Georgia, "Times New Roman", Times, serif;
	font-size: 14px;
	text-decoration: none;
	background-color: #EEEEEE;
	border: thin dotted #06D900;
	color: #039300;
	font-weight: bold;
}
.link1 {
	font-family: Georgia, "Times New Roman", Times, serif;
	font-size: 14px;
	color: #0033FF;
	text-decoration: none;
}
.link2 {
	font-family: Georgia, "Times New Roman", Times, serif;
	font-size: 14px;
	color: #0066FF;
	text-decoration: underline;
}
-->
</style></head>

<body>
<br><br>
<table width="779" height="522" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="img_0_0.jpg" width="153" height="108"></td>
    <td width="626" background="img_0_1.jpg" class="titulo"><div align="center">Amigo O'Cult Sala&ETH;ois </div></td>
  </tr>
  <tr>
    <td><img src="img_1_0.jpg" width="153" height="414"></td>
    <td width="626" height="414" valign="top" background="img_1_1.jpg" class="texto1"><p>Participe do amigo oculto (o'cult)  Sala&ETH;ois!! Os bilhetes ser&atilde;o sorteados pela internet e os presentes ser&atilde;o trocados no dia do churrasco! Se vc &eacute; <strong>Fair-play</strong>, n&atilde;o deixe de participar!! </p>
    <p>Para obter informa&ccedil;&otilde;es mais atualizadas, visite nossa central na <a href="http://www.orkut.com/Community.aspx?cmm=492120" class="link1" onMouseOver="this.className='link2'" onMouseOut="this.className='link1'">comunidade do orkut</a>. </p>
    <table border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
    <% If Not RsMembros.EOF Or Not RsMembros.BOF Then %>
        <td><div align="center"><strong>Fair-players inscritos<br>
          <br></strong></div></td>
    <% End If ' end Not RsMembros.EOF Or NOT RsMembros.BOF %>
        <td><div align="center"><strong>Sou fair-play e quero participar! <br>
          <br></strong></div></td>
      </tr>
      <tr>
    <% If Not RsMembros.EOF Or Not RsMembros.BOF Then %>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
		  Do Until RsMembros.EOF
		  	%>
          <tr>
            <td><div align="center"><%=RsMembros("NomeMembro")%></div></td>
          </tr>
	      <%
		  	RsMembros.MoveNext()
		  Loop
		  %>
    </table>
		</td>
    <% End If ' end Not RsMembros.EOF Or NOT RsMembros.BOF %>
        <td valign="middle"><form action="<%=MM_editAction%>" method="POST" name="form1" id="">
          <div align="center">Nome: 
              <input name="TxbNome" type="text" class="form1" id="TxbNome" size="22" maxlength="30">
              <input type="submit" class="form1" value="Ok!">
          </div>
        
          <input type="hidden" name="MM_insert" value="form1">
        </form></td>
      </tr>
    </table>    
    <p>&nbsp;</p></td>
  </tr>
</table>

</body>
</html>
<%
RsMembros.Close()
Set RsMembros = Nothing
%>
