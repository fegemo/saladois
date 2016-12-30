<%
'########################################
'# AutoFotoAlbum 1.0
'#
'# (c) 2001 WoschNet, Osterhofen
'# www.woschnet.com
'########################################

Dim bildzahl, bildspacing, bildpadding, bildtabellenbreite, linkschriftart, linkschriftgroesse, linkschriftfarbe
Dim standarschriftfarbe, standardschriftart, standardschriftgroesse, hintergrundfarbe
Dim navitabellenbreite, navispacing, navipadding, navirahmenbreite, navirahmenfarbe, navihgfarbe

'########################################
'# Variablen definieren
'########################################

'# Anzahl der Bilder nebeneinander
bildzahl = 3

'# Navibreite
navitabellenbreite = 250
'# Zellenabstand
navispacing = 2
'# Zellenauffüllung
navipadding = 2
'# Rahmenbreite
navirahmenbreite = 2
'# Rahmenfarbe
navirahmenfarbe = "#FFFFFF"
'# Farbe Zellenhintergrund
navihgfarbe = "#666666"

'# Albumbreite
bildtabellenbreite = 500
'# Zellenabstand
bildspacing = 2
'# Zellenauffüllung
bildpadding = 3

'# Hintergrundfarbe
hintergrundfarbe = "#000097"

'# Standardschriftart
standardschriftart = "Tahoma,Verdana,Arial"
'# Standardschriftgröße
standardschriftgroesse = 12 'in Pixel
'# Standardschriftfarbe
standardschriftfarbe = "#FFFFFF"

'# Schriftart des Navigations Links
linkschriftart = "Tahoma,Verdana,Arial"
'# Schriftgröße des Navigations Links
linkschriftgroesse = 12 'in Pixel
'# Schriftfarbe des Navigations Links
linkschriftfarbe = "#99CD00"

'########################################
%>
<html>
<head>
<title>Fotogalerie</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {  font-family: <% = standardschriftart %>; font-size: <% = standardschriftgroesse %>px; color: <% = standardschriftfarbe %>}
a {  font-family: <% = linkschriftart %>; font-size: <% = standardschriftgroesse %>px; color: <% = standardschriftfarbe %>}
a:hover {  font-family: <% = linkschriftart %>; font-size: <% = standardschriftgroesse %>px; color: <% = standardschriftfarbe %>}
a:visited {  font-family: <% = linkschriftart %>; font-size: <% = standardschriftgroesse %>px; color: <% = standardschriftfarbe %>}
-->
</style>
</head>

<body bgcolor="<% = hintergrundfarbe %>">
<p>Hier k&ouml;nnte Ihr Text stehen.</p>
<p> Layout Einstellungen (Farben, Schriftarten, uvm.) nehmen Sie bitte direkt 
  im HTML Code unter dem Punkt &quot;Variablen definieren vor&quot;.</p>
<%
'########################################
'# Ab hier nichts mehr ändern !
'########################################
Dim bildpfad, bildcount

' aktuellen Ordner auslesen
strCurFolder = Request.QueryString("level")

If len(strCurFolder) = 0 Then
	strPath = "."
	bildpfad = ""
Else
	strPath = strCurFolder
	
	' Bildpfad bearbeiten
	bildpfad = RIGHT(strPath, len(strPath) - 2)
End If

set objFS = Server.CreateObject("Scripting.FileSystemObject")
set objFolder = objFS.GetFolder(Server.MapPath(strPath))

' Navigations Links erstellen
%>
<table width="<% = navitabellenbreite %>" align="center" cellpadding="<% = navipadding %>" cellspacing="<% = navispacing %>" bordercolor="<% = navirahmenfarbe %>" border="<% = navirahmenbreite %>">
  <%
for each strFolder in objFolder.SubFolders
	' versteckte Ordner ausschließen
	If left(strFolder.Name, 1) <> "-" Then
%>
  <tr>
		
    <td bgcolor="<% = navihgfarbe %>" align="center"><a href="album.asp?level=<% = strPath %>\<% = Server.URLEncode(strFolder.Name) %>">
      <% = strFolder.Name %>
      </a><br>
    </td>
	</tr>
<%
	End If
next
%>
</table>
<br>
<table width="<% = bildtabellenbreite %>" border="0" align="center" cellspacing="<% = bildspacing %>" cellpadding="<% = bildpadding %>">
	<tr>
<%
' Bilder anzeigen die im gewählten Ordner liegen
bildcount = 0

For each bild in objFolder.Files

	bildExtension = right(bild.Name, 3)
	If bildExtension = "gif" or bildExtension = "jpg" or bildExtension = "png" Then
%>
		<td align="center"><img src="<% = bildpfad %>/<% = bild.Name %>"></td>
<%
	End If
	
	If bildcount = bildzahl Then
%>
	</tr>
	<tr>
<%
	bildcount = 0
	End If
	
	bildcount = bildcount + 1
next
%>
	</tr>
</table>
</body>
</html>
