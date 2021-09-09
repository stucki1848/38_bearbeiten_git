<%
 Option Explicit
 Dim useFrames, httpHost, zielURL, domList(), domZiel()
 Dim aktHost, aktZiel, i
 ReDim domList(6)
 ReDim domZiel(6)
 
 ' ---- Einstellung -----------------------------------
 ' Sollen Frames verwendet werden?
 ' useFrames = true       Frames sind aktiviert
 ' useFrames = false      direkte Weiterleitung (standard)
 useFrames = true
 
 ' ---- Einstellung -----------------------------------
 ' Geben Sie hier Ihre Domain-Namen an. Wenn der entsprechende
 ' Domain-Platz nicht belegt ist, dann setzen Sie ein "" ein.
 ' Bitte geben Sie kein "http://" vor der Domain an.
 ' z.B. domList(4) = ""
 
 domList(0) = "www.kulturstiftung-masthoff.de"    '(Haupt-Domain) 
 domList(1) = "www.sythener-gitarrentage.de"    '(Alias-Domain 1)
 domList(2) = "kulturstiftung-masthoff.de"
 domList(3) = "sythener-gitarrentage.de"
 domList(4) = ""
 domList(5) = ""
 domList(6) = "GH38.s7.domainkunden.de" '(Hier sollten Sie Ihre Server-ID eintragen, 
                                     ' damit auch diese abgefangen wird.)

 ' ---- Einstellung -----------------------------------
 ' Geben Sie hier die Ziel-Verzeichnisse an. Es können sowohl
 ' vollständige URLs mit http:// als auch Verzeichnisse sein
 ' z.B. domZiel(4) = "http://www.domain.de/verzeichnis/"
 ' oder domZiel(4) = "verzeichnis/"
 domZiel(0) = "/ksm/"
 domZiel(1) = "/sg/"
 domZiel(2) = "/ksm/"
 domZiel(3) = "/sg/"
 domZiel(4) = ""
 domZiel(5) = ""
 domZiel(6) = "/ksm/"

 ' Variablen initialisieren
 httpHost = Request.ServerVariables("HTTP_HOST")
 zielURL = ""
 
 For i = 0 to 6
  aktHost = domList(i)
  aktZiel = domZiel(i)
  if lCase(httpHost) = lCase(aktHost) then
   zielURL = aktZiel
  end if
 Next
 
 if zielURL = "" then
  %>
   <h1>Domain nicht gefunden</h1>
   <h3>Die Domain <%= httpHost %> wurde nicht gefunden oder es liegt ein Konfigurationsfehler vor.
  <%
 else
  if useFrames = true then
   %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html><head><title><%= httpHost %></title></head><frameset rows="0,*" border="0" frameborder="0" framespacing="0"><frame src="leer" scrolling="no"><frame src="<%= zielURL %>"></frameset></html><%
  else
   response.redirect zielURL
  end if
 end if
%>