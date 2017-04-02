option explicit
dim URL, strXML, xmlhttp, RequestHeaderContentType, RequestHeaderSOAPAction, proxy, response

'============================================================
'AB HIER PARAMETER ANPASSEN
'============================================================

'URL festlegen
'Developer Tools for UPnP Technologies von http://opentools.homeip.net/ installieren
'"Device Spy.exe" starten, ein Gerät auswählen, darunter ein Service mit "Play" finden
'Help, Show Debug Information, Events, Show Information Messages
'Im Hauptfenster: Rechtsklick auf die gewünschte Aktion, Invoke Action, Parameter angeben, Invoke
'Im Debugfenster: Aus dem HTTP POST den Host samt Port und die POST Action extrahieren
'URL = "http://[Host samt Port][POST]"
URL = "http://192.168.1.103:2869/upnphost/udhisapi.dll?control=uuid:de86cf12-71e0-49da-9b09-3b85c791e300+urn:upnp-org:serviceId:AVTransport"


'RequestHeader festlegen
'Developer Tools for UPnP Technologies von http://opentools.homeip.net/ installieren
'"Device Spy.exe" starten, ein Gerät auswählen, darunter ein Service mit "Play" finden
'Help, Show Debug Information, Events, Show Information Messages
'Im Hauptfenster: Rechtsklick auf die gewünschte Aktion, Invoke Action, Parameter angeben, Invoke
'Im Debugfenster: Aus dem HTTP POST den Content Type und die SOAP Action extrahieren
RequestHeaderContentType = "text/xml; charset=""utf-8"""
RequestHeaderSOAPAction = """urn:schemas-upnp-org:service:AVTransport:1#Play"""


'XML festlegen
'Developer Tools for UPnP Technologies von http://opentools.homeip.net/ installieren
'"Device Spy.exe" starten, ein Gerät auswählen, darunter ein Service mit "Play" finden
'Help, Show Debug Information, Events, Show Information Messages
'Im Hauptfenster: Rechtsklick auf die gewünschte Aktion, Invoke Action, Parameter angeben, Invoke
'Im Debugfenster: Aus dem HTTP POST das XML extrahieren
strXML = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf & _
	"<s:Envelope s:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf & _
	"	<s:Body>" & vbcrlf & _
	"		<u:Play xmlns:u=""urn:schemas-upnp-org:service:AVTransport:1"">" & vbcrlf & _
	"			<InstanceID>0</InstanceID>" & vbcrlf & _
	"			<Speed>1</Speed>" & vbcrlf & _
	"		</u:Play>" & vbcrlf & _
	"	</s:Body>" & vbcrlf & _
	"</s:Envelope>"


'Proxy festlegen
'"http://server:8888" oder ""
proxy = ""

'============================================================
'AB HIER NICHTS MEHR ÄNDERN
'============================================================

'Parameter anzeigen
wscript.echo "Parameter" & vbcrlf & "============================================================"
wscript.echo "URL" & vbcrlf & URL & vbcrlf & "--"
wscript.echo "Request Header Content-Type" & vbcrlf & RequestHeaderContentType & vbcrlf & "--"
wscript.echo "Request Header SOAPAction" & vbcrlf & RequestHeaderSOAPAction & vbcrlf & "--"
wscript.echo "XML" & vbcrlf & strXML & vbcrlf & "--"
wscript.echo
wscript.echo


'Befehl absetzen
wscript.echo "Befehl absetzen" & vbcrlf & "============================================================"
on error resume next
set xmlhttp = WScript.CreateObject("MSXML2.ServerXMLHTTP.6.0")
if proxy <> "" then xmlhttp.setProxy 2, proxy
xmlhttp.Open "POST", URL, False
xmlhttp.setRequestHeader "Content-Type", RequestHeaderContentType
xmlhttp.setRequestHeader "SOAPAction", RequestHeaderSOAPAction
xmlhttp.send strXML
if err.number = 0 then
	wscript.echo "Erfolgreich."
else
	wscript.echo "Fehler Nummer " & err.number & " in " & err.source & ": " & err.description
end if
on error goto 0
wscript.echo
wscript.echo


'Antwort am Bildschirm ausgeben
wscript.echo "Antwort" & vbcrlf & "============================================================"
if xmlhttp.responseXML.xml = "" then
	wscript.echo "Keine XML-Antwort erhalten."
	wscript.echo "Wenn das Absetzen des Befehls erfolgreich war: Parameter prüfen!"
else
	wscript.Echo  xmlhttp.responseXML.xml
end if
'the response is an MSXML2.DOMDocument.6.0
'set response = xmlhttp.responseXML