<%'Response.CharSet="UTF-8"%> 
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/google_conn.asp"-->

<body>
    <form name="googlesync" action="googlepost_full.asp" method="post">
  <%
Function Sint(nr)
if IsNumeric(nr) then
Sint = cint(nr)
else
Sint = -1
end if
end function
Function CompareName(name1,name2)
if name1 <> "" and name2 <> "" then
name1 = SForm(name1)
name2 = SForm(name2)
if not Instr(name2,"a/s") AND Instr(name1,"a/s") then
name2 = Replace(name2,"a/s","")
name2 = trim(name2)
elseif not instr(name2,"aps") AND instr(name2,"aps") then
name2 = Replace(name2,"aps","")
name2 = Trim(name2)
end if

if Instr(name2," ") then
char = " "
elseif Instr(name2,"-") then
char = "-"
elseif Instr(name2,"/") then
char = "/"
elseif Instr(name2,",") then
char = ","
elseif Instr(name2,".") then
char = "."
end if
if char <> "" then
name2Split = split(name2,char)
splitpoint = cint(1/ubound(name2Split))
For i = 0 to ubound(name2Split)
if Instr(name1,name2Split(i)) then
splitpoints = cint(splitpoints + splitpoint)
end if
next
if splitpoints => 0.5 then
CompareName = splitpoints
else
CompareName = 0
end if
else
if Instr(name1,name2) then
CompareName = 1
end if
end if
else
CompareName = 0
end if
End Function
Function SForm(streng)
if not IsNull(streng) then
streng = Server.HTMLEncode(Lcase(Trim(streng)))
streng = replace(streng,"æ",Chr(230))
streng = replace(streng,"Ã¦",Chr(230))
streng = replace(streng,"ø",Chr(248))
streng = replace(streng,"Ã¸",Chr(248))
streng = replace(streng,"å",Chr(229))
streng = replace(streng,"Ã¥",Chr(229))
streng = replace(streng,"amp;","")
streng = replace(streng,"'","''")
SForm = streng
end if
end function
Function AddPoint(id,point)
if point > 0 then
if Instr(result,id) then
'skil strengen ad ved id output streng bliver ...|id,point|(split),point|...
SplitResult = split(result,id)
'split strengen ved | output streng bliver ,point(split)id,point|...
SplitSplitResult = split(SplitResult(1),"|")
'fjern 1 længde fra højre output bliver point
existpoint = right(SplitSplitResult(0),cint(len(SplitSplitResult(0))-1))
'smid den brugte delstreng ud
SplitSplitResult(0) = ""
'tilføj point
newpoint = cint(existpoint + point)
'join strengen output bliver (...|id,point|)&(id,point|...)&(|id,point)
thisresult = SplitResult(0) & id & "," & newpoint & join(SplitSplitResult,"|")   
result = thisresult
else
result = result & "|"&id&","&point
end if
end if
end function
Response.Write result
Function MyMatch(Match)
If (Match = "true") then
Response.Write (Match & ": true<br>")
Match = Null
end if
End Function
Function fromGT(tidsstreng)
	Tsplit = split(tidsstreng,"T")
    klokkeslet = left(Tsplit(1),8)
    DatoSplit = split(Tsplit(0),"-")
    dato = DatoSplit(2) & "-" & DatoSplit(1) & "-" & DatoSplit(0)
	fromGT = dato & " " & klokkeslet
end function
Function arrFromGT(tidsstreng)
arrFromGT = split(fromGT(tidsstreng)," ")
end function
Function toGT(klokkeslet, dato)
if Instr(klokkeslet," ") then
splitklokken = split(klokkeslet," ")
klokkeslet = splitklokken(1)
end if
    DatoSplit = split(dato,"-")
    Gdate = DatoSplit(2) & "-" & DatoSplit(1) & "-" & DatoSplit(0)
    Gtime = klokkeslet & ".000+02.00"
    toGT = Gdate & "T" & Gtime
end function

'Påbegynd synkronisering
if Sint(request.Form("unsync")) > 0 then


'Split DataCells fra nyposter op
ArrTitles = Split(Request.Form("titles"),"][#TO:DataCell:")
ArrDetails = Split(Request.Form("details"),"][#TO:DataCell:")
ArrAdress = Split(Request.Form("adresses"),"][#TO:DataCell:")
ArrTimes = Split(Request.Form("times"),"][#TO:DataCell:")
ArrParticipants = Split(Request.Form("participants"),"][#TO:DataCell:")
'split DataCells fra gamle poster op
ArrsTitles = Split(Request.Form("stitles"),"][#TO:DataCell:")
ArrsDetails = Split(Request.Form("sdetails"),"][#TO:DataCell:")
ArrsAdress = Split(Request.Form("sadresses"),"][#TO:DataCell:")
ArrsTimes = Split(Request.Form("stimes"),"][#TO:DataCell:")
ArrsId = Split(Request.Form("sids"),"][#TO:DataCell:")
ArrsParticipants = Split(Request.Form("sparticipants"),"][#TO:DataCell:")

For Ogp = 1 to cint(Request.Form("existsync"))

'definer det nuværende indhold af strengene
if (ArrsTitles(Ogp) <> "") and (ArrsTitles(Ogp) <> "undefined") then
Titles = SForm(ArrsTitles(Ogp))
else
Titles = ""
end if

if (ArrsDetails(Ogp) <> "") and (ArrsDetails(Ogp) <> "undefined") then
Details = SForm(ArrsDetails(Ogp))
else
Details = ""
end if

if (ArrsAdress(Ogp) <> "") and (ArrsAdress(Ogp) <> "undefined") then
Adress = SForm(ArrsAdress(Ogp))
Adress = Right(Adress, cint(len(Adress)-19))
else
Adress = ""
end if

if (ArrsParticipants(Ogp) <> "") and (ArrsParticipants(Ogp) <> "undefined") then
Participants = Split(ArrsParticipants(Ogp),"][#TO:PartDataString:")
else
Participants = ""
end if

if (ArrsTimes(Ogp) <> "") and (ArrsTimes(Ogp) <> "undefined") then
Times = Split(ArrsTimes(Ogp),"][#TO:PartDataString:")
else
Times = ""
end if

'split startdato
start = arrFromGT(Times(1))
startdato = start(0)
startklokkeslet = start(1)
    
'split slutdato
slut = arrFromGT(Times(2))
slutdato = slut(0)
slutklokkeslet = slut(1)

strUpdSQL = "UPDATE crmhistorik SET navn ='"& Titles &"', komm = '" & Details &"', crmdato = '" & startdato &"', crmklokkeslet = '" & startklokkeslet &"', crmdato_slut = '" & slutdato &"', crmklokkeslet_slut = '" & slutklokkeslet &"' WHERE id = "& ArrsId(Ogp) &""
oConn.Execute(strUpdSQL)
Next

'For hver post gør følgende:
For Ngp = 1 to cint(Request.Form("unsync"))

'definer det nuværende indhold af strengene
if (ArrTitles(Ngp) <> "") and (ArrTitles(Ngp) <> "undefined") then
Titles = SForm(ArrTitles(Ngp))
else
Titles = ""
end if

if (ArrDetails(Ngp) <> "") and (ArrDetails(Ngp) <> "undefined") then
Details = SForm(ArrDetails(Ngp))
else
Details = ""
end if

if (ArrAdress(Ngp) <> "") and (ArrAdress(Ngp) <> "undefined") then
Adress = SForm(ArrAdress(Ngp))
Adress = Right(Adress, cint(len(Adress)-19))
else
Adress = ""
end if
if (ArrParticipants(Ngp) <> "") and (ArrParticipants(Ngp) <> "undefined") then
Participants = Split(ArrParticipants(Ngp),"][#TO:PartDataString:")
else
Participants = ""
end if

'tilpasning af tider
if ArrTimes(Ngp) <> "" then
Times = Split(ArrTimes(Ngp),"][#TO:PartDataString:")
'split startdato
start = arrFromGT(Times(1))
startdato = start(0)
startklokkeslet = start(1)
    
'split slutdato
slut = arrFromGT(Times(2))
slutdato = slut(0)
slutklokkeslet = start(1)
end if

result = " "

'Tjek om Hotkey er tilstede
if Instr(Details,"[#to:") then
SplitDetails = Split(Details,"[#to:")
SplitSplitDetails = Split(SplitDetails(1),"]")
Hotkey = SplitSplitDetails(0)

CountHotkeyResults = "SELECT COUNT(*) AS nr FROM kunder where (lCase(Kkundenavn) = '" & Hotkey & "') OR (Kkundenr = '" & Hotkey & "') OR (lCase(email) = '" & Hotkey &"') OR (telefon = '" & Hotkey & "')"
SearchByHotkey = "select Kid, Kkundenavn from kunder where (lCase(Kkundenavn) = '" & Hotkey & "') OR (Kkundenr = '" & Hotkey & "') OR (lCase(email) = '" & Hotkey &"') OR (telefon = '" & Hotkey & "')"
set HotkeyResultsSet = oConn.Execute(CountHotkeyResults)
HotkeyResults = HotkeyResultsSet("nr")
if HotkeyResults = 1 then
oRec.open SearchByHotkey, oConn, 3
if not oRec.EOF then
				IntKundeId = oRec("Kid")
				kundenavn = oRec("Kkundenavn")
				Command = "insert"
oRec.close
end if
elseif HotkeyResults > 1 then
oRec.open SearchByHotkey, oConn, 3
while not oRec.EOF
               kundenavne = kundenavne & "|" & oRec("Kid")
               Command = "prompt"
oRec.movenext
wend
oRec.close
else
CountLikeHotkeyResults = "SELECT COUNT(*) FROM kunder WHERE KKundenavn LIKE %" & Hotkey & "% OR email LIKE "& left(Hotkey,len(Hotkey)-3) & "%"
SearchLikeHotkey = "SELECT Kid, Kkundenavn FROM kunder WHERE KKundenavn LIKE %" & Hotkey & "% OR email LIKE "& left(Hotkey,len(Hotkey)-3) & "%"
set LikeHotkeyResults = oConn.Execute(CountLikeHotkeyResults)
if LikeHotkeyResults = 1 then
oRec.open SearchLikeHotkey, oConn, 3
if not oRec.EOF then
				IntKundeId = oRec("Kid")
				kundenavn = oRec("Kkundenavn")
				Command = "insert"
end if
oRec.close
elseif LikeHotkeyResults > 1 then
oRec.open SearchLikeHotkey, oConn, 3
while not oRec.EOF
               kundenavne = kundenavne & "|" & oRec("Kid")
               Command = "prompt"
oRec.movenext
wend
oRec.close
end if
end if
else

' søg i kontaktpers
FindByEmailSQL = "select email, navn, privattlf, mobiltlf, dirtlf, kundeid from kontaktpers"
oRec.open FindByEmailSQL, oConn, 3
while not oRec.EOF

'Foretag søgning i participants
For Npp = 0 to Ubound(Participants)
CurParticipant = Lcase(trim(Participants(Npp)))
if (CurParticipant <> "") and (CurParticipant <> "undefined") then

'led efter mail i db
if (CurParticipant = Lcase(Trim(oRec("email"))) and oRec("email") <> "") and EmailLocalFound = "" then
Call AddPoint(oRec("kundeid"),5)
EmailContactMatch = "true"
EmailLocalFound = "true"
end if

'led efter navn db
if (Instr(Trim(oRec("navn")),CurParticipant) and CurParticipant <> "") and NameLocalFound = "" then
Call AddPoint(oRec("kundeid"),cint(len(CurParticipant) * 0.15))
NameContactMatch = "true"
NameLocalFound = "true"
end if

'led i Details efter tlf numre tilknyttet kontaktpersoner
if (Instr(Details,oRec("dirtlf")) and oRec("dirtlf") <> "") or (Instr(Details,oRec("mobiltlf")) and oRec("mobiltlf") <> "") or (Instr(Details,oRec("privattlf")) and oRec("privattlf") <> "") and PhoneLocalFound = "" then
Call AddPoint(oRec("kundeid"),5)
PhoneLocalFound = "true"
end if
end if
Next
'nulstil lokale variabler til forhindring af dobbeltpointgivning
NameLocalFound = Null
PhoneLocalFound = Null
EmailLocalFound = Null
oRec.movenext
wend
oRec.close


'Foretag søgning i kunder
FindByKunde = "select Kid, email, Kkundenavn, postnr, adresse, city, kpersmobil1, kpersmobil2, kpersmobil3, kpersmobil4, kpersmobil5, kperstlf1, kperstlf2, kperstlf3, kperstlf4, kperstlf5,"
FindByKunde = FindByKunde & "kpersemail1, kpersemail2, kpersemail3, kpersemail4, kpersemail5, kontaktpers1, kontaktpers2, kontaktpers3, kontaktpers4, kontaktpers5 from kunder"
oRec.open FindByKunde, oConn, 3
while not oRec.EOF

'Match title med kundenavn
if Instr(Titles,trim(oRec("Kkundenavn"))) then
Call AddPoint(oRec("Kid"),5)
TitleFirmMatch = "true"
end if

'//Tjek details for indhold
If (Details <> "") and Details <> "undefined" then
'Match details med kundenavn
Call AddPoint(oRec("Kid"),(1*CompareName(Details,oRec("Kkundenavn"))))
Call AddPoint(oRec("Kid"),(1*CompareName(Details,oRec("Kkundenavn"))))
DetailFirmMatch = "true"


'Match details med telefonnumre fra firmaer
if (Instr(Details,oRec("kperstlf1")) and oRec("kperstlf1") <> "") or (Instr(Details,oRec("kperstlf2")) and oRec("kperstlf2") <> "") or (Instr(Details,oRec("kperstlf3")) and oRec("kperstlf3") <> "") or (Instr(Details,oRec("kperstlf4")) and oRec("kperstlf4") <> "") or (Instr(Details,oRec("kperstlf5")) and oRec("kperstlf5") <> "")_
or (Instr(Details,oRec("kpersmobil1")) and oRec("kpersmobil1") <> "") or (Instr(Details,oRec("kpersmobil2")) and oRec("kpersmobil2") <> "") or (Instr(Details,oRec("kpersmobil3")) and oRec("kpersmobil3") <> "") or (Instr(Details,oRec("kpersmobil4")) and oRec("kpersmobil4") <> "") or (Instr(Details,oRec("kpersmobil5")) and oRec("kpersmobil5") <> "") then
Call AddPoint(oRec("Kid"),5)
end if
end if
'//Tjek

'Foretag søgning i participants
For Npp = 0 to Ubound(Participants)
CurParticipant = trim(Participants(Npp))
if (CurParticipant <> "") and (CurParticipant <> "undefined") then
'led efter mail i db
if (SForm(CurParticipant) = SForm(oRec("kpersemail1")) and SForm(oRec("kpersemail1") <> "")) or (SForm(CurParticipant) = SForm(oRec("kpersemail2")) and oRec("kpersemail2") <> "") or (SForm(CurParticipant) = SForm(oRec("kpersemail3")) and oRec("kpersemail3") <> "") or (SForm(CurParticipant) = SForm(oRec("kpersemail4")) and oRec("kpersemail4") <> "") or (SForm(CurParticipant) = SForm(oRec("kpersemail5")) and oRec("kpersemail5") <> "") and LocalMailMatch = "" then
Call AddPoint(oRec("Kid"),5)
EmailFirmMatch = "true"
LocalMailMatch = "true"
end if

'led efter navn db
if (Instr(SForm(oRec("kontaktpers1")),SForm(CurParticipant)) and oRec("kontaktpers1") <> "") or (Instr(SForm(oRec("kontaktpers2")),SForm(CurParticipant)) and oRec("kontaktpers2") <> "") or (Instr(SForm(oRec("kontaktpers3")),SForm(CurParticipant)) and oRec("kontaktpers3") <> "") or (Instr(SForm(oRec("kontaktpers4")),SForm(CurParticipant)) and oRec("kontaktpers4") <> "") or (Instr(SForm(oRec("kontaktpers5")),SForm(CurParticipant)) and oRec("kontaktpers5") <> "") and LocalNameMatch = "" then
Call AddPoint(oRec("Kid"),1)
NameFirmMatch = "true"
LocalNameMatch = "true"
end if
end if
next
'Foretag søgning i Adress hvis adresse er mere end inting
if (Adress <> "") And (Adress <> "undefined") And oRec("adresse") <> "" then
AdresseMatchPoint = cint(AdresseMatchPoint)
Adress = Replace(Adress,".","")
Adress = Replace(Adress,",","")
Adress = SForm(Adress)
CompAdress = Replace(oRec("adresse"),".","")
CompAdress = Replace(CompAdress,",","")
CompAdress = SForm(CompAdress)
CompAdress2 = Replace(CompAdress,"."," ")
CompAdress2 = Replace(CompAdress2,","," ")
CompAdress2 = SForm(CompAdress2)
AddrPostNr = SForm(oRec("postnr"))
AddrCity = SForm(oRec("city"))
if (Instr(Adress,CompAdress) and CompAdress <> "") OR (Instr(Adress,CompAdress2) and CompAdress2 <> "") then
AdresseMatchPoint = 2
end if
if ((Instr(Adress, AddrPostNr) and AddrPostNr <> "") Or (instr(Adress, AddrCity) and  AddrCity <> "")) then
AdresseMatchPoint = int(AdresseMatchPoint^2)
end if
if AdresseMatchPoint > 0 then
AdresseMatchPoint = int(AdresseMatchPoint^1.5)
Call AddPoint(oRec("Kid"),AdresseMatchPoint)
AdresseMatchPoint = 0
AdresseFirmMatch = "true"
end if
end if

'nulstil lokale variable til forhinding af dobbeltpointgivning
LocalNameMatch = Null
LocalMailMatch = Null
resultSum = 0
oRec.movenext
wend
oRec.close

'returnering af strenge
Response.Write "Title: " & Titles & "<br>"
Response.Write "Details: " & Details & "<br>"
Response.Write "participants: " & join(Participants,",") & "<br>"
Response.Write "Adress: " & Adress

Response.Write "<br>resultat = " & result & "<br>"

ResultList = split(result,"|")
For i = 1 to Ubound(ResultList)
CurrPoint = split(ResultList(i),",")
ResultPoint = cint(CurrPoint(1))
Response.Write "point = " & ResultPoint
'procedure til at tælle point
if ResultPoint > 15 then
Command = "insert"
IntKundeId = CurrPoint(0)
exit for
elseif ResultPoint > 0 then
Command = "prompt"
kundenavne = kundenavne & "|" & CurrPoint(0)
end if
next
'luk if paranteser der har med Hotkey at gøre
end if

'registrer kommando fra pointgivning
if Command = "insert" then
response.Write "insert går igennem"
If (AdresseFirmMatch = "true") then
dbAdresse = Adress
else
'dbAdresse =find i db
end if
If EmailFirmMatch = "true" then
dbEmail = procedure
else
'dbEmail = find i db
end if
If Not DetailFirmMatch = "true" then
'dbDetails = husk at skrive firmanavnet ind i googleposten
end if
oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, komm, navn, kundeid) VALUES('" & session("user") & "', '"& strDato &"', '"& startdato &"', '"& startklokkeslet &"', '"& slutklokkeslet &"', '"& slutdato &"', '"& Replace(Details,"'","''") &"', '"& Titles &"', "& intKundeId &")")
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close%>
			<img src="anything.gif" style="position: absolute; top: -1000px;" />
<script language="javascript">var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';var searchText = <%="'"&Titles&"'"%>;var query = new google.gdata.calendar.CalendarEventQuery(feedUri);query.setFullTextQuery(searchText);var eventFound = false;var callback = function(result) {var entries = result.feed.entry;if (entries.length > 0) {var event = entries[0]; var GoogleAktionsId = new google.gdata.ExtendedProperty();GoogleAktionsId.setName('aktionsid'); GoogleAktionsId.setValue(<%="'"&aktionsid&"'"%>); event.addExtendedProperty(GoogleAktionsId); event.updateEntry(function(result) {document.getElementById('information').innerHTML += ('event updated! - extq =' + query.getParam('aktionsid'));},handleError);} else {document.getElementById('information').innerHTML += ('Der opstod en fejl i kommunikationen med google. Handlingen er ikke redigeret på google\'s server. Der kræves en manuel redigering af posten.');}}; var handleError = function(error) {document.getElementById('information').innerHTML += (error);}; calendarService.getEventsFeed(query, callback, handleError);</script>
<%
elseif Command = "prompt" then
NGPsToCheck = NGPsToCheck & "," & Ngp
ConfirmPosts = "j"
    %>
    Timeout Intellisync er i tvivl om hvilke af disse kunder der passer til din googlepost:
    <select name="selCustomer_<%=Ngp%>" id="selCustomer_<%=Ngp%>">
        <% PotentialCustomers = split(kundenavne,"|")
For i = 1 to uBound(PotentialCustomers)
strFindPotentialSQL = "select Kid, KKundenavn from kunder where Kid = " & PotentialCustomers(i)
oRec.open strFindPotentialSQL, oConn, 3
if not oRec.EOF then%>
<option value="<%=oRec("Kid") %>"><%=oRec("KKundenavn")%></option>
<%oRec.close
end if
next %>
    </select><br />
    <%
kundenavne = Null
else
Response.Write "Timeout Intellisync var ude af stand til at matche kunde-/kontaktpersonsoplysninger på din googlepost"
end if

MyMatch(AdresseFirmMatch)
MyMatch(NameFirmMatch)
MyMatch(NameContactMatch)
MyMatch(DetailFirmMatch)
MyMatch(DetailContactMatch)
MyMatch(TitleFirmMatch)
MyMatch(EmailContactMatch)
MyMatch(EmailFirmMatch)
'nulstil matchvariabler
AdresseFirmMatch = Null
NameContactMatch = Null
DetailContactMatch = Null
DetailFirmMatch = Null
TitleFirmMatch = Null
EmailContactMatch = Null
EmailFirmMatch = Null
AdresseFirmMatch = Null
NameFirmMatch = Null
'nulstil andre variable
result = Null
PotentialCustomers = Null
ResultList = Null
Adress = Null
Participants = Null
Details = Null
Titles = Null
Command = Null
ResultPoint = Null
AdresseMatchPoint = 0
Next


elseif Request.Form("confirmposthidden") = "j" then%>
<%
'Split DataCells fra nyposter op
ArrTitles = Split(Request.Form("titles"),"][#TO:DataCell:")
ArrDetails = Split(Request.Form("details"),"][#TO:DataCell:")
ArrTimes = Split(Request.Form("times"),"][#TO:DataCell:")

IdsToConfirm = split(Request.Form("confirmposts"),",")
'start indsæt procedure med de udvalgte ID's
For i = 1 to Ubound(IdsToConfirm)

'definer det nuværende indhold af strengene
if (ArrTitles(IdsToConfirm(i)) <> "") and (ArrTitles(IdsToConfirm(i)) <> "undefined") then
Titles = Server.HTMLEncode(ArrTitles(IdsToConfirm(i)))
else
Titles = ""
end if

if (ArrDetails(IdsToConfirm(i)) <> "") and (ArrDetails(IdsToConfirm(i)) <> "undefined") then
Details = Server.HTMLEncode(ArrDetails(IdsToConfirm(i)))
else
Details = ""
end if

IntKundeId = Request.Form("selCustomer_"&IdsToConfirm(i))
if ArrTimes(IDsToConfirm(i)) <> "" then
Times = Split(ArrTimes(IDsToConfirm(i)),"][#TO:PartDataString:")
'split startdato
start = arrFromGT(Times(1))
startdato = start(0)
startklokkeslet = start(1)
'split slutdato
slut = arrFromGT(Times(2))
slutdato = slut(0)
slutklokkeslet = start(1)

end if
oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, komm, navn, kundeid) VALUES('" & session("user") & "', '"& strDato &"', '"& startdato &"', '"& startklokkeslet &"', '"& slutklokkeset &"', '"& slutdato &"', '"& Replace(Details,"'","''") &"', '"& Titles &"', "& IntKundeId &")")

			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
			response.Write ArrTitles(IdsToConfirm(i))
%>
<img src="anything.gif" style="position: absolute; top: -1000px;" />
<script language="javascript">var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';var searchText = <%="'"&Replace(ArrTitles(IdsToConfirm(i)),"'","\'")&"'"%>;var query = new google.gdata.calendar.CalendarEventQuery(feedUri);query.setFullTextQuery(searchText);var eventFound = false;var callback = function(result) {var entries = result.feed.entry;if (entries.length > 0) {var event = entries[0]; var GoogleAktionsId = new google.gdata.ExtendedProperty();GoogleAktionsId.setName('aktionsid'); GoogleAktionsId.setValue(<%="'"&aktionsid&"'"%>); event.addExtendedProperty(GoogleAktionsId); event.updateEntry(function(result) {document.getElementById('information').innerHTML += ('event updated! - extq =' + query.getParam('aktionsid'));},handleError);} else {document.getElementById('information').innerHTML += ('Der opstod en fejl i kommunikationen med google. Handlingen er ikke redigeret på google\'s server. Der kræves en manuel redigering af posten.');}}; var handleError = function(error) {document.getElementById('information').innerHTML += (error);}; calendarService.getEventsFeed(query, callback, handleError);</script>
 <%
Next
%>  
<%end if%>

    <script language="javascript">
      
          function logMeIn() {
              myService = new google.gdata.calendar.CalendarService('TimeoutCal');
              scope = "http://www.google.com/calendar/feeds/";
              var token = google.accounts.user.login(scope);
              if (google.accounts.user.getInfo()) {
              document.getElementById("information").innerHTML += ('du er logget ind');}
              else {
              document.getElementById("information").innerHTML += ('du er ikke logget ind');}
              document.getElementById("information").innerHTML += ('<br>' + google.accounts.user.checkLogin(scope) + "<br>");
              usercredentials = myService.getUserCredentials();
              document.getElementById("information").innerHTML += (usercredentials + "," + usercredentials[1]);
          }
          function setupMyService() {
              myService = new google.gdata.calendar.CalendarService('TimeoutCal');
              logMeIn();
          }
          function logMeOut() {
              google.accounts.user.logout();  
          }
          function InsertPost() {
          }
          function getAllEvents() {
              // Create the calendar service object
              var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');

              // The default "private/full" feed is used to retrieve events from
              // the primary private calendar with full projection
              var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';
              
              // The callback method that will be called when getEventsFeed() returns feed data
              var callback = function(result) {
                  // Obtain the array of CalendarEventEntry
                  var entries = result.feed.entry;
                  var Unsynchronized = 0;
                  var Synchronized = 0;
                  for (var i = 0; i < entries.length; i++) {
                      var event = entries[i];   
                 if(event.getExtendedProperties().length == 0){
                Unsynchronized ++;
                document.getElementById("information").innerHTML +=('Event title = ' + event.getTitle().getText());
                document.googlesync.titles.value += ('][#TO:DataCell:' + event.getTitle().getText());
                document.googlesync.details.value += ('][#TO:DataCell:' + event.getContent().getText());
                //fang locations array
                document.googlesync.adresses.value += ('][#TO:DataCell:');
                for(var l = 0; l < event.getLocations().length; l++){
                document.googlesync.adresses.value += ('][#TO:PartDataCell:' + event.getLocations()[l].getValueString());}
                //fang times
                document.googlesync.times.value += ('][#TO:DataCell:');
                for(var t = 0; t < event.getTimes().length; t++){ 
                document.googlesync.times.value += ('][#TO:PartDataString:' + google.gdata.DateTime.toIso8601(event.getTimes()[t].getStartTime()));
                document.googlesync.times.value += ('][#TO:PartDataString:' + google.gdata.DateTime.toIso8601(event.getTimes()[t].getEndTime()));}
                //fang participants
                document.googlesync.participants.value += ('][#TO:DataCell:');
                for(var p = 0; p < event.getParticipants().length; p++){
                document.googlesync.participants.value += ('][#TO:PartDataString:' + event.getParticipants()[p].getValueString());
                document.googlesync.participants.value += ('][#TO:PartDataString:' + event.getParticipants()[p].getEmail());}
                document.getElementById("events").innerHTML += (' - this event has extended property: ' + event.getExtendedProperties().length + '<br>');
                }
                else{
                document.googlesync.sids.value += ('][#TO:DataCell:' + event.getExtendedProperties()[0].getValue());
                Synchronized ++;
                document.getElementById("information").innerHTML +=('Event title = ' + event.getTitle().getText());
                document.googlesync.stitles.value += ('][#TO:DataCell:' + event.getTitle().getText());
                document.googlesync.sdetails.value += ('][#TO:DataCell:' + event.getContent().getText());
                //fang locations array
                document.googlesync.sadresses.value += ('][#TO:DataCell:');
                for(var l = 0; l < event.getLocations().length; l++){
                document.googlesync.sadresses.value += ('][#TO:PartDataCell:' + event.getLocations()[l].getValueString());}
                //fang times
                document.googlesync.stimes.value += ('][#TO:DataCell:');
                for(var t = 0; t < event.getTimes().length; t++){ 
                document.googlesync.stimes.value += ('][#TO:PartDataString:' + google.gdata.DateTime.toIso8601(event.getTimes()[t].getStartTime()));
                document.googlesync.stimes.value += ('][#TO:PartDataString:' + google.gdata.DateTime.toIso8601(event.getTimes()[t].getEndTime()));}
                //fang participants
                document.googlesync.sparticipants.value += ('][#TO:DataCell:');
                for(var p = 0; p < event.getParticipants().length; p++){
                document.googlesync.sparticipants.value += ('][#TO:PartDataString:' + event.getParticipants()[p].getValueString());
                document.googlesync.sparticipants.value += ('][#TO:PartDataString:' + event.getParticipants()[p].getEmail());}
                }
                }
                // Print the total number of events
                  document.getElementById("events").innerHTML += ('Total of ' + entries.length + ' event(s) of which ' + Unsynchronized + ' are not synchronized');
                  document.googlesync.posthidden.value = ('j');
                  document.googlesync.unsync.value = (Unsynchronized);
                  document.googlesync.existsync.value = (Synchronized);
                  document.googlesync.submit();
                  }

                // Error handler to be invoked when getEventsFeed() produces an error
                var handleError = function(error) {
                document.getElementById("events").innerHTML += (error);
              }

              // Submit the request using the calendar service object
              calendarService.getEventsFeed(feedUri, callback, handleError);}
    </script>

    <br />
    Action information:
    <div id="information">
    </div>
    <br />
    <div id="calendarTitle">
    </div>
    <div id="events">
    </div>
    <img src="anything.gif" style="position: absolute; top: -1000px;" />
    <div id="panel" />
    <input type="button" value="login" onclick="logMeIn()" /><input type="button" value="logout"
        onclick="logMeOut()" /><input type="button" value="doGetInfo" onclick="doGetInfo()" />
    <input type="button" value="insert" onclick="InsertPost()" /><input type="button"
        value="Update" onclick="UpdatePost()" /><input type="button" value="hent events"
            onclick="getAllEvents()" />
    <input type="button" value="delete" onclick="DeletePost()" />
    <%
    'skriv information til felterne
    FormTitles = Request.Form("titles")
    FormDetails = Request.Form("details")
    FormAdresses = Request.Form("adresses")
    FormTimes = Request.Form("times")
    FormParticipants = Request.Form("participants")
    
    'skriv btnValue
    if NGPsToCheck <> "" then
    btnText = "Valider oplysninger"
    btnOnClick = "submit()"
    else
    btnText = "Synkroniser"
    btnOnClick = "getAllEvents()"
    end if
    %>
    <input type="button" value="<%=btnText %>" onclick="<%=btnOnClick%>" />
    <input type="text" id="confirmposthidden" name="confirmposthidden" value="<%=ConfirmPosts%>"
        style="width:600px;" /><br />
    <input type="text" id="confirmposts" name="confirmposts" value=" <%=NGPsToCheck %>"
        style="width:600px;" /><br />
    <input type="text" id="posthidden" name="posthidden" value=" " style="width: 600px;" /><br />
    <input type="text" id="unsync" name="unsync" value=" " style="width: 600px;" /><br />
    <input type="text" id="titles" name="titles" value=" <%=FormTitles %>" style="width: 600px;" /><br />
    <input type="text" id="details" name="details" value=" <%=FormDetails %>" style="width: 600px;" /><br />
    <input type="text" id="adresses" name="adresses" value=" <%=FormAdresses %>" style="width: 600px;" /><br />
    <input type="text" id="times" name="times" value=" <%=FormTimes %>" style="width: 600px;" /><br />
    <input type="text" id="participants" name="participants" value=" <%=FormParticipants %>"
        style="width: 600px;" /><br />
    <!--data for allerede oprettede posts-->
    <input type="text" id="existsync" name="existsunsync" value=" " style="width: 600px;" /><br />
    <input type="text" id="sids" name="sids" value=" " style="width: 600px;" /><br />
    <input type="text" id="stitles" name="stitles" value=" " style="width: 600px;" /><br />
    <input type="text" id="sdetails" name="sdetails" value=" " style="width: 600px;" /><br />
    <input type="text" id="sadresses" name="sadresses" value=" " style="width: 600px;" /><br />
    <input type="text" id="stimes" name="stimes" value=" " style="width: 600px;" /><br />
    <input type="text" id="sparticipants" name="sparticipants" value=" " style="width: 600px;" /><br />
    </form>
</body>
</html>
<%strGoogleCase = Null %>