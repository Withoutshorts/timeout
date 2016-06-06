<%
response.CharSet="UTF-8"
if request.QueryString("googlecase") <> "" then
strGoogleCase = request.QueryString("googlecase")
end if %>     
     <!--#include file="../inc/connection/conn_db_inc.asp"-->
<% if strGoogleCase = "synchronize" then %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>Timeout Intellisync</title>
    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
    <script src="http://www.google.com/jsapi?key=ABQIAAAA-4DPu8eLSXMGJBa1LCchkhQH9W39-xhJgRxDecUFR532IY3yLBRMWPH6xjGJncGUW0pt06u6fNf4bw"
        type="text/javascript"></script>
    <script language="javascript">
        google.load("gdata", "1");
    </script>
</head>

  <body>
  <img src="anything.gif" style="position: absolute; top: -1000px;" />
  <div id="information"></div><br />
<div id="events"></div><br />
<% end if %>
      <script type="text/javascript">
          function logMeIn() {
              scope = "http://www.google.com/calendar/feeds/";
              var token = google.accounts.user.login(scope);
          }
          function setupMyService() {
              myService = new google.gdata.calendar.CalendarService('TimeoutCal');
              logMeIn();
          }
          function logMeOut() {
              document.getElementById("login").disabled = false;
              document.getElementById("logout").disabled = true;
              document.getElementById("synchronize").disabled = true;
              google.accounts.user.logout();
          }
          </script>
          
<%
if Trim(Request.Form("unsync")) <> "" or Trim(Request.Form("existsync")) <> "" or Trim(Request.Form("confirmposthidden")) <> "" then
strGoogleCase = "synchronize"
end if
select case strGoogleCase
case "create" %>
 </td></tr></table>
<script>

              // Create the calendar service object
              var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');

              // The default "private/full" feed is used to insert event to the
              // primary calendar of the authenticated user
              var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';

              // Create an instance of CalendarEventEntry representing the new event
              var entry = new google.gdata.calendar.CalendarEventEntry();


<%//Start bestemmelse af titel
if Request.Form("GoogleStringEmne") <> "" then
strGoogleStringEmne = Request.Form("GoogleStringEmne")
else
strGoogleStringEmne = "Andet"
end if
 if strNavn <> "" then
strTitle = ("'" & strNavn & " - " & strGoogleStringEmne & "'")
else
strTitle = "'" & strGoogleStringEmne & "'"
end if
%>
              entry.setTitle(google.gdata.Text.create(<%=strTitle%>));
//Slut bestemmelse af titel
//Start Kommentering af event
<% if strBesk <> "" then %>
              entry.setContent(google.gdata.Text.create(<%= "'" & Replace(strBesk,vbCrLf,"<br>") & "'" %>)); 
<% end if %>
//Slut kommentering
//Start tidsbestemmelse
<%
strStarttime = "'" & Request.Form("FM_start_aar") & "-" & Request.Form("FM_start_mrd") & "-" & Request.Form("FM_start_dag") & "T" & Request.Form("klok_timer") & ":" & Request.Form("klok_min") & ":00.000+01:00'"
strEndtime = "'" & Request.Form("FM_slut_aar") & "-" & Request.Form("FM_slut_mrd") & "-" & Request.Form("FM_slut_dag") & "T" & Request.Form("klok_timer") & ":" & Request.Form("klok_min_slut") & ":00.000+01:00'"
%>
              // Createa When object that will be attached to the event
              var when = new google.gdata.When();
              // Set the start and end time of the When object
              var startTime = google.gdata.DateTime.fromIso8601(<%=strStarttime %>);
              var endTime = google.gdata.DateTime.fromIso8601(<%=strEndtime %>);
              when.setStartTime(startTime);
              when.setEndTime(endTime);
              // Add the When object to the event
              entry.addTime(when);
//Slut tidsbestemmelse
//Start stedsbestemmelse
<% if Request.Form("GoogleStringAdresse") <> "" then
strWhere = "'"&Request.Form("GoogleStringAdresse")&"'"
else
strWhere = "'here'"
end if %>
              var where = new google.gdata.Where();
              
             where.setValueString(<%=strWhere%>);
             entry.addLocation(where);
//Slut stedsbestemmelse
//start tilføjning af deltagere
<%
if Request.Form("GoogleStringEmail") <> "" then
strEmail = "'"& Request.Form("GoogleStringEmail") & "'" 
strParticipantNavn = "'" & Request.Form("GoogleStringKunde") & "'" %>
             
             var who = new google.gdata.Who();
             who.setEmail(<%=strEmail%>);
             who.setValueString(<%=strParticipantNavn %>)
             entry.addParticipant(who);
<%
end if
medarbRelationer = Split(request("FM_medarbrel"), ", ")
For b = 0 to Ubound(medarbRelationer)
strSQL = "Select email, Mnavn from medarbejdere where mid = " & medarbRelationer(b)
oRec.open strSQL, oConn
if not oRec.EOF then
strEmail = "'" & oRec("email") & "'"
strParticipantNavn = "'" & oRec("Mnavn") & "'"%>
             
             var who = new google.gdata.Who();
             who.setEmail(<%=strEmail%>);
             who.setValueString(<%=strParticipantNavn %>)
             entry.addParticipant(who);
<%
strEmail = null
strParticipantNavn = null
end if
oRec.Close				
next
%>
//Slut tilføjning af deltagere      
//tilføj extendedproperty
      var GoogleAktionsId = new google.gdata.ExtendedProperty();
      GoogleAktionsId.setName('aktionsid'); 
      GoogleAktionsId.setValue(<%="'"&GoogleStringAktionsId&"'"%>); 
      entry.addExtendedProperty(GoogleAktionsId); 
//slut tilføjning af extendedproperty
      
              var callback = function(result) {
              window.opener.document.getElementById('GoogleIframeCalendar').src = window.opener.document.getElementById('GoogleIframeCalendar').src;
              window.close();
              }

              var handleError = function(error) {
                  alert(error);
              }
              calendarService.insertEntry(feedUri, entry, callback,
    handleError, google.gdata.calendar.CalendarEventEntry);
    </script>
   

<% case "delete" %>
<script>
/*
* Delete an event
behøver kun GoogleStringAktionsId
*/
// Create the calendar service object
var calendarService = new google.gdata.calendar.CalendarService('GoogleInc-jsguide-1.0');

// The default "private/full" feed is used to delete existing event from the
// primary calendar of the authenticated user
var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';

              // Create a CalendarEventQuery, and specify that this query is
              // applied toward the "private/full" feed
              var query = new google.gdata.calendar.CalendarEventQuery(feedUri);
              
              //søg på extendedproperty
              query.setParam('extq', <%="'[aktionsid:"&GoogleStringAktionsId&"]'"%>);
// This callback method that will be called when getEventsFeed() returns feed data
var callback = function(result) {

  // Obtain the array of matched CalendarEventEntry
  var entries = result.feed.entry;

  // If there is matches for the full text query
  if (entries.length == 1) {

    // delete the first matched event entry  
    var event = entries[0];
    event.deleteEntry(
        function(result) {
          document.getElementById("events").innerHTML = 'aktion slettet!';
        },
        handleError);
  } else {
    // No match is found for the full text query
    document.getElementById("events").innerHTML = 'Der opstod en fejl i kommunikationen med google. Handlingen er ikke redigeret på google\'s server. Der kræves en manuel redigering af posten.';
  }
}

// Error handler to be invoked when getEventsFeed() or updateEntry()
// produces an error
var handleError = function(error) {
  document.getElementById("events").innerHTML = error;
}

// Submit the request using the calendar service object
calendarService.getEventsFeed(query, callback, handleError);
</script>
 <% case "getevents" %>
 <script>
     // Create the calendar service object
     var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');

     // The default "private/full" feed is used to retrieve events from
     // the primary private calendar with full projection
     var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';

     // The callback method that will be called when getEventsFeed() returns feed data
     var callback = function(result) {

         // Obtain the array of CalendarEventEntry
         var entries = result.feed.entry;

         // Print the total number of events
         document.getElementById("events").innerHTML += ('Total of ' + entries.length + ' event(s)');

         for (var i = 0; i < entries.length; i++) {
             var eventEntry = entries[i];
             var eventTitle = eventEntry.getTitle().getText();
             document.getElementById("events").innerHTML += ('Event title = ' + eventTitle + '<br>');
         }
     }

     // Error handler to be invoked when getEventsFeed() produces an error
     var handleError = function(error) {
         document.getElementById("events").innerHTML += (error);
     }

     // Submit the request using the calendar service object
     calendarService.getEventsFeed(feedUri, callback, handleError);
              </script>
  <% case "update" %>
  <script>
/*Updateprocedurer behøver følgende variabler fra siden der inkluderer
GoogleStringOrgTime
GoogleStringOrgTimeCalc
strGoogleStringNavn
strGoogleString
GoogleStringAktionsId*/

              // Create the calendar service object
              var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');

              // The default "private/full" feed is used to update existing event from the
              // primary calendar of the authenticated user
              var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';

              // Create a CalendarEventQuery, and specify that this query is
              // applied toward the "private/full" feed
              var query = new google.gdata.calendar.CalendarEventQuery(feedUri);

              // Create and set the minimum and maximum start time for the date query
              var startMin = google.gdata.DateTime.fromIso8601(<%="'"&Request.Form("GoogleStringOrgTime")&"'" %>);
              var startMax = google.gdata.DateTime.fromIso8601(<%="'"&Request.Form("GoogleStringOrgTimeCalc")&"'" %>);
                query.setMinimumStartTime(startMin);
                query.setMaximumStartTime(startMax);
                
              //søg på extendedproperty
              query.setParam('extq', <%="'[aktionsid:"&GoogleStringAktionsId&"]'"%>); 

              // Flag to indicate whether a match is found
              var eventFound = false;

              // This callback method that will be called when getEventsFeed() returns feed data
              var callback = function(result) {

                  // Obtain the array of matched CalendarEventEntry
                  var entries = result.feed.entry;

                  // If there is matches for the full text query
                  if (entries.length == 1) {

                      // update the first matched event's title
                      var entry = entries[0];
//foretag den egentlige opdatering af første feed
<%//Start bestemmelse af titel
if Request.Form("GoogleStringEmne") <> "" then
strGoogleStringEmne = Request.Form("GoogleStringEmne")
else
strGoogleStringEmne = "Andet"
end if
 if strNavn <> "" then
strTitle = ("'" & strNavn & " - " & strGoogleStringEmne & "'")
else
strTitle = "'" & strGoogleStringEmne & "'"
end if
%>
              entry.setTitle(google.gdata.Text.create(<%=strTitle%>));
//Slut bestemmelse af titel
//Start Kommentering af event
<% if strBesk <> "" then %>
              entry.setContent(google.gdata.Text.create(<%= "'" & Replace(strBesk,vbCrLf,"<br>") & "'" %>)); 
<% end if %>
//Slut kommentering
//Start tidsbestemmelse
<%
strStarttime = "'" & Request.Form("FM_start_aar") & "-" & Request.Form("FM_start_mrd") & "-" & Request.Form("FM_start_dag") & "T" & Request.Form("klok_timer") & ":" & Request.Form("klok_min") & ":00.000+02:00'"
strEndtime = "'" & Request.Form("FM_slut_aar") & "-" & Request.Form("FM_slut_mrd") & "-" & Request.Form("FM_slut_dag") & "T" & Request.Form("klok_timer") & ":" & Request.Form("klok_min_slut") & ":00.000+02:00'"
%>
              // Createa When object that will be attached to the event
              var when = new google.gdata.When();
              // Set the start and end time of the When object
              var startTime = google.gdata.DateTime.fromIso8601(<%=strStarttime %>);
              var endTime = google.gdata.DateTime.fromIso8601(<%=strEndtime %>);
              when.setStartTime(startTime);
              when.setEndTime(endTime);
              // Add the When object to the event
              entry.addTime(when);
//Slut tidsbestemmelse
//Start stedsbestemmelse
<% if Request.Form("GoogleStringAdresse") <> "" then
strWhere = "'"&Request.Form("GoogleStringAdresse")&"'"
else
strWhere = "'here'"
end if %>
              var where = new google.gdata.Where();
              
             where.setValueString(<%=strWhere%>);
             entry.addLocation(where);
//Slut stedsbestemmelse
//start tilføjning af deltagere
<%
if Request.Form("GoogleStringEmail") <> "" then
strEmail = "'"& Request.Form("GoogleStringEmail") & "'" 
strParticipantNavn = "'" & Request.Form("GoogleStringKunde") & "'" %>
             
             var who = new google.gdata.Who();
             who.setEmail(<%=strEmail%>);
             who.setValueString(<%=strParticipantNavn %>)
             entry.addParticipant(who);
<%
end if
medarbRelationer = Split(request("FM_medarbrel"), ", ")
For b = 0 to Ubound(medarbRelationer)
strSQL = "Select email, Mnavn from medarbejdere where mid = " & medarbRelationer(b)
oRec.open strSQL, oConn
if not oRec.EOF then
strEmail = "'" & oRec("email") & "'"
strParticipantNavn = "'" & oRec("Mnavn") & "'"%>
             
             var who = new google.gdata.Who();
             who.setEmail(<%=strEmail%>);
             who.setValueString(<%=strParticipantNavn %>)
             entry.addParticipant(who);
<%
strEmail = null
strParticipantNavn = null
end if
oRec.Close				
next
%>
//Slut tilføjning af deltagere 
entry.updateEntry(function(result) {document.getElementById("information").innerHTML += ('event updated! - extq =' + query.getParam('aktionsid'));},handleError);} else {
                      // No match is found for the full text query
                  document.getElementById("information").innerHTML += (' ' + 'Der opstod en fejl i kommunikationen med google. Handlingen er ikke redigeret på google\'s server. Der kræves en manuel redigering af posten.');
                  }
              }

              // Error handler to be invoked when getEventsFeed() or updateEntry()
              // produces an error
              var handleError = function(error) {
                  document.getElementById("information").innerHTML += error;
              }

              // Submit the request using the calendar service object
              calendarService.getEventsFeed(query, callback, handleError);
              </script>
<% case "synchronize" %>
<script>
<%
if Instr(Request.ServerVariables("HTTP_REFERER"),"googlepost.asp") then%>
function getAllEvents(){
<%end if
%>
              // Create the calendar service object
              var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');

              // The default "private/full" feed is used to retrieve events from
              // the primary private calendar with full projection
              var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';

            var query = new google.gdata.calendar.CalendarEventQuery(feedUri);

             query.setMaxResults(50);
             query.setSingleEvents(true);     
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
                document.googlesync.participants.value += ('][#TO:PartDataString:' + event.getParticipants()[p].getEmail());}}
                else{
                document.googlesync.sids.value += ('][#TO:DataCell:' + event.getExtendedProperties()[0].getValue());
                Synchronized ++;
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
              calendarService.getEventsFeed(query, callback, handleError);
              <%if Instr(Request.ServerVariables("HTTP_REFERER"),"googlepost.asp") then
Response.Write "}"
end if %>
</script>
<%
Function ReturnPosts(PostsNumber,PostType)
if IsNumeric(PostsNumber) then
if PostsNumber = 1 then
PostNrExpr = "post"
elseif PostsNumber = 0 then
PostNrExpr = "0 poster"
else
PostNrExpr = "poster"
end if
Response.Write ("<br />Timeout Intellisync " & PostType & " " & cstr(PostsNumber) & " " & PostNrExpr)
end if
end function
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
else
char = ""
end if
if char <> "" then
name2Split = split(name2,char)
splitpoint = cint(1/ubound(name2Split))
For i = 0 to ubound(name2Split)
if Instr(name1,name2Split(i)) then
splitpoints = cint(splitpoints + splitpoint)
end if
next
if splitpoints > 0.4 then
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
streng = replace(streng,"'","")
streng = Server.HTMLEncode(Lcase(Trim(streng)))
streng = replace(streng,"&#227;&#166;",Chr(230))
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
if not IsNumeric(point) then
point = 0
end if
if point > 0 then
if Instr(result,id) then
'skil strengen ad ved id output streng bliver ...|id,point|(split),point|...
SplitResult = split(result,id)
'split strengen ved | output streng bliver ,point(split)id,point|...
SplitSplitResult = split(SplitResult(1),"|")
'fjern 1 længde fra højre output bliver point
existpoint = right(SplitSplitResult(0),cint(len(SplitSplitResult(0))-1))
if Not IsNumeric(existpoint) then
existpoint = 0
end if
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
Function MyMatch(Match)
If (Match = "true") then
Response.Write (Match & ": true<br>")
Match = Null
end if
End Function
Function fromGT(tidsstreng)
	Tsplit = split(tidsstreng,"T")
    klokkeslet = left(Tsplit(1),8)
    dato = Tsplit(0)
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
    Gtime = klokkeslet & ".000+01.00"
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
'tjek om posten skal tjekkes
if (Instr(Details,"[#to:off]") = false) and (Instr(Details,"[#to: off]") = false) then
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
CountLikeHotkeyResults = "SELECT COUNT(*) as nr FROM kunder WHERE KKundenavn LIKE '%" & Hotkey & "%' OR email LIKE '"& left(Hotkey,(len(Hotkey)-Instr(Hotkey,"."))) & "%'"
SearchLikeHotkey = "SELECT Kid, Kkundenavn FROM kunder WHERE KKundenavn LIKE '%" & Hotkey & "%' OR email LIKE '"& left(Hotkey,(len(Hotkey)-Instr(Hotkey,"."))) & "%'"
set LikeHotkeyResultsC = oConn.Execute(CountLikeHotkeyResults)
LikeHotKeyResults = LikeHotkeyResultsC("nr")
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
CurParticipant = SForm(Participants(Npp))
if (CurParticipant <> "") and (CurParticipant <> "undefined") then

'led efter mail i db
if (CurParticipant = SForm(oRec("email")) and oRec("email") <> "") and EmailLocalFound = "" then
Call AddPoint(oRec("kundeid"),5)
EmailContactMatch = "true"
EmailLocalFound = "true"
end if

'led efter navn db
if (Instr(SForm(oRec("navn")),CurParticipant) and CurParticipant <> "") and NameLocalFound = "" then
Call AddPoint(oRec("kundeid"),cint(len(CurParticipant) * 0.15))
NameContactMatch = "true"
NameLocalFound = "true"
end if

if Instr(Details,SForm(oRec("navn"))) and (oRec("navn") <> "") and (Details <> "") and (NameDetailsFound = "") then
Call AddPoint(oRec("kundeid"),cint(len(oRec("navn")) * 0.15))
NameDetailsFound = "true"
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
NameDetailsFound = Null
oRec.movenext
wend
oRec.close


'Foretag søgning i kunder
FindByKunde = "select Kid, email, Kkundenavn, postnr, adresse, city, kpersmobil1, kpersmobil2, kpersmobil3, kpersmobil4, kpersmobil5, kperstlf1, kperstlf2, kperstlf3, kperstlf4, kperstlf5,"
FindByKunde = FindByKunde & "kpersemail1, kpersemail2, kpersemail3, kpersemail4, kpersemail5, kontaktpers1, kontaktpers2, kontaktpers3, kontaktpers4, kontaktpers5 from kunder"
oRec.open FindByKunde, oConn, 3
while not oRec.EOF

'Match title med kundenavn
Call AddPoint(oRec("Kid"),3*CompareName(Titles,oRec("Kkundenavn")))

'//Tjek details for indhold
If (Details <> "") and Details <> "undefined" then
'Match details med kundenavn
Call AddPoint(oRec("Kid"),(2*CompareName(Details,oRec("Kkundenavn"))))

'tjek details for kontaktpersoner
if (Instr(Details,SForm(oRec("kontaktpers1")))) and (oRec("kontaktpers1") <> "") or (Instr(Details,SForm(oRec("kontaktpers2")))) and (oRec("kontaktpers2") <> "") or (Instr(Details,SForm(oRec("kontaktpers3")))) and (oRec("kontaktpers3") <> "") or (Instr(Details,SForm(oRec("kontaktpers4")))) and (oRec("kontaktpers4") <> "") or (Instr(Details,SForm(oRec("kontaktpers5")))) and (oRec("kontaktpers5") <> "") then
Call AddPoint(oRec("Kid"),2)
end if

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

ResultList = split(result,"|")
if Ubound(ResultList) > 0 then
For i = 1 to Ubound(ResultList)
CurrPoint = split(ResultList(i),",")
ResultPoint = cint(CurrPoint(1))
'procedure til at tælle point
if ResultPoint > 7 then
Command = "insert"
IntKundeId = CurrPoint(0)
exit for
elseif ResultPoint > 1 then
Command = "prompt"
kundenavne = kundenavne & "|" & CurrPoint(0)
end if
next
end if
'luk if paranteser der har med Hotkey at gøre
end if


'registrer kommando fra pointgivning
if Command = "insert" then
InsertPosts = cint(InsertPosts + 1)
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
<script>var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';var searchText = <%="'"&replace(ArrTitles(Ngp),"'","\'")&"'"%>;var query = new google.gdata.calendar.CalendarEventQuery(feedUri);query.setFullTextQuery(searchText);query.setMaxResults(50);query.setSingleEvents(true);var eventFound = false;var callback = function(result) {var entries = result.feed.entry;if (entries.length > 0) {for(i = 0; i < entries.length; i++){if(entries[i].getExtendedProperties().length == 0){event = entries[i];if (event != null){continue;}}else{event = entries[entries.length]}}var GoogleAktionsId = new google.gdata.ExtendedProperty();GoogleAktionsId.setName('aktionsid'); GoogleAktionsId.setValue(<%="'"&aktionsid&"'"%>); event.addExtendedProperty(GoogleAktionsId); event.updateEntry(function(result) {},handleError);} else {document.getElementById('information').innerHTML += 'der opstod en fejl i kommunikationen med google der er afhængig af den information der står i din aftale. Kontakt os venligst med information om den/de aftale du prøver at synkronisere, så vi kan få tilpasset intellisync';}}; var handleError = function(error) {document.getElementById('information').innerHTML += (error);}; calendarService.getEventsFeed(query, callback, handleError);</script>
<%
elseif Command = "prompt" then

'returnering af strenge
if Titles <> "" then
Response.Write "<b>Title:</b> " & Titles & "<br>"
end if
if Details <> "" then
Response.Write "<b>Details:</b> " & Details & "<br>"
end if
Response.Write "<b>participants:</b> " & join(Participants,",") & "<br>"
if Adress <> "" then
Response.Write "<b>Adress:</b> " & Adress & "<br>"
end if
Response.Write "<b>Planlagt:</b> " & FromGT(Times(1)) & " - " & FromGT(Times(2)) & "<br />"
PromptPosts = cint(PromptPosts + 1)
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
<option value="">ingen</option>
    </select>
    <hr />
    <%
kundenavne = Null
else
FailPosts = cint(FailPosts + 1)
end if
'hvis posten indeholder [#to:off]
else
IgnorePosts = cint(IgnorePosts + 1)
end if
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

'returner poster
Call ReturnPosts(InsertPosts,"fandt perfekt match til")
Call ReturnPosts(PromptPosts,"fandt mulige kunder til")
Call ReturnPosts(FailPosts,"var ude af stand til at matche")
Call ReturnPosts(IgnorePosts,"ignorerede")

elseif Request.Form("confirmposthidden") = "j" then%>
<%
'Split DataCells fra nyposter op
ArrTitles = Split(Request.Form("titles"),"][#TO:DataCell:")
ArrDetails = Split(Request.Form("details"),"][#TO:DataCell:")
ArrTimes = Split(Request.Form("times"),"][#TO:DataCell:")

IdsToConfirm = split(Request.Form("confirmposts"),",")
%>
<script>
UpdatePost0();
<%
'start indsæt procedure med de udvalgte IDs
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
if IsNumeric(Trim(IntKundeId)) and (Trim(IntKundeId) <> "") then
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
ConfirmedPosts = cint(ConfirmedPosts + 1)
oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, komm, navn, kundeid) VALUES('" & session("user") & "', '"& strDato &"', '"& startdato &"', '"& startklokkeslet &"', '"& slutklokkeset &"', '"& slutdato &"', '"& Replace(Details,"'","''") &"', '"& Replace(Titles,"'","''") &"', "& IntKundeId &")")

			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
%>
function UpdatePost<%=cint(ConfirmedPosts - 1)%>(){
var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');var feedUri = 'http://www.google.com/calendar/feeds/default/private/full';var searchText = <%="'"&replace(ArrTitles(IdsToConfirm(i)),"'","\'")&"'"%>;var query = new google.gdata.calendar.CalendarEventQuery(feedUri);query.setFullTextQuery(searchText);query.setMaxResults(50);query.setSingleEvents(true);var eventFound = false;var callback = function(result) {var entries = result.feed.entry;if (entries.length > 0) {for(i = 0; i < entries.length; i++){if(entries[i].getExtendedProperties().length == 0){event = entries[i];if (event != null){break;}}else{event = entries[0]}}var GoogleAktionsId = new google.gdata.ExtendedProperty();GoogleAktionsId.setName('aktionsid'); GoogleAktionsId.setValue(<%="'"&aktionsid&"'"%>); event.addExtendedProperty(GoogleAktionsId); event.updateEntry(function(result) {UpdatePost<%=ConfirmedPosts%>()},handleError);} else {document.getElementById('information').innerHTML += 'der opstod en fejl i kommunikationen med google der er afhængig af den information der står i din aftale. Kontakt os venligst med information om den/de aftale du prøver at synkronisere, så vi kan få tilpasset intellisync';}}; var handleError = function(error) {document.getElementById('information').innerHTML += (error);}; calendarService.getEventsFeed(query, callback, handleError);}
<%
end if
Next
%>
function UpdatePost<%=cint(ConfirmedPosts)%>(){
window.opener.document.getElementById('GoogleIframeCalendar').src = window.opener.document.getElementById('GoogleIframeCalendar').src;
window.close();}
</script>
<%
Call ReturnPosts(ConfirmedPosts,"bekræftede matches til")
end if%>
<% end select %> <form name="googlesync" action="" method="post"><% if strGoogleCase = "synchronize" then
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
    if strGoogleCase = "synchronize" then
    btnOnClick = "submit()"
    btnText = "luk vindue"
    else
    btnOnClick = "NewWin_popupaktion('googlepost.asp?googlecase=synchronize')"
    end if
    end if
    if strGoogleCase = "synchronize" then
    btnPost = "location='crmhistorik.asp?menu=crm&shownumofdays=5&func=opret&id=0&ketype=e&selpkt=kal&showinwin=j'"
    else
    btnPost = "NewWin_popupaktion('crmhistorik.asp?menu=crm&shownumofdays=5&func=opret&id=0&ketype=e&selpkt=kal&showinwin=j')"
    end if
    %>
      <input type="button" value="opret ny aftale" id="newevent" name="newevent" onclick="<%=btnPost %>" style="visibility:hidden"/><br />
    
    <input type="button" id="synchronize" name="synchronize" value="<%=btnText %>" onclick="<%=btnOnClick%>"/><br />
  <%end if %>  <input type="hidden" id="confirmposthidden" name="confirmposthidden" value="<%=ConfirmPosts%>"
        style="width:600px;" />
    <input type="hidden" id="confirmposts" name="confirmposts" value=" <%=NGPsToCheck %>"
        style="width:600px;" />
    <input type="hidden" id="posthidden" name="posthidden" value=" " style="width: 600px;" />
    <input type="hidden" id="unsync" name="unsync" value=" " style="width: 600px;" />
    <input type="hidden" id="titles" name="titles" value=" <%=FormTitles %>" style="width: 600px;" />
    <input type="hidden" id="details" name="details" value=" <%=FormDetails %>" style="width: 600px;" />
    <input type="hidden" id="adresses" name="adresses" value=" <%=FormAdresses %>" style="width: 600px;" />
    <input type="hidden" id="times" name="times" value=" <%=FormTimes %>" style="width: 600px;" />
   <input type="hidden" id="participants" name="participants" value=" <%=FormParticipants %>"
        style="width: 600px;" />
    <!--data for allerede oprettede posts-->
    <input type="hidden" id="existsync" name="existsunsync" value=" " style="width: 600px;" />
    <input type="hidden" id="sids" name="sids" value=" " style="width: 600px;" />
    <input type="hidden" id="stitles" name="stitles" value=" " style="width: 600px;" />
    <input type="hidden" id="sdetails" name="sdetails" value=" " style="width: 600px;" />
    <input type="hidden" id="sadresses" name="sadresses" value=" " style="width: 600px;" />
    <input type="hidden" id="stimes" name="stimes" value=" " style="width: 600px;" />
    <input type="hidden" id="sparticipants" name="sparticipants" value=" " style="width: 600px;" />
    </form>
  <% if strGoogleCase = "synchronize" then %>
  </body>
  </html>
  <%end if %>