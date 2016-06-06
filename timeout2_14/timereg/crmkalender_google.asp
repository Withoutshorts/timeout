	<%response.buffer = true
	'response.CharSet="UTF-8"
	%>
	
	
	
	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="../inc/regular/google_conn.asp"-->
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	func = request("func")
	thisfile = "crmkalender"
	%>
			<!--#include file="../inc/regular/header_lysblaa2_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call crmmainmenu(3)%>
	</div>
	
	<div style="position: absolute; top:167px; left:40px; border:1px #C3D9FF solid;">
		
    <table width="100" cellspacing=0 cellpadding=4 border=0>
<tr>
<td id="login" name="login" onclick="logMeIn()" BGCOLOR="#E8EEF7" onmouseover="style.backgroundColor='#C3D9FF';" onmouseout="style.backgroundColor='#E8EEF7';">
<a href="#" class="vmenu">Login</a>
</td>
</tr>
<tr><td height="1" bgcolor="#C3D9FF"></td></tr>
<tr>
<td id="logout" name="logout" onclick="logMeOut()" BGCOLOR="#E8EEF7" onmouseover="style.backgroundColor='#C3D9FF';" onmouseout="style.backgroundColor='#E8EEF7';">
<a href="#" class="vmenu">Logud</a>
</td>
</tr>
<tr><td height="1" bgcolor="#C3D9FF"></td></tr>
<tr>
<td id="newevent" name="newevent" onclick="NewWin_popupaktion('crmhistorik.asp?menu=crm&shownumofdays=5&func=opret&id=0&ketype=e&selpkt=kal&showinwin=j')" BGCOLOR="#E8EEF7" onmouseover="style.backgroundColor='#C3D9FF';" onmouseout="style.backgroundColor='#E8EEF7';">
<a href="#" class="vmenu">Opret ny aftale</a>
</td>
</tr>
<tr><td height="1" bgcolor="#C3D9FF"></td></tr>
<tr>
<td id="synchronize" name="synchronize" onclick="NewWin_popupaktion('googlepost.asp?googlecase=synchronize')" BGCOLOR="#E8EEF7" onmouseover="style.backgroundColor='#C3D9FF';" onmouseout="style.backgroundColor='#E8EEF7';">
Synkroniser
</td>
</tr>
<tr><td height="1" bgcolor="#C3D9FF"></td></tr>
<tr>
<td id="help" name="help" onclick="NewWin_help('inc/googlefaq.htm')" BGCOLOR="#E8EEF7" onmouseover="style.backgroundColor='#C3D9FF';" onmouseout="style.backgroundColor='#E8EEF7';">
Om IntelliSync
</td>
</tr>
</table>
    
    </div>
		<script type="text/javascript">
		    var calendarService = new google.gdata.calendar.CalendarService('TimeoutCal');
		    var feedUri = 'http://www.google.com/calendar/feeds/default/allcalendars/full';
		    var callback = function(result) {
		        var entries = result.feed.entry;
		        var Callink = 'http://www.google.com/calendar/embed?bgcolor=%23d6dff5';
		        for (var i = 0; i < entries.length; i++) {
		            var calendarEntry = entries[i];
		            var CalPartLink = calendarEntry.getId().getValue().split("www.google.com/calendar/feeds/default/allcalendars/full/");
		            Callink += ('&src=' + CalPartLink[1] + '&color=' + calendarEntry.getColor().getValue().replace('#', '%23'));
		        }
		        document.getElementById('GoogleIframeCalendar').src = Callink;
		    }
		    var handleError = function(error) {}
		    calendarService.getAllCalendarsFeed(feedUri, callback, handleError);
		    
		    if (google.accounts.user.getInfo()) {
		        document.getElementById('login').disabled = true;
		    }
		    else {
		        document.getElementById('logout').disabled = true;
		        document.getElementById('synchronize').disabled = true;
		    };
		     </script>		
		
	
	<div style="position: absolute; visibility:hidden; top:120px; left:20px; border:1px #000000 solid;">
			<%strGoogleCase = "ShowPanel" %>
		<!--#include file="googlepost.asp"-->
		</div>
		<div name="content" style="position:absolute; top:120px; left:150px;">
		<iframe src="" style="border-width:0;" width="800" height="600" id="GoogleIframeCalendar" frameborder="0" scrolling="no">Vent venligst mens kalenderen hentes</iframe>
<div id="information"></div><br />
<div id="events"></div>
        </div>
		<%
end if
%>


<!--#include file="../inc/regular/footer_inc.asp"-->
 



