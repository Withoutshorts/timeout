
<%
'response.Buffer = true
'**** Hvis LTO skal køre på gammel version skal det ændres her ****'
'key = request("key")
'lto = request("lto")
'Response.redirect "../ver2_14/login.asp?key="&key&"&lto="&lto


'**** Sletter alle cookies ***'
'Dim x,y

'For Each x in Request.Cookies
'    Response.Write("<p>")
'    If Request.Cookies(x).HasKeys Then
'        For Each y in Request.Cookies(x)
'            'Response.Write(x & ":" & y & "=" & Request.Cookies(x)(y))
'            'Response.Cookies(x).expires = date - 1
'            'Response.Write("<br />")
'        Next
'    Else
'        'Response.Write(x & "=" & Request.Cookies(x) & "<br />")
'        'Response.Cookies(x).expires = date - 1
'    End If
'    Response.Write "</p>"
'Next
%>


<noscript>
  <META HTTP-EQUIV="Refresh" CONTENT="0;URL=nojava.asp">
</noscript>
    



<!--
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="Mon, 22 Jul 2002 11:12:01 GMT">
    -->
    

    




<%
thisfile = "login.asp" %>
<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/errors/error_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->






<%


'** Fra email om nyt job TO mobile ***'
if len(trim(request("tomobjid"))) <> 0 then
session("tomobjid") = request("tomobjid")
end if

'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030




session("spmettanigol") = session("spmettanigol") + request("attempt")

	If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
	
	'************************************ Login side *********************************************
	
	strSQL = "SELECT licens FROM licens WHERE id = 1"
	oRec.Open strSQL, oConn, 0, 1, 1
	if not oRec.EOF then
	licensto = oRec("licens")
	end if
	oRec.close


   

   

    if lto = "demo" then
	uval = "Guest"
	
    		else
			    if Request.Cookies("login")("usrval") <> "" then
			    uval = Request.Cookies("login")("usrval")
                else
			    uval = ""
                end if
			end if


     '*** Special rdir side, forvalgt initialer TEC / ESN **'
     if lto = "tec" OR lto = "esn" then

        if len(trim(request("usr"))) <> 0 then
        uval = request("usr") 
        uval = replace(uval, "ESN-", "")

        end if

    end if
   

    if lto = "demo" then
			pawval = "1234"
			else
			    if Request.Cookies("login")("pwval") <> "" then
			    pawval = Request.Cookies("login")("pwval")
			    else
			    pawval = ""
			    end if
			end if


      if len(pawval) <> 0 then
			    huskCHK = "CHECKED"
			    else
			    huskCHK = ""
			    end if

	
	'*** tjekker om det er PDA eller PC ****
	
   

      

    call browsertype()
    
        
        
    'if instr(request.servervariables("HTTP_USER_AGENT"), "Smartphone") <> 0 then
	if browstype_client = "ip" then
    'pixLeft = 20
	'pixTop = 40
	'tboxsize = 100

  
		
			    if request.Cookies("timeout")("mobileuser") <> "" AND lto <> "outz" then 
                '** Er der cookie og Du allerede er loggget ind, behøver Du ikke logge ind igen på din telefon
                '** For at skifte bruger på mobil SKAL man slette sine cookies ****'
                
                
				    
                    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp"

                end if
				
			


    %>
     <link href="inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>

	<script src="inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
    <script src="inc/jquery/jquery.timer.js" type="text/javascript"></script>
    <script src="inc/jquery/jquery.corner.js" type="text/javascript"></script>

<script src="login_jav.js" type="text/javascript"></script>

  

    <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />-->
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%=login_txt_012 %></title>
<link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
<link href="timetag_web/css/style.less" rel="stylesheet/less" type="text/css" />
<script>
  less = {
    env: "development",
    async: false,
    fileAsync: false,
    poll: 1000,
    functions: {},
    dumpLineNumbers: "comments",
    relativeUrls: false,
    rootpath: ""
  };
</script>
<script src="timetag_web/js/less.js" type="text/javascript"></script>
</head>
<body>
    <div id="header"><%=login_txt_010 %></div>
    <form id="container" action="login.asp" method="POST">
        <input type="text" id="m_login" name="login" placeholder="<%=login_txt_019 %>" value="<%=uval %>"/>
        <input type="password" id="m_pw" name="pw" placeholder="Password" value="<%=pawval %>"/>
         
        <span><input id="Checkbox1" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /><%=login_txt_011 %></span>
        <input type="submit" class="active" value="Logind >> "/>
          <span style="font-size:10px; color:#999999;">[<%=lto %>]</span>
    </form>

    <!--
    <br /><br /><br /><br /><br />
    <div style="padding:20px;"><span style="font-size:8px; color:#999999;"><%=request.servervariables("HTTP_USER_AGENT") %></span></div>
    -->
  
</body>






    <%
	else


    %>

       

	<!--#include file="inc/regular/header_login_inc.asp"-->

            <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
            <link href="../inc/menu/css/chronograph_01.less" rel="stylesheet/less" type="text/css" />

       



	<%
     

	pixLeft = 120
	pixTop = 120
	tboxsize = 200
	'topgif = "<img src='ill/login_timeout.gif' alt='' border='0'>"
	%>
	
		
	<center>
	<div id="content" style="position:relative; width:850px; padding:0px; left:0px; top:40px; background-color:#FFFFFF; border:1px #cccccc solid;">
	
		<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="login.asp" method="POST">
   
	<tr bgcolor="#004E90">
		<td align=left style="padding:20px;"><img src="ill/outzource_logo_neg_300.gif" /><!--<img src="ill/jul_outz_2012.png" />-->
		</td>
		
		<td>
            <!--
            <div style='position:absolute; top:0px; left:160px; margin:0px; z-index : 1250;'>
            <OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="750" HEIGHT="53" id="OBJECT1" ALIGN="">
 	<PARAM NAME=movie VALUE="godjul.swf"> 
	<PARAM NAME=quality VALUE=high>
	<PARAM NAME=wmode VALUE=transparent> 
	<PARAM NAME=bgcolor VALUE=#000000>
	<EMBED src="godjul.swf" quality=high wmode=transparent bgcolor=#000000 WIDTH="750" HEIGHT="53" NAME="godjul" ALIGN="" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
	</OBJECT>
                </div>
            -->

		<!--<img SRC="ill/Segl_100X72.gif" BORDER="0" ALT="Denne side er SSL 128 bit krypteret." width="97" height="67">-->
		
		
		<!-- GeoTrust QuickSSL [tm] Smart Icon tag. Do not edit. -->
        <!--
        <SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="//smarticon.geotrust.com/si.js"></SCRIPT>
        -->
        <!-- end GeoTrust Smart Icon tag -->
		

		
		</td>
	</tr>
	<tr>
	    <td bgcolor="#ffffff" style="height:5px;"><img src="ill/blank.gif" alt="" border="0"></td>
	</tr>
	</table>
		   
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<tr>
		<td valign="top" style="padding:20px 10px 10px 40px; width:350px;">
		
		
		<div style="position:relative; left:0px; top:0px; width:340px; overflow:hidden; border:10px yellowgreen solid; padding:20px; background-color:#FFFFFF;">
		
        <%select case lto 
        case "demo", "intranet - local" 
       ' Response.Write "<b>Velkommen til vores online TimeOut demo.</b>"_
		'	&" <br>"_
	'		&" Login og password til demoen er: <b>guest / 1234 </b> .<br><br>"_
'			&" Husk! I er altid velkommen til at kontakte OutZourCE på "_
	'		&" +45 2684 2000 eller på email: <a href='mailto:timeout@outzource.dk' class=vmenu>timeout@outzource.dk</a>"
			'&" <br><br>God fornøjelse...<br>&nbsp;"
        case "kejd_pb2", "wwf2", "epi_cati"
        Response.write "<h4 style=""color:red;"">DEMO - DEMO - DEMO - DEMO</h4>"
        case else
        end select'OR lto = "intranet - local"  %>
        
        <table border=0 cellspacing=0 cellpadding=0 width=100%>
	    <tr><td colspan=2><h4><%=login_txt_013 %></h4></td></tr>
	
			

            	<%if lto = "demo" OR lto = "intranet - local" then %>

                <tr><td colspan=2 style="padding:2px 0px 20px 0px; line-height:14px;"><%=login_txt_014 %></td></tr>

                     <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b><%=login_txt_015 %></b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_ref" id="FM_ref" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>

			   <tr><td align=right valign=top style="padding:4px 0px 0px 0px;"><b><%=login_txt_016 %></b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_firma" id="FM_firma" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>

          <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b><%=login_txt_017 %></b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_email" id="FM_email" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>
  <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b><%=login_txt_018 %></b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_tel" id="Text1" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>


     

            <%else %>

            	<!--<tr>
		    <td colspan=2 style="padding:10px;" class=lille>Logind <%=weekdayname(weekday(now, 1)) %> d. <%=formatdatetime(now, 1) %> kl. <%=formatdatetime(now, 3) %> </td>
		</tr>-->
			
			<%end if %>
		
        <input type="hidden" id="lto" value="<%=lto %>" />
       

        <tr>
			<td align=right valign=top style="padding:16px 0px 0px 0px;"><b><%=login_txt_019 %>:</b>
			
			
             
			</td>
			<td valign="top" style="padding:12px 0px 2px 0px;">&nbsp;<input type="Text" name="login" id="login" value="<%=uval%>" placeholder="<%=login_txt_019 %>" style="width:<%=tboxsize%>;"></td>
		</tr>
		<tr>
			<td align=right valign=top style="padding-top:4px;"><b><%=login_txt_020 %></b>
			</td>
			<td valign="top" style="padding:0px 0px 2px 0px;">
                &nbsp;<input type="Password" name="pw" id="pw" value="<%=pawval %>" placeholder="Password" style="width:<%=tboxsize%>;">
			&nbsp;<input type="hidden" name="attempt" value="1">
			
                <br /><input id="huskmig" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> <%=login_txt_011 %>




               
			    
			</td>
			</tr>
			
            <!-- Nyhedsbrev 
            <tr>
            <td valign=top colspan=2 style="padding:4px 0px 2px 20px;"> <%
                
                monthnow = month(now)
                
                select case monthnow
                case 2,4,8,12
                shownyhb = 1
                case else
                shownyhb = 0
                end select

                if request.cookies("tsa")("nyhansw") <> "" then
                shownyhb = 0
                end if
                
                if cint(shownyhb) = 1 then %>
                <br /><br />
               
                <b>Ønsker du at modtage TimeOut nyhedsbrev:</b><br />
                <input id="Radio1" name="nyhedsbrev" type="checkbox" value="1"/> Ja tak, det vil jeg gerne.
                (afmelding sker fra din medarbejderprofil)
                <%end if%></td>
            </tr>
            -->
	
			
			
			
		<tr>
	
		<td valign=top colspan=2 style="padding-top:30px;">	
		
			
			<%
			'**************** Stempelur ********************************
			session("stempelur") = 0
			strSQL = "SELECT stempelur FROM licens WHERE id = 1"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			
				session("stempelur") = oRec("stempelur")
			
			end if
			oRec.close  
			
			if session("stempelur") <> 0 then
			
			%>
		    <h4><%=login_txt_021 %></h4>
			<%
			
				
				strSQL = "SELECT id, navn, faktor, forvalgt FROM stempelur ORDER BY navn"
				oRec.open strSQL, oConn, 3 
				st = 0
                while not oRec.EOF 
				
				if oRec("forvalgt") = 1 then
				chk = "CHECKED"
				else
				chk = ""
				end if
				%>
				<input type="radio" name="FM_stempelur" value="<%=oRec("id")%>" <%=chk%>> <%=oRec("navn")%><br />
				<%

                st = st + 1
				oRec.movenext
				wend
				oRec.close 

                if st = 0 then
                %>
                <br />(ingen)
			    <%
                end if
			
			else
			%>
			<input type="hidden" name="FM_stempelur" id="FM_stempelur" value="0">
			
			<%
			end if
			'************************************************************
			
            
              select case lto 
              case "demo", "intranet - local" 
              submitDsp = "none"
              submitVzb = "hidden"
              case else
              submitDsp = ""
              submitVzb = "visible"
              end select
            %>
		
		
		
		
		</td>
		</tr>
        <tr>
		<td align=right colspan=2 style="padding:30px 20px 2px 0px;">
        <div id="logindiv" style="visibility:<%=submitVzb%>; display:<%=submitDsp%>;">
		<input id="loginsubmit" type="submit" value="Login >>" />
        
        </div>	

      
		<br />&nbsp;
		
		</td>
		</tr>

        <tr><td colspan=2><br /><br />
           
            
               

        <b><%=login_txt_028 %></b><br />
              <a href="timereg/sendpw.asp?lto=<%=lto %>" target="_blank" style="color:#999999; font-size:10px; text-decoration:underline;"><%=login_txt_022 %></a><br />
        
        <a href="javascript:window.external.AddFavorite('https://outzource.dk/<%=lto%>','TimeOut - Tid, Overblik & Fakturering')" style="color:#999999; font-size:10px; text-decoration:underline;"><%=login_txt_023 %></a>&nbsp;<span style="font-size:8px;">[CTRL] + [D]</span> 
        
        
 <br />
 <a  href="#" onclick="this.style.behavior='url(#default#homepage)';this.sethomepage('https://outzource.dk/<%=lto%>');" style="color:#999999; font-size:10px; text-decoration:underline;"><%=login_txt_024 %></a><br />

  <a  href="https://outzource.dk/<%=lto%>" style="color:#999999; font-size:10px; text-decoration:underline;"><%=login_txt_025 %></a> &nbsp;<span style="font-size:10px;"><%=login_txt_026 %></span>
  <br /><br />&nbsp;
        
        </td></tr>
		
		
			<% '*** Logo og licensindehaver ***'
			strSQL = "SELECT useasfak, logo, id, filnavn FROM kunder, filer WHERE useasfak = 1 AND filer.id = kunder.logo"
			
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			logonavn = "<img src='inc/upload/"&lto&"/"&oRec("filnavn")&"' alt='' style='border:0px #000000 solid;'>"
			else
			logonavn = ""
			end if
			oRec.close
			
		%>
		
			 
		    <%if logonavn <> "" then %>
		    <tr>
			<td colspan=2 bgcolor="#ffffff" style="padding:10px;" align=center>
			<%=logonavn %>
		    </td>
		    </tr>
		    
		    <%end if %>
		    <tr>
			<td colspan=2 style="padding:4px 0px 2px 15px;">
			<br /><%=login_txt_027 %><b>&nbsp;<%=licensto%></b>
			</td>
			</tr>
		
		</table>
		</div>
		
              </form>
		
		<br />&nbsp;
		
			
	   </td>
		
		
		<td style="width:20px">&nbsp;</td>	
			
		
		
		
		
		
		<td valign="top" style="padding:20px 40px 10px 0px;">
		    
		    <table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td><br />
		            <!--<h4>Seneste:</h4>
		            
                <b class=news>24 maj 2013</b>
                <p><b>Vi forventer at være klar med de nye forbedringer iløbet af mandag...</b>
                <br />Vi vil i løbet af weekenden og mandag uploade de nye funktioner og forbedringer. <br /><br />
                Det drejer sig bla. om forbedrede loadtider, et forenklet søgefilter på statistikkerne og så en ny kørsels registrering, der er bruger Google til selv at beregne afstand.
    <br /><br />
                    Mvh.<br />
                      OutZourCE

                    </p>
                     -->


                <!--
                <b class=news>12 mar. 2014</b> 
                
              
                      
                      <p><b>Ny menu</b><br />
                      <br /> 
                      I løbet af kort tid vil TimeOut blive iklædt en ny moderne menu, således at det bliver nemmere at navigere fra et hjørne i TimeOut til et andet. 
                      Menuen vil desuden være optimeret til tablets, så TimeOut kan afvikles optimalt på alle enheder.<br />  
                      I vil få nærmere besked når menuen ligger til demo i 3_99 versionen. 
                      <br /><br /><br />
                      </p>    
                -->

                <%select case lto
                 case "epi_uk", "nt"%>
                
                <h4>Dear TimeOut user</h4> 
                As part of the optimization of our support, we hope that you will respond positively to our initiative, and use the following support phone number and email address.
                <br /><br />
                <b>Phone: + 45 25 36 55 00 <br />
                E-mail: support@outzource.dk <br />
                The support is open from 9 a.m. to 4 p.m.</b>
                <br /><br />
                This initiative is part of our objective to structure internal processes so everyone gets a good and professional service.
                <br /><br />
                Out target is 90% of all queries are resolved on the same day or the day after receiving an error message.
                <br /><br />
                Thank you in advance and best regards 
                <br /><br />
                Outzource

                
                
                <%

                case else %>

                <h4>23.5.2016 - in-aktiv af mere end 12 timer</h4>
                Du har nogle gange i løbet af se seneste dage måske observeret at "Du har været in-aktiv af mere end 12 timer", selvom Du kun har været på TimeOut i ganske kort tid, eller Du har modtaget en fejl 500.<br />
                <br />
                Fejlen er rettet.
                <br /><br />
                Mvh.<br />
                Outzource


                <h4>Nyt server miljø</h4>
                Vi er påbegyndt flytningen til et nyt server miljø. Dette medfører at I vil opleve en markant forbedret performance i den daglige afvikling af Timeout.<br />
                Det vil specielt komme til udtryk når der laves store dataudtræk. Vi ser meget frem til at kunne tilbyde jer denne forbedrede oplevelse i den daglige brug af TimeOut.<br /><br />
                Vi flytter vores kunder løbende og I vil får besked når i bliver flyttet. I kommer ikke til at opleve nedetid i forbindelse med flytningen, da vi flytter data i aften, weekend og natte timer.
                <br /><br />
                Mvh.<br />
                Outzource

                <h4>Chrome - PDF problem</h4>
                Der er opstået et problem i Chrome browseren med at generere PDF filer, og når man forsøger at åbne PDF'en får ma nen besked om at den er beskadiget.<br /><br />
                Fejlen skyldes en konflikt mellem Chrome og Adobe's PDF plugins.<br /><br />
                Vi har lavet en lille <a href="chrome_pdf.pdf" target="_blank">guide</a> der beskriver hvordan i løser problemet.<br /><br />
             

                <!--
                <h4>TimeOut – november release 2014</h4>

                TimeOut er blevet opdateret på flere områder – primært er der tale om ny og mere direkte adgang til funktionalitet.<br /><br />

                Læs mere <a href="http://www.outzource.dk/?p=2001" target="_blank">her</a> og kom gerne med kommentarer, ris og ros.<br /><br />

                Hilsen Outzource
                <br /><br />

                <h4>Kære TimeOut bruger</h4>
                Som led i optimering af vores support, håber vi, at I vil tage imod vores initiativ, og benytte nedenstående support telefon samt support mailadresse.
                <br /><br />
                <b>Tel.:  25 36 55 00<br />
                E-mail: support@outzource.dk<br />
                Supporttelefonen er åben kl. 9.00 - 16.00.</b>
                <br /><br />
                Med dette initiativ, sætter vi lidt struktur på processen, således at alle oplever den gode service vi allerede leverer.
                <br /><br />
                Målsætningen er, at 90% af alle henvendelser løses samme dag eller senest dagen efter fejlmeldingen er modtaget.
                <br /><br />
                På forhånd tak, og venlig hilsen 
                <br /><br />
                Outzource
                -->

                






                <%end select %>



            

			
			</td></tr></table>
			
		 
        
		
		
		</td>
			<td style="width:20px">&nbsp;</td>	
		
		</tr>
		</table>
		
		</div> <!-- cointent -->
		</center>
		<br /><br />
		
		
        <%
        'call browsertype()
        'if browstype_client <> "ip" then %>
        
        <script>
		
        
        
        function NewWin(url)    {
		window.open(url, 'Disclaimer', 'width=600,height=580,scrollbars=yes,toolbar=no');    
		}
		</script>
		<br /><br /><br />


		<center>
            <span style="font-size:11px; color:#999999;">
		© 2002 - <%=year(now) &" "& login_txt_029 %>&nbsp;&nbsp;|&nbsp;&nbsp; 
		<%=login_txt_030 %>&nbsp;&nbsp;|&nbsp;&nbsp; <%=login_txt_031 %><br>
		<%=login_txt_032 %>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript:NewWin('timereg/disclaimer.asp');" target="_self" class="vmenu"><u><%=login_txt_033 %></u></a></font>&nbsp;&nbsp;|&nbsp;&nbsp;Optimeret til IE 5+, FireFox 3+ og Chrome 
        
		<% 
		select case lto
		case "infow_demo", "infow", "assurator", "viatech", "pcmanden", "dreist", "rage", "enho"
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b><%=login_txt_034 %></b> salg@informationworker.dk &nbsp;&nbsp;|&nbsp;&nbsp;Tlf: 70212128 
		<%case else%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b><%=login_txt_034 %></b> support@outzource.dk&nbsp;&nbsp;|&nbsp;&nbsp;Tlf: +45 25 36 55 00&nbsp;&nbsp;|&nbsp;&nbsp;Skype: OutZourCE
		
		<%end select%>
		</span>
        <br /><span style="font-size:10px; color:#999999;"><%=request.servervariables("HTTP_USER_AGENT") %></span>
        <br /><br>&nbsp;
               
		</center>

        <%end if %>
		
		
		
		<%'end if ' Iphone
		
		
	
	
else '** POST *****

	'************ Hvis der er brugt mere end 3 login forsøg *********
	if session("spmettanigol") > 3 then '3
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<%
	errortype = 25
	call showError(errortype)
	else
	varLogin = Request.Form("login")
	varPW = Request.Form("pw")


    
	if varlogin = "" Then
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<%
	errortype = 1
	call showError(errortype)
	
	else
	'***************** login og pw modtages og valideres ***********************************************
		
		if varPW = "" then
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 2
		call showError(errortype)
		
		else
		




        '*** HØSTE PW *****'
        hostPw = 0
        hostMid = 0
        if lto = "tec" OR lto = "esn" then
                
               

                strSQL = "SELECT pw AS pw, mid FROM medarbejdere WHERE login = '"& varLogin &"'" 
                
                
            
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then

                'response.write "pw: "& oRec("pw") & "<br>"

               
                if oRec("pw") = "xxx01092014" then
                hostPw = 1
                hostMid = oRec("mid")
                end if

                end if
                oRec.close

                
                if cint(hostPw) = 1 then

                strSQL = "UPDATE medarbejdere SET pw = MD5('"& trim(Request.Form("pw")) &"') WHERE mid = " & hostMid
                oConn.execute(strSQL)

                end if

        end if
        '***** End høste ************'

		
		
		'**** Bruges på timregside til at nulstille bruger ***'
		session("forste") = "j"
		
		'*** Er det en ansat eller en kunde der logger på?***'
		'*** Altid ansat, Kunder har deres  egen login fil **'

        call TimeOutVersion()
	
		
		stransatkunde = "1" 'request("FM_ansat_kunde")
		
		if stransatkunde = "1" then
		'DECODE(
            if cint(hostPw) = 1 then
            strSQL = "SELECT m.mansat, m.mnavn, b.rettigheder, m.mid, m.tsacrm, m.pw, m.timereg FROM medarbejdere m, brugergrupper b WHERE "_
		    &" login = '"& trim(varLogin)&"' AND (mansat <> 2 AND mansat <> 3) AND b.id = m.brugergruppe"
		    
            else
		    strSQL = "SELECT m.mansat, m.mnavn, b.rettigheder, m.mid, m.tsacrm, m.pw, m.timereg FROM medarbejdere m, brugergrupper b WHERE "_
		    &" login = '"& trim(varLogin)&"' AND (mansat <> 2 AND mansat <> 3) AND m.pw = MD5('"& trim(Request.Form("pw")) &"') AND b.id = m.brugergruppe"
		    end if

        else
		strSQL = "SELECT id, kundeid, email, password AS pw, navn AS Mnavn FROM kontaktpers WHERE email='"& request.Form("login") &"'"
		end if
		
		founduser = 0

        'response.write "strSQL "& strSQL
        'response.end
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF Then
		
		            founduser = 1		
            			
			            '*******************************************************************************************'
			            '*** Sætter sesions var. alt efter om det er en medarbejder eller en kunde der logger på****'
            			
            			
			            if stransatkunde = "1" then
			            strUsrId = Trim(oRec("Mid"))
			            session("user") = Trim(oRec("Mnavn"))
			            session("login") = strUsrId
			            session("mid") = strUsrId
			            session("rettigheder") = oRec("rettigheder")

                        
                            '**** Sætter mobile cookies("") til at forblive logget ind på mobil *****
                            response.Cookies("timeout")("mobilemid") = session("mid") 
                            response.Cookies("timeout")("mobileuser") = session("user")
                            response.Cookies("timeout")("rettigheder") = session("rettigheder")
                            response.Cookies("timeout").expires = date + 180
                            '************************************************************************


				            startside = oRec("tsacrm")
            				
            				
				            '*** Hent timrregside fra db eller Form *** 
				            'if len(trim(request("FM_0206"))) <> 0 then
				            'treg0206 = request("FM_0206")
				            'else
				            treg0206 = 1 'oRec("timereg") ALTID NY VERSION
				            'end if
            				
            			
            				
            				
			            else

                            '*** Kunde der er loggget på ****'
			                strUsrId = 0
			                thisKpid = oRec("id")
			                session("user") = Trim(oRec("Mnavn"))
			                session("login") = strUsrId
			                session("mid") = oRec("kundeid") 'Kundeid
			                session("rettigheder") = 0
			                startside = 2 'joblog_k.asp
			                treg0206 = 0
			
                        end if
			            '*******************************************************************************************'
            		
            		
            		
            		
            		
            		
		            '********************* Skriver til logfil ********************************************************'
		            if request.servervariables("PATH_TRANSLATED") <> "E:\www\timeout_xp\wwwroot\ver2_1\login.asp" AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
                
                    
                    'response.write "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt"
                    'response.end    

				            Set objFSO = server.createobject("Scripting.FileSystemObject")
				            select case lto
				            '*** Bruger case else undt. for nedenstående, der dog burde kunne skiftes unden probs. Minus måske "net" pga lto.
            				
				            case "titoonic"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt", 8)
				            case "skybrud"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt", 8)
            				
				            case else
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt", 8)
            				
				            end select
            				
				            objF.writeLine(session("user") &chr(009)&chr(009)&chr(009)& date &chr(009)& time&chr(009)& request.servervariables("REMOTE_ADDR"))
				            objF.close	
		            end if
		            '*******************************************************************************************'
            		
            		
            		
		            '******************* Sætter sidste login dato *********************************************'
		            strMthNow = month(now)
            		
		            select case strMthNow
		            case 1
		            strMth = "Jan. "
		            case 2
		            strMth = "Feb. "
		            case 3
		            strMth = "Mar. "
		            case 4
		            strMth = "Apr. "
		            case 5
		            strMth = "Maj. "
		            case 6
		            strMth = "Jun. "
		            case 7
		            strMth = "Jul. "
		            case 8
		            strMth = "Aug. "
		            case 9
		            strMth = "Sep."
		            case 10
		            strMth = "Okt. "
		            case 11
		            strMth = "Nov. "
		            case 12
		            strMth = "Dec. "
		            end select
            		
		            strLastLogin =  day(now)& " " & strMth & " " & year(now)
		            session("dato") = strLastLogin
            		
		  end if
		  oRec.close
            		
		
		
		if founduser = 1 then
		
		
    		
		    if stransatkunde = "1" then
		    strSQL = "SELECT lastlogin FROM medarbejdere WHERE Mnavn ='" & session("user") &"'"
		    else
		    strSQL = "SELECT lastlogin AS lastlogin FROM kontaktpers WHERE id="& thisKpid 
		    end if
    		
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    session("strLastlogin") = oRec("lastlogin")
		    end if
		    oRec.close
		    
		    
		     
		    
		    
		    '************************************************************************'
		    '***** Medarbejder kunde login ***'
		    if stransatkunde = "1" then
			
			    oConn.execute("UPDATE medarbejdere SET lastlogin = '"& strLastLogin &"', Mnavn = '"& session("user") &"', timereg = "& treg0206 &" WHERE Mid ="& strUsrId &"")
		        
                '** Tilmeld til nyhedsbrev ***'
		        if request("nyhedsbrev") = "1" then
                oConn.execute("UPDATE medarbejdere SET nyhedsbrev = 1 WHERE Mid ="& strUsrId &"")
                response.cookies("tsa")("nyhansw") = 1
                response.cookies("tsa").expires = date + 180
		        end if
		       
		       '*** Ferie Pl ==> Ferie afholdt, AldersRed, omSorg PL ==> Afholdt **'
                select case lto
                case "tec", "esn"
                highV = 6
                case else
                highV = 2
                end select
             
                   select case lto
                   
                   case "epi", "epi_no", "epi_sta", "epi_ab", "intranet - local"
                
                        if session("rettigheder") = 1 AND ((datepart("w", now, 2,2) = 1 AND _
                        datepart("w", session("strLastlogin"), 2,2) <> 1)) then  'HVER MANDAG    
                          
                            'call opdaterFeriePl(session("rettigheder"), highV)
                        
                        end if

                   case else
                        '*** ÆNDRE TIL kun første login hver dag
                        call opdaterFeriePl(session("rettigheder"), highV)
                   end select

		        
               '*** Kosoliderer timer *****'

               'Response.Write "<br><br>lto:" & lto & "dt: "& datepart("d", now, 2,2)  & " //last: "& datepart("d", session("strLastlogin"), 2,2) &" level: " & session("rettigheder")
               'Response.end

               if session("rettigheder") = 1 AND ((datepart("d", now, 2,2) = 18 AND _
               datepart("d", session("strLastlogin"), 2,2) <> 18) OR (datepart("d", now, 2,2) = 27 AND datepart("d", session("strLastlogin"), 2,2) <> 27)) then
                   select case lto
                   case "xx"
                   case "dencker", "epi", "epi_no", "epi_sta", "epi_ab"
                   call timer_konsolider(lto,0)
                   case "xintranet - local", "xoutz"
                   call timer_konsolider(lto,0)
                   end select
               end if

                
		        'Response.end
		        
		        
		       
		        
				
				'*** Opdater login_historik (Stempelur) ****'		
				if len(request("FM_stempelur")) <> 0 then
				intStempelur = request("FM_stempelur")
				else
				intStempelur = 0
				end if
				
                session("dontLogind") = 0

                '**** Selvom stempelur er slået til skal følgende grupper ikke have et logind **'
                select case lto 
				case "cst"
                        
                        strSQLmansat = " mansat <> 0"
                        call medarbiprojgrp(21, session("mid"), 0, -1)
                        if instr(instrMedidProgrp, "#"& session("mid") &"#,") <> 0 then
                        session("dontLogind") = 1
                        else
                        session("dontLogind") = 0
                        end if

                
              
                        
                case else
                session("dontLogind") = 0
                end select


                '**** Log ikke ind automatisk ****'
                
               

                call erStempelurOn()
               
                if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom
				    
                else

                    if session("dontLogind") <> 1 then
				    '**** Tjekker om der skla laves en auto-logud eller om der er et igangværende logind på dagen ***'    
				    call logindStatus(strUsrId, intStempelur, 1, now)
				    end if

                end if
                'end select

                'response.end
				
				

                '******************* Minimumslager notificering ******''
                if session("rettigheder") = 1 then

                    'select case lto
                    'case "dencker", "intranet - local"

                    call minimumslageremail_fn()

                        if cint(minimumslageremail) = 1 then

                            if datePart("w", now, 2,2) = 4 then 'torsdag / OR datePart("w", now, 2,2) = 3  onsdag
                                call matUnderMinLager(lto)
                            end if

                        end if
			
                    'end select	

                end if


				'Response.End
		
		else
		oConn.execute("UPDATE kontaktpers SET lastlogin = '"& strLastLogin &"' WHERE id="& thisKpid  &"")
		end if
		'***********************************'
		'***** Medarbejder kunde login  END '
		
		
        '********************* Sebnder mail hvis det er demo **************************************
        if lto = "demo" then
        email = request("FM_email")
        firma = request("FM_firma")
        ref = request("FM_ref")
        tel = request("FM_tel")

                                    

                                    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                
                                    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut | " & ref &"("& firma &")" 
		                            Mailer.FromAddress = email
		                            Mailer.RemoteHost = "webout.smtp.nu" '195.242.131.254 '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            Mailer.AddRecipient "Søren Karlsen", "sk@outzource.dk"
                                    Mailer.AddRecipient "Outzource Salg", "salg@outzource.dk"
		                          

                                 
            					    Mailer.Subject = "Der har været en ny potentiel kunde i demoen"
		                            strBody = "Hej OutZourCE" & vbCrLf & vbCrLf
                                    strBody = strBody &"Der har været en ny potentiel kunde i demoen : "& vbCrLf & vbCrLf
                                    strBody = strBody &"Firma: "& firma & vbCrLf 
						            strBody = strBody &"Navn: "& ref & vbCrLf & vbCrLf
                                    strBody = strBody &"Email: "& email &""& vbCrLf & vbCrLf
                                    strBody = strBody &"Tel: "& tel &""& vbCrLf & vbCrLf
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing


                                    end if

        end if


			
		
		'*********** redirect til den valgte side *************************************************
		'Response.Cookies("usrname")("ansatkunde") = request("FM_ansat_kunde")
		
            '** Den skal nulstilles , så man altid starter med tom liste. PGA performance EPI
		response.cookies("tsa")("visAlleMedarb") = ""

		Response.Cookies("loginlto")("lto") = lto
		Response.Cookies("loginlto").expires = date + 90
       
		
		if len(request("huskmig")) <> 0 then
		Response.Cookies("login")("usrval") = request("login")
		Response.Cookies("login")("pwval") = request("pw")
		Response.Cookies("login").expires = date + 45
		else
		Response.Cookies("login").expires = date - 1
		end if
		
		
		
		
		'*** nulstiller antal logind forsøg, eller kan man ikke logge ind igen
		session("spmettanigol") = 0

		'Response.Cookies("loginlto")("lto") = lto
			'*** tjekker om det er PDA eller PC ****

   
    'if lto = "outz" OR lto = "hestia" then
            call browsertype()
            browstype_client = browstype_client
    'else
     '       browstype_client = "xx"
    'end if
 

			if browstype_client = "ip" then
			
				'New server
                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                else
                response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                end if

                
			    	

			else
				
				'*********************************************************
				'**** Tjekker om det er localhost (udviklingsmaskine) ****
				'**** Hvis nej: SSL 								  ****
				'*********************************************************
				'if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
				'	if stransatkunde = "1" then
				'		select case startside 
				'		case 1 
				'			if lto = "kasters" then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/kunder.asp?menu=crm"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
				'			end if
				'		case 2
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/jbpla_w.asp?menu=res"
				'		case else '0
				'			'if lto = "bowe" then
				'			'Response.write "treg0206" & treg0206
				'			if cint(treg0206) = 1 then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			'response.redirect "http://81.19.249.35/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg.asp?showstempelur=1"
				'			end if
				'		end select
				'	else
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/joblog_k.asp"
				'	end if
				'else

                    '***** KUNDE/EKSTERN *****'
					if stransatkunde = "1" then


                                if lto = "demo" then
                    
                                response.redirect "to_2015/home_dashboard_demo.asp"

                                else


						        select case startside 
						        case 1 
						            if lto = "essens" OR lto = "fkfk" then
						            response.redirect "timereg/crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
						            else
						            response.redirect "timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
						            end if
						        case 2
						        response.redirect "timereg/ressource_belaeg_jbpla.asp?menu=webblik"
                                case 3

                                call sonmaniuge(now)

                                'Response.Write "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'Response.end
                                response.redirect "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'response.redirect "timereg/logindhist_2011.asp"    

                                case 4
						        response.redirect "timereg/webblik_joblisten.asp?menu=webblik&fromvmenu=1"
                                case 5

                                    if lto = "nt" OR lto = "intranet - local" then
                                    response.redirect "to_2015/job_nt.asp"
                                    else
						            response.redirect "timereg/jobs.asp?menu=job&shokselector=1&fromvemenu=j"
                                    end if
                    
                                case 6



                                call sonmaniuge(now)


                                'Response.Write "timereg/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'Response.end
                                response.redirect "timereg/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                            
                                case 7 ' DASHBOARD
						        response.redirect "to_2015/medarbdashboard.asp"

						        case else '0



							            if cint(treg0206) = 1 then
                            
                            

                                            'Response.write timeout_version
                                            'Response.end

                                            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                                            

                                            select case timeout_version
                                            case "ver2_1"
							                response.redirect "timereg/timereg_akt_2006.asp"
                                            case "ver3_99"
                                            response.redirect "../ver3_99/timereg/timereg_akt_2006.asp"
                                            case "ver2_10"
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_10/timereg/timereg_akt_2006.asp"
                                            case "ver2_14"
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_14/timereg/timereg_akt_2006.asp"
                                            case else
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_14/timereg/timereg_akt_2006.asp"
                                            end select
							                'response.redirect "timereg/timereg_2006_fs.asp"

                                        else

                                        'select case lto
                                        'case "dencker", "intranet - local"
                                        'response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp"
                                        response.redirect "http://localhost/timereg/timereg_akt_2006.asp"
                                        'case else
                                        'response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                        'end select

                                        end if


							            else
							            response.redirect "timereg/timereg.asp?showstempelur=1"
							            end if
						        end select

                        
                                end if 'demo

					else
						response.redirect "timereg/joblog_k.asp"
					end if
				'end if
				
				
			end if
		'********************************************************************************************
		
		
		
		else 'founduser
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 4
		call showError(errortype)
		
		End if
		
		
		
		
		
		'call closeDB
		
		End if '* Validering '
		End if '* Validering '
	end if '* Validering '
	

end if
%>
<!--#include file="inc/regular/footer_login_inc.asp"-->


