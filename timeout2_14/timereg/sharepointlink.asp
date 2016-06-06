<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>TimeOut</title>
</head>
<body>
  <!--#include file="../inc/connection/conn_db_inc.asp"-->
<%

 '**** Hvis der linkes direkte fra Intranet / Sharepoint ******'
	if len(trim(request("key"))) <> 0 then
	
	
	
	
	
	if len(trim(request("shpinit"))) <> 0 then
	shpinit = request("shpinit")
	else
	shpinit = "a6196f3646c2"
	end if
	
	
	
	 if len(trim(request("shpjobnr"))) <> 0 then
	 shpjobnr = request("shpjobnr")
	 else
	 shpjobnr = 0
	 end if
	    
	    
	
	
	    
	    '** Henter bruger ***'
	    strSQL = "SELECT mid, mnavn, b.rettigheder, lastlogin FROM medarbejdere m "_
	    &" LEFT JOIN brugergrupper b ON (b.id = m.brugergruppe) WHERE init = '"& shpinit &"'"
	    oRec.open strSQL, oConn, 3
	    if not oRec.EOF then
	                    
	                    strUsrId = Trim(oRec("mid"))
			            session("user") = Trim(oRec("mnavn"))
			            session("login") = strUsrId
			            session("mid") = strUsrId
			            session("rettigheder") = oRec("rettigheder")
			            session("strLastlogin") = oRec("lastlogin")
			            session("fromsharepoint") = "a6196f3646c2e60eddb95cd2134d457f"
			            
			            intStempelur = 0
				
				
				        LoginDateTime = year(now)&"/"& month(now)&"/"&day(now)&" "& datepart("h", now) &":"& datepart("n", now) &":"& datepart("s", now) 
				        LoginDato = year(now)&"/"& month(now)&"/"&day(now)
				
			            
			            '*** Opdaterer lastlogin og loginhistorik ****'
			            
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
			            
			             oConn.execute("UPDATE medarbejdere SET lastlogin = '"& strLastLogin &"' WHERE Mid ="& strUsrId &"")
		
			            
			            strSQL = "INSERT INTO login_historik (dato, login, mid, stempelurindstilling) VALUES ('"& LoginDato &"', '"& LoginDateTime &"', "& strUsrId &", "& intStempelur &")"
				        oConn.execute(strSQL)
				        
				        
	    
	    end if
	    oRec.close
	    
	    
	    '*** Henter jobid ***'
	   
	    
	    session("shpjobid") = 0
	    strSQLj = "SELECT id FROM job WHERE jobnr = " & shpjobnr
	    oRec.open strSQLj, oConn, 3
	    if not oRec.EOF then
	    session("shpjobid") = oRec("id")
	    end if
	    
	
	'Response.Write "session(shpjobid)" & session("shpjobid")
	'Response.end
	
	Response.Redirect "timereg_2006_fs.asp"
	
	else
	
	
	%>
	<!-- 
	Link til TimeOut fra Sharepoint. <a href="sharepointlink.asp?key=665f644e43731ff9db3d341da5c827e1&shpinit=sk&shpjobnr=440" target="_blank">Klik her</a>.<br />
	<br />LINK:<br />
	sharepointlink.asp?key=665f644e43731ff9db3d341da5c827e1&shpinit=sk&shpjobnr=417
	<br /><br />
	-->
	<%
	
	Response.Write "Der er ikke adgang til TimeOut"
	Response.end
	
	end if


%>

</body>
</html>
