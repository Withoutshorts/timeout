
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<% 
session.lcid = 1030


func = request("func")

select case func
case "slet"
            

    id = request("id")

        '*** Sletter timeregistrering ***
	strSQL = "DELETE FROM timer WHERE Tid = " & id
	oConn.execute(strSQL) 


    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
    Response.Write("<script language=""JavaScript"">window.close();</script>")

case else




function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
end function

function SQLBless3(s3)
dim tmp3
tmp3 = s3
tmp3 = replace(tmp3, ":", "")
SQLBless3 = tmp3
end function

antalErr = 0
idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)

			

		'*** Komm
		if len(Request("Timerkom")) >= 1255 then 
			antalErr = 1
			errortype = 34
			useleftdiv = "m"
			%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
			call showError(errortype)
			response.end
		end if	
		

	 	'** timer
				call erDetInt(request("timer"))
				if isInt > 0 then
					antalErr = 1
					errortype = 28
					useleftdiv = "m"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
		
		if len(trim(request("FM_sttid"))) AND len(trim(request("FM_sltid"))) <> 0 then
		
		'** sttid
				
				sttidTemp = request("FM_sttid")&":00"
				call erDetInt(SQLBless3(trim(sttidTemp)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "m"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'Response.write idagErrTjek &" "&sttidTemp
				
				if len(trim(sttidTemp)) <> 0 then
					if isdate(idagErrTjek&" "&sttidTemp) = false then
						antalErr = 1
						errortype = 64
						useleftdiv = "m"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
		
		'** sltid
				
				sltidTemp = request("FM_sltid")&":00"
				call erDetInt(SQLBless3(trim(sltidTemp)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "m"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				if len(trim(sltidTemp)) <> 0 then
					if isdate(idagErrTjek&" "&sltidTemp) = false then
						antalErr = 1
						errortype = 64
						useleftdiv = "m"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
	
		
		end if		
					
		
		
		intAktId = 0
	    intAktId = request("intAktId")
    	
	    aktNavn = ""
	    tfaktim = 0
    	
	    strSQLaktnavn = "SELECT navn, fakturerbar FROM aktiviteter WHERE id = "& intAktId
        oRec.open strSQLaktnavn, oConn, 3
        if not oRec.EOF then
        aktNavn = replace(oRec("navn"), "'", "''")
        tfaktim = oRec("fakturerbar")
        end if
        oRec.close
        
        call akttyper2009prop(tfaktim)
        
        'Response.Write aty_pre & ": "& request("timer")
        
        if len(trim(Request("Timerkom"))) = 0 AND aty_pre <> 0 AND cdbl(aty_pre) <> cdbl(replace(request("timer"), ".", ",")) then
		
		    antalErr = 1
		    errortype = 131
		    useleftdiv = "m"
		    %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
		    call showError(errortype)
		    response.end
    				
        end if			
						
					
		if antalErr = 0 then '** Validering ok!
		
					if len(trim(request("FM_sttid"))) AND len(trim(request("FM_sltid"))) <> 0 then 
						
						sTtid = request("FM_sttid")
						sLtid = request("FM_sltid")
						
						idag = day(now)&"/"&month(now)&"/"&year(now)
						totalmin = datediff("n", idag &" "& sTtid, idag &" "& sLtid)
						call timerogminutberegning(totalmin)
						
						varTimerkomma = thoursTot&"."&tminProcent 'tminTot
						
						sTtid = sTtid &":00"
						sLtid = sLtid &":00"
						
					
					else
					
					varTimerkomma = SQLBless(request.form("Timer"))
					sTtid = ""
					sLtid = ""
					
					end if
	
	jobnr = request("jobnr")
	medid = request("medid")
	
	
	strDato = request.form("aar")&"/"& request.form("mrd")&"/"& request.form("dag")
    strDatoTjk = request.form("dag")&"/"& request.form("mrd")&"/"& request.form("aar")
	dagsdato = year(now) &"/"& month(now) &"/"& day(now)
	str_dagsdato = formatdatetime(dagsdato, 1)
	intOff = request("FM_off")
	timerKom = replace(Request("Timerkom"), "'","''")
	
	timepris = SQLBless(request.form("timepris"))
	
	if len(trim(request("FM_bopal"))) <> 0 then
	bopal = 1
	else
	bopal = 0
	end if
	
	
	intValuta = request("FM_valuta")
	call valutaKurs(intValuta)
      
	
	if varTimerkomma <> 0 then
	

    '** Kan ikke opdatere ferie til en dato uden norm tid ***'
    datoNomrtidOk = 1
    select case tfaktim
    case 11,12,13,14,15,16,17,18,19,20,21 '** Ferie typer + syg og barn syg **'
      
    	    ntimPer = 0
    	    'gnstimerprdag = 0
           
		    call normtimerPer(medid, strDatoTjk, 0, 0)
    	   
    	    'Response.write "ntimPer " & ntimPer & " strDato" & strDatoTjk
            'Response.end

	        if cdbl(ntimPer) > 0 then
    	  
    	       datoNomrtidOk = 1
	            'Response.Write "timerthis " & timerthis & ", timercalc: " & timercalc & ", gnstimerprdag "& gnstimerprdag & ", ntimPer" & ntimPer & "<br>"
	            'Response.flush
	            'timercalc = 0
    	    
	        else
    	    
	        datoNomrtidOk = 0
    	    
	        end if
    	    
       

    case else
    datoNomrtidOk = 1
    end select


	'*** Opdaterer timeregistrering ***'
        if cint(datoNomrtidOk) = 1 then

	    strSQL = "UPDATE timer" &" SET editor = '"& session("user") &"', Tdato = '" & strDato &"', Timer = "& varTimerkomma &", Timerkom = '" & timerKom &"', "_
	    &" offentlig = "& intOff &", timepris = "& timepris &", "_
	    &" tastedato = '"& dagsdato &"', sttid = '"& sTtid &"', "_
	    &" sltid = '"&sLtid&"', tfaktim = "& tfaktim &", valuta = "& intValuta &", kurs = "& dblKurs &", "_
	    &" taktivitetid = "& intAktId &", taktivitetnavn = '"& aktNavn &"', bopal = "& bopal &""_
	    &" WHERE Tid = " & Request.Form("id") 
	
	    oConn.execute(strSQL) 

        else
            
           
           
            useleftdiv = "m"
            errortype = 158
		    %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
		    call showError(errortype)
		    response.end
	
	

          	
                        
            	      

        end if
	
	else
	
	'*** Sletter timeregistrering ***
	strSQL = "DELETE FROM timer WHERE Tid = " & Request.Form("id") 
	oConn.execute(strSQL) 
	
	end if
	
	
	
	'*** Opdaterer eksisterende timepriser siden vlgt dato ****
		
        if request("FM_opdater_timepriser") = "1" then

		'*** Jobid ***
		jobid = 0
		strSQL = "SELECT id FROM job WHERE jobnr = "& jobnr
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		jobid = oRec("id")
		
		oRec.movenext
		wend
		
		oRec.close
		
		
        fraDato = request("FM_opdatertpfra") 
        if isDate(fraDato) = false then
        fraDato = year(now) &"/"& month(now) & "/"& day(now)
        else
        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
        end if

		
		'** Opdaterer alle timereg til ny timepris og valuta ** ( Uanset lastfak ) **'
		strSQL = "UPDATE timer SET timepris = "& timepris &", valuta = "& intValuta &", kurs = "& dblKurs &" WHERE tmnr = "& medid &" AND tjobnr = '"& jobnr &"' AND taktivitetid = " & intAktId & " AND tdato >= '"& fraDato &"'"
		oConn.execute(strSQL)
		
		'*** Opdaterer timepris på aktivitet ****
		oConn.execute("DELETE FROM timepriser WHERE jobid = "& jobid &" AND aktid = "& intAktId &" AND medarbid = "& medid &"")
							
		
		strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
		&" VALUES ("& jobid &", "& intAktId &", "& medid &", 6, "& timepris &", "& intValuta &")"
		oConn.execute (strSQLtp)
		
        end if

	
	    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")
	
	
	end if '** Validering ok
	
			
end select 'func


%>

