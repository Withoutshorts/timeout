<%
  

function updFil(filnavn, filertxt, editor, dato, doktype, jobid, kid, folderid, id)


    strSQLupdfil = "UPDATE filer SET filnavn = '"& filnavn &"', "_
	&" filertxt = '"& filertxt &"', editor = '"& editor &"', dato = '"& dato &"', type = "& dokType &", "_
	&" jobid = "& jobid &", kundeid = "& kid &", folderid = "& folderid &""_
	&" WHERE id = " & id

    
	'Response.Write strSQLupdfil
	'Response.end
    oConn.execute(strSQLupdfil)

end function



function insFil(filnavn, filertxt, editor, dato, doktype, jobid, kid, folderid)


    filertxt = replace(filertxt, "'", "''")
    filnavn = replace(filnavn, "'", "''")

    strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	&" adg_kunde, adg_alle, incidentid) VALUES ('"& filnavn &"', '"& filertxt &"', '"& editor &"', "_
	&" '"& dato &"', "& dokType &", "& jobid &", "& kid &", "& folderid &", 0, 1, 0)"

    'Response.Write strSQLupdfil
	'Response.end
    oConn.execute(strSQLinsfil)

end function
    

public jobBeskTxt, jobbeskTxthold
function gemJobBesk(jobbeskTxt, overskriv)
    
    
    jobBeskTxt_from = instr(jobbeskTxt, "##job_besk##")
    jobBeskTxt_to = instr(jobbeskTxt, "##job_besk_slut##")
    jobBeskTxt_len = (jobBeskTxt_to-jobBeskTxt_from) 

    

    if jobBeskTxt_from <> 0 AND jobBeskTxt_to <> 0 then 

    jobBeskTxt = mid(jobbeskTxt, jobBeskTxt_from+12, jobBeskTxt_len-12)
    jobBeskTxt = replace(jobBeskTxt, "jobpr16", "")
    
    if cint(overskriv) <> 2 then 'gem txt til hold
    jobbeskTxthold = jobBeskTxt
    'Response.write "<br><br>jobbeskTxthold: <br><br>" & jobbeskTxthold
    end if


        if cint(overskriv) = 1 then '** opdater **'
        strSQLupdJobB = "UPDATE job SET beskrivelse = '"&jobbeskTxt&"' WHERE id = " & id

        'Response.write strSQLupdJobB
        'Response.end

        oConn.execute(strSQLupdJobB)
        
        end if

    end if

end function
    
     
    
%>