 	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	
<%
	 
	session("strDato") = request.form("mrd")&"/"&request.form("dag")&"/"&request.form("aar")
	session("strDag") = request.form("dag")
	session("strMrd") = request.form("mrd")
	session("straar") = request.form("aar")
	 
%>
<% 
	function SQLBless(s)
	dim tmp
	
	tmp = s
	tmp = replace(tmp, ",", ".")

	SQLBless = tmp
	end function
	varTimerkomma = SQLBless(request.form("FTimer"))
 
	 
	
if len(request.form("FTimer")) = 0 then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
errortype = 6
call showError(errortype)
	
	else
	
	strJobnrAktId = request("Fjobnr")
	antalCharStr = len(strJobnrAktId)
	strKundeKomma = instr(strJobnrAktId, ",")
				
	strJobnr = left(strJobnrAktId, strKundeKomma -1)
	strAktId = right(strJobnrAktId, antalCharStr - strKundeKomma)
	
	'now open it
	oRec.Open "SELECT * FROM job WHERE Jobnr ="& strJobnr &"", oConn, 3
	 	
	'now loop through the records
	While Not oRec.EOF
	
		 strJobnavn = oRec("Jobnavn")
 		 strJobknavn = oRec("Jobknavn")
		 strJobknr = oRec("Jobknr") 	
		 strFakturerbart = oRec("fakturerbart")	
			
	oRec.MoveNext
	Wend 
	
	oRec.Close

	strJobnrCheck = strJobnr

	'Her testes for dubletter i primary index
	
	oRec.Open "Select MTjobnr From Mtid WHERE MTAktivitetId = "& strAktId &" AND MTjobnr = "&strJobnrCheck&" AND MTmnavn = '"&Session("user")&"' AND MTdato = #"&session("strDato")&"#", oConn, 3  
	
	if oRec.EOF Then
	
	oRec.Close
	
	if not len(session("user")) > 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showerror(errortype)
	else
	
	
	'Her indsættes den nye record i mtid
	strFTimer = varTimerkomma
	strTdato = session("strDato") 
	
	oConn.Execute("INSERT INTO mtid (MTJobnr, MTJobnavn, MTmnr, MTmnavn, MTdato, MTimer, MTimerkom, MTknavn, MTknr, MTAktivitetId, MTfaktim) VALUES (" & Cint(strJobnr) & ", '"&strJobnavn&"', " & Cint(request.form("FMnr")) & ", '" & Cstr(Session("user")) & "', #"&strTdato&"# , "& strFTimer &", '" & Cstr(request.form("FTimerkom")) & "', '" & strJobknavn & "', " & strJobknr & ", "& strAktId &", "& strFakturerbart &")")
	  
	Response.redirect "ny_indtastning.asp"

end if
else 
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
errortype = 7
call showerror(errortype)
 
end if
end if


%> 
 
<!--#include file="../inc/regular/footer_inc.asp"-->