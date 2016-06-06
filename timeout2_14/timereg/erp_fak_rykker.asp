<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
    thisfile = "erp_fak_rykker"
    kid = request("FM_kunde")
   
    if len(trim(request("FM_job"))) <> 0 then
    jobid = request("FM_job")
    else
    jobid = 0
    end if
    
    if len(trim(request("FM_aftale"))) <> 0 then
    aftid = request("FM_aftale")
    else
    aftid = 0
    end if
    
    jobelaft = request("jobelaft")
	
	func = request("func")
	err = 0
	
	
	
	dontshowDD = 1 
	%>
	<!--#include file="inc/weekselector_s.asp"-->
	<%  
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag                'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut  'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
    
	
	
	
	if func = "reddb" then
	
	antalx = request("antalx")
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
	<%
	
	for x = 0 to antalx
	
	rykkertxt = replace(request("FM_txt_rykker_"&x&""),"'","''")
	rykkerantal = replace(request("FM_antal_rykker_"&x&""),",",".")
	rykkerbelob = replace(request("FM_belob_rykker_"&x&""),",",".")
	
	'Response.Write "rykkerantal " & rykkerantal &"<br>"
	'Response.flush
	
	editdato = year(now) &"/"& month(now) &"/"& day(now)
	rykkerdato = request("FM_aar_"&x&"") &"/"& request("FM_mrd_"&x&"") &"/"& request("FM_dag_"&x&"")
	
	useleftdiv = "t"
	
	if isdate(rykkerdato) = false then
	err = 85
	end if
	
	
    call erDetInt(rykkerantal)
	if isInt > 0 then
    err = 86
    end if
    isInt = 0
    
    call erDetInt(rykkerbelob)
	if isInt > 0 then
    err = 87
    end if
    isInt = 0
    
    next
    
    
	
	if err = 0 then
	
	
	for x = 0 to antalx
	
	rykkerid = request("rykkerid_"&x&"")
	
	rykkertxt = replace(request("FM_txt_rykker_"&x&""),"'","''")
	rykkerantal = replace(request("FM_antal_rykker_"&x&""),",",".")
	rykkerbelob = replace(request("FM_belob_rykker_"&x&""),",",".")
	
	editdato = year(now) &"/"& month(now) &"/"& day(now)
	rykkerdato = request("FM_aar_"&x&"") &"/"& request("FM_mrd_"&x&"") &"/"& request("FM_dag_"&x&"")
	
	rvaluta = request("FM_valuta_"&x&"") 
	rkurs = 100
	                            
    '*** Kurs ****'
    strSQL5 = "SELECT kurs FROM valutaer WHERE id = "& rvaluta
    
    oRec5.open strSQL5, oConn, 3 
    if not oRec5.EOF then 
    rkurs = replace(oRec5("kurs"), ",",".")
    end if
    oRec5.close
		           
		                             
	
	if cint(rykkerantal) <> 0 then
	        
	        if cint(rykkerid) = 0 then
	        
	        strSQL = "INSERT INTO faktura_rykker (rykkerantal, rykkertxt, rykkerbelob, "_
	        &" rykkerdato, dato, editor, fakid, valuta, kurs) VALUES ("_
	        &""& rykkerantal &", '"& rykkertxt&"', "&rykkerbelob &", '"& rykkerdato &"', "_
	        &" '"& editdato &"', '"& session("user") &"', "& id &", "& rvaluta &", "& rkurs &")"
	        
	        else
	        
	        strSQL = "UPDATE faktura_rykker SET rykkerantal = "& rykkerantal &", "_
	        &" rykkertxt = '"& rykkertxt&"', rykkerbelob = "&rykkerbelob &", rykkerdato = '"& rykkerdato &"', "_
	        &" dato = '"& editdato &"', editor = '"& session("user") &"', fakid = "& id &", "_
	        &" valuta = "& rvaluta &", kurs = "& rkurs &" WHERE id = "& rykkerid
	        
	        end if
	        
	
	else
	'** DELETE **
	
	    strSQL = "DELETE FROM faktura_rykker WHERE id = "& rykkerid
	
	end if
	
	
	'Response.Write strSQL
	'Response.flush
	
	oConn.execute(strSQL)
	
	next
	
	
	stDag = request("FM_start_dag")
	stMrd = request("FM_start_mrd")
	stAar = request("FM_start_aar")
	
	slDag = request("FM_slut_dag")
	slMrd = request("FM_slut_mrd")
	slAar = request("FM_slut_aar")
	
	
	%>
	<script>
	window.top.frames['erp3'].location.href = "erp_fak_godkendt_2007.asp?jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&FM_usedatointerval=1&FM_start_dag=<%=stDag%>&FM_start_mrd=<%=stMrd%>&FM_start_aar=<%=stAar%>&FM_slut_dag=<%=slDag%>&FM_slut_mrd=<%=slMrd%>&FM_slut_aar=<%=slAar%>&kid=<%=kid%>"
    window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?FM_type=<%=intType%>&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&jobelaft=<%=jobelaft%>&FM_kunde=<%=kid%>&FM_start_dag=<%=stDag%>&FM_start_mrd=<%=stMrd%>&FM_start_aar=<%=stAar%>&FM_slut_dag=<%=slDag%>&FM_slut_mrd=<%=slMrd%>&FM_slut_aar=<%=slAar%>"
    </script>
											
	
	<%
	else
	
	    errortype = err
		call showError(errortype)
	
	
	end if
	
	
	else
	
	
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<form name=main id=main action="erp_fak_rykker.asp?func=reddb" method="post" target="erp3">
	
	     <input id="id" name="id" type="hidden" value="<%=id %>"/>
        <input id="FM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
        <input id="FM_job" name="FM_job" type="hidden" value="<%=jobid %>"/>
        <input id="FM_aftale" name="FM_aftale" type="hidden" value="<%=aftid %>"/>
         <input id="jobelaft" name="jobelaft" type="hidden" value="<%=jobelaft %>"/>
         <input id="FM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>"/>
         
         <input id="FM_start_dag" name="FM_start_dag" type="hidden" value="<%=strDag%>"/>
         <input id="FM_start_mrd" name="FM_start_mrd" type="hidden" value="<%=strMrd%>"/>
         <input id="FM_start_aar" name="FM_start_aar" type="hidden" value="<%=strAar%>"/>
         
          <input id="FM_slut_dag" name="FM_slut_dag" type="hidden" value="<%=strDag_slut%>"/>
         <input id="FM_slut_mrd" name="FM_slut_mrd" type="hidden" value="<%=strMrd_slut%>"/>
         <input id="FM_slut_aar" name="FM_slut_aar" type="hidden" value="<%=strAar_slut%>"/>
	
	
	 <br /><h4>Tilføj / opdater rykker gebyr på faktura:</h4>          
	<div id=rykkerergebyrDiv style="position:relative; width:700; visibility:visible; border:1px red solid; left:5px; top:10px; padding:10px 10px 10px 10px; background-color:#ffffff;">
    <table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr><td colspan=4><br /><b>Rykker gebyr(er):</b></td></tr>
	
	<%
	strSQL = "SELECT rykkerantal, rykkertxt, rykkerbelob, id AS rykkerid, rykkerdato, valuta AS rvaluta, kurs FROM faktura_rykker WHERE fakid = "& id
	
	'Response.Write strSQL
	'Response.Flush
	oRec.open strSQL, oConn, 3
	
	x = 0
    while not oRec.EOF
        
        call rykkerfelter(oRec("rykkerid"))
        
    x = x + 1
    oRec.movenext
    wend
    oRec.close
	
	
	
	 call rykkerfelter(0)%>
        <input type="hidden" id="antalx" name="antalx" value="<%=x %>" />
    <tr>
	<td colspan=4 align=right><br /><br />
	
     <input id="Submit1" type="submit" value="Tilføj gebyr >> " /></td>
      </tr>
	</table>
	
	
	</div>
	
	
	    
	   </form> 
	
	     
	     
	 <%function rykkerfelter(rykkerid) %>
    <input id="rykkerid_<%=x %>" name="rykkerid_<%=x %>" type="hidden" value="<%=rykkerid %>" />
	
	<%if cint(rykkerid) = 0 then %>
	<tr><td colspan=4><br /><b>Tilføj nyt rykker gebyr:</b></td></tr>
	<%end if %>
	<tr>
	<td>Dato: (dd.mm.yyyy)<br />
	    <%if cint(rykkerid) = 0 then 
	    dagVal = strDag_slut
	    mrdVal = strMrd_slut
	    aarVal = strAar_slut
	    else
	    dagVal = day(oRec("rykkerdato"))
	    mrdVal = month(oRec("rykkerdato"))
	    aarVal = datepart("yyyy",oRec("rykkerdato"), 2,3)
	    end if
	    %>
       
        <input id="dag_<%=x %>" name="FM_dag_<%=x %>" type="text" style="width:20px; font-size:12px;" value="<%=dagVal%>" />.
        <input id="mrd_<%=x %>" name="FM_mrd_<%=x %>" type="text" style="width:20px; font-size:12px;" value="<%=mrdVal%>"/>.
        <input id="aar_<%=x %>" name="FM_aar_<%=x %>" type="text" style="width:40px; font-size:12px;" value="<%=aarVal%>"/>
       
	</td>
	<td>Tekst:<br />
	<input type="Text" name="FM_txt_rykker_<%=x %>" value="Rykker gebyr" style="width:250px; font-size:12px;">
	</td>
	<td>
	<%
	if cint(rykkerid) = 0 then 
	antal = 0
	else
	antal = oRec("rykkerantal")
	end if%>
	
	Antal: (0 = slet / ingen) <br /> <input type="Text" name="FM_antal_rykker_<%=x %>" value="<%=antal %>" style="width:50px; font-size:12px;">
	</td>
	<td>
	<%
	if cint(rykkerid) = 0 then 
	belob = 0
	else
	belob = oRec("rykkerbelob")
	end if%>
	
	Beløb:<br /> <input type="Text" name="FM_belob_rykker_<%=x %>" value="<%=belob %>" style="width:80px; font-size:12px;">
	
	</td>
	
	<%if cint(rykkerid) = 0 then 
	valutaId = basisValId  'DKK
	else
	valutaId = oRec("rvaluta")
	end if
	
	%>
	
	<td>Valuta:<br />
		
	<!-- Valuta -->
								<select name="FM_valuta_<%=x %>" id="FM_valuta_<%=x %>" style="width:50px; font-size:9px;">
								
		                        <%
		                        strSQL5 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
                        		
                        		
		                        oRec5.open strSQL5, oConn, 3 
		                        while not oRec5.EOF 
                        		
		                        if oRec5("id") = valutaId then
		                        valGrpCHK = "SELECTED"
		                        else
		                        valGrpCHK = ""
		                        end if
                        		
                        		
		                        %>
		                        <option value="<%=oRec5("id")%>" <%=valGrpCHK %>><%=oRec5("valutakode")%></option>
		                        <%
		                        oRec5.movenext
		                        wend
		                        oRec5.close
		                        %>
		                        </select>
	</td>
	</tr>
	<%end function %>          
									

<%end if

end if





%>
<!--#include file="../inc/regular/footer_inc.asp"-->


