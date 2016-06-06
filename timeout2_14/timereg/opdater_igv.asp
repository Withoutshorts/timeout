<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	%>

    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
   



    <%
    

    jobids = split(request("i_jobids"), ", ")
    dddato = day(now) &" "& left(monthname(month(now)), 3) &". "& year(now) 

    for i = 0 to UBOUND(jobids)
    ibrutto = request("i_brutto_"& jobids(i)) 
    irest = request("i_restestimat_"& jobids(i))
    isand = request("i_sandsynlighed_"& jobids(i))

    call erDetInt(ibrutto)
    call erDetInt(irest)
    call erDetInt(isand)
    if isInt = 0 then 

    ibrutto = replace(ibrutto, ".","")
    ibrutto = replace(ibrutto, ",",".")

    if len(trim(ibrutto)) <> 0 then
    ibrutto = ibrutto
    else
    ibrutto = 0
    end if
   
    irest = replace(irest, ".","")
    irest = replace(irest, ",",".")

    if len(trim(irest)) <> 0 then
    irest = irest
    else
    irest = 0
    end if

    isand = replace(isand, ".","")
    isand = replace(isand, ",",".")

    if len(trim(isand)) <> 0 then
    isand = isand
    else
    isand = 0
    end if

    ist_tim_proc = request("i_stade_tim_proc_"& jobids(i))

 

    strSQL = "UPDATE job SET editor = '"& session("user") &"', dato = '"& dddato &"', jo_bruttooms = "& ibrutto &", restestimat = "& irest &", stade_tim_proc = "& ist_tim_proc &", sandsynlighed = '"& isand &"' WHERE id = "& jobids(i)
    'Response.write strSQL & "<br><br>"
    oConn.execute(strSQL) 

    '******************************************************************************
    '****** WIP historik **********************************************************
    strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& irest &", "& ist_tim_proc &", "& session("mid") &", "& jobids(i) &")"
    oConn.Execute(strSQLUpdjWiphist)


    i_md1_bel = request("i_md1_"& jobids(i))

    call erDetInt(i_md1_bel)
    if isInt = 0 then 
    
    i_md1_bel = replace(i_md1_bel, ".","")
    i_md1_bel = replace(i_md1_bel, ",",".")

    if len(trim(i_md1_bel)) <> 0 then
    i_md1_bel = i_md1_bel
    else
    i_md1_bel = 0
    end if

    milepal_dato_1 = request("i_md1_dt_"& jobids(i)) 
    milepal_dato_1 = year(milepal_dato_1) &"/"& month(milepal_dato_1) &"/15"
    
   findes_1 = request("i_findes1_"& jobids(i))

    if cdbl(findes_1) = 0 then
        if cdbl(i_md1_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_1 &"', "& jobids(i) &", "& i_md1_bel &")"
        end if
    else
        if cdbl(i_md1_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_1 &"', belob = "& i_md1_bel &" WHERE id = "& findes_1
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_1
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0

    i_md2_bel = request("i_md2_"& jobids(i))

    call erDetInt(i_md2_bel)
    if isInt = 0 then 
    
    i_md2_bel = replace(i_md2_bel, ".","")
    i_md2_bel = replace(i_md2_bel, ",",".")

    if len(trim(i_md2_bel)) <> 0 then
    i_md2_bel = i_md2_bel
    else
    i_md2_bel = 0
    end if

    milepal_dato_2 = request("i_md2_dt_"& jobids(i)) 
    milepal_dato_2 = year(milepal_dato_2) &"/"& month(milepal_dato_2) &"/15"
    
    findes_2 = request("i_findes2_"& jobids(i))

    if cdbl(findes_2) = 0 then
        if cdbl(i_md2_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_2 &"', "& jobids(i) &", "& i_md2_bel &")"
        end if
    else
        if cdbl(i_md2_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_2 &"', belob = "& i_md2_bel &" WHERE id = "& findes_2
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_2
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0


    i_md3_bel = request("i_md3_"& jobids(i))

    call erDetInt(i_md3_bel)
    if isInt = 0 then 
    
    i_md3_bel = replace(i_md3_bel, ".","")
    i_md3_bel = replace(i_md3_bel, ",",".")

    if len(trim(i_md3_bel)) <> 0 then
    i_md3_bel = i_md3_bel
    else
    i_md3_bel = 0
    end if

    milepal_dato_3 = request("i_md3_dt_"& jobids(i)) 
    milepal_dato_3 = year(milepal_dato_3) &"/"& month(milepal_dato_3) &"/15"
    
    findes_3 = request("i_findes3_"& jobids(i))

    if cdbl(findes_3) = 0 then
        if cdbl(i_md3_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_3 &"', "& jobids(i) &", "& i_md3_bel &")"
        end if
    else
        if cdbl(i_md3_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_3 &"', belob = "& i_md3_bel &" WHERE id = "& findes_3
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_3
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0


    end if
    isInt = 0

    next



    tTop = 40
	tLeft = 20
	tWdth = 200
	
	
	call tableDiv(tTop,tLeft,tWdth)
     %>
      <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            Din jober opdateret
	            </td></tr>
	            </table>

    </div> <!-- table div -->
	
    <br /><br /><br /><br /><br />&nbsp;
	
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
