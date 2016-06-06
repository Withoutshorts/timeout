<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%



 level = session("rettigheder")

sub crmaktemnerfheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Milepæle typer</b><br>
	Tilføj, fjern eller rediger milepæle typer.</td>
	</tr>
	</table><br>
<%
end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes  ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190px; top:102px; visibility:visible;">
	
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en milepæl type. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="milepale_typer.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM milepale_typer WHERE id = "& id &"")
	Response.redirect "milepale_typer.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		strIkon = request("FM_ikon")

        if len(trim(request("FM_mpt_fak"))) <> 0 then
        mpt_fak = request("FM_mpt_fak")
        else
        mpt_fak = 0
        end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO milepale_typer (navn, editor, dato, ikon, mpt_fak) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strIkon &"', "& mpt_fak &")")
		else
		oConn.execute("UPDATE milepale_typer SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', ikon = '"& strIkon &"', mpt_fak = "& mpt_fak &" WHERE id = "&id&"")
		end if
		
		Response.redirect "milepale_typer.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	strIkon = "gron"
	
	else
	strSQL = "SELECT navn, editor, dato, ikon, mpt_fak FROM milepale_typer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strIkon = oRec("ikon")
    mpt_fak = oRec("mpt_fak")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	
        <% 

        call menu_2014()

	'oimg = "ikon_milepale_48.png"
	oleft = 10
	otop = 20
	owdt = 200
	oskrift = "Milepæle"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    %>


   

	<div id="Div2" style="position:absolute; left:90px; top:102px; visibility:visible;">
	

    
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>

	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr><form action="milepale_typer.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><h4>Opret / Rediger milepæl type</h4></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b><br /><br />&nbsp;</td>
	</tr>
	<%end if%>
	
	<tr>
		<td style=padding-right:3;><b>Milepælnavn:</b><br /><input type="text" name="FM_navn" value="<%=strNavn%>" size="30"></td>
	</tr>

    <%if cint(mpt_fak) = 1 then
    mpt_fakCHK = "CHECKED"
    else
    mpt_fakCHK = ""
    end if%>
    <tr>
		<td style=padding-right:3;>
            <input id="Checkbox1" name="FM_mpt_fak" value="1" type="checkbox" <%=mpt_fakCHK%> />
		Vis denne type under ERP --> "Til fakturering"</td>
	</tr>
     
	<!--
	<tr>
		<td colspan=2><b>Ikon:</b><br>
		Vælg et ikon der passer til den type milepæl du vil oprette:<br>
		<
		if strIkon = "gron" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="gron" <%=chthis%>>&nbsp;<img src="../ill/mp_gron.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "graae" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="graae" <%=chthis%>>&nbsp;<img src="../ill/mp_graae.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "gul" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="gul" <%=chthis%>>&nbsp;<img src="../ill/mp_gul.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "orange" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="orange" <%=chthis%>>&nbsp;<img src="../ill/mp_orange.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "rod" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="rod" <%=chthis%>>&nbsp;<img src="../ill/mp_rod.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "sort" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="sort" <%=chthis%>>&nbsp;<img src="../ill/mp_sort.gif" width="10" height="10" alt="" border="0"><br>
		
		<
		if strIkon = "end" then
		chthis = "checked"
		else
		chthis = ""
		end if
		%>
		<input type="radio" name="FM_ikon" value="end" <%=chthis%>>&nbsp;<img src="../ill/mp_end.gif" width="10" height="10" alt="" border="0"><br></td>
	</tr>
	-->
	<tr>
		<td colspan="2" align="right"><br><br><input id="Submit1" type="submit" value=" Opdater >> " /></td>
	</tr>
	</form>
	</table>
    </div>

	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
   
    <!-------------------------------Sideindhold------------------------------------->

    


     <%

         

         call menu_2014()
    
    
	
	
    
    %>

   <div style="position:absolute; left:740px; top:102px;"><% 
                nWdt = 180
                nTxt = "Opret ny Milepæl-type"
                nLnk = "milepale_typer.asp?menu=tok&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %> 
	        </div>


	<div id="sidediv" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
         
    
	
	<%
	




	tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	
	'oimg = "ikon_milepale_48.png"
	oleft = 0
	otop = 0
	owdt = 200
	oskrift = "Milepæle-typer"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    %>

        <br /><br />
    
   
    
    <table cellspacing="0" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#5582d2">
	<td class=alt><b>Id</b></td>
	<td class=alt><b>Emne</b></td>
	<td class=alt><b>Vis under "Til fakturering"</b></td>
	<td>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	'if sort = "navn" then
	strSQL = "SELECT id, navn, mpt_fak FROM milepale_typer ORDER BY navn"
	'else
	'strSQL = "SELECT id, navn, mpt_fak FROM milepale_typer ORDER BY id"
	'end if
	
    c = 1

	oRec.open strSQL, oConn, 3
	while not oRec.EOF 

    select case right(c, 1)
    case 0,2,4,6,8
    bgt = "#EFF3FF"
    case else
    bgt = "#FFFFFF"
    end select

	%>
	<tr>
		<td bgcolor="#cccccc" colspan="5" style="padding:0px 0px 0px 0px;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr style="background-color:<%=bgt%>;">
		<td><%=oRec("id")%></td>
		<td height="20"><a href="milepale_typer.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
        <td>
        <%if cint(oRec("mpt_fak")) = 1 then
        %>
        Ja
        <%
        end if %>
        </td>
		<td><a href="milepale_typer.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
	</tr>
	<%
	x = 0
    c = c + 1
	oRec.movenext
	wend
	%>	
	</table>
	
    </div>

	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
