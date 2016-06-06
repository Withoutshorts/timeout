<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	strInternt = request("int") 
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
    if len(trim(request("rdir"))) <> 0 then
	rdir = request("rdir")
    else
    rdir = "milep"
    end if
	
	jid = request("jid")
	if len(jid) <> 0 then
	jid = jid 
	else
	jid = 0
	end if
	
	thisfile = "milepale"
	headergif = "../ill/job_milepale_header.gif"
	
	strKnr = request("kundeid")
	

    if len(trim(request("type"))) then
    typ = request("type")
    else
    typ = 0
    end if
	
	
	
	'************************** funktioner *********************************''
	function topfaneblade()
	%>
	
				<span id=visakt name=visakt style="position:absolute; left:600; top:148; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="aktiv.asp?menu=job&jobid=<%=jid%>&jobnavn=<%=jobnavn%>&rdir=<%=rdir%>" class='alt'>Aktiviteter</a></td>
						</tr>
						</table>
				</span>
				
				<span id=visfak name=visfak style="position:absolute; left:330; top:150; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="milepale.asp?menu=job&jid=<%=jid%>&rdir=<%=rdir%>&kundeid=<%=strKnr%>" class='alt'>Milepæle</a></td><!--Fakturering-->
						</tr>
						</table>
				</span>
				
				
				<span id=visgannt name=visgannt style="position:absolute; left:210; top:148; visibility:visible; z-index:100;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="104" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="javascript:NewWin_large('jobplanner.asp?menu=job&id=<%=jid%>')" class='alt'>Gannt</a></td>
						</tr>
						</table>
				</span>
				
				
				
				<span id=visres name=visres style="position:absolute; left:120; top:148; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="javascript:NewWin_large('ressource_belaeg_jbpla.asp?id=<%=jid%>&showonejob=1')" class='alt'>Ressourcer</a></td>
						</tr>
						</table>
				</span>
				
				<span id=visstamdata name=visstamdata style="position:absolute; left:30; top:148; visibility:visible; z-index:250;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="jobs.asp?menu=job&func=red&id=<%=jid%>&int=1&rdir=<%=rdir%>" class='alt'>Stamdata</a></td>
						</tr>
						</table>
				</span>
				
				<span id=vistpris name=vistpris style="position:absolute; left:510; top:148; visibility:visible; z-index:250;">
				<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
					<tr>
						<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
						<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
						<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><a href="jobs.asp?menu=job&func=red&id=<%=jid%>&int=1&rdir=<%=rdir%>&showdiv=tpriser" class='alt'>Timepriser</a></td>
					</tr>
					</table>
			</span>
			
			<span id=filearkiv name=filearkiv style="position:absolute; left:650; top:128; visibility:visible; z-index:0;">
			<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
				<tr>
					<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
					<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
					<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#5582D2">
					<td><a href="javascript:popUp('filer.asp?kundeid=<%=strKnr%>&jobid=<%=jid%>', '600', '600','200', '50')" target="_self" class='alt'>Filarkiv</a></td>
				</tr>
				</table>
		</span>
				
	<%
	end function
	
	
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function
	
	'************************************************************************************************
	
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	
	<%

    slttxt = "<b>Slet milepæl?</b><br />"_
	&"Du er ved at <b>slette</b> en milepæl. Er dette korrekt?"
	slturl = "milepale.asp?menu=job&func=sletok&id="&id&"&rdir="&rdir&"&jid="&jid
	
	call sltquePopup(slturl,slttxt,slturlalt,slttxtalt,110,90)

	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM milepale WHERE id = "& id &"")
	'Response.redirect "milepale.asp?menu=job&shokselector=1&rdir="&rdir&"&jid="&jid

    select case rdir
    case "milep"
    Response.Write("<script language=""JavaScript"">window.opener.location.href(""webblik_milepale.asp"");</script>")
    Response.Write("<script language=""JavaScript"">window.close();</script>")
    case "wip"
    Response.Write("<script language=""JavaScript"">window.opener.location.href(""webblik_joblisten.asp"");</script>")
    Response.Write("<script language=""JavaScript"">window.close();</script>")
    case else
    Response.Write("<script language=""JavaScript"">window.opener.location.href(""jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=milep"");</script>")
    Response.Write("<script language=""JavaScript"">window.close();</script>")
	end select


	case "dbopr", "dbred"
			
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 61
		call showError(errortype)
		
		else
			
			'*tjekker om startdag eksisterer ** 
				if Request("FM_start_dag") > 28 then 
				select case Request("FM_start_mrd")
				case "2"
				strStartDay = 28
				case "4", "6", "9", "11"
				strStartDay = 30
				case else
				strStartDay = Request("FM_start_dag")
				end select
				else
				strStartDay = Request("FM_start_dag")
				end if
			
			strNavn = request("FM_navn")
			intType = request("FM_type")
			milepal_dato = Request("FM_start_aar") &"/" & Request("FM_start_mrd") & "/" & strStartDay 
			strBesk = SQLBless2(request("FM_besk"))
			jid = request("jid")
			timestamp = year(now) & "/" & month(now) &"/" & day(now)

            if len(trim(request("FM_belob"))) <> 0 then
            dblBelob = request("FM_belob")
            dblBelob = replace(dblBelob, ".", "")
            dblBelob = replace(dblBelob, ",", ".")
            else
			dblBelob = 0
            end if

            call erDetInt(dblBelob)
            if isInt = 1 then
            %>
		    <!--#include file="../inc/regular/header_inc.asp"-->
		    <%

            errortype = 61
		    call showError(errortype)

            Response.end
            end if

			if func = "dbopr" then
			strSQL = "INSERT INTO milepale (navn, type, milepal_dato, beskrivelse, editor, dato, jid, belob) VALUES ('"& strNavn &"', "& intType &", '"& milepal_dato &"', '"& strBesk &"', '"& session("user") &"', '"& timestamp &"', "& jid &", "& dblBelob &")"
			oConn.execute(strSQL)
			
			else
			strSQL = "UPDATE milepale SET navn = '"& strNavn &"', type = "& intType &", milepal_dato = '"& milepal_dato &"', beskrivelse = '"& strBesk &"', editor = '"& session("user") &"', dato = '"& timestamp &"', belob = "& dblBelob &", jid = "& jid &" WHERE id = "& id
			oConn.execute(strSQL)
			
			end if
			
			select case rdir 
            case "webblik"
            Response.Write("<script language=""JavaScript"">window.opener.location.href(""webblik_joblisten21.asp"");</script>")
            Response.Write("<script language=""JavaScript"">window.close();</script>") 
            case "milep", "wip"
            Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
            Response.Write("<script language=""JavaScript"">window.close();</script>")
            case else
            Response.Write("<script language=""JavaScript"">window.opener.location.href(""jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=milep"");</script>")
            Response.Write("<script language=""JavaScript"">window.close();</script>")
	        end select
			
		end if 'validering
	
	
	
	
	case "opr", "red" 
		
		
		if func = "opr" then
		jid = request("jid")
		func = "dbopr"
		subm = "Opret"
        dblBelob = "0"
		else
		
		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, type, ikon, beskrivelse, belob FROM milepale LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) WHERE milepale.id = "& id 
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
			strNavn = oRec("navn")
			strTdato = oRec("milepal_dato")
			strBesk = oRec("beskrivelse")
			intType = oRec("type")
            dblBelob = oRec("belob")
		end if
		oRec.close
		
		jid = request("jid")
		func = "dbred"
		subm = "Opdater"
		end if
	%>
	
	<script>
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	
	
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	dtop = 20
	dleft = 10
	%>
	<!--#include file="inc/dato2.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	

	<div id="sindhold" style="position:absolute; left:<%=dleft%>; top:<%=dtop%>; visibility:visible; z-index:50;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="600">
	<tr><form action="milepale.asp?menu=job&func=<%=func%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="rdir" value="<%=rdir%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 valign="top" class="alt" style="padding-top:4;">&nbsp;</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td colspan=2 valign="top" style="padding-top:4;">
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>

    <tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" align=right style="padding-top:8; padding-right:4;"><b>Type / Status:</b></td>
		<td valign="top" style="padding-top:4;">
		
		<%

        if func <> "dbred" then

            if typ = 1 then
            strSQLtyp = " mpt_fak = 1"
            else
            strSQLtyp = " mpt_fak = 0"
            end if

        else

        strSQLtyp = " mpt_fak <> -1"

        end if

		strSQL = "SELECT id, navn FROM milepale_typer WHERE "& strSQLtyp &" ORDER BY navn"
		
        'Response.Write strSQL & "func:" & func

        %>
        <select name="FM_type" style="width:200px; color:#999999;">
        <%


        oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if intType = oRec("id") then
		selthis = "SELECTED"
		else
		selthis = ""
		end if
		%>
		<option value="<%=oRec("id")%>" <%=selthis%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close%>
		</select>&nbsp;(Sættes i kontrolpanelet)
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>

	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" width=150 align=right style="padding-top:7; padding-right:4;"><b>Navn:</b></td>
		<td valign="top" style="padding-top:4;"><input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" style="width:200px;"></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%  if typ = 1 then %>
    </tr>
    	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" style="padding-top:10; padding-right:4;" align=right><b>Beløb:</b></td>
        <td valign="top" style="padding-top:4;"><input type="text" id="FM_belob" name="FM_belob" value="<%=formatnumber(dblBelob,2) %>" style="width:80px;" /> <%=basisValISO %></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<% end if %>
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td>&nbsp;</td>
		<td valign="top" style="padding-top:4;"><b>Kunde og job:</b><br />Kun aktive job, som du er tilknyttet via dine projektgrupper</td>
		<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td>&nbsp;</td>
		<td valign="top" style="padding-top:4;">
			
			<%
			call hentbgrppamedarb(session("mid"))
			
			strSQL = "SELECT j.jobnavn, j.jobnr, j.id, j.jobknr, kkundenavn, kkundenr, kid FROM job j "_
			&" LEFT JOIN kunder ON (kid = j.jobknr) WHERE (j.jobstatus = 1 "& strPgrpSQLkri &") OR id = "& jid &" ORDER BY kkundenavn, j.jobnavn"
			
            'Response.Write strSQL

            %>
            <select name="jid" id="jid" style="width:450px;">
            <%
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
            if cdbl(jid) = oRec("id") then
            jsel = "SELECTED"
			else
			jsel = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=jsel%>><%=oRec("kkundenavn")%> | <%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select>&nbsp;</td>
		<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	
	
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td align=right style="padding-top:4; padding-right:4;"><b>Dato:</b></td>
		<td><select name="FM_start_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>&nbsp;
		
		<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		
        	<%for x = -5 to 10
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 4)%></option>
		<%next %>
        
        </select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		</tr>
	</tr>
	
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
		<td valign="top" align=right style="padding-top:4; padding-right:4;"><b>Evt. beskrivelse / kommentar:</b></td>
		<td valign="top" style="padding-top:4;"><textarea cols="60" rows="4" name="FM_besk" id="FM_besk" id="FM_besk" style="font-size:9px;"><%=strBesk%></textarea></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
		<td colspan=2 valign="top" align=center style="padding-top:4;">
		<input type="submit" value=" <%=subm %> >>" />
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=2 valign="bottom"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</form>
</table>
<br><br>
<!--
<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a>
-->
</div>
<%case else%>
    
    
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->

	$(document).ready(function() {
	$("#vismp").css("background-color", "#5582d2"); 
	});
	
	</script>
	
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	
	
	
	<%
	
	'**** henter job info ***************
		strSQL = "SELECT jobstartdato, jobslutdato, jobnavn FROM job WHERE id = "& jid
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		 	jstdato = oRec("jobstartdato")
			jenddato = oRec("jobslutdato")
			jobnavn = oRec("jobnavn")
		end if
		oRec.close
	
	
	
	oimg = "ikon_milepale_48.png"
	oleft = 10
	otop = 50
	owdt = 200
	oskrift = "Milepæle"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	%>	
	
	

<%call faneblade("mile", "") %>
	
	<div id="sindhold" style="position:absolute; left:20; top:200; visibility:visible; z-index:50;">
    
    <% tTop = 0
	tLeft = 0
	tWdth = 700
	
	
	call tableDiv(tTop,tLeft,tWdth)%>
    <table cellspacing="0" cellpadding="3" border="0" bgcolor="#EFF3FF" width="100%">
	<tr>
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan=4 valign="top"><img src="../ill/blank.gif" width="586" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b><%=jobnavn%></b></td><td colspan=3 valign="top" class="alt" align=right style="padding-top:4; padding-bottom:2; padding-right:20;">
		<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=jid%>&rdir=milep','650','500','250','120');" target="_self" class=alt>Opret ny milepæl&nbsp;<img src="../ill/pillillexp.gif" width="16" height="18" alt="" border="0"></a></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td valign="top" height=10>&nbsp;</td>
		<td colspan=4 valign="top"><b>Milepæle</b><br>Dette kan være en deadline, en del-aflevering eller måske en del-fakturering.<br>Opret nye milepæle typer i kontrolpanelet.<br>&nbsp;</td>
		<td valign="top" align=right>&nbsp;</td>
	</tr>
	    
       
		<tr>
			<td valign="top" height=25 style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">
			<b><%=formatdatetime(jstdato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">Job startdato</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right>&nbsp;</td>
		</tr>
		
		<%
		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, "_
		&" milepale_typer.navn AS type, ikon, beskrivelse, milepale.editor, belob FROM milepale "_
		&" LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) "_
		&" WHERE jid = "& jid &" ORDER BY milepal_dato"
		x = 0
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		
		select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFFFF3"
		end select
		%>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">
			<b><%=formatdatetime(oRec("milepal_dato"), 1)%></b><br />
			<i><%=oRec("editor") %></i></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">
			<span style="color:#999999;"><%=oRec("type")%></span>
            <%if oRec("belob") <> 0 then %><br />
            <b><%=formatnumber(oRec("belob")) &" "& basisValISO %> </b>
            <%end if %></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">
			<a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("id")%>&jid=<%=jid%>&rdir=milep','650','500','250','120');" target="_self" class=vmenu><%=oRec("navn")%></a>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<br>
			<%=oRec("beskrivelse")%>
			<%end if%></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;"><a href="javascript:popUp('milepale.asp?menu=job&id=<%=oRec("id")%>&jid=<%=jid%>&func=slet&rdir=milep','650','500','250','120');" target="_self"><img src="../ill/delete2.gif" alt="Slet" border="0"></a></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right>&nbsp;</td>
		</tr>
		
		<%
		x = x + 1
		oRec.movenext
		wend
		oRec.close
		%>
		
		<tr>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;"><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" height=25><b><%=formatdatetime(jenddato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;">Job slutdato</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		
	</table>
	</div>
	
	<br><br>
	<a href="#" onclick="javacript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a>
	<br><br>&nbsp;
	</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
