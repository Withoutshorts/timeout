
<!--#include file="../xml/global_xml_inc.asp"-->

<%
'*** Er global inc_ inkluderet på den aktuelle side?? 
global_inc = "j"


'*** Afsluttede uger *****
public afslUgerMedab
redim afslUgerMedab(208) '4 år
function afsluger(medarbid, stdato, sldato)
		
		strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, YEAR(u.uge) AS aar, u.id, u.mid FROM ugestatus u WHERE mid =  "& medarbid &""_
		&" AND uge BETWEEN '"& stDato &"' AND '"& slDato &"' GROUP BY u.mid, uge"
		'Response.write strSQL2
		'Response.end
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
		
			afslUgerMedab(oRec2("mid")) = afslUgerMedab(oRec2("mid")) & "#"& oRec2("ugenr")&"_"& oRec2("aar") &"#,"
		
		oRec2.movenext
		wend
		oRec2.close 
		
		
		'Response.Write afslUgerMedab(medarbid) 
		
end function

public startDatoAar, startDatoMd, startDatoDag
function licensStartDato()

key = "2.151-3112-B000"
	
	strSQL = "SELECT l.key FROM licens l WHERE id = 1"
	oRec.open strSQL, oConn, 3
	If not oRec.EOF then
	
	key = oRec("key")
	
	end if
	oRec.close
	
	'Response.write key
	'Response.flush
	
	licensnb = right(key, 3)
	'Response.Write licensnb &"<br>"
	
	if cint(licensnb) < 69 then
	startDatoAar = "200" & mid(key, 5,1)
	startDatoMd = mid(key, 9,2)
	startDatoDag = mid(key, 7,2)
	else
	startDatoAar = "20" & mid(key, 5,2)
	startDatoMd = mid(key, 10,2)
	startDatoDag = mid(key, 8,2)
	end if



end function

public browstype_client
function browsertype()

	if instr(Request.ServerVariables("HTTP_USER_AGENT") , "MSIE") then
	browstype_client = "ie"
	else
	browstype_client = "mz"
	end if

end function

public dsksOnOff
function erSDSKaktiv()
'** SerivceDesk ordning aktiv **'
	dsksOnOff = 0
	strSQL = "SELECT sdsk FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	dsksOnOff = oRec("sdsk") 
	end if
	oRec.close 
end function

public erpOnOff
function erERPaktiv()
'** ERP ordning aktiv 
	erpOnOff = 0
	strSQL = "SELECT erp FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	erpOnOff = oRec("erp") 
	end if
	oRec.close 
end function

public crmOnOff
function erCRMaktiv()
'** CRM ordning aktiv 
	crmOnOff = 0
	strSQL = "SELECT crm FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	crmOnOff = oRec("crm") 
	end if
	oRec.close 
end function

public bgtOnOff
function erBGTaktiv()
'** BGT ordning aktiv 
	crmOnOff = 0
	strSQL = "SELECT bgt FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	bgtOnOff = oRec("bgt") 
	end if
	oRec.close 
end function

public smilaktiv, autogk, autolukvdato, autolukvdatodato 
function ersmileyaktiv()
'** Smiley ordning aktiv 
	smilaktiv = 0
	strSQL = "SELECT smiley, autogk, autolukvdato, autolukvdatodato FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	smilaktiv = oRec("smiley") 
	autogk = oRec("autogk")
	autolukvdato = oRec("autolukvdato")
	autolukvdatodato = oRec("autolukvdatodato")
	end if
	oRec.close 
end function

function medrabSmilord(usemid)
	strSQL = "SELECT smilord FROM medarbejdere WHERE mid = "& usemid
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	smilaktiv = oRec("smilord") 
	end if
	oRec.close 
end function


'** Timeregside ***
public treg0206thisMid
function treg0206use(usemid)
	
	treg0206thisMid = 1
	strSQL = "SELECT timereg FROM medarbejdere WHERE mid = "& usemid
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	treg0206thisMid = oRec("timereg") 
	end if
	oRec.close 
	
end function

public strStartDato, strSlutDato, strAar, strMrd, strDag, strDag_slut, strMrd_slut, strAar_slut
sub datocookie
'******************************* Datoer ******************************************
	if len(request("FM_start_dag")) <> 0 then
	strMrd = request("FM_start_mrd")
	strDag = request("FM_start_dag")
	strAar = right(request("FM_start_aar"),2) 
	strDag_slut = request("FM_slut_dag")
	strMrd_slut = request("FM_slut_mrd")
	strAar_slut = right(request("FM_slut_aar"),2)
	else
		
		if len(Request.Cookies("datoer")("st_md")) <> 0 then
		strMrd = Request.Cookies("datoer")("st_md")
		strDag = Request.Cookies("datoer")("st_dag")
		strAar = Request.Cookies("datoer")("st_aar") 
		strDag_slut = Request.Cookies("datoer")("sl_dag")
		strMrd_slut = Request.Cookies("datoer")("sl_md")
		strAar_slut = Request.Cookies("datoer")("sl_aar")
		else
		strMrd = month(now)
		strDag = day(now)
		strAar = year(now) 
		strDag_slut = strDag
		strMrd_slut = strMrd
		strAar_slut = strAar
		end if
	end if
	
	'** Indsætter cookie **
	Response.Cookies("datoer")("st_dag") = strDag
	Response.Cookies("datoer")("st_md") = strMrd
	Response.Cookies("datoer")("st_aar") = strAar
	Response.Cookies("datoer")("sl_dag") = strDag_slut
	Response.Cookies("datoer")("sl_md") = strMrd_slut
	Response.Cookies("datoer")("sl_aar") = strAar_slut
	Response.Cookies("datoer").Expires = date + 10	
	
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
end sub

function grafik(FM_id, strPic, pictype, txt)
		strSelPic = ""
		strHiddenPic = ""
		%>
		<tr>
		<td align=right valign=top><%=txt%>:&nbsp;&nbsp;</td>
		<td>
		<select name="FM_pic_<%=FM_id%>" id="FM_pic_<%=FM_id%>" onChange="chpic('<%=FM_id%>')" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:200;">
			<option value="1">Ingen</option>
				<%
				strSelPic = "<input type='hidden' name='FM_pic_"&FM_id&"_hidden' id='FM_pic_"&FM_id&"_hidden' value='blank.gif'>"
				strHiddenPic = "<input type='hidden' name='FM_pic_"&FM_id&"_hidden_1' id='FM_pic_"&FM_id&"_hidden_1' value='blank.gif'>"	
				
				
					strSQL = "SELECT id, filnavn AS navn FROM filer WHERE type = "& pictype &" ORDER BY navn"
					oRec.open strSQL, oConn, 3
					
					while not oRec.EOF 
					
						if cint(strPic) = oRec("id") then
						tSelB = "SELECTED"
						strSelPic = "<input type='hidden' name='FM_pic_"&FM_id&"_hidden' id='FM_pic_"&FM_id&"_hidden' value='"&oRec("navn")&"'>"
						
						else
						tSelB = ""
						end if%>
					<option value="<%=oRec("id")%>" <%=tSelB%>><%=oRec("navn")%></option>
					<%
					strHiddenPic = strHiddenPic & "<input type='hidden' name='FM_pic_"&FM_id&"_hidden_"&oRec("id")&"' id='FM_pic_"&FM_id&"_hidden_"&oRec("id")&"' value='"&oRec("navn")&"'>"
					
					oRec.movenext
					wend
					oRec.Close
					%>
			</select>&nbsp;&nbsp;
			|&nbsp;&nbsp;<a href="Javascript:NewWin_cal('upload.asp?func=onthefly&type=<%=pictype%>')" target="_self" class='vmenu'>Upload ny fil</a>
			&nbsp;&nbsp;
			|&nbsp;&nbsp;<a href="javascript:preview('preview.asp?id=<%=FM_id%>');" class=vmenu target="_self">Preview</A>&nbsp;&nbsp;|
			<%=strSelPic%>
			<%=strHiddenPic%>
			<!--<input type="text" name="FM_pic_100" id="FM_pic_100" size=20>
			<input type="text" name="FM_pic_100_h" id="FM_pic_100_h">-->
			</td></tr>
	  	<%end function
		
		
		
		function resstopmenu()
		%>
			<br>
			<a href="jbpla_w.asp?menu=res" class=rmenu>Ressource Kalender</a>
			&nbsp;&nbsp;|&nbsp;&nbsp;<a href="ressource_belaeg_jbpla.asp" class=rmenu>Ressourcetimer (Forecast)</a>
			<!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href="ressource_belaeg.asp" class=rmenu>Ressource Belægning</a>
			&nbsp;&nbsp;|&nbsp;&nbsp;<a href="jbpla_k.asp" class=rmenu>Job Belastning</a><br>&nbsp;-->
		<%
		end function
		
		
		function tsamainmenu(pkt)
		
		call treg0206use(session("mid"))
		%>
		
		<div style="position:relative; z-index:0;">
		<%
		call mmenuTableSt()
		
		if cint(treg0206thisMid) = 0 then
			tregLink = "timereg.asp?menu=timereg"
		else
			tregLink = "timereg_2006_fs.asp"
		end if
		
		call mmenuPkt(1, "102", tregLink,""& global_txt_120 &"",pkt)
		
		if level <= 2 OR level = 6 then
            call mmenuPkt(2, "101", "webblik_joblisten.asp?menu=webblik",""& global_txt_121 &"",pkt)
		end if
		
		if level <= 3 OR level = 6 then
		    call mmenuPkt(3, "63", "jobs.asp?menu=job&shokselector=1&fromvemenu=j",""& global_txt_122 &"",pkt)
		end if
		
		if level <= 3 OR level = 6 OR lto = "bminds" AND (level = 3 OR level = 7) then
		    call mmenuPkt(4, "20", "ressource_belaeg_jbpla.asp?menu=res",""& global_txt_123 &"",pkt)
		end if
		
		
		if level <= 3 OR level = 6 then
		   call mmenuPkt(5, "09", "kunder.asp?menu=kund&visikkekunder=1",""& global_txt_124 &"",pkt)
		end if
		
		
		
		call mmenuPkt(6, "19", "medarb.asp?menu=medarb",""& global_txt_125 &"",pkt)
		
		if level <= 2 OR level = 6 then
		    call mmenuPkt(7, "33", "joblog_timetotaler.asp",""& global_txt_126 &"",pkt)
		end if
		
		
		call mmenuPkt(10, "38", "filer.asp?kundeid=0&jobid=0",""& global_txt_127 &"",pkt)
		
		
		if level <= 3 OR level = 6 then
            call mmenuPkt(11, "30", "materialer.asp?menu=mat",""& global_txt_128 &"",pkt)
		end if
		
		call mmenuTableEnd(1)%>
		
		</div>
		
		<%
		end function
		
		
		function sdskmainmenu(pkt)
		%>
		<div style="position:relative; z-index:0;">
		<%
		call mmenuTableSt()
		call mmenuPkt(1, "12", "sdsk.asp?menu=sdsk","ServiceDesk (Incidents)",pkt)
		call mmenuPkt(4, "33", "sdsk_stat.asp?menu=sdsk","Incident Statistik",pkt)
		call mmenuPkt(5, "56", "sdsk_knowledge.asp?menu=sdsk&visikke=1","Knowledgebase (søg)",pkt)
		call mmenuPkt(10, "38", "filer.asp?kundeid=0&jobid=0","Filarkiv",pkt)
		
		'for x = 1 to 19
		'Response.Write "<td bgcolor=#eff3ff>&nbsp;</td>"
		'next
		
		call mmenuTableEnd(1)%>
		</div>
		<%
		end function
		
		function crmmainmenu(pkt)
		%>
		<div style="position:relative; z-index:0;">
		<%
		call mmenuTableSt()
		call mmenuPkt(1, "12", "crmkalender2.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0","Kalender",pkt)
		call mmenuPkt(2, "09", "kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt","Kontakter",pkt)
		call mmenuPkt(3, "56", "crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist","Aktions historik",pkt)
		call mmenuPkt(10, "38", "filer.asp?kundeid=0&jobid=0","Filarkiv",pkt)
		
		'for x = 1 to 19
		'Response.Write "<td bgcolor=#eff3ff>&nbsp;</td>"
		'next
		
		call mmenuTableEnd(1)%>
		</div>
		<%
		end function
		
		
		
		function erpmainmenu(pkt)
		%>
		<div style="position:relative; z-index:0;">
		<%
		call mmenuTableSt()
		call mmenuPkt(1, "10", "erp_tilfakturering.asp?menu=erp","Fakturering",pkt)
		call mmenuPkt(2, "45", "erp_serviceaft_saldo.asp?menu=erp","Afstemning",pkt)
		
		if level <= 2 OR level = 6 then
		call mmenuPkt(3, "44", "kontoplan.asp?menu=erp","Bogføring",pkt)
		end if
		
		call mmenuPkt(4, "47", "budget_aar_dato.asp?menu=erp","Budget",pkt)
		
		
		'for x = 1 to 18
		'Response.Write "<td bgcolor=#ffffff>&nbsp;</td>"
		'next
		call mmenuTableEnd(1)%>
		</div>
		<%
		end function
		
		
		
		
		
		function mmenuTableSt()
		%>
		<table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#ffffff">
		<tr>
			<td colspan=32 height=25 valign=top><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
		</tr>
		
		
		<tr bgcolor="#ffffff">
		<%
		end function
		
		function mmenuTableEnd(lysblaaOnOff)
		%>
		<td bgcolor="#ffffff">&nbsp;</td>
		</tr>
		<%if lysblaaOnOff = 1 then%>
		<tr bgcolor="#5C75AA">
		     <td colspan=32 valign=top style="border-top:3px #8CAAE6 solid; height:1px;">
               <img src="../ill/blank.gif" width="1" height="1" alt="" border="0"><br /></td>
				</tr>
		<%end if%>
		</table>
		
		<%
		call browsertype()
	    if browstype_client = "mz" then
		%>
		<table cellspacing=0 cellpadding=0 border=0>
		<tr>
		    <td id="screenw"><img src="../ill/blank.gif" width="100" height="1" alt="" border="0"><br /></td>
		</tr>
		</table>
		
		<%
		end if
		end function
		
		
		function mmenuPkt(nb, img, lnk, lnkTxt, valgtPkt)
		
		if cint(nb) = cint(valgtPkt) then
		bgthis = "#ffff99"
		bgr = 1
		else
		bgthis = "#EFF3FF" '"#eff3ff" 
		bgr = 0	
		end if
		
		tdwdt = (len(lnkTxt) * 9)
		%>
		<td align=center width="<%=tdwdt %>" id="mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgthis%>; border:<%=bgr %>px orange solid; border-bottom:0px;" onmouseover="bgcolthisMON('mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>');" onmouseout="bgcolthisOFF('mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>','<%=bgthis%>');">
	    <a href="<%=lnk%>" class="mainmenu" target="_top"><%=lnkTxt%></a>
		</td>
		<td bgcolor="#ffffff" style="width:2px;"><img src="../ill/blank.gif" width="1" height="3" alt="" border="0"></td>
		<%
		end function
		
		
		
		
		
		public SideHeader
		public pkt1oskrift
		function kundelogin_mainmenu(pkt, lto, kundeid)
		
		
		'*** Konfigurerer menupkt navne efter lto. ***
			select case lto 
			case "kringit"
			spkt1 = "Seviceordre"
			spkt2 = "Aftaler"
			spkt3 = "Filarkiv"
			spkt4 = "Infobase"
			pkt1oskrift = "Seviceordre"
			case else
			spkt1 = "Timeregistreringer"
			spkt2 = "Aftaler"
			spkt3 = "Filarkiv"
			spkt4 = "Infobase"
			pkt1oskrift = "Job"
			end select
			
			select case pkt 
			case 1
			SideHeader = "<img src='../ill/ac00102-32.gif' alt='' border='0'>&nbsp;Timeregistreringer."
			case 2
			SideHeader = "<img src='../ill/ac0052-24.gif' width='24' height='24' alt='' border='0'>&nbsp;Aftaler."
			end select
			
			call erSDSKaktiv()
			if cint(dsksOnOff) = 1 then
			spkt5 = "ServiceDesk"
			else
			spkt5 = ""
			end if
	
			'*** Hvis der ikke vises print side udskrives menu her **************
			
	
					'*************** Henter kunde oplysninger ************************************
					strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, logo FROM kunder WHERE Kid =" & kundeid		
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then 
						intKnr = oRec("kkundenr")
						strKnavn = oRec("kkundenavn")
						strKadr = oRec("adresse")
						strKpostnr = oRec("postnr")
						strBy = oRec("city")
						strLand = oRec("land")
						intKid = oRec("Kid")
						intCVR = oRec("cvr")
						intTlf = oRec("telefon")
						logo = oRec("logo")
					end if
					
					oRec.close
					%>
					
					<%
					topknavn = 65
					%>
					<table border="0" cellpadding="0" cellspacing="0" style="position:absolute; left:20; top:<%=topknavn%>; z-index:0;">
					<tr>
						<td colspan="3"><img src="../ill/ac0009-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Kontakt:</b></td>
					</tr>
					<tr>
						<td valign="top" colspan="3"><%=strKnavn%><br>
						<%=strKadr%><br><%=strKpostnr%>&nbsp;<%=strBy%>
						<%if len(trim(intTlf)) <> 0 then%>
						<br>Tlf:&nbsp;<%=intTlf%><br>
						<%end if%>
						&nbsp;</td>
					</tr>
					</table>
					
					
				<!-- slut ventre menu -->
		
		<%
		call mmenuTableSt()
		
		if len(spkt5) <> 0 then
			call mmenuPkt(5, "12", "sdsk.asp?usekview=j&FM_kontakt="&intKid,spkt5,pkt)
		end if
		
		
		if lto <> "spritelab" then
		    call mmenuPkt(1, "45", "joblog_k.asp?func=tim&usekid="&intKid&"&FM_seljob="&jobnr,spkt1,pkt)
		end if
		
		if lto <> "dencker" AND lto <> "spritelab" then
			call mmenuPkt(2, "52", "joblog_k.asp?func=aft&usekid="&intKid&"&FM_seljob="&jobnr,spkt2,pkt)
			call mmenuPkt(3, "56", "javascript:popUp('filer.asp?kundeid="&intKid&"&jobid=0&kundelogin=1', '950', '600','100', '50')",spkt3,pkt)
			'call mmenuPkt(4, "38", "infobase.asp?menu=kund&usekview=j&id="&intKid&"&kontaktid="&intKid&"&FM_seljob="&jobnr,spkt4,pkt)
		
		else
		
		call mmenuPkt(2,"52", "#","&nbsp;",0)
	    call mmenuPkt(2,"52", "#","&nbsp;",0)
	    call mmenuPkt(2,"52", "#","&nbsp;",0)
	    call mmenuPkt(2,"52", "#","&nbsp;",0)
	    
		
		
		end if
		
		
		
		
		
		call mmenuTableEnd(0)%>
		
		<%
		end function
		
		
		
	
	
	
	
	'*****************************************************************************************
	'** Antal dage i md funktion  ***
	public mthDays, workingMthDays, lastWday
	function dageimd(md,ye)
	
	
			Select case md 'monththis
			case "4", "6", "9", "11"
				mthDays = 30
			case 2
					select case ye 'yearthis
					case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040, 2044
					mthDays = 29
					case else 
					mthDays = 28
					end select
			case else
				mthDays = 31
			end select
			
			workingMthDays = 0
			 
			For CheckDay = 1 To mthDays
				if weekday(CheckDay&"/"&md&"/"&ye, 2) < 6 then
					
					call helligdage(CheckDay&"/"&md&"/"&ye)
					
					if erHellig <> 1 then
						workingMthDays = workingMthDays + 1
						lastWday = CheckDay 'weekday(CheckDay&"/"&md&"/"&ye, 2)
					end if
				
				end if
			next
	
	
	end function
	'*********************************************************************************************
	
	public thoursTot, tminTot, tminProcent
	function timerogminutberegning(totalmin)
	
	'Response.write "totalmin:" & totalmin 
	'Response.flush
		
		
		
		thoursTot = 0 
		tminTot = 0
		
		if totalmin  <> 0 then
		
		tminTot = totalmin
		timTemp = formatnumber(tminTot/60, 3)
		
		select case session.LCID
		case "1030" 
		timTemp_komma = split(timTemp, ",")
		case "2057"
		timTemp_komma = split(timTemp, ".")
		case else
		timTemp_komma = split(timTemp, ",")
		end select
		
		for x = 0 to UBOUND(timTemp_komma)
			
			if x = 0 then
			thoursTot = timTemp_komma(x)
			end if
			
			if x = 1 then
			tminTot = totalmin - (thoursTot * 60)
			tminTot = Replace(tminTot, "-", "")
			
				'** Omregner til hel timer (60 min = 100)
				tminProcent = formatnumber((tminTot/60) * 100, 0)
				if tminProcent < 10 then
				tminProcent = "0"&tminProcent
				end if
			end if
			
	    next
		
		
		
		if len(tminTot) <> 0 then
			tminTot = tminTot
			
			if instr(tminTot, "-") <> 0 then
			tminTot = replace(tminTot, "-", "")
			end if
			
			if tminTot = 0 then
			tminTot = "00"
			end if
			
			if len(tminTot) = 1 then
			tminTot = "0"&tminTot
			end if
			
			
		else
		tminTot = "00"
		end if
		
	end if
	
	end function
	
	
	public trclass, tdclass_left, tdclass_right, tdclass
	function tdbgcol_to_1(countrow)
					
					select case right(countrow, 1)
					case 2,4,6,8,0
					trclass = "first"
					tdclass_left = "firstleft"
					tdclass = "first"
					tdclass_right = "firstright"
					case else
					trclass = "second"
					tdclass_left = "secondleft"
					tdclass = "second"
					tdclass_right = "secondright"
					end select
					
	end function
	
	
	
	public strPgrpSQLkri, antalpgrp, strSQLkri3
	function hentbgrppamedarb(medarbid)
		
		f = 0
		dim bgrpId
		Redim bgrpId(f)
		
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& medarbid &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid GROUP BY ProjektgruppeId" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve bgrpId(f) 
				bgrpId(f) = oRec("ProjektgruppeId")
				'Response.write oRec("ProjektgruppeId") & "<br>"
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close
		
		antalpgrp = f
		
			'Rettigheds tjeck på job
			'***********************************************************************
	  	if antalpgrp > 0 then
		
		
				'********************************************************************
				'************ Job ***************************************************
				strPgrpSQLkri =  ""
				strPgrpSQLkri = strPgrpSQLkri &" AND ( "
				
				for intcounter = 0 to f - 1  
			  
			  	strPgrpSQLkri = strPgrpSQLkri &" j.projektgruppe1 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe2 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe3 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe4 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe5 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe6 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe7 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe8 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe9 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe10 = "&bgrpId(intcounter)&""_
				&" OR "
				
				next
			  	
				'** Trimmer de 2 sql states ***
				strPgrpSQLkri_len = len(strPgrpSQLkri)
				strPgrpSQLkri_left = strPgrpSQLkri_len - 3
				strPgrpSQLkri_use = left(strPgrpSQLkri, strPgrpSQLkri_left) 
			  	strPgrpSQLkri = strPgrpSQLkri_use & ") "
		
		
		
		'**************************** Rettigheds tjeck aktiviteter **************************
		  	strSQLkri3 = " a.job = j.id AND aktstatus = 1 AND ( "
			
			for intcounter = 0 to f - 1  
		  
			strSQLkri3 = strSQLkri3 &" a.projektgruppe1 = "& bgrpId(intcounter) &""_
			&" OR a.projektgruppe2 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe3 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe4 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe5 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe6 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe7 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe8 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe9 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe10 = "& bgrpId(intcounter) &" OR "
			
			next
		  	
			'** Trimmer sql states ***
			strSQLkri3_len = len(strSQLkri3)
			strSQLkri3_left = strSQLkri3_len - 3
			strSQLkri3_use = left(strSQLkri3, strSQLkri3_left)  
		  	strSQLkri3 = strSQLkri3_use &") "
		
		
		else

		
		
		
		'*** Job ***
		strPgrpSQLkri = strPgrpSQLkri &" AND ( "
		strPgrpSQLkri = strPgrpSQLkri &" j.projektgruppe1 = 0"_
		&" OR "_
		&" j.projektgruppe2 = 0"_
		&" OR "_
		&" j.projektgruppe3 = 0"_
		&" OR "_
		&" j.projektgruppe4 = 0"_
		&" OR "_
	  	&" j.projektgruppe5 = 0"_
		&" OR "_
	  	&" j.projektgruppe6 = 0"_
		&" OR "_
	  	&" j.projektgruppe7 = 0"_
		&" OR "_
	  	&" j.projektgruppe8 = 0"_
		&" OR "_
	  	&" j.projektgruppe9 = 0"_
		&" OR "_
	  	&" j.projektgruppe10 = 0"_
		&" )"
		
		
		
		'***  Aktiviteter ***
		strSQLkri3 = " a.job = j.id AND aktstatus = 1 AND ( "
		strSQLkri3 = strSQLkri3 &" a.projektgruppe1 = 0"_
		&" OR a.projektgruppe2 = 0"_
		&" OR a.projektgruppe3 = 0"_
		&" OR a.projektgruppe4 = 0"_
		&" OR a.projektgruppe5 = 0"_
		&" OR a.projektgruppe6 = 0"_
		&" OR a.projektgruppe7 = 0"_
		&" OR a.projektgruppe8 = 0"_
		&" OR a.projektgruppe9 = 0"_
		&" OR a.projektgruppe10 = 0)"
		
		end if
		
		
	end function
	
	
	
		
	
	
		function visfirmalogo(plefther, ptopher, kundeid)
		
		%>
		<div id="firmalogo" style="position:absolute; left:<%=plefther%>; top:<%=ptopher%>; z-index=10; width:150px; height:100px; overflow:hidden;">
			<%
			strSQL = "SELECT useasfak, logo, id, filnavn FROM kunder, filer WHERE kid = "& kundeid &" AND filer.id = kunder.logo"
			'Response.Write strSQL
			'Response.flush
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			logonavn = "<img src='../inc/upload/"&lto&"/"&oRec("filnavn")&"' alt='' border='0'>"
			'logonavn = "<img src='../ill/test_logo.gif' alt='' border='0'>"
			nologo = "y"
			else
			logonavn = ""'<br><br><font class=megetlillesilver>(Firma logo kan uploades her.)</font>"
			nologo = "n"
			end if
			oRec.close
			
			Response.write logonavn
			
			%>
		</div>
		
	<% end function
	
	'**************************************************************
	
	public ntimPer 
	function normtimerPer(medid, stDato, interval)
	
	'Response.Write "Kalder normtidpper<br>"
	
	'** Maks 30 typer. **'
	dim mtyperIntvDato, mtyperIntvTyp 
	redim mtyperIntvDato(30), mtyperIntvTyp(30)
	
	'** Aktuel medarbejdertype **'
	strSQLtype = "SELECT medarbejdertype, ansatdato FROM medarbejdere WHERE mid = "& medid
	oRec2.open strSQLtype, oConn, 3
	if not oRec2.EOF then
	
	mtyperIntvTyp(0) = oRec2("medarbejdertype")
    mtyperIntvDato(0) = oRec2("ansatdato")
	
	end if
	oRec2.close
	
	'if cdate(stDato) < cdate(mtyperIntvDato(0)) then
	'stDato = mtyperIntvDato(0)
	'end if
	
	
	'Response.Write "stdato: "& stDato & " interval: "& interval &"<br>"
    'Response.Write "mtyperIntvTyp(0): "& mtyperIntvTyp(0) &"<br>"
    'Response.Write "mtyperIntvDato(0): "& mtyperIntvDato(0) &"<br>"
    'Response.flush
	
	 
	 
	'*** Finder medarbejdertyper_historik ***'
    strSQLmth = "SELECT mid, mtype, mtypedato FROM medarbejdertyper_historik WHERE "_
    &" mid = "& medid &" ORDER BY mtypedato, id"
    
    'Response.Write strSQLmth
    'Response.flush
    
    t = 1
    oRec2.open strSQLmth, oConn, 3
    while not oRec2.EOF 
    
    mtyperIntvTyp(t) = oRec2("mtype")
    mtyperIntvDato(t) = oRec2("mtypedato")
    
    t = t + 1
    oRec2.movenext
    wend
    oRec2.close
    
    'Response.Write "t: " & t & "<br>"
    '*****************************************'
	
			
			stDatoDag = weekday(stDato, 2)
			ntimPer = 0
			
			for c = 0 to interval
			
			            if c = 0 then
			            n = stDatoDag
			            datoCount = stDato
			            else
			                if n < 7 then
			                n = n + 1
			                else
			                n = 1
			                end if
			            datoCount = dateAdd("d", 1, datoCount)
			            end if
			            
			            
			            '*** Tjekker ansatDato **'
			            '*** Skal frst tjekke dage efter ansat dato ***'
			            if cDate(datoCount) >= cDate(mtyperIntvDato(0)) then
			            
			            'Response.Write "<u>cdate(datoCount) </u>"& cdate(datoCount) & "<br>"
			            
			            
			            '*** Finder den medarbejder typer der passer til den valgte dato ***'
			            if t = 1 then
			            mtypeUse = mtyperIntvTyp(0) 
			            'Response.Write  " A mtypeUse: "&  mtypeUse & "<br>"
			            else
			               
			                for d = 2 to t  
			                    
			                    
			                    if cdate(datoCount) >= mtyperIntvDato(d-1) then 
    			                mtypeUse = mtyperIntvTyp(d-1)
    			                'Response.Write  cdate(datoCount) & " >= "& mtyperIntvDato(d-1)  &" - B mtypeUse: "&  mtypeUse & "<br>"
    			                else
    			                    if d = 2 then
    			                    mtypeUse = mtyperIntvTyp(d-1)
    			                    'Response.Write  cdate(datoCount) & " - C1 mtypeUse: "&  mtypeUse & "<br>"
			                    
    			                    else
			                        mtypeUse = mtypeUse
			                        'Response.Write  cdate(datoCount) & " - C mtypeUse: "&  mtypeUse & "<br>"
			                    
			                        end if
			                    end if
			            
			                next
			            
			            end if
			            
			           
			
			            select case n 
						case 7
						normdag = "normtimer_son"
						nd1 = 0
						case 1
						normdag = "normtimer_man"
						nd2 = 0
						case 2
						normdag = "normtimer_tir"
						nd3 = 0
						case 3
						normdag = "normtimer_ons"
						nd4 = 0
						case 4
						normdag = "normtimer_tor"
						nd5 = 0
						case 5
						normdag = "normtimer_fre"
						nd6 = 0
						case 6
						normdag = "normtimer_lor"
						nd7 = 0
						end select
						
					
					'** Finder normtimer p� den valgte dag for den vagte type **'	
					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				
					
					'if medid = 1 then
					'Response.write strSQLnt & "<hr>"
					'Response.flush
					'end if 
					
					oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						ntimPerThis = oRec3("timer")
						
					end if
					oRec3.close 
					
					'Response.write "datoCount " & formatdatetime(datoCount, 2) & "<br>"
					'Response.flush
					
				    call helligdage(datoCount)
					
					if erHellig <> 1 then
					ntimPer = ntimPer + ntimPerThis	
					else
					ntimPer = ntimPer
					end if
					
					
					
					end if '** Ansat Dato **'
					
							
		    next
			
			if len(ntimPer) <> 0 then
			ntimPer = ntimPer
			else
			ntimPer = 0
			end if
			
			
	end function
	
	
	
	
	
	function medarbafstem(intMid, startDato, slutDato, visning)
	
	'Response.Write "startdato " & startdato
	
	stdato = day(startdato)&"/"&month(startdato)&"/"&year(startdato)
	sldato = day(slutdato)&"/"&month(slutdato)&"/"&year(slutdato)
	dageDiff = dateDiff("d", stdato, sldato, 2, 2)
	weekDiff = dateDiff("ww", stdato, sldato, 2, 2)
	
	if len(weekDiff) <> 0 AND weekDiff <> 0 then
	weekDiff = cint(weekDiff)
	else
	weekDiff = 1
	end if
	
	'Response.Write "weekDiff" & weekDiff & "<br>"
	'weekDiff = 1
	
	if intMid <> 0 then
	medarbSQL = " mid = " & intMid & ""
	else
	medarbSQL = " mid <> 0 "
	end if
	
	x = 0
	dim realTimer, medarbNavn, medarbNr, medarbInit, normTimer, normTimerUge
	redim realTimer(200), medarbNavn(200), medarbNr(200), medarbInit(200), normTimer(200), normTimerUge(200)
	
	if visning = 1 then 
	dim  ifTimer, sygTimer, BarnsygTimer, sTimer
	redim  ifTimer(200), sygTimer(200), BarnsygTimer(200), sTimer(200)
	end if
	
	 if visning <> 3 then 
	 dim feriePLTimer, ferieAFTimer, afspTimer, afspTimerBr
	 dim fakTimer, mfForbrug , resTimer, fpTimer
	 redim resTimer(200), fakTimer(200), mfForbrug(200)
	 redim afspTimer(200), afspTimerBr(200), feriePLTimer(200), ferieAFTimer(200), fpTimer(200)
	 end if
	
	
	'** Timer **'
	strSQL = "SELECT t.tid, sum(t.timer) AS realtimer, m.mnavn, m.mnr, m.init,"_
	&" t.tdato, m.mid FROM medarbejdere m "_
	&" LEFT JOIN  timer t ON (t.tmnr = m.mid AND (t.tfaktim = 1 OR t.tfaktim = 2 OR t.tfaktim = 6 OR t.tfaktim = 13 OR t.tfaktim = 14 OR t.tfaktim = 20 OR t.tfaktim = 21) "_
	&" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"') "_
	&" WHERE "& medarbSQL &" AND mansat <> '2' "_
	&" GROUP BY m.mid ORDER BY m.mnavn"
	
	'Response.Write "<br>"& strSQL &"<br>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	ntimPer  = 0
	'Response.Write oRec("tid") & " - " & oRec("tdato") & " - " & oRec("realtimer") & " - " & oRec("mnavn") & "<br>"
	
	    medarbNavn(x) = oRec("mnavn")
	    medarbNr(x) = oRec("mnr")
	    medarbInit(x) = oRec("init")
	    realTimer(x) = oRec("realtimer")
	   
	    'Response.Write "stDato::" & stDato & "#:"& dageDiff &"<br>"
	    call normtimerPer(oRec("mid"), stDato, dageDiff)
	    normTimer(x) = ntimPer 
	    
	    '*** Normeret uge  til ferie beregning (gns. fuld dag) ***'
	    call normtimerPer(oRec("mid"), stDato, 6)
	    normTimerUge(x) = formatnumber(ntimPer/5,2) 
	    
	    if normTimerUge(x) <> 0 then
	    normTimerUge(x) = normTimerUge(x)
	    else
	    normTimerUge(x) = 1
	    end if
	    
	    
	   if visning <> 3 then 
	    
	    '* Afspadsering ***'
	    strSQLaf = "SELECT sum(af.timer) AS afsptimer, af.tdato FROM timer af WHERE af.tmnr = "& oRec("mid") &" AND (af.tfaktim = 12) "_
	    &" AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tmnr "
	    
	    'Response.Write "<br>"& strSQLaf &"<br>"
	    'Response.flush
	   
	    oRec2.open strSQLaf, oConn, 3
	    while not oRec2.EOF
	     afspTimer(x) = oRec2("afsptimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	     strSQLafBr = "SELECT sum(af.timer) AS afsptimerBr, af.tdato FROM timer af WHERE af.tmnr = "& oRec("mid") &" AND (af.tfaktim = 13) "_
	     &" AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tmnr "
	    
	    'Response.Write "<br>"& strSQLaf &"<br>"
	    'Response.flush
	   
	    oRec2.open strSQLafBr, oConn, 3
	    while not oRec2.EOF
	     afspTimerBr(x) = oRec2("afsptimerBr")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	    '** ferie **'
	    strSQLf = "SELECT sum(f.timer) AS feriepltimer, f.tdato FROM timer f WHERE f.tmnr = "& oRec("mid") &" AND (f.tfaktim = 11) "_
	    &" AND f.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY f.tmnr "
	    
	    'Response.Write "<br>"& strSQLaf &"<br>"
	    'Response.flush
	   
	    oRec2.open strSQLf, oConn, 3
	    while not oRec2.EOF
	     feriePLTimer(x) = oRec2("feriepltimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    strSQLf = "SELECT sum(f.timer) AS ferieaftimer, f.tdato FROM timer f WHERE f.tmnr = "& oRec("mid") &" AND (f.tfaktim = 14) "_
	    &" AND f.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY f.tmnr "
	    
	   
	    oRec2.open strSQLf, oConn, 3
	    while not oRec2.EOF
	    ferieAFTimer(x) = oRec2("ferieaftimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	    '*** Ressource Timer ***'
	    strSQLres = "SELECT sum(r.timer) AS restimer FROM ressourcer_md r WHERE r.medid = "& oRec("mid") &" AND "_
	    &" (aar >= YEAR('"& startdato &"') AND md >= MONTH('"& startdato &"') "_
	    &" AND aar <= YEAR('"& slutdato &"') AND md <= MONTH ('"& slutdato &"'))"
	    
	    'Response.Write strSQLres & "<br>"
	    
	    oRec2.open strSQLres, oConn, 3
	    while not oRec2.EOF
	    resTimer(x) = oRec2("restimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	   '*** Faktureret ***'
	    strSQLfak = "SELECT sum(fms.fak) AS faktimer, f.fid, f.fakdato FROM fakturaer f "_
	    &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid")&")"_
	    &" WHERE f.fakdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY fms.mid"
	    
	    'Response.Write strSQLfak & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLfak, oConn, 3
	    while not oRec2.EOF
	    fakTimer(x) = oRec2("faktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	  
	    '*** Materiale forbrug ***'
	    strSQLmf = "SELECT sum(mf.matantal * (mf.matsalgspris * (kurs/100))) AS mfforbrug FROM materiale_forbrug mf "_
	    &" WHERE mf.forbrugsdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND usrid = "& oRec("mid") &" GROUP BY mf.usrid"
	    
	    'Response.Write strSQLmf & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLmf, oConn, 3
	    while not oRec2.EOF
	    mfForbrug(x) = oRec2("mfforbrug")/1
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	    
	    
	    end if
	    
	    
	    
	    if visning = 1 then
	   
	    '*** Interne Timer ***'
	    strSQLif = "SELECT t.tid, sum(t.timer) AS ikkefaktimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND (t.tfaktim = 2 OR t.tfaktim = 13 OR t.tfaktim = 14 OR t.tfaktim = 20 OR t.tfaktim = 21)"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	    oRec2.open strSQLif, oConn, 3
	    while not oRec2.EOF
	    ifTimer(x) = oRec2("ikkefaktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    
	         if len(trim(ifTimer(x))) <> 0 then
	         ifTimer(x) = ifTimer(x)
	         else
	         ifTimer(x) = 0
	         end if
	         
	    
	    '*** Salg & Newbizz timer***'
	    strSQLsalg = "SELECT t.tid, sum(t.timer) AS satimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND (t.tfaktim = 6)"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	    oRec2.open strSQLsalg, oConn, 3
	    while not oRec2.EOF
	    sTimer(x) = oRec2("satimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    
	         if len(trim(sTimer(x))) <> 0 then
	         sTimer(x) = sTimer(x)
	         else
	         sTimer(x) = 0
	         end if
	         
	         
	     '*** Frokost/Pause***'
	    strSQLsalg = "SELECT t.tid, sum(t.timer) AS frotimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND (t.tfaktim = 10)"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	    oRec2.open strSQLsalg, oConn, 3
	    while not oRec2.EOF
	    fpTimer(x) = oRec2("frotimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    
	         if len(trim(fpTimer(x))) <> 0 then
	         fpTimer(x) = fpTimer(x)
	         else
	         fpTimer(x) = 0
	         end if
	         
	         
	    '*** Syg ****'
	    strSQLsyg = "SELECT t.tid, sum(t.timer) AS sygtimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND t.tfaktim = 20 "_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	    oRec2.open strSQLsyg, oConn, 3
	    while not oRec2.EOF
	    sygTimer(x) = oRec2("sygTimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    
	         if len(trim(sygTimer(x))) <> 0 then
	         sygTimer(x) = sygTimer(x)
	         else
	         sygTimer(x) = 0
	         end if
	  
	  
	  
	    
	    '*** Barn syg ***'
	    strSQLbsyg = "SELECT t.tid, sum(t.timer) AS bsygtimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND t.tfaktim = 21 "_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	    oRec2.open strSQLbsyg, oConn, 3
	    while not oRec2.EOF
	    BarnsygTimer(x) = oRec2("bsygTimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    
	         if len(trim(BarnsygTimer(x))) <> 0 then
	         BarnsygTimer(x) = BarnsygTimer(x)
	         else
	         BarnsygTimer(x) = 0
	         end if
	    
	    end if
	    
	    
	     if len(trim(realTimer(x))) <> 0 then
	     realTimer(x) = realTimer(x)
	     else
	     realTimer(x) = 0
	     end if
	     
	    
	   
	     if len(trim(normTimer(x))) <> 0 then
	     normTimer(x) = normTimer(x)
	     else
	     normTimer(x) = 0
	     end if
	     
    	 
    	 if visning <> 3 then
    	 
    	 
    	  if len(trim(afspTimer(x))) <> 0 then
	     afspTimer(x) = afspTimer(x)
	     else
	     afspTimer(x) = 0
	     end if
	     
	     if len(trim(afspTimerBr(x))) <> 0 then
	     afspTimerBr(x) = afspTimerBr(x)
	     else
	     afspTimerBr(x) = 0
	     end if
	     
	     
	     if len(trim(feriePLTimer(x))) <> 0 then
	     feriePLTimer(x) = feriePLTimer(x)
	     else
	     feriePLTimer(x) = 0
	     end if
	     
	     if len(trim(ferieAFTimer(x))) <> 0 then
	     ferieAFTimer(x) = ferieAFTimer(x)
	     else
	     ferieAFTimer(x) = 0
	     end if
    	  
    	 
	     if len(trim(resTimer(x))) <> 0 then
	     resTimer(x) = resTimer(x)
	     else
	     resTimer(x) = 0
	     end if
    	 
	     if len(trim(fakTimer(x))) <> 0 then
	     fakTimer(x) = fakTimer(x)
	     else
	     fakTimer(x) = 0
	     end if
	     
	     if len(trim(mfForbrug(x))) <> 0 then
	     mfForbrug(x) = mfForbrug(x)
	     else
	     mfForbrug(x) = 0
	     end if
	    
	     end if
	    
	x = x + 1
	oRec.movenext
    wend
	oRec.close
	
	
	
	if visning = 1 then%>
	 <table cellspacing=1 cellpadding=3 border=0 bgcolor="#5C75AA" width=100%>
	 <tr bgcolor="#5582d2">
	 <td class=alt><b><%=tsa_txt_147%></b></td>
	 <td class=alt colspan=8><b><%=tsa_txt_148%></b> <br />
	 <%=tsa_txt_149%></td>
	 <td class=alt><b><%=tsa_txt_150%></b></td>
	 <td class=alt colspan=3><b><%=tsa_txt_174%></b></td>
	 <td class=alt colspan=3><b><%=tsa_txt_152%></b></td>
	 <td class=alt colspan=2><b><%=tsa_txt_153%></b></td>
	 <td class=alt><b><%=tsa_txt_154%></b></td>
	 <td class=alt><b><%=tsa_txt_155%></b></td>
	  <td class=alt><b><%=tsa_txt_156%></b></td>
	 </tr>
	 <tr bgcolor="#8CAAe6">
	  <td class=alt>
          &nbsp;</td>
	 <td class=alt_lille align=right><%=tsa_txt_157%></td>
	 <td class=alt_lille align=right><%=tsa_txt_158%></td>
	 <td class=alt_lille align=right><%=tsa_txt_259%></td>
	 <td class=alt_lille align=right><%=tsa_txt_159%></td>
	 <td class=alt_lille align=right><%=tsa_txt_160%></td>
	 <td class=alt_lille align=right><%=tsa_txt_161%></td>
	 <td class=alt_lille align=right><%=tsa_txt_162%></td>
	  <td class=alt_lille align=right><%=tsa_txt_163%></td>
	  <td class=alt_lille align=right><%=tsa_txt_150%></td>
	 <td class=alt_lille align=right><%=tsa_txt_164%></td>
	 <td class=alt_lille align=right><%=tsa_txt_165%></td>
	 <td class=alt_lille align=right><%=tsa_txt_166%></td>
	 <td class=alt_lille align=right><%=tsa_txt_167%></td>
	 <td class=alt_lille align=right><%=tsa_txt_168%></td>
	 <td class=alt_lille align=right><%=tsa_txt_166%></td>
	 <td class=alt_lille align=right><%=tsa_txt_169%></td>
	 <td class=alt_lille align=right><%=tsa_txt_170%></td>
	 <td class=alt_lille align=right><%=tsa_txt_148%></td>
	 <td class=alt_lille align=right><%=tsa_txt_148%></td>
	  <td class=alt_lille align=right><%=tsa_txt_171%></td>
	 </tr>
	 
	 <%
	 for x = 0 to x - 1 
	 
	 select case right(x, 1)
	 case 0,2,4,6,8
	 bgcl = "#FFFFFF"
	 case else
	 bgcl = "#D6DFf5"
	 end select%>
	 <tr bgcolor="<%=bgcl %>">
	 <td class=lille><b><%=left(medarbNavn(x), 10) %> (<%=medarbNr(x) %>) </b> <%=medarbInit(x) %></td>
	 <td align=right class=lille><b><%=formatnumber(realTimer(x),2)%></b> (<%=formatnumber(realTimer(x)/weekDiff,2)%>&nbsp;<%=tsa_txt_179%>)</td>
	 <td align=right class=lille><%=formatnumber(normTimer(x),2)%></td>
	 
	 <td align=right class=lille>
	 <%if normTimerUge(x) > 1 then %>
	 <%=formatnumber(normTimerUge(x), 1) %>
	 <%else%>
	 0
	 <%end if %>
	 </td>
	 
	 <td align=right class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 <td align=right class=lille><%=formatnumber((realTimer(x) - (ifTimer(x) + sTimer(x))),2)%></td>
	 <td align=right class=lille><%=formatnumber(ifTimer(x),2)%></td>
	  <td align=right class=lille><%=formatnumber(sTimer(x),2)%></td>
	  <%
	  if realTimer(x) <> 0 then
	  iebal = ((ifTimer(x) + sTimer(x))/realTimer(x)) * 100 
	  else
	  iebal = 0
	  end if%>
	  
	  <td align=right class=lille><%=formatnumber(iebal,2)%>%</td>
	  <td align=right class=lille><%=formatnumber(fpTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(afspTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(afspTimerBr(x),2) %></td>
	 <td align=right class=lille><b><%=formatnumber(afspTimer(x) - afspTimerBr(x),2) %></b></td>
	 <td align=right class=lille><%=formatnumber(feriePLTimer(x)/(normTimerUge(x)),2) %></td>
	 <td align=right class=lille><%=formatnumber(ferieAFTimer(x)/(normTimerUge(x)),2) %></td>
	 <td align=right class=lille><b><%=formatnumber(feriePLTimer(x)/(normTimerUge(x)) - ferieAFTimer(x)/(normTimerUge(x)),2) %></b></td>
	 <td align=right class=lille><%=formatnumber(sygTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(BarnsygTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(resTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(fakTimer(x),2) %></td>
	 <td align=right class=lille><%=formatnumber(mfForbrug(x)) & " DKK" %></td>
	 </tr>
	 
	 
	 
	 <%next%>
	 </table>
	 <%
	 
	 else
	 
	 %>
	 <table cellspacing=0 cellpadding=0 border=0 width=100%>
	 
	 <%
	 for x = 0 to x - 1 
	 
	 %>
	 <tr>
	 <td colspan=2 class=lille style="border-bottom:1px silver dashed;"><%=tsa_txt_147%></td>
	 <td colspan=2 class=lille style="border-bottom:1px silver dashed;" align=right><%=medarbNavn(x) %> (<%=medarbNr(x) %>)  <%=medarbInit(x) %></td>
	 </tr>
	 <tr>
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_148%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_172%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b></td>
	 </tr>
	 <tr>
	 <td style="border-bottom:1px silver dashed;">
         &nbsp;</td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber(realTimer(x),2) %></b>
	 <%if visning <> 3 then %>
	  (<%=formatnumber(realTimer(x)/weekDiff,2)%>&nbsp;<%=tsa_txt_179%>)
	  <%end if %></td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 </tr>
	
	 <%if visning <> 3 then %>
	  <tr>
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_174%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_164%></b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_165%></b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_166%></b></td>
	 </tr>
	  <tr>
	 <td style="border-bottom:1px silver dashed;">
         &nbsp;</td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x),2) %></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x),2) %></td>
	  <td  align=right  class=lille style="border-bottom:1px silver dashed;"><b><%=formatnumber(afspTimer(x) - afspTimerBr(x),2) %></b></td>
	 </tr>
	  <tr>
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_152%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_175%></b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_168%></b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_166%></b></td>
	 </tr>
	 <tr>
	 <td style="border-bottom:1px silver dashed;">
         &nbsp;</td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(feriePLTimer(x),2) %></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(ferieAFTimer(x),2) %></td>
	  <td  align=right  class=lille style="border-bottom:1px silver dashed;"><b><%=formatnumber(feriePLTimer(x) - ferieAFTimer(x),2) %></b></td>
	 </tr>
	 <tr>
	 <td colspan=2 class=lille style="border-bottom:1px silver dashed;"><%=tsa_txt_176%></td>
	    <td colspan=2 align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(resTimer(x),2) %></td>
	 </tr>
	 <tr>
	 <td colspan=2 class=lille style="border-bottom:1px silver dashed;"><%=tsa_txt_177%></td>
	 <td colspan=2 align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(fakTimer(x),2) %></td>
	 </tr>
	 <tr>
	   <td class=lille colspan=2 style="border-bottom:1px silver dashed;"><%=tsa_txt_178%></td>
	 <td colspan=2 align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(mfForbrug(x)) %></td>
	 </tr>
	 <%end if %>
	 
	 
	 
	 <%next%>
	 </table>
	 <%
	 
	 
	 end if
	
	end function
	
	
	public akttypenavn
	function akttyper(akttype, visning)
	
	select case visning
	case 1, 3
	    select case akttype 
	    case 1 
	    akttypenavn = global_txt_129 '"Fakturerbar"
	    case 2
	    akttypenavn = global_txt_130 '"Km"
	    case 0
	    akttypenavn = global_txt_131 '"Ej fakturerbar"
	    case 6
	    akttypenavn = replace(global_txt_132, "|", "&") '"Salg & NewBizz"
	    case 10
	    akttypenavn = global_txt_133 '"Frokost / Pause"
	    case 11
	    akttypenavn = global_txt_134 '"Ferie planlagt"
	    case 14
	    akttypenavn = global_txt_135 '"Ferie afholdt"
	    case 12
	    akttypenavn = global_txt_136 '"Afspad. optjent"
        case 13
	    akttypenavn = global_txt_137 '"Afspad. brugt"
	    case 20
	    akttypenavn = global_txt_138 '"Syg"
	    case 21
	    akttypenavn = global_txt_139 '"Barn syg"
	    end select
	
	case 2
	            
	            select case akttype 'aktdata(iRowLoop, 11)
				case 0
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_131 &")</font> "
				case 1
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_129 &")</font> "
				case 2
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_130 &")</font>"
				case 6
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& replace(global_txt_132, "|", "&") &")</font>"
				case 10
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_133 &") </font>"
				case 11
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_134 &") </font>"
				case 14
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_135 &") </font>"
				case 12
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_136 &") </font>"
				case 13
				akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_137 &") </font>"
				case 20
	            akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_138 &")</font> "
	            case 21
	            akttypenavn = "&nbsp;<font class=megetlilleblaa>("& global_txt_139 &")</font> "
				end select              
	
	case 4
	    '*** I timereg tabellen skifter Km og Ej fakbar type fra 2 til 5 og fra 0 til 2. **'
	    select case akttype 
	    case 1 
	    akttypenavn = global_txt_129 '"Fakturerbar"
	    case 5
	    akttypenavn = global_txt_130 '"Km"
	    case 2
	    akttypenavn = global_txt_131 '"Ej fakturerbar"
	    case 6
	    akttypenavn = replace(global_txt_132, "|", "&") '"Salg & NewBizz"
	    case 10
	    akttypenavn = global_txt_133 '"Frokost / Pause"
	    case 11
	    akttypenavn = global_txt_134 '"Ferie planlagt"
	    case 14
	    akttypenavn = global_txt_135 '"Ferie afholdt"
	    case 12
	    akttypenavn = global_txt_136 '"Afspad. optjent"
        case 13
	    akttypenavn = global_txt_137 '"Afspad. brugt"
	    case 20
	    akttypenavn = global_txt_138 '"Syg"
	    case 21
	    akttypenavn = global_txt_139 '"Barn syg"
	    end select
	
	end select
	
	end function
	
	
    function sltque(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="slet" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffe1; visibility:visible; border:2px #8cAAe6 dashed;">
	<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1">
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><h4>Slet?</h4> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - slet" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %>><img src="../ill/<%=usejaimg %>.gif" alt="Ja - slet" border="0"></a>
		</td>
		<td align=right>
		<a href="Javascript:history.back()"><img src="../ill/stop.gif" alt="Nej - tilbage" border="0"></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function eksportogprint(ptop,pleft,pwdt)
	%>
	<div id=eksport style="position:absolute; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px silver solid; padding:3px 3px 3px 3px;">
    <table cellpadding=5 cellspacing=0 border=0 width=100%>
    <tr>
    <td height=30 bgcolor="#FFFFe1" style="border-bottom:1px #ffff99 solid;" colspan=2><b>Eksport & Print:</b></td>
    </tr>
	<%
	end function
	
	
	function filterheader(ptop,pleft,pwdt,pTxt)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="filter" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function
   
     
     
    function tableDiv(tTop,tLeft,tWdth)
	%>
	<div id="maintable" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:1px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible;">
    <%
	end function
    
     
     
    function sideinfo(itop,ileft,iwdt)
	%>
	<div id="sideinfo" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:2px red dashed; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFFFF" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td bgcolor="#FFFFFF" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Hj�ælp & Sideinfo:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFE1" style="padding:10px;">
	
    
	
	<%
	end function
	
	
	function opretNy(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px; 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td>
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	
	function valutakoder(i, valuta)
	%>
	<select name="FM_valuta_<%=i %>" id="FM_valuta_<%=i %>" style="width:65px; font-size:9px;">
		    
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = oRec3("id") then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
		    
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
	<%
	end function
	
	
	public dblKurs
	function valutaKurs(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       dblKurs = replace(oRec("kurs"), ",", ".")
       end if 
       oRec.close
	end function
	
	
	public valBelobBeregnet
	function beregnValuta(belob,frakurs,tilkurs)
	
	        valBelobBeregnet = belob * (frakurs/tilkurs)
    
    if len(valBelobBeregnet) <> 0 then
    valBelobBeregnet = valBelobBeregnet
    else
    valBelobBeregnet = 0
    end if
    
    
    
    valBelobBeregnet = valBelobBeregnet/1
    'Response.Write valBelobBeregnet & "<br>"
    'Response.flush
	end function
	
    %>