
<!--#include file="../xml/global_xml_inc.asp"-->

<!--#include file="cls_menu.aspx"-->
<!--#include file="cls_help.asp"-->
<!--#include file="cls_projektgrp.asp"-->
<!--#include file="cls_medarb.asp"-->
<!--#include file="cls_afstem.asp"-->

<!--#include file="global_func2.asp"-->
<!--#include file="global_func3.asp"-->
<!--#include file="global_func4.asp"-->

<%
'*** Er global inc_ inkluderet på den aktuelle side?? 
global_inc = "j"


'*** Afsluttede uger *****
public afslUgerMedab
redim afslUgerMedab(208) '4 år
function afsluger(medarbid, stdato, sldato)
		
		strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, YEAR(u.uge) AS aar, "_
		&" u.id, u.mid FROM ugestatus u WHERE mid =  "& medarbid &""_
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




public licensindehaverKid, licensindehaverKnavn
function licKid()

    strSQL4 = "SELECT kid, kkundenavn FROM kunder WHERE useasfak = 1"
	oRec4.open strSQL4, oConn, 3
	If not oRec4.EOF then
	
	licensindehaverKid = oRec4("kid")
	licensindehaverKnavn = oRec4("kkundenavn")
	
	end if
	oRec4.close

end function

public startDatoAar, startDatoMd, startDatoDag, licensklienter
function licensStartDato()

key = "2.151-3112-B000"
	
	strSQL4 = "SELECT l.key, licensstdato, klienter FROM licens l WHERE id = 1"
	oRec4.open strSQL4, oConn, 3
	If not oRec4.EOF then
	
	key = oRec4("key")
	licensstdato = oRec4("licensstdato")
	licensklienter = oRec4("klienter")
	
	end if
	oRec4.close
	
	
	'licensnb = right(key, 3)
	'Response.Write licensnb &"<br>"
	
	'if cint(licensnb) < 69 then
	'startDatoAar = "200" & mid(key, 5,1)
	'startDatoMd = mid(key, 9,2)
	'startDatoDag = mid(key, 7,2)
	'else
	'startDatoAar = "20" & mid(key, 5,2)
	'startDatoMd = mid(key, 10,2)
	'startDatoDag = mid(key, 8,2)
	'end if
    
    
    startDatoDag = day(licensstdato)
    startDatoMd = month(licensstdato)
    startDatoAar = year(licensstdato)

end function

public browstype_client, user_agent_txt
function browsertype()

	if instr(Request.ServerVariables("HTTP_USER_AGENT") , "Firefox") <> 0 then
	browstype_client = "mz"
	else
	browstype_client = "ie"
	end if
	
	user_agent_txt = request.servervariables("HTTP_USER_AGENT")

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

public kmDialogOnOff
function erkmDialog()
'** Km dialog **'
	kmDialogOnOff = 0
	strSQL = "SELECT kmdialog FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	kmDialogOnOff = oRec("kmdialog") 
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

public stempelurOn
function erStempelurOn()
'** Stempelur ***'
	stempelurOn = 0
	strSQL = "SELECT stempelur FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	stempelurOn = oRec("stempelur") 
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
		
		
		
		
		
		
		
		
		public SideHeader
		'public pkt1oskrift
		function kundelogin_mainmenu(pkt, lto, kundeid)
		
		 
		
		
		'*** Konfigurerer menupkt navne efter lto. ***
			select case lto 
			case "kringit"
			spkt1 = "Seviceordre"
			spkt2 = "Aftaler"
			spkt3 = "Filarkiv"
			spkt4 = "Infobase"
			'pkt1oskrift = "Seviceordre"
			case else
			spkt1 = "Timeregistreringer"
			spkt2 = "Aftaler"
			spkt3 = "Filarkiv"
			spkt4 = "Infobase"
			'pkt1oskrift = "Job"
			end select
			
			select case pkt 
			case 1
			SideHeader = "Timeregistrerings log"
			case 2
			SideHeader = "Aftaleoversigt"
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
					
					
					
						<div id=Div1 style="position:absolute; background-color:#ffffff; width:200px; left:20px; top:80px; border:1px silver solid; padding:3px 3px 3px 3px;">
                        <table cellpadding=5 cellspacing=0 border=0 width=100%>
                        <tr>
                        <td height=30 bgcolor="#ffccff" style="border-bottom:1px #ff00ff solid;" colspan=2><b>Registrerings-log for:</b></td>
                        </tr>
					
					<tr>
						<td valign="top" style="padding:5px;"><%=strKnavn%><br>
						<%=strKadr%><br><%=strKpostnr%>&nbsp;<%=strBy%>
						<%if len(trim(intTlf)) <> 0 then%>
						<br>Tlf:&nbsp;<%=intTlf%><br>
						<%end if%>
						&nbsp;</td>
					</tr>
					</table>
					</div>
					
					
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
			    call mmenuPkt(3, "56", "filer.asp?kundeid="&intKid&"&jobid=0&kundelogin=1",spkt3,pkt)
			    'call mmenuPkt(4, "38", "infobase.asp?menu=kund&usekview=j&id="&intKid&"&kontaktid="&intKid&"&FM_seljob="&jobnr,spkt4,pkt)
    		
		    else
    		
		    call mmenuPkt(2,"52", "#","&nbsp;",0)
	        call mmenuPkt(2,"52", "#","&nbsp;",0)
	        call mmenuPkt(2,"52", "#","&nbsp;",0)
	        call mmenuPkt(2,"52", "#","&nbsp;",0)
    	    
    		
    		
		    end if
		
		
		
		
		
		
		call mmenuTableEnd(1)%>
		
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
					
					call helligdage(CheckDay&"/"&md&"/"&ye, 0)
					
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
	
	if len(trim(totalmin)) <> 0 then
	totalmin = formatnumber(totalmin, 0)
	totalmin = replace(totalmin, ".", "")/1
	else
	totalmin = 0
	end if
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
		
		for tomb = 0 to UBOUND(timTemp_komma)
			
			if tomb = 0 then
			thoursTot = timTemp_komma(tomb)
			end if
			
			if tomb = 1 then
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
			
			if cint(tminTot) = 0 then
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
		<div id="firmalogo" style="position:absolute; left:<%=plefther%>; top:<%=ptopher%>; z-index=10; padding:0px; border:0px silver solid; overflow:hidden;">
			<%
			
			
			
			
			strSQL = "SELECT useasfak, logo, id, filnavn, kkundenavn FROM kunder "_
			&" LEFT JOIN filer ON (filer.id = kunder.logo) WHERE useasfak = 1" 
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
			
			kundenavn = oRec("kkundenavn")
			
			oRec.close
			
			Response.Write "<br>Denne service er leveret af:<br><b> " & kundenavn & "</b><br>"
			
			Response.write logonavn
			
			
			%>
		</div>
		
	<% end function
	
	'**************************************************************
	
	public ntimPer, ntimMan, ntimTir, ntimOns, ntimTor, ntimFre, ntimLor, ntimSon
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
	
	
	'** Hvis medarbjeder ansat dato er før licens startdato benyttes licens startDato **'
	call licensStartDato()
	ltoStDato = startDatoDag &"/"& startDatoMd &"/"& startDatoAar
	if cdate(ltoStDato) > cDate(mtyperIntvDato(0)) then
	mtyperIntvDato(0) = ltoStDato
	end if
	 
	 
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
			            '*** Skal først tjekke dage efter ansat dato ***'
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
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				
					
					'if medid = 1 then
					'Response.write strSQLnt & " n:"& n & "<hr>"
					'Response.flush
					'end if 
					
					oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						ntimPerThis = oRec3("timer")
						
						select case n 
						case 7
						ntimSon = ntimPerThis
						case 1
						ntimMan = ntimPerThis
						case 2
						ntimTir = ntimPerThis
						case 3
						ntimOns = ntimPerThis
						case 4
						ntimTor = ntimPerThis
						case 5
						ntimFre = ntimPerThis
						case 6
						ntimLor = ntimPerThis
						end select
						
						
					end if
					oRec3.close 
					
					'Response.write "datoCount " & formatdatetime(datoCount, 2) & "<br>"
					'Response.flush
					
				    call helligdage(datoCount, 0)
				        
				        if cint(erHellig) = 1 then
				            select case n 
						    case 7
						    ntimSon = 0
						    case 1
						    ntimMan = 0
						    case 2
						    ntimTir = 0
						    case 3
						    ntimOns = 0
						    case 4
						    ntimTor = 0
						    case 5
						    ntimFre = 0
						    case 6
						    ntimLor = 0
						    end select
						end if
					
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
	
	
	
	public normtimerStDag
	function nortimerStandardDag(mtypeUse)
	
	
	ntimStdag = 0
	
	    for n = 1 to 7 
	                    select case n 
						case 7
						normdag = "normtimer_son"
						case 1
						normdag = "normtimer_man"
						case 2
						normdag = "normtimer_tir"
						case 3
						normdag = "normtimer_ons"
						case 4
						normdag = "normtimer_tor"
						case 5
						normdag = "normtimer_fre"
						case 6
						normdag = "normtimer_lor"
						end select
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				
		            oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						ntimStdag = ntimStdag + oRec3("timer")
						
					end if
					oRec3.close 
				
				next
	
	            normtimerStDag = ntimStdag/5
	
	end function
	
	function AjaxUpdate(table, column, msg)
	if request.Form("value") <> "" AND table <> "" AND column <> "" then
	strSQL = "Update " & table & " set " & column & " = '" & request.Form("value") & "' where id = " & request.Form("id")
	
	'Response.Write strSQL
	'Response.flush
	oConn.execute(strSQL)
	Response.Write "timeout notifikation: <br />" & msg
	Response.End
	end if
	end function
	
	function infoUnisport(uWdt, uTxt)%>
	 <div style="width:<%=uWdt%>px; padding:3px 7px 3px 7px; border:1px #6CAE1C solid; background-color:#DCF5BD;"><%=uTxt %></div>
    <%end function
    
    function infoUnisportSp(uWdt, uTxt)%>
	 <span style="width:<%=uWdt%>px; padding:3px 7px 3px 7px; border:1px #6CAE1C solid; background-color:#DCF5BD;"><%=uTxt %></span>
    <%end function
    
    function infoUnisportAB(uWdt, uTxt, uTop, uLeft)%>
	 <div style="position: absolute; top:<%=uTop%>; left:<%=uLeft%>; width:<%=uWdt%>px; padding:3px 7px 3px 7px; border:1px #6CAE1C solid; background-color:#DCF5BD;"><%=uTxt %></div>
    <%end function
	
	
	
	
	public jquerystrTxt
    function jquery_repl_spec(jquerystr)
    
    jquerystr = replace(jquerystr, "ø", "&oslash;")
    jquerystr = replace(jquerystr, "æ", "&aelig;")
    jquerystr = replace(jquerystr, "å", "&aring;")
    jquerystr = replace(jquerystr, "Ø", "&Oslash;")
    jquerystr = replace(jquerystr, "Æ", "&AElig;")
    jquerystr = replace(jquerystr, "Å", "&Aring;")
    jquerystr = replace(jquerystr, "Ö", "&Ouml;")
    jquerystr = replace(jquerystr, "ö", "&ouml;")
    jquerystr = replace(jquerystr, "Ü", "&Uuml;")
    jquerystr = replace(jquerystr, "ü", "&uuml;")  
    
    jquerystrTxt = jquerystr
    
    end function
	
	

%>
    