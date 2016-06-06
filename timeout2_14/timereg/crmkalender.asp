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
	func = request("func")
	thisfile = "crmkalender"
	%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<script LANGUAGE="javascript">	
				<!--
				function nyaktion(thisdag, thismrd, thisaar, klslet){
				document.all["nyaktion"].style.display = ""
				document.all["FM_start_dag_per"].value = thisdag;
				document.all["FM_start_mrd_per"].value = thismrd;
				document.all["FM_start_aar_per"].value = thisaar;
				document.all["FM_klslet_personal"].value = klslet;
				document.all["func"].value = "dbopr";
				document.all["submit"].value = "Opret!" 
				document.all["FM_note_personal"].focus()
				document.all["FM_note_personal"].select()
				}
				
				function redaktion(thisdag, thismrd, thisaar, klslet, textvalue, crmaktion){
				document.all["nyaktion"].style.display = ""
				document.all["FM_start_dag_per"].value = thisdag;
				document.all["FM_start_mrd_per"].value = thismrd;
				document.all["FM_start_aar_per"].value = thisaar;
				document.all["FM_klslet_personal"].value = klslet;
				document.all["FM_note_personal"].value = document.getElementById("text_"+textvalue+"").value
				//textvalue;
				document.all["func"].value = "dbred";
				document.all["submit"].value = "Opdater!"
				document.all["crmaktion"].value = crmaktion;
				document.all["FM_note_personal"].focus()
				document.all["FM_note_personal"].select()
				} 
				
				function closepersonal() {
				document.all["nyaktion"].style.display = "none"
				}
				// End -->
			</script>
			<!--#include file="inc/dato.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<!--#include file="inc/convertDate.asp"-->
			<!--#include file="inc/isint_func.asp"-->
										
			<%
					'************ Vmenu/kalender værdier på filtre *****************
					shownumofdays = request("shownumofdays") 
					if shownumofdays = 1 then
					viewnumofdays = 2
					shownumofdays = 1
					fiveoronewidth = 500
					else
					viewnumofdays = 6
					shownumofdays = 5
					fiveoronewidth = 99
					end if
					
					
					if id = 0 OR len(id) = 0 then
					useidKri = "AND kundeid <> " & id & ""
					else
					useidKri = "AND kundeid = " & id & ""
					end if
					
					'*** Denne bruges bl.a i kalender
					select case level 
					case "1", "2"
							if medarb = 0 OR len(medarb) = 0 then
							usemedarbKri = " AND editorid <> 0 " & useidKri
							else
							usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & medarb & "' " & useidKri
							end if
					case else
					usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & session("Mid") & "' " & useidKri
					end select
					
			%>
			<!--#include file="../inc/regular/calender.asp"-->
			<!--#include file="inc/timereg_dage_inc.asp"-->
			<!--#include file="../inc/regular/vmenu.asp"-->
			<%
			dato = useDate 'fra kalender
			sqlDato = convertDateYMD(useDate)
			
		'***** Ny aktions div  ****%>
		<div id="nyaktion" style="position:absolute; left:300; top:200; width:300; display:none; z-index:2000; background-color:#FFFFFF; border-left: 1px solid #003399; border-top: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399;">
		
		<table cellspacing="0" cellpadding="0" border="0" bgcolor="#d6dff5" width="300">
		<tr><form action="crmhistorik.asp?menu=crm&personal=y" method="post">
		<td style="padding-left:5; padding-top:2;"><b>Personlig note:</b></td>
		<td align="right"><a href="#" onClick="closepersonal();"><img src="../ill/luk_xp.gif" width="30" height="28" alt="Luk vindue" border="0"></a></td>
		</tr>
		<tr>
		<td colspan="2" style="padding-left:5; padding-top:2;">
		<textarea cols="33" rows="7" name="FM_note_personal"></textarea><br>
		<input type="hidden" name="FM_start_dag_per" value="0">
		<input type="hidden" name="FM_start_mrd_per" value="0">
		<input type="hidden" name="FM_start_aar_per" value="0">
		<input type="hidden" name="FM_klslet_personal" value="0">
		<input type="hidden" name="crmaktion" value="0">
		<input type="hidden" name="func" value="dbopr">
		<input type="hidden" name="kalviskunde" value="0">
		</td></tr>
		<tr>
			<td colspan="2" align="center" height="30" valign="middle"><input type="submit" name="submit" id="submit" value="Opret!"></td></tr>
		</tr></form>
		</table>
		</div>
		<%'************************
		'***** Side indhold ****%>	
		<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
			<table cellspacing="0" cellpadding="0" border="0" width="550">
			<tr>
		    <td valign="top"><br>
			<img src="../ill/header_crmkal.gif" alt="" border="0"><hr align="left" width="530" size="1" color="#000000" noshade>
			</td></tr></table>
			<div id="opretny" style="position:absolute; left:390; top:30; visibility:visible;">
			<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&shownumofdays=<%=shownumofdays%>&func=opret&id=0&ketype=e&selpkt=kal&strdag=<%=strDag%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&showinwin=j')">Opret ny aktion <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
		</div>
		<br>
		<table cellspacing="0" cellpadding="0" border="0" width="532">
			<tr bgcolor="#5582D2">
			<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
			<td width="524" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;" class="alt">
			<font class="stor-hvid">&nbsp;Uge <%=(datepart("ww", dato, 2, 2))%></font>
			<img src="../ill/blank.gif" width="300" height="1" alt="" border="0">Vis: <a href="crmkalender.asp?menu=crm&shownumofdays=1&strdag=<%=strDag%>&strmrd=<%=strMrd%>&straar=<%=strAar%>" class=alt><b><u>1 dag</u></b></a> / <a href="crmkalender.asp?menu=crm&shownumofdays=5&strdag=<%=strDag%>&strmrd=<%=strMrd%>&straar=<%=strAar%>" class=alt><b><u>5 dage's</u></b></a> oversigt.</td>
			<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
		</tr>
		</table>
		
		<table cellspacing="0" cellpadding="0" border="0" width="532">
		<%
		p = 1
		
		call dageDatoer_crmKal(viewnumofdays)
		%>
		</tr>
		<%
					'******* daglige crm aktiviteter ****
					counter = 7
					klokkeslet = counter&":00:00"
				
					for counter = 7 to 21
					klokkeslet = counter&":00:00"
					klokkesletNext = (counter+1)&":00:00"
					
					
					
					'******* crm aktiviteter der løber over flere dage ****
					'if counter = 7 then
					'	
					'	if shownumofdays = 1 then
					'	startdag_fld = convertDateYMD(dato)
					'	slutdag_fld = convertDateYMD(dato)
					'	else
					'	startdag_fld = convertDateYMD(tjekdag(1))
					'	slutdag_fld = convertDateYMD(tjekdag(7))
					'	end if
						
					'	strSQL = "SELECT DISTINCT(crmhistorik.id) AS crmaktionid, "_
					'	&" crmhistorik.kundeid, crmdato, crmstatus.navn AS statusnavn, crmdato_slut, kontaktemne, crmemne.navn AS enavn,"_
					'	&" kunder.kkundenavn, kunder.kid"_
					'	&" FROM crmhistorik, aktionsrelationer "_
					'	&" LEFT JOIN crmemne ON (crmemne.id = kontaktemne) "_
					'	&" LEFT JOIN crmstatus ON (crmstatus.id = status) "_
					'	&" LEFT JOIN kunder ON (kunder.kid = kundeid) "_
					'	&" WHERE crmdato <= '"& slutdag_fld &"' AND crmdato_slut > '"& startdag_fld &"' AND crmdato <> crmdato_slut "& usemedarbKri  
						
					'	aktwritten = 0
					'	oRec.open strSQL, oConn, 3
					'	While not oRec.EOF
					'	numofdaysmisty = oRec("crmdato_slut") - oRec("crmdato") 
					'	>						
					'	<tr>
					'		<td height="20" bgcolor="#EFF3FF" style="padding-left:5; border-left:#5582d2 1px solid; border-right:#5582d2 1px solid; border-bottom:#5582d2 1px solid;">&nbsp;</td>
					'	   	<%for intcounter = 2 to 6
					'		'alldaydatothis = 
					'		if cint(aktwritten) = 0 then
					'			if cdate(convertDateYMD(oRec("crmdato"))) <= cdate(convertDateYMD(tjekdag(intcounter))) then>
					'			<td valign="bottom" bgcolor="#FFFFFF" style="padding-left:2; border-right:#d6dff5 1px solid; border-bottom:#d6dff5 1px solid;">
					'				<table cellspacing=0 cellpadding=0 border=0 height=100% width=100><tr>
					'					<td bgcolor="Thistle" valign="bottom" style="padding-left:2; padding-right:2; padding-top:2; padding-bottom:2; border-bottom:darkred 1px dashed; border-top:darkred 1px dashed;">
					'					<img src="ill/scalag4.gif" width="2" height="10" alt="" border="0">&nbsp;<font size=1><%=FormatDateTime(oRec("crmdato"),2)><br>
					'					<img src="ill/scalag1.gif" width="2" height="10" alt="" border="0">&nbsp;<%=FormatDateTime(oRec("crmdato_slut"),2)><br></font>
					'					<a href="kunder.asp?menu=crm&func=red&id=<=oRec("kid")>" class="kalblue"><=left(oRec("kkundenavn"),12)></a><br>
					'					<font class=megetlillesort><u>Emne:</u>&nbsp;<=oRec("enavn")%
					'					<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=red&id=<%=oRec("kid")>&crmaktion=<%=oRec("crmaktionid")>&selpkt=kal&kalemner=<%=emner>&kalstatus=<%=status>&kalviskunde=<%=id>&kaldato=<%=dato>&kalmedarb=<%=medarb>&showinwin=j')" class="vmenuglobal"><img src="../ill/blyant.gif" width="12" height="11" alt="" border="0"></a>
					'					<br>
					'					<font class=megetlillesort><u>Status:</u>&nbsp;<%=oRec("statusnavn")></font>
					'					<br>
					'					<%if cdate(convertDateYMD(oRec("crmdato"))) < cdate(convertDateYMD(tjekdag(intcounter))) OR shownumofdays = 1 AND cdate(convertDateYMD(oRec("crmdato"))) < cdate(convertDateYMD(dato)) then>
					'					<br><img src="../ill/pile_tilbage.gif" width="10" height="6" alt="<%=formatdatetime(oRec("crmdato"),1)>" border="0"><%end if>
					'					</td>
					'					<td bgcolor="Thistle" valign="bottom" style="padding-top:2; padding-bottom:2; padding-left:2; padding-right:2; border-bottom:darkred 1px dashed; border-top:darkred 1px dashed;"><a href="crmhistorik.asp?menu=crm&ketype=e&func=hist&id=<%=oRec("kid")>&selpkt=hist"><img src="../ill/hist_lille.gif" width="12" height="12" alt="Se historik" border="0"></a><br><img src="../ill/blank.gif" width="1" height="3" alt="" border="0"><br>
					'					<a href="javascript:NewWin_large('../inc/regular/kundelogview.asp?useKid=<%=oRec("kid")>')" target="_self"><img src="../ill/joblog_small.gif" width="12" height="12" alt="Se kundelog" border="0"></a><br><img src="../ill/blank.gif" width="1" height="3" alt="" border="0"><br>
					'					<a href="crmhistorik.asp?menu=crm&func=slet&id=<%=id>&crmaktion=<%=oRec("crmaktionid")>&medarb=<%=medarb>&selpkt=kal"><img src="../ill/slet_crmkal.gif" width="12" height="12" alt="" border="0"></a>
					'					<%if ((intcounter = 6 OR shownumofdays = 1) AND cdate(convertDateYMD(oRec("crmdato_slut"))) > cdate(convertDateYMD(tjekdag(intcounter)))) then>
					'					<br><img src="../ill/pile_selected.gif" width="10" height="6" alt="<%=formatdatetime(oRec("crmdato_slut"),1)>" border="0">
					'					<%end if>
					'					</td></tr>
					'				</table>
					'			</td>
					'			<%aktwritten = 1%
					'			<%
					'			else
					'				if shownumofdays <> 1 then>
					'			<td bgcolor="#FFFFFF" style="padding-left:2; border-right:#d6dff5 1px solid; border-bottom:#d6dff5 1px solid;">&nbsp;</td>
					'			<%  end if
					'			end if
					'		
					'		
					'		else
					'			if shownumofdays <> 1 then
					'			if cdate(convertDateYMD(oRec("crmdato_slut"))) >= cdate(convertDateYMD(tjekdag(intcounter))) then>
					'			<td valign="bottom" bgcolor="#FFFFFF" style="border-right:#d6dff5 1px solid; border-bottom:#d6dff5 1px solid;">
					'				<table cellspacing=0 cellpadding=0 border=0 height=100% width=100>
					'					<tr>
					'					<td valign="bottom" align="right" bgcolor="Thistle" style="padding-top:2; padding-bottom:2; padding-left:2; padding-right:6; border-bottom:darkred 1px dashed; border-top:darkred 1px dashed;"><br><img src="../ill/blank.gif" width="70" height="1" alt="" border="0"><br>
					'					<%
					'					if intcounter = 6 AND cdate(convertDateYMD(oRec("crmdato_slut"))) > cdate(convertDateYMD(tjekdag(intcounter))) then>
					'					<img src="../ill/pile_selected.gif" width="10" height="6" alt="<%=formatdatetime(oRec("crmdato_slut"),1)>" border="0">
					'					<%end if></td>
					'				 </tr>
					'				</table>
					'			</td>
					'			<%else>
					'			<td bgcolor="#FFFFFF" style="border-right:#d6dff5 1px solid; border-bottom:#d6dff5 1px solid;">&nbsp;</td>
					'			<%end if
					'			end if
					'		
					'		end if
							
							
					'		next>
						
					'	</tr>
					'	<%
					'	aktwritten = 0
					'	oRec.movenext
					'	wend
					'	oRec.close
		
					'end if
					
					
					
					%>
					<tr bgcolor="#ffffe1">
					
					
					<%
					for incounter = 0 to viewnumofdays - 1
					
					
					if incounter = 0 then%>
							<td valign="top" width=32 style="padding-left:6; padding-right:1; padding-top:4; background-color:WhiteSmoke; border-bottom: 1px solid #5582D2; border-left: 1px solid #5582D2; border-right: 1px solid #5582D2;" height="27"><b><%=FormatDateTime(klokkeslet,4)%></b>&nbsp;</td>
					<%
					else
								if shownumofdays = 1 then
								thisdato = sqlDato
								else
								thisdato = tjekdag(incounter)
								'Response.write thisdato & "<br>"
								end if
								
								strDag_link = day(thisdato)
								strMrd_link = month(thisdato)
								strAar_link = year(thisdato)
								sqlDatoThis = convertDateYMD(thisdato)
								'Response.write sqlDatoThis
							
								if formatdatetime(date(), 0) = formatdatetime(thisdato, 0) then
										marker = left(FormatDateTime(Now,4) , 2)
										if cint(marker) = cint(counter) then
										bordercolor = "Darkred"
										bleft = 1
										bbot = 1
										bright = 1
										btop = 1
										else
										bleft = 0
										bordercolor = "#d6dff5"
										bbot = 1
										bright = 1
										btop = 0
										end if
								else
								bleft = 0
								bbot = 1
								bright = 1
								btop = 0
								bordercolor = "#d6dff5"
								end if%>
							<td width="<%=fiveoronewidth%>" valign="top" style="padding-top:2; padding-left:5; padding-right:3; background-color:#FFFFFF; border-bottom: <%=bbot%>px solid <%=bordercolor%>;  border-left: <%=bleft%>px solid <%=bordercolor%>; border-top: <%=btop%>px solid <%=bordercolor%>;  border-right: <%=bright%>px solid <%=bordercolor%>; z-index:1000;"><a href="#" onClick="nyaktion('<%=strDag_link%>', '<%=strMrd_link%>', '<%=strAar_link%>', '<%=FormatDateTime(klokkeslet,4)%>');" class="vmenu">p+</a>&nbsp;&nbsp;<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&shownumofdays=<%=shownumofdays%>&func=opret&id=0&selpkt=kal&ketype=e&kl=<%=FormatDateTime(klokkeslet,4)%>&strdag=<%=strDag_link%>&strmrd=<%=strMrd_link%>&straar=<%=strAar_link%>&showinwin=j')" class="vmenuglobal" target="_self">a+</a><br>
							<%
								strSQL = "SELECT DISTINCT(crmhistorik.id) AS crmaktionid, crmhistorik.komm AS besk,"_
								&" crmhistorik.editor, crmhistorik.kundeid, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enavn,"_
								&" crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers,"_
								&" crmklokkeslet, crmklokkeslet_slut, kunder.kkundenavn, kunder.kid, kunder.telefon, editorid,"_
								&" medarbejdere.email, serialnb FROM "_ 
								&" crmhistorik, aktionsrelationer "_
								&" LEFT JOIN crmemne ON (crmemne.id = kontaktemne) "_
								&" LEFT JOIN crmstatus ON (crmstatus.id = status) "_
								&" LEFT JOIN crmkontaktform ON (crmkontaktform.id = kontaktform) "_
								&" LEFT JOIN kunder ON (kunder.kid = kundeid) "_
								&" LEFT JOIN medarbejdere ON (medarbejdere.mid = editorid) "_
								&" WHERE crmdato = '"& sqlDatoThis &"' AND (crmdato = crmdato_slut OR COALESCE(crmdato_slut,0) = 0) AND crmklokkeslet BETWEEN '"& klokkeslet &"' AND '"& klokkesletNext &"' "& usemedarbKri &" ORDER BY crmklokkeslet" 
								
								
								'AND (crmdato = crmdato_slut OR crmdato_slut = Null)
								
								if shownumofdays = 1 then
								tdw1 = "100%"
								tdw2 = "100%"
								else
								tdw1 = "80"
								tdw2 = "90"
								end if
								
								
								%>
								<form>
								<%
								'Response.write strSQL
								'Response.flush
								
								oRec.open strSQL, oConn, 3
								foundone = 0
								
								While not oRec.EOF
								
								
								if oRec("kundeid") = -1 then
									strBesk = oRec("besk")
									'Response.write strBesk%>
									<table border=0 cellspacing=0 cellpadding=0 width="<%=tdw1%>" style="border:1px darkred solid;"><tr>
									<td bgcolor="Cornsilk" colspan="2" style="padding-top:2; padding-left:3; padding-right:3; padding-bottom:3;"><a href="#" onClick="redaktion('<%=strDag_link%>', '<%=strMrd_link%>', '<%=strAar_link%>', '<%=FormatDateTime(klokkeslet,4)%>','<%=p%>', '<%=oRec("crmaktionid")%>')"; class=kal_g><%=strBesk%></a><br>
									<a href="crmhistorik.asp?menu=crm&shownumofdays=<%=shownumofdays%>&func=slet&id=<%=id%>&crmaktion=<%=oRec("crmaktionid")%>&medarb=<%=medarb%>&selpkt=kal"><img src="../ill/slet_crmkal.gif" width="12" height="12" alt="" border="0"></a>
									<input type="hidden" name="text_<%=p%>" id="text_<%=p%>" value="<%=strBesk%>">
									<%
									p = p + 1
								else
									%>
									<table border=0 cellspacing=0 style="border:1px darkred solid;" cellpadding=0 width="<%=tdw2%>">
									<tr>
										<td bgcolor="MistyRose" style="padding-top:2; padding-left:3; padding-right:3; padding-bottom:3;" colspan="2">
										<img src="ill/scalag4.gif" width="2" height="10" alt="" border="0">&nbsp;<%=FormatDateTime(oRec("crmklokkeslet"),4)%> - <img src="ill/scalag1.gif" width="2" height="10" alt="" border="0">&nbsp;<%=FormatDateTime(oRec("crmklokkeslet_slut"),4)%></td>
									</tr>
									<tr><td colspan=2 bgcolor="MistyRose" style="padding-top:2; padding-left:3; padding-right:1; padding-bottom:0;">
										<a href="kunder.asp?menu=crm&func=red&id=<%=oRec("kid")%>" class="kalblue"><%=left(oRec("kkundenavn"),12)%></a></td>
									</tr>
									<tr>
										<td colspan=2 bgcolor="MistyRose" style="padding-top:1; padding-left:3; padding-right:1; padding-bottom:0;">
										t:<font class=megetlillesort><%=oRec("telefon")%></font></td>
									</tr>
									<tr>
										<td colspan=2 bgcolor="MistyRose" valign="top" style="padding-top:1; padding-left:3; padding-right:1; padding-bottom:0;">
										k:<font class=megetlillesort ><%
										
										'**************************************************
										'* finder kontaktpers afhængig af om aktionen er
										'* fra før der belv oprettet kontaktpers tabel.
										'**************************************************
										call erDetInt(oRec("kontaktpers"))
										if isInt > 0 then
											if len(oRec("kontaktpers")) > 12 then
											Response.write left(oRec("kontaktpers"), 12) &".." 
											else
											Response.write oRec("kontaktpers")
											end if
										else
										
											strSQL2 = "SELECT id, navn FROM kontaktpers WHERE id = "& oRec("kontaktpers")
											oRec2.open strSQL2, oConn, 3
										
											if not oRec2.EOF then
												if len(oRec2("navn")) > 12 then
												Response.write left(oRec2("navn"), 12) &".." 
												else
												Response.write oRec2("navn")
												end if
											else
												Response.write "Ingen"
											end if
											oRec2.close
										end if
										isInt = 0
										%></font></td>
									</tr>
									<tr>
										<td colspan="2" bgcolor="#FFFFFF" style="padding-top:3; padding-left:3; padding-right:3; padding-bottom:1; border-left:0px; border-right:0px; border-bottom:1px darkred solid; border-top:1px darkred solid;"><a href="crmhistorik.asp?menu=crm&ketype=e&func=hist&id=<%=oRec("kid")%>&selpkt=hist"><img src="../ill/hist_lille.gif" width="12" height="12" alt="Se historik" border="0"></a>&nbsp;<a href="javascript:NewWin_large('../inc/regular/kundelogview.asp?useKid=<%=oRec("kid")%>')" target="_self"><img src="../ill/joblog_small.gif" width="12" height="12" alt="Se kundelog" border="0"></a>&nbsp;<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=red&id=<%=oRec("kid")%>&crmaktion=<%=oRec("crmaktionid")%>&selpkt=kal&kalemner=<%=emner%>&kalstatus=<%=status%>&kalviskunde=<%=id%>&kaldato=<%=dato%>&kalmedarb=<%=medarb%>&showinwin=j')" class="vmenuglobal">
									<img src="../ill/blank.gif" width="19" height="1" alt="" border="0"><img src="../ill/blyant.gif" width="12" height="11" alt="" border="0"></a>&nbsp;<a href="crmhistorik.asp?menu=crm&func=slet&id=<%=id%>&crmaktion=<%=oRec("crmaktionid")%>&medarb=<%=medarb%>&selpkt=kal&serial=<%=oRec("serialnb")%>"><img src="../ill/slet_crmkal.gif" width="12" height="12" alt="" border="0"></a></td>
									</tr>
									<tr><td bgcolor="MistyRose" style="padding-top:2; padding-left:3; padding-right:3; padding-bottom:3;" colspan="2">
									<img src="../ill/<%=oRec("ikon")%>" alt="<%=oRec("kontaktform")%>" border="0">&nbsp;
									
									<%if shownumofdays = 1 then
									emne = oRec("enavn")
									else
									emne = left(oRec("enavn"),14)
									end if%>
									
									<%if cint(oRec("serialnb")) <> 0 then%>
									¤&nbsp;
									<%end if%>
									<font class=megetlillesort><u>Emne:</u><br><%=emne%>
									</td>
									</tr>
									<tr><td bgcolor="MistyRose" style="padding-top:2; padding-left:3; padding-right:3; padding-bottom:3;" colspan="2">
									<font class=megetlillesort><u>Status:</u><br><%=oRec("statusnavn")%><br>
									<%
									if len(oRec("navn")) <> 0 then
									Response.write "<u>Overskrift:</u><br>"
										if shownumofdays = 1 then
										Response.write "<b>"&oRec("navn")&"</b>" & "<br>"
										else
										Response.write oRec("navn") & "<br>"
										end if
									end if%>
									</font>
									<%if shownumofdays = 1 then%>
									<%=oRec("besk")%>
									<%else%>
										<%if len(oRec("besk")) <> 0 then%>
										<font title="<%=oRec("besk")%>" class=lillegray style="cursor:crosshair;"><u>Beskrivelse...</u></font>
										<%end if%>
									<%end if%>
								
								<%end if%>
								</td></tr></table><br>
								<%	
								foundone = 1
								oRec.movenext
								wend
								oRec.close
								
								%>
								</form>
								<%
								
								
								
								
								'**** Markerer hvis aktiviteten løber over flere timer ***
								if foundone = 0 then
								
								strSQL = "SELECT DISTINCT(crmhistorik.id) AS crmaktionid, kkundenavn, crmklokkeslet_slut"_
								&" FROM "_ 
								&" crmhistorik, aktionsrelationer "_
								&" LEFT JOIN kunder ON (kunder.kid = kundeid) "_
								&" WHERE crmdato = '"& sqlDatoThis &"' AND (crmdato = crmdato_slut OR COALESCE(crmdato_slut,0) = 0) AND crmklokkeslet < '"& klokkeslet &"' AND crmklokkeslet_slut > '"& klokkeslet &"' "& usemedarbKri &" AND kundeid <> -1 ORDER BY crmklokkeslet" 
								
								y = 0
								oRec.open strSQL, oConn, 3
								While not oRec.EOF
								if y = 0 then
								%><table border=0 cellspacing=0 cellpadding=0 width="100%">
								<%
								y = 1
								end if%>
								<tr><td bgcolor="MistyRose" valign="top" style="padding-top:4; padding-left:5; padding-right:3; background-color:mistyRose; border-left: 1px dashed darkred; border-right: 1px dashed darkred;">
								<font size=1><%
								Response.write left(oRec("kkundenavn"),12) &"<br>&nbsp;"
								%>
								</font></td></tr>
								<%
								oRec.movenext
								wend
								oRec.close
								
								if y > 0 then
								%>
								</table>
								<%
								end if
								
								end if
							%>
							</td>
							<%
					end if 'intcounter
					next '1 eller 5 dages view
				%>
				</tr>
				<%
				next 'For 7 til 21
				%>
		
		
		<tr><td colspan="6" height="30"><br>&nbsp;</td></tr>
		</table>
			
		</div>
		<%
end if
%>


<!--#include file="../inc/regular/footer_inc.asp"-->
 



