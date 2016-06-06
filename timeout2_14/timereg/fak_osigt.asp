<%
if len(request("FM_chkmed")) <> 0 then
	if request("FM_chkmed") = "1" then
	response.cookies("tvmedarb") = "j"
	else
	response.cookies("tvmedarb") = "n"
	end if
	response.cookies("tvmedarb").expires = date + 40
end if

%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	'else
	%>
	<!--include file="inc/isint_func.asp"-->
	<%
	'call erDetInt(request("FM_fakint"))
	'if isInt > 0 AND len(trim(request("FM_fakint"))) > 0 then
	%>
	<!--include file="../inc/regular/header_inc.asp"-->
	<!--include file="../inc/regular/topmenu_inc.asp"-->
	<%
	'errortype = 42
	'call showError(errortype)
	
	'isInt = 0
	else
	
	session("erOprettet") = "n"
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%if request("print") <> "j" then%>			
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
	}
	
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	 if (document.images){
			plus = new Image(200, 200);
			plus.src = "ill/plus.gif";
			minus = new Image(200, 200);
			minus.src = "ill/minus2.gif";
			}
	
			function expand (de){
				if (document.all("D" + de)){
					if (document.all("D" + de).style.display == "none"){
					document.all("D" + de).style.display = "";
					document.images["plus" + de].src = minus.src;
				}else{
					document.all("D" + de).style.display = "none";
					document.images["plus" + de].src = plus.src;
				}
			}
		}
	
	
	//function hidefaknr(){
	//document.all["FM_fakint"].value = ""
	//}
	
	//function hidedatointerval(){
	//document.all["FM_usedatointerval"].checked = ""
	//}
	//-->
	</script>
	<%
	'** Function til at fjerne decimaler
	Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = SQLBless(instr(x, "."))
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function
	
	public notfakbeloeb
	public intjobTpris
	public fastpris
	
	public function ikkafaktimerPaaJob(startDato, lastFakdato)
	
	'findesderfakpaaJob = findes der fakturaer på job 
	'fo = findes deer fakturaer på job i periode
	
	'Response.write startDato &"#"& lastFakdato & " fo " & foundone &" f" & f & "findesderfakpaaJob "& findesderfakpaaJob
	if cdate(startDato) > cdate(lastFakdato) then
	'Response.Write "-A-"
	stdatoKri = startDato
	else
	    'if foundone = 1 then 'fo
	    if cdate(lastFakdato) => cdate(startDato) then
	    'Response.Write "-- her - "& cdate(lastFakdato) &" >= "& cdate(startDato) 
	    'Response.Write "-B-"
	    stdatoKri_temp = dateadd("d", 1, lastFakdato)
	    stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
	    else
	    stdatoKri = lastFakdato 'startDato
	    'Response.Write "-C-"
	    end if
	end if
	
	'Response.Write "stDatoKri: " & stDatoKri
	
	
				'* Finder ikke fakturerede timer på det pågældende job **
				strSQL3 = "SELECT timer.Tid, timer.timer, timer.tdato, timer.tjobnr, timer.timepris, "_
				&" job.jobnr, job.id, tfaktim, timer.fastpris, jobTpris, budgettimer, "_
				&" timer.TAktivitetId, timer.tidspunkt FROM job "_
				&" LEFT JOIN timer ON (timer.tjobnr = job.jobnr "_
				&" AND tfaktim = 1 AND timer.tdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"') "_
				&" WHERE job.id = "& intArrJobnr(a) 
				
				'Response.write strSQL3
				oRec3.open strSQL3, oConn, 3
				
				while not oRec3.EOF
				
				diff = DateDiff("d", oRec3("tdato"), stdatoKri)
				'Response.Write diff
				'Response.flush
				if diff <= 0 then
							notfaktim = notfaktim + oRec3("timer")
							if oRec3("fastpris") = "1" then
							fastpris = "ya"
							notfakbeloeb = notfakbeloeb + (oRec3("timer") * (oRec3("jobTpris")/oRec3("budgettimer")))
							intjobTpris = oRec3("jobTpris")
							else
							intjobTpris = 0
							fastpris = "no"
							notfakbeloeb = notfakbeloeb + (oRec3("timer") * oRec3("timepris"))
							fakbar = "y"
							end if
				end if
					
				oRec3.movenext
				wend
				oRec3.close
				'****
				
				'lastFakdato = "01/01/2002"
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	fakbart = 0
	a = 0
	Redim intArrJobnr(a)
	strJobnrPrinted = "0#"
	'linket = request("linket")
	lastJobnr = 0
	fakTotalprJob = 0
	totalnotfakbeloeb = 0
	
	'strYear = request("year")
	'strReqAar = strYear 
	'strReqMrd = request("mrd")
	
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	if len(request("use_nr_or_interval")) <> 0 OR len(intJobnr) = 0 then
	use_nr_or_interval = 1
	useniCHK = "CHECKED"
	else
	use_nr_or_interval = 0 
	useniCHK = ""
	end if
	
	useFaknr = trim(request("FM_fakint"))
	lenuseFaknr = len(trim(request("FM_fakint")))
	barwhere = instr(request("FM_fakint"), "-")
	
	if cint(barwhere) > 0 then
	usebarwhere = (barwhere-1) 
	faknrfirstkri = left(useFaknr, usebarwhere)
	faknrlastkri = right(useFaknr, (lenuseFaknr-barwhere))
	end if
	
	
	'** Hvis der ikke er valgt et faknummer/ faknummerinterval, bruges dato interval ***
	if len(useFaknr) <> 0 then
	        
	    if cint(barwhere) = 0 then
	    usefakinterval = "fakturaer.faknr = '"& useFaknr &"'"
	    else
	    
	    %>
	    <!--#include file="inc/isint_func.asp"-->
		<%
			call erDetInt(faknrfirstkri)
			call erDetInt(faknrlastkri)  
			if isInt > 0 then
			usefakinterval = "fakturaer.faknr = '"& useFaknr &"'"
			isInt = 0
			else
	        usefakinterval = "fakturaer.faknr >= "& faknrfirstkri &" AND fakturaer.faknr <= "& faknrlastkri &""
	        end if
	        
	    end if 
	
	startDato = "2001/1/1"
	slutDato = "2014/1/1"
	usefakinterval = usefakinterval
	usefaknr = "j"
	else
		'** Bruger altid datointerval (Periode afg.)
		'if cint(request("FM_usedatointerval")) = 1 then
			
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		startDato = "20"&right(request("FM_start_aar"),2)&"/"& request("FM_start_mrd") &"/"& dagparset
		
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutDato = "20"&right(request("FM_slut_aar"),2)&"/"& request("FM_slut_mrd") &"/"& dagparset
		
		'else
		'startDato = "2001/1/1"
		'slutDato = "2014/1/1"
		'end if
	'usefakinterval = ""
	usefakinterval = "fakturaer.fakdato BETWEEN '"& startDato &"' AND '"& slutDato &"'"
    usefaknr = "j" 'n
	end if
	
	
	
	
	selmedarb = request("selmedarb")
	selaktid = request("selaktid")
	
	
	
	
	'************************************************************
	'**** Valgte jobnr kriterie 							*****
	'************************************************************
	if len(intJobnr) = 0 OR cint(use_nr_or_interval) = 1 then
		strSQL = "SELECT id, faknr FROM job, fakturaer WHERE "& usefakinterval &" AND job.id = fakturaer.jobid ORDER BY jobknr, jobnavn"
	else
		if intJobnr = "0" OR len(intJobnr) = "0" then
		intJobnrKri = "jobnr = 0"
		else
			Dim intJobnrKriValues 
			Dim i
			intJobnrKriValues = Split(intJobnr, ", ")
			For i = 0 to Ubound(intJobnrKriValues)
			intJobnrKri = intJobnrKri & "jobnr = " & intJobnrKriValues(i) &" OR "
			next
			intJobnrKriCount = len(intJobnrKri)
			trimintJobnrKri = left(intJobnrKri, (intJobnrKriCount-3))
			intJobnrKri = trimintJobnrKri 
		end if
		
		strSQL = "SELECT id FROM job WHERE "& intJobnrKri &" ORDER BY jobknr, jobnavn"
	end if
	
	
	
	'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
	'Response.Write strSQL
	'Response.Write "<br>jobnr valgt: " & intJobnr		
	
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			'*** Opbygning af jobnr til fakturaer *****
			areJobnrP = instr(strJobnrPrinted, ", " &oRec("id")&"#")
			if areJobnrP = "0" then 
			Redim preserve intArrJobnr(a)
			intArrJobnr(a) = oRec("id")
			a = a + 1
			end if
			
			strJobnrPrinted = strJobnrPrinted & ", " & oRec("id") &"#" 
			
			oRec.movenext
			wend
		oRec.close
	' *** '	
		
	
	kriBetalt = " "
	
	thisfile = "fak_osigt"
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<%
	if request("print") <> "j" then
	pleft = 190
	ptop = 55
	else
	pleft = 20
	ptop = 10
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><img src="../ill/header_fak.gif" alt="" border="0" width="279" height="34">
	<hr align="left" width="760" size="1" color="#000000" noshade>
	Oversigt over udsendte fakturaer, rykkere og kreditnotaer på de valgte <b>fakturerbare job</b>.<br>
	Der kan oprettes nye fakturer og de eksisterende kan redigeres eller slettes.<br>
	Kun de <b>jobansvarlige</b> og administratorer kan oprette nye fakturaer, alle med adgang til <br>
	bogføring og fakturering kan redigere og oprette rykkere eller kreditnotaer.<br>
	
	
	<br>Fakturaer og timer omfatter time-registreringer for <b>samtlige medarbejdere</b>.<br>
	Farvekoder: 
	<table cellpadding=2 border=0><tr><td bgcolor="LightPink" width=50 style="padding:3px;">Kreditnota</td>
	<td bgcolor="#FFFFE1" width=50 style="padding:3px;">Rykker</td>
	<td bgcolor="#ffffff" width=50 style="padding:3px;">Faktura</td>
	<td bgcolor="#cccccc" width=150 style="padding:3px;">Inaktiv (Oprindelig faktura)</td></tr></table>
	<br>
<%
	lastFakID = 0
	notfaktim = 0
	notfakbeloeb = 0
	fakTotal = 0
	
	yesterdayDate = date 
	dagsDatomySQLformat = convertDateYMD(yesterdayDate)



'*********************************************************
'** Henter fakturaer på de valgte job 					**
'*********************************************************
foundone = 0
di = 0
venter = 0

eksportFid = 0

for a = 0 to a - 1


f = "no"
divwritten = "n"
antalFak = 0
	
	'*** Vælger hvilket kriterie der skal bruges ***
	if usefaknr = "j" then
	usefakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND job.id = fakturaer.jobid AND Kid = jobknr AND " & usefakinterval &""
	else
	usefakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND job.id = fakturaer.jobid AND Kid = jobknr "& kriBetalt &" AND fakturaer.fakdato BETWEEN '"& startDato &"' AND '"& slutDato &"'"
	end if
	
	'*** Fakturaer SQL ***
	strSQL = "SELECT fakturaer.Fid, fakturaer.faknr, fakturaer.dato AS fdato, "_
	&" fakturaer.fakdato, fakturaer.jobid, fakturaer.timer, fakturaer.beloeb, "_
	&" fakturaer.kommentar, fakturaer.faktype, fakturaer.parentfak, jobnavn, "_
	&" jobnr, id, jobknr, fakturerbart, Kkundenavn, Kkundenr, tidspunkt AS faktidspkt, "_
	&" betalt FROM fakturaer, job, kunder WHERE "& usefakkri &" ORDER BY fakDato DESC, tidspunkt DESC"
	
	
	'Response.Write strSQL
	'Response.Write "<hr>"
		
		'*** Finder sidste fak uafhængig af valgt datointerval. ***
		findesderfakpaaJob = 0
		strSQL3 = "SELECT fakturaer.fakdato AS lastFakEver FROM fakturaer WHERE fakturaer.jobid=" & intArrJobnr(a) &" ORDER BY fakDato DESC, tidspunkt DESC"
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then
		lastFakdato = oRec3("lastFakEver")
		findesderfakpaaJob = 1
		end if
		oRec3.close
		
		'Response.Write lastFakdato
		
		if usefaknr = "j" then
		useThisFakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND " & usefakinterval &""
		else
		useThisFakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND fakturaer.fakdato BETWEEN '"& startDato &"' AND '"& slutDato &"'"
		end if
		
		strSQL3 = "SELECT count(fakturaer.fid) AS antalFak FROM fakturaer WHERE "& useThisFakkri
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then
		antalFak = oRec3("antalFak")
		end if
		oRec3.close
		
		if len(lastFakdato) <> 0 then
		lastFakdato = lastFakdato
		else
		lastFakdato = "1/1/2002"
		end if
		
		oRec.open strSQL, oConn, 0, 1
		x = 0
		while not oRec.EOF
		foundone = 1 
		eksportFid = eksportFid &","& oRec("fid")
					
		
		'*** Finder vente timer på faktura. ** 
		strSQL2 = "SELECT sum(venter) AS venter FROM fak_med_spec WHERE fakid = "& oRec("fid") 

		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then 
			venter = venter + oRec2("venter")
		end if
		oRec2.close
		
		
		'** Skriver Kunde og jobnavn *** 
		if intArrJobnr(a) <> lastJobnr then
		
		
			if x = 0 then
				lastFakdato = convertDateYMD(oRec("fakdato")) 
				strfaktidspkt = formatdatetime(oRec("faktidspkt"), 3)
				x = x + 1
			end if	
			
			call ikkafaktimerPaaJob(startDato, lastFakdato)
			
		
		
				'*** Hvis vis alle med registrerede timer siden sidste faktura er valgt *****	
				if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
				
					'*** Skal div være slået ud?? ***
					disp = "none"
					useplusmnusgif = "plus"
					di = 1
					
					%>
					<table border="0" cellpadding="0" cellspacing="0" width="500">
					<tr>
						<td height=1 bgcolor="#003399" colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td height="30" style="padding-left : 4px; padding-right : 1px;"><a href="javascript:expand('<%=oRec("jobnr")%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec("jobnr")%>"></a>&nbsp;
						<font class='storhvid'><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</font>&nbsp;<a href="javascript:expand('<%=oRec("jobnr")%>');" class='alt'>&nbsp;&nbsp;<b><%=oRec("jobnavn")%></b> (<%=oRec("jobnr")%>)</a>&nbsp;&nbsp;&nbsp;
						<b><%=antalFak%></b> faktura(er)</td>
						<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<tr>
						<td height=1 bgcolor="#003399" colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					</table>
					<%
					divwritten = "y"
					%>
					<DIV ID="D<%=oRec("jobnr")%>" Style="position: relative; display: <%=disp%>;">
					<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#ffffff">
					<tr bgcolor="#66CC33">
						<td width=10><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td width="50" valign="bottom" style="padding-left:5px; " class='alt'><b>Fakturanr.</b></td>
						<td width="50" valign="bottom" style="padding-left:5px; " class='alt'><b>Status</b></td>
						<td width="70"  valign="bottom" align="right" style="padding-right:3px;" class='alt'><b>Fak. dato</b></td>
						<td width="70" valign="bottom" class='alt' style="padding-left:10px;"><b>Type</b>&nbsp;(Opr.)</td>
						<td width="170" colspan=2 valign="bottom" align="right" style="padding-right:10px;" class='alt'><b>Beløb</b></td>
						<td width=10 align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<%
					recJobnr_ya = oRec("jobnr")
					fakbart = oRec("fakturerbart") 'Er jobbet fakturerbart?
					f = "ya" 'Er der udskrevet fakturaer på jobbet?
				end if
		end if%>
		
		
		<%
		'************************************************************************************
		'*** Skal fakturaerne på det valgte job vises? lever de op til filter kriterier?? ****
		'*** Hvis vis alle med registrerede timer siden sidste faktura er valgt. 		*****
		'************************************************************************************
		
		if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
		
			'*** Skriver Fakturaer på det pågældende job ***
			if oRec("Fid") <> lastFakID then
		
				%>
				<tr bgcolor="#003399">
					<td height=1 colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<%
					strSQL2 = "SELECT Fid, parentfak FROM fakturaer WHERE parentfak = " &oRec("Fid")&" AND faktype = 2"
					oRec2.open strSQL2, oConn, 3 
					hasfak = 0
					if not oRec2.EOF then
					hasfak = 1
					end if
					oRec2.close
				
				
				Select case oRec("faktype")
					case 0
						if hasfak = 1 then
						bgcol = "#cccccc"
						else
						bgcol = "#ffffff"
						end if
					case 1
					bgcol = "LightPink"
					case 2
					bgcol = "#ffffe1"
					end select
				%>
				<tr bgcolor="<%=bgcol%>">
					<td style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"></td>
					<td valign=top style="padding-left:5;">
					
					
						<%if request("print") <> "j" then
							
							'*** Rediger via gammel fak fil. (gamle fakturaer!)
							thisfakdato = oRec("fdato")
							fakdag = left(trim(thisfakdato), 2)
							fakaar = right(trim(thisfakdato), 4)
							punkt = instr(trim(thisfakdato), ".")
							
							
								if punkt <> 0 then
								'Response.write thisfakdato
								fakmd_temp = mid(trim(thisfakdato), punkt - 3, punkt - 1)
								fakmd = left(fakmd_temp, 3)
								
								select case fakmd
								case "Jan"
								fakmdnr = 1
								case "Feb"
								fakmdnr = 2
								case "Mar"
								fakmdnr = 3
								case "Apr"
								fakmdnr = 4
								case "Maj"
								fakmdnr = 5
								case "Jun"
								fakmdnr = 6
								case "Jul"
								fakmdnr = 7
								case "Aug"
								fakmdnr = 8
								case "Sep"
								fakmdnr = 9
								case "Okt"
								fakmdnr = 10
								case "Nov"
								fakmdnr = 11
								case "Dec"
								fakmdnr = 12
								case else
								fakmdnr = month(now)
								end select 
								
								fakdatothis = fakdag&"/"&fakmdnr&"/"&fakaar
								
								else
								fakdatothis = oRec("fakdato")
								end if
								
								
								'*** Gamle fakturaer åbnes i gammle fil. ***
								if cdate(fakdatothis) < cdate("18/10/2004") then
								fakfil = "fak_org.asp"
								else
								fakfil = "fak.asp"
								end if
					
						if cdate(fakdatothis) > cdate("1/1/2005") then 'oRec("betalt") <> 0 AND%>
								
								<a href="fak_godkendt.asp?menu=stat_fak&id=<%=oRec("Fid")%>&histback=1&jobid=<%=oRec("jobid")%>" class=vmenu>
								<%=oRec("faknr")%></a>
						<%
						else	
						%>		
								
								<a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=<%=oRec("faktype")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=vmenu>
								<%=oRec("faknr")%></a>
								
						<%end if%>
					
					<%else%>
					<%=oRec("faknr")%>
					<%end if%>
					
					
					&nbsp;&nbsp;</td>
					
					<td valign=top style="padding-left:5px;"><%if oRec("betalt") <> 0 then%>
					<font class=lillesort>Godkendt</font>
					<%else%>
					<font class=lillegray>Kladde</font>
					<%end if%>
					</td>
					
					<td valign=top align=right style="padding-right:3;"><b><%=formatdatetime(oRec("fakdato"), 0)%></b>&nbsp;</td>
					<td valign=top style="padding-right:3; padding-left:10;"><%
					if oRec("faktype") <> 0 then
							strSQL3 = "SELECT faknr FROM fakturaer WHERE Fid = " & oRec("parentfak")
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							oprFaknr = oRec3("faknr")
							end if
							oRec3.close
					end if
					
					Select case oRec("faktype")
					case 0
					strFaktypeShow = "<b>Faktura</b>"
					case 1
					strFaktypeShow = "<b>Kredit</b>&nbsp;("&oprFaknr&")"
					case 2
					strFaktypeShow = "<b>Rykker</b>&nbsp;("&oprFaknr&")"
					end select
					Response.write strFaktypeShow 
					%>
					</td>
					<td align="right" colspan=2 valign=top style="padding-right:3; padding-left:3;"><%=formatnumber(oRec("beloeb"), 2)%> dkr</td>
					<td align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"></td>
				</tr>
				<tr bgcolor="<%=bgcol%>">
					<td colspan="8" height="25" valign="top" style="padding-right:3; padding-left:12; border-left:1px #003399 solid; border-right:1px #003399 solid;">
					
					
					<%
					if oRec("betalt") <> 0 then
					
						if (cdate(fakdatothis) > cdate("18/10/2004")) then%>
						|&nbsp;<a href="fak_godkendt.asp?menu=stat_fak&id=<%=oRec("Fid")%>&histback=1&jobid=<%=oRec("jobid")%>" class=rmenu>Gennemse</a>&nbsp;|
							<%if hasfak = 0 AND oRec("faktype") <> 1 then%>
							&nbsp;<a href="fak.asp?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=2&faknr=<%=oRec("faknr")%>&rykkreopr=j&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_gr>Opret Rykker</a>&nbsp;|&nbsp;
							<a href="fak.asp?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=1&faknr=<%=oRec("faknr")%>&rykkreopr=j&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_g>Opret Kreditnota</a>&nbsp;|
							
							<%if level <= 2 AND (cdate(fakdatothis) > cdate("3/5/2006")) then%>
							&nbsp;<a href="fak.asp?menu=stat_fak&func=fortryd&id=<%=oRec("Fid")%>&jobid=<%=oRec("jobid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=1&faknr=<%=oRec("faknr")%>&rykkreopr=j&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_r>Fortryd godkend og posteringer</a> &nbsp;|
							<%end if%>
							
							<%end if%>
						<%else%>
						|&nbsp;<a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=<%=oRec("faktype")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=rmenu>Gennemse</a>&nbsp;|
						<%end if
						
					else
					%>
					|&nbsp;<a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=<%=oRec("faktype")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=rmenu>Rediger</a>&nbsp;|
						<%if oRec("betalt") = 0 then%>
						&nbsp;<a href="<%=fakfil%>?menu=stat_fak&func=slet&id=<%=oRec("Fid")%>&FM_job=<%=oRec("jobnr")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_r>Slet</a>&nbsp;|
						<%end if%>
					<%end if%>
					</td>
				</tr>
				<%
				Select case oRec("faktype") 
				case 0
					if hasfak = 0 then
					fakTotalprJob = fakTotalprJob + oRec("beloeb")
					else
					fakTotalprJob = fakTotalprJob
					end if
				case 1
				fakTotalprJob = fakTotalprJob - oRec("beloeb")
				case 2
				fakTotalprJob = fakTotalprJob + oRec("beloeb")
				end select
				
				
				lastJobnr = intArrJobnr(a)
				lastFakID = oRec("Fid")
			end if
		end if
		oRec.movenext
		wend
		oRec.close
		%>
		<!--
		<tr bgcolor="#003399">
			<td height=1 colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>-->
		<%
		
		'fakTotal = fakTotal + fakTotalprJob 
		
		
		
		
		
		
		'***************************************************************************************
		'**** Udkskriver total omsætning på job og viser knap til oprettelse af ny faktura. ****
		'***************************************************************************************
		
		if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
		
				if fakbart = "1" then
				%>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="35" alt="" border="0"></td>
					<td colspan="2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					<td colspan="4" valign="bottom" colspan="2" align="right" style="padding-right : 5px;"><u>Balance:&nbsp;&nbsp;<b><%=formatnumber(fakTotalprJob,2)%> dkr.</b></u></td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="35" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
					<td valign="bottom">&nbsp;</td>
					<td valign="bottom" align="right" class=lille>Timer</td>
					<td valign="bottom" align="right" class=lille>Venter</td>
					<td valign="bottom" align="right" class=lille colspan="3">Omsætning (Excl. vente timer.)&nbsp;</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					<td valign="bottom" class=lille style="padding-left:5; padding-right:5;">Fakturerbare timer til fakturering:&nbsp;&nbsp;</td>
					<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;"><%=SQLBless(notfaktim)%></td>
					<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;">&nbsp;(<%=venter%>)</td>
					<td valign="bottom" colspan="3" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; padding-left : 4px; padding-right : 3px;">
					<%
					Response.write formatnumber(notfakbeloeb, 2)
					totalnotfakbeloeb = totalnotfakbeloeb + notfakbeloeb
					%> dkr</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
					<td colspan="6" height="30" align="right"><font size=1 color="#708090">
						<%
						'*** Tjekker rettigehder eller om man er jobanssvarlig ***
						strSQL4 = "SELECT id, jobans1, jobans2 FROM job WHERE id = "& intArrJobnr(a)
						oRec4.open strSQL4, oConn, 3 
						if not oRec4.EOF then 
						intJobans1 = oRec4("jobans1")
						intJobans2 = oRec4("jobans2")
					 	end if 
						oRec4.close 
						
						editok = 0
						
						if level = 1 then
						editok = 1
						else
								if cint(session("mid")) = cint(intJobans1) OR cint(session("mid")) = cint(intJobans2) OR cint(intJobans1) = 0 AND cint(intJobans2) = 0 then
								editok = 1
								end if
						end if
						
						if editok = 1 then
							if request("print") <> "j" then
							%>
							<br><a href="fak.asp?menu=stat_fak&ttf=<%=notfaktim%>&ktf=<%=kommaFunc(notfakbeloeb)%>&jobid=<%=intArrJobnr(a)%>&jobnr=<%=recJobnr_ya%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=0&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">&nbsp;<img src="../ill/nyfak.gif" width="114" height="28" alt="" border="0"></a>&nbsp;&nbsp;
							<%
							end if
						else
						%>
						Du har ikke rettigheder til at oprette fakturaer på dette job.
						<%
						end if
						%>
					</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				</tr>
				<tr><td colspan="8" height="10" bgcolor="#5582D2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
				<%end if
				
				notfaktim = 0
				notfakbeloeb = 0
		end if
		
		
		'**********************************************************************************
		'** Hvis der er ikke oprettet fakturaer på det pågældende job.					  *
		'** Og der er valgt en kørsel hvor der ikke er brugt en af følgende kritertier:   *
		'** 1) Ikke faktura interval 													  *
		'** 3) Status skal være = 'vis alle'											  *
		'**********************************************************************************
		
		'AND cint(request("FM_usedatointerval")) <> 1 
		if f = "no" AND (len(useFaknr) <> 0 AND cint(intbetalt) = 0) then
			strSQL2 = "SELECT id, jobnavn, jobnr, fakturerbart, Kkundenavn, Kkundenr FROM job, kunder WHERE id =" & intArrJobnr(a) &" AND Kid = jobknr" 
			oRec2.open strSQL2, oConn, 0, 1
				if not oRec2.EOF then
					if oRec2("fakturerbart") = "1" then
					%>
					<table border="0" cellpadding="0" cellspacing="0" width="500">
					<tr>
						<td height=1 bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td colspan="4" height="30" style="padding-left : 4px; padding-right : 1px;">
						<a href="javascript:expand('<%=oRec2("jobnr")%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec2("jobnr")%>"></a>&nbsp;
						<font class='storhvid'><%=oRec2("Kkundenavn")%>&nbsp;(<%=oRec2("Kkundenr")%>)</font>&nbsp;<a href="javascript:expand('<%=oRec2("jobnr")%>');" class='alt'>&nbsp;&nbsp;<b><%=oRec2("jobnavn")%></b> (<%=oRec2("jobnr")%>)</a>&nbsp;&nbsp;&nbsp;
						<b><%=antalFak%></b> faktura(er)</td>
						<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<tr>
						<td height=1 bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					</table>
					<%
					divwritten = "y"
					foundone = 1
					%>
					<DIV ID="D<%=oRec2("jobnr")%>" Style="position: relative; display: none;">
					<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#ffffff">
					<%
					recJobnr_no = oRec2("jobnr")
					fakbart = 1
					end if
				end if
			strfaktidspkt = "00:00:01 AM"
			
			jobnrwww = oRec2("jobnr")
			oRec2.close
			
			
			call ikkafaktimerPaaJob(startDato, "2001-01-01") 'lastFakdato
			
			
			'** 03-02-2006
			'** Da der jo ikke findes fakturaer på dette job og der derfor ikke er en lastFakdato
			'** og der opstod en fejl hvis jobbet før havde en lastFakdato der var senere end startDato.
			
			if fakbart = "1" then%>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				<td colspan="2" width=200 style="padding-left:5px;">Der er ikke oprettet fakturaer på dette job i den angivne periode.</td>
				<td valign="bottom" align="right"><font size=1>Timer</font></td>
				<td valign="bottom" align="right"><font size=1>Venter</font></td>
				<td valign="bottom" align="right" style="padding-right:4px;"><font size=1>Omsætning (Excl. vente timer.)</font></td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
				<td colspan="2" valign="bottom" align="right" style="padding-right:4px;">Endnu ikke fakturede, fakturerbare timer:&nbsp;&nbsp;</td>
				<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;"><%=SQLBless(notfaktim)%></td>
				<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;">(<%=venter%>)&nbsp;</td>
				<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; padding-left : 4px; padding-right : 4px;">
				<%
				Response.write formatnumber(notfakbeloeb,2) 
				totalnotfakbeloeb = totalnotfakbeloeb + notfakbeloeb
				%> dkr</td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				<td colspan="5" height="30" align="right">
				<%
				'*** Tjekker rettigehder eller om man er jobanssvarlig ***
						strSQL4 = "SELECT id, jobans1, jobans2 FROM job WHERE id = "& intArrJobnr(a)
						oRec4.open strSQL4, oConn, 3 
						if not oRec4.EOF then 
						intJobans1 = oRec4("jobans1")
						intJobans2 = oRec4("jobans2")
					 	end if 
						oRec4.close 
						
						editok = 0
						
						if level = 1 then
						editok = 1
						else
								if cint(session("mid")) = cint(intJobans1) OR cint(session("mid")) = cint(intJobans2) OR cint(intJobans1) = 0 AND cint(intJobans2) = 0 then
								editok = 1
								end if
						end if
				
				if editok = 1 then
					if request("print") <> "j" then%>		
					<br><a href="fak.asp?menu=stat_fak&ttf=<%=notfaktim%>&ktf=<%=kommaFunc(notfakbeloeb)%>&jobid=<%=intArrJobnr(a)%>&jobnr=<%=recJobnr_no%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">&nbsp;<img src="../ill/nyfak.gif" width="114" height="28" alt="" border="0"></a>&nbsp;&nbsp;
					<%
					end if
				else
					%>
					Du har ikke rettigheder til at oprette fakturaer på dette job.&nbsp;&nbsp;
					<%
				end if%></td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
			</tr>
			<tr><td colspan="7" height="10" bgcolor="#5582D2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
			<%end if
			
			notfaktim = 0
			notfakbeloeb = 0
			lastJobnr = intArrJobnr(a)
		
		end if

			fakbart = 0
			fakTotalprJob = 0
			
			'** afsluttter inter tabel og div hvis de er påbegyndt ***
			if divwritten = "y" then ' OR cint(intbetalt) = 0 then%>
			</table>
			</div><img src="../ill/blank.gif" width="1" height="4" alt="" border="0"><br>
			<%
			end if

			venter = 0
next
%><br><br><br>
<%if foundone <> 0 then%>
<br>
<a href="fakturaer_eksport.asp?fakids=<%=eksportFid%>" class=vmenu target="_blank">Eksporter ovenstående liste af fakturaer.&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br>
<a href="javascript:NewWin_help('help.asp?menu=stat_fak');" target="_self" class='helplink'>+ Hjælp til eksport fil. +</a>
<br><br><br>&nbsp;
<%end if%>
</div>


<%if cint(foundone) = 0 then%>
<div style="position:absolute; left:250; top:230; z-index:10000; background-color:#ffffe1; border:1px darkred solid; padding:5;">
<table cellspacing="0" cellpadding="0" border="0" width=330>
<tr><td valign="top"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"></td>
<td style="padding-left:5;"><br>Der er ikke fundet nogen <b>Fakturaer</b> eller<b> job</b> der matcher de valgte kriterier.<br>
<br>Dette skyldes en af to ting:<br>
<b>1)</b>&nbsp;Der er ikke valgt nogen job.<br>
<b>2)</b>&nbsp;Der er angivet et faktura nummer der ikke findes.<br><br>&nbsp;</td></tr>
</table>
</div>
<%end if%>




<div style="position:absolute; left:720; top:130; z-index:10000;">
<table cellspacing="0" cellpadding="0" border="0" width=230>
<form action="fak_osigt.asp?menu=stat_fak" method="post" name="statselector" id="statselector">
<input type="hidden" name="FM_job" value="<%=FM_job%>">
	<tr bgcolor="#5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="218" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
			<td colspan=2 class="alt">Kør ny oversigt</td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td  style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td colspan="2" valign="top">
	<input type="hidden" name="FM_usedatointerval" id="FM_usedatointerval" value="1">
	<br><b>A) Periodeafgrænsning.</b><br>
	Finder fakturaer og timer til fakturering i den valgte periode.<br> 
	<!--#include file="inc/weekselector_s.asp"-->
	
	
	<!--<br />
    <input id="use_nr_or_interval" name="use_nr_or_interval" value=0 type="radio" /> Brug valgte job og periode kriterie. 
    <br /><input id="use_nr_or_interval" name="use_nr_or_interval" value=1 type="radio" /> Brug Fakturanr kriterie. 
    <br /><input id="use_nr_or_interval" name="use_nr_or_interval" value=2 type="radio" /> Brug Periode kriterie. 
	-->
	
	<input type="hidden" name="FM_vis_betalte" value="0">
	<br><br><br>
	<b>B)</b> Angiv et <b>faktura nr</b> / <b>interval</b> her, hvis du ønsker at finde bestemte fakturaer.<br>
	<input type="text" name="FM_fakint" size="12" value="<%=trim(request("FM_fakint"))%>">&nbsp;<font class=megetlillesort>f.eks. 1045 ell. 1045-1067</font><br>
	
	<br />
    <input id="use_nr_or_interval" name="use_nr_or_interval" <%=useniCHK%> value=1 type="checkbox" />
    <b>Ignorer valgte kontakter og job.</b><br />
    Hvis der er angivet et faktura nr /interval (B) bruges dette, ellers bruges periode (A).<br><br>
	
	<font class=megetlillesort>(Den i pkt. A valgte periode-afgrænsning bruges stadigvæk til visning af timer til fakturering, samt som afgrænsning ved oprettelse af nye fakturaer.)&nbsp;</td>
	<td  style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    </tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
		<td colspan="2" align=center><br>
		<input type="image" src="../ill/statpil.gif" alt="Kør Statistik!"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom" style="padding-left:40;"><br>
		<table><tr><td>&nbsp;</td>
		<td style="padding-top:4;">&nbsp;<a href="javascript:NewWin_help('help.asp?menu=stat_fak');" target="_self" class='helplink'>+ Hjælp til fakturering. +</a></td></tr></table>
		<br><br>
		<font class=megetlillesort>NB)<br>
		Pga. ændrede egenskaber i faktura-filen, er følgende gældende:<br>
		Fakturaer oprettet efter d. 1/1-2005 kan ses direkte i deres printlayout ved at klikke på faktura nummeret.<br><br>
		Fakturaer der er oprettret før 18/10-2004 kan ikke redigeres, der kan ej heller oprettes rykkere eller kredit notaer.
		Hvis der er brug for at oprette en kredit nota på en af disse fakturaer kontakt da <u>support@outzource.dk</u>.
		</font>
		<!--
		<font class=megetlillesort><b>Note(r):</b><br> Hvis der benyttes et datointerval hvor den valgte startdato ligger før datoen for den sidst oprettede faktura på det pågældende job, bruges faktura datoen som startdato når nye fakturerbare timer skal findes. <br><br>
		Begge de valgte dage er inklusive.<br><br>
		Faktura datoen oprettes således, at alle timer til og med faktura datoen er indbefattet. Når en faktura oprettes fjernes muligheden for at indtaste timer, frem til og med faktura datoen.
		--></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
</div>

<%
end if 
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
