<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 AND request("menu") <> "no" then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	public chk
	function checked(val)
	if val = "j" then
	chk = "CHECKED"
	else
	chk = ""
	end if
	end function
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:290; top:120; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en <b>Signoff</b>. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="signoff.asp?menu=kund&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM signoff WHERE id = "& id &"")
	Response.redirect "signoff.asp?menu=kund"
	
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
				
				function SQLBless2(s2)
				dim tmp2
				tmp2 = s2
				tmp2 = replace(tmp2, ",", ".")
				SQLBless2 = tmp2
				end function
				
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		instadr = request("FM_instadr")
		postnr = request("FM_postnr")
		town = request("FM_town") 
		kpers = request("FM_kpers") 
		kpersstilling = request("FM_kpersstilling")
		kperstlf = request("FM_kperstlf")
		kpersfax = request("FM_kpersfax")
		kpersmobil = request("FM_kpersmobil")
		kpersemail = request("FM_kpersemail")
		itans = request("FM_itans")
		itanstlf = request("FM_itanstlf")
		itansfax = request("FM_itansfax")
		itansmobil = request("FM_itansmobil")
		itansemail = request("FM_itansemail")
		
		virktype =	request("FM_virktype")
		virktypeandet = request("FM_virktype_andet")
		pafgift = request("FM_pafgift")
		
		tohsalgskon = request("FM_tohsalgskon")
		tohsyskon = request("FM_tohsyskon")
		
		placering = request("FM_placering")
		placeringandet = request("FM_placering_andet")
		model = request("FM_model")
		usednew = request("FM_usednew")
		proveops = request("FM_proveops")
		
		levdag = request("FM_levdag")
		levmd = request("FM_levmd")
		levaar = request("FM_levaar")
		levdato = levaar&"/"&levmd&"/"&levdag
		
		levtid = request("FM_levtid")
		
		insdag = request("FM_insdag")
		insmd = Request("FM_insmd")
		insaar = Request("FM_insaar")
		insdato = insaar&"/"&insmd&"/"&insdag
		
		instid = request("FM_instid")
		
		kurdag = Request("FM_kurdag")
		kurmd = request("FM_kurmd")
		kuraar = request("FM_kuraar")
		kurdato = kuraar&"/"&kurmd&"/"&kurdag
		
		
		kurtid = Request("FM_kurtid")
		kurskonsul = Request("FM_kurskonsul")
		
		rj45 = Request("FM_rj45")
		ethernet = Request("FM_Ethernet")
		token = Request("FM_Token")
		IPX_SPX = Request("FM_IPX_SPX")
		TCP_IP = Request("FM_TCP_IP") 
		appleTalk = Request("FM_AppleTalk")
		
		novell3 = Request("FM_Novell3")
		novell4 = Request("FM_Novell4")
		novell5 = Request("FM_Novell5")
		novell6 = request("FM_Novell6")
		nt4 = request("FM_nt4")
		k2 = request("FM_2000")
		k03 = request("FM_2003")
		
		ipadr = Request("FM_ipadr1") & Request("FM_ipadr2") & Request("FM_ipadr3") & Request("FM_ipadr4")
		subnet = Request("FM_subnet")
		gateway = Request("FM_gateway")
		dns = Request("FM_dns")
		smtp = Request("FM_smtp")
		ftp = Request("FM_ftp")
		bnavn = Request("FM_bnavn")
		pw = Request("FM_pw")
		
		arb_win98 = Request("FM_arb_win98") 
		arb_winME = Request("FM_arb_winME")
		arb_winNT = Request("FM_arb_winNT")
		arb_win2000 = Request("FM_arb_win2000")
		arb_winXP = Request("FM_arb_winXP")
		arb_mac8 = Request("FM_arb_mac8")
		arb_mac9 = Request("FM_arb_mac9")
		arb_mac10 = Request("FM_arb_mac10") 
		arb_dos = Request("FM_arb_dos")
		
		kopi_farve = Request("FM_kopi_farve") 
		kopi_sh = Request("FM_kopi_sh")
		kopi_it = Request("FM_kopi_it")
		kopi_t1 = Request("FM_kopi_t1")
		kopi_t2 = Request("FM_kopi_t2") 
		
		'**** feltnavne
		kopi_b1_a5 = Request("FM_kopi_b1_a5")
		kopi_b1_a4 = Request("FM_kopi_b1_a4")
		kopi_b1_a3 = Request("FM_kopi_b1_a3")
		
		kopi_b2_a5 = Request("FM_kopi_b2_a5")
		kopi_b2_a4 = Request("FM_kopi_b2_a4")
		kopi_b2_a3 = Request("FM_kopi_b2_a3")
		
		kopi_b3_a5 = Request("FM_kopi_b3_a5")
		kopi_b3_a4 = Request("FM_kopi_b3_a4")
		kopi_b3_a3 = Request("FM_kopi_b3_a3")
		 '****
		
		faxdriver = Request("FM_faxdriver")
		faxnr = Request("FM_faxnr")
		faxid = Request("FM_faxid")
		faxfarve = Request("FM_faxfarve")
		fax_papribk_1 = Request("FM_fax_papribk_1")
		fax_papribk_2 = Request("FM_fax_papribk_2")
		fax_papribk_3 = Request("FM_fax_papribk_3")
		fax_papribk_4 = Request("FM_fax_papribk_4")
		fax_it = Request("FM_fax_it")
		fax_t1 = Request("FM_fax_t1")
		fax_t2 = Request("FM_fax_t2")
		fax_rap = Request("FM_fax_rap")
		
		print_plc5 = Request("FM_print_plc5")
		print_plc6 = request("FM_print_plc6") 
		print_psl2 = Request("FM_print_psl2")
		print_psl3 = Request("FM_print_psl3")
		print_farve = Request("FM_print_farve")
		print_sh = Request("FM_print_sh")
		print_it = Request("FM_print_it")
		print_t1 = Request("FM_print_t1")
		print_t2 = request("FM_print_t2")
		print_admin = Request("FM_print_admin")
		
		scan_email = Request("FM_scan_email") 
		scan_efile = Request("FM_scan_e-file")
		scan_ftp = request("FM_scan_ftp")
		scan_esame = request("FM_scan_e-same")
		scan_twain = Request("FM_scan_twain")
		scan_filedownload = Request("FM_scan_filedownload")
		scan_topacc = Request("FM_scan_topacc")
		scan_agfa = Request("FM_scan_agfa")
		
		bemaerk = Request("FM_bemaerk")
		dag = Request("FM_dag")	
		md = request("FM_md")
		aar = request("FM_aar")
		signoffdato = aar&"/"&md&"/"&dag
		underskrift = request("FM_underskrift")
		
		pandet = request("FM_pandet")
		'****'
				
				
				
				
				
				if func = "dbopr" then
				
			    strSQL = ("INSERT INTO signoff (navn, editor, dato, "_
				&" adresse, postnr, town, kpersnavn, kpersstilling, kperstlf, kpersfax, kpersmobil, "_
				&" kpersemail, itansnavn, itanstlf, itansfax, itansmobil, itansemail, virktype, "_
				&" virktypeandet, pafgift, tohsalgskon, tohsyskon, placering, placeringandet, "_
				&" model, usednew, proveops, levdato, levtid, insdato, instid, kurdato, kurtid, "_ 
				&" kurskonsul, rj45, ethernet, token, IPX_SPX, TCP_IP, appleTalk, novell3, novell4, "_ 
				&" novell5, novell6, nt4, k2, k03, ipadr, subnet, gateway, dns, smtp, ftp, bnavn, pw, "_ 
				&" arb_win98, arb_winME, arb_winNT, arb_win2000, arb_winXP, arb_mac8, arb_mac9, arb_mac10, "_
		        &" arb_dos, kopi_farve, kopi_sh, kopi_it, kopi_t1, kopi_t2, kopi_b1_a5, kopi_b1_a4, "_
				&" kopi_b1_a3, kopi_b2_a5, kopi_b2_a4, kopi_b2_a3, kopi_b3_a5, kopi_b3_a4, kopi_b3_a3, "_
		        &" faxdriver, faxnr, faxid, faxfarve, fax_papribk_1, fax_papribk_2, fax_papribk_3, "_
				&" fax_papribk_4, fax_it, fax_t1, fax_t2, fax_rap, print_plc5, print_plc6, print_psl2, "_
				&" print_psl3, print_farve, print_sh, print_it, print_t1, print_t2, print_admin, "_
				&" scan_email, scan_efile, scan_ftp, scan_esame, scan_twain, scan_filedownload, "_
				&" scan_topacc, scan_agfa, bemaerk, signoffdato, underskrift, pandet) VALUES ("_
				&" '"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
				&" '"&instadr&"', '"& postnr &"', '"& town &"', '"& kpers &"', '"& kpersstilling &"', '"& kperstlf&"', '"& kpersfax&"', '"& kpersmobil&"', "_
				&" '"&kpersemail&"', '"& itans &"', '"& itanstlf&"', '"& itansfax&"', '"& itansmobil&"', '"& itansemail&"', '"& virktype&"', "_
				&" '"&virktypeandet&"', '"& pafgift&"', '"& tohsalgskon&"', '"& tohsyskon&"', '"& placering&"', '"& placeringandet&"', "_
				&" '"&model&"', '"& usednew&"', '"& proveops&"', '"& levdato&"', '"& levtid&"', '"& insdato&"', '"& instid&"', '"& kurdato&"', '"& kurtid&"', "_ 
				&" '"&kurskonsul&"', '"& rj45&"', '"& ethernet&"', '"& token&"', '"& IPX_SPX&"', '"& TCP_IP&"', '"& appleTalk&"', '"& novell3&"', '"& novell4&"', "_ 
				&" '"&novell5&"', '"& novell6&"', '"& nt4&"', '"& k2&"', '"& k03&"', '"& ipadr&"', '"& subnet&"', '"& gateway&"', '"& dns&"', '"& smtp&"', '"& ftp&"', '"& bnavn&"', '"& pw&"', "_ 
				&" '"&arb_win98&"', '"& arb_winME&"', '"& arb_winNT&"', '"& arb_win2000&"', '"& arb_winXP &"', '"& arb_mac8&"', '"& arb_mac9&"', '"& arb_mac10&"', "_
		        &" '"&arb_dos&"', '"& kopi_farve&"', '"& kopi_sh&"', '"& kopi_it&"', '"& kopi_t1&"', '"& kopi_t2&"', '"& kopi_b1_a5&"', '"& kopi_b1_a4&"', "_
				&" '"&kopi_b1_a3&"', '"& kopi_b2_a5&"', '"& kopi_b2_a4&"', '"& kopi_b2_a3&"', '"& kopi_b3_a5&"', '"& kopi_b3_a4&"', '"& kopi_b3_a3&"', "_
		        &" '"&faxdriver&"', '"& faxnr&"', '"& faxid&"', '"& faxfarve&"', '"& fax_papribk_1&"', '"& fax_papribk_2&"', '"& fax_papribk_3&"', "_
				&" '"&fax_papribk_4&"', '"& fax_it&"', '"& fax_t1&"', '"& fax_t2&"', '"& fax_rap&"', '"& print_plc5&"', '"& print_plc6&"', '"& print_psl2&"', "_
				&" '"&print_psl3&"', '"& print_farve&"', '"& print_sh&"', '"& print_it&"', '"& print_t1&"', '"& print_t2&"', '"& print_admin&"', "_
				&" '"&scan_email&"', '"& scan_efile&"', '"& scan_ftp&"', '"& scan_esame&"', '"& scan_twain&"', '"& scan_filedownload&"', "_
				&" '"&scan_topacc&"', '"& scan_agfa&"', '"& bemaerk&"', '"& signoffdato&"', '"& underskrift &"', '"& pandet &"')")
				
				
				Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
					' Sætter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "TimeOut Sign-off."
					' Afsenderens e-mail 
					Mailer.FromAddress = "signoff@outzource.dk"
					Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient "RavnIT", "sign-off@ravnit.dk"
					Mailer.AddBCC "Support", "support@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "Nyt Sign-off. (" & strNavn &")"
					
					' Selve teksten
					Mailer.BodyText = "Hej RavnIT."& vbCrLf _
					&"Der er oprettet et nyt Sign-off."& vbCrLf & vbCrLf _ 
					&"Med venlig hilsen" & vbCrLf _ 
					&"TimeOut Sign-off email service."
					
					Mailer.SendMail
				
				oConn.execute(strSQL)
				
				else
				
				strSQL = ("UPDATE signoff SET navn = '"& strNavn &"', editor = '"& strEditor &"', dato = '"& strDato &"', "_
				&" adresse = '"&instadr&"', postnr = '"& postnr &"', town = '"& town &"', kpersnavn = '"& kpers &"', "_
				&" kpersstilling = '"& kpersstilling &"', kperstlf = '"& kperstlf&"', kpersfax = '"& kpersfax&"', kpersmobil = '"& kpersmobil&"', "_
				&" kpersemail = '"&kpersemail&"', itansnavn = '"& itans &"', itanstlf = '"& itanstlf&"', itansfax = '"& itansfax&"', itansmobil = '"& itansmobil&"', "_
				&" itansemail = '"& itansemail&"', virktype = '"& virktype&"', "_
				&" virktypeandet = '"&virktypeandet&"', pafgift = '"& pafgift&"', tohsalgskon = '"& tohsalgskon&"', tohsyskon = '"& tohsyskon&"', "_
				&" placering = '"& placering&"', placeringandet = '"& placeringandet&"', "_
				&" model = '"&model&"', usednew = '"& usednew&"', proveops = '"& proveops&"', levdato = '"& levdato&"', "_
				&" levtid = '"& levtid&"', insdato = '"& insdato&"', instid = '"& instid&"', kurdato = '"& kurdato&"', kurtid = '"& kurtid&"', "_ 
				&" kurskonsul = '"&kurskonsul&"', rj45 = '"& rj45&"', ethernet = '"& ethernet&"', token = '"& token&"', IPX_SPX = '"& IPX_SPX&"', TCP_IP = '"& TCP_IP&"', "_
				&" appleTalk = '"& appleTalk&"', novell3 = '"& novell3&"', novell4 = '"& novell4&"', "_ 
				&" novell5 = '"&novell5&"', novell6 = '"& novell6&"', nt4 = '"& nt4&"', k2 = '"& k2&"', k03 = '"& k03&"', "_
				&" ipadr = '"& ipadr&"', subnet = '"& subnet&"', gateway = '"& gateway&"', dns = '"& dns&"', smtp = '"& smtp&"', ftp = '"& ftp&"', "_
				&" bnavn = '"& bnavn&"', pw = '"& pw&"', "_ 
				&" arb_win98 = '"&arb_win98&"', arb_winME = '"& arb_winME&"', arb_winNT = '"& arb_winNT&"', arb_win2000 = '"& arb_win2000&"', arb_winXP = '"& arb_winXP &"', "_
				&" arb_mac8 = '"& arb_mac8&"', arb_mac9 = '"& arb_mac9&"', arb_mac10 = '"& arb_mac10&"', "_
		        &" arb_dos = '"&arb_dos&"', kopi_farve = '"& kopi_farve&"', kopi_sh = '"& kopi_sh&"', kopi_it = '"& kopi_it&"', kopi_t1 = '"& kopi_t1&"', kopi_t2 = '"& kopi_t2&"', kopi_b1_a5 = '"& kopi_b1_a5&"', kopi_b1_a4 = '"& kopi_b1_a4&"', "_
				&" kopi_b1_a3 = '"&kopi_b1_a3&"', kopi_b2_a5 = '"& kopi_b2_a5&"', kopi_b2_a4 = '"& kopi_b2_a4&"', kopi_b2_a3 = '"& kopi_b2_a3&"', "_
				&" kopi_b3_a5 = '"& kopi_b3_a5&"', kopi_b3_a4 = '"& kopi_b3_a4&"', kopi_b3_a3 = '"& kopi_b3_a3&"', "_
		        &" faxdriver = '"&faxdriver&"', faxnr = '"& faxnr&"', faxid = '"& faxid&"', faxfarve = '"& faxfarve&"', fax_papribk_1 = '"& fax_papribk_1&"', fax_papribk_2 = '"& fax_papribk_2&"', fax_papribk_3 = '"& fax_papribk_3&"', "_
				&" fax_papribk_4 = '"&fax_papribk_4&"', fax_it = '"& fax_it&"', fax_t1 = '"& fax_t1&"', fax_t2 = '"& fax_t2&"', "_
				&" fax_rap = '"& fax_rap&"', print_plc5 = '"& print_plc5&"', print_plc6 = '"& print_plc6&"', print_psl2 = '"& print_psl2&"', "_
				&" print_psl3 = '"&print_psl3&"', print_farve = '"& print_farve&"', print_sh = '"& print_sh&"', print_it = '"& print_it&"', "_
				&" print_t1 = '"& print_t1&"', print_t2 = '"& print_t2&"', print_admin = '"& print_admin&"', "_
				&" scan_email = '"&scan_email&"', scan_efile = '"& scan_efile&"', scan_ftp = '"& scan_ftp&"', scan_esame = '"& scan_esame&"', "_
				&" scan_twain = '"& scan_twain&"', scan_filedownload = '"& scan_filedownload&"', "_
				&" scan_topacc = '"&scan_topacc&"', scan_agfa = '"& scan_agfa&"', bemaerk = '"& bemaerk&"', signoffdato = '"& signoffdato&"', underskrift = '"& underskrift &"', pandet = '"& pandet &"' WHERE id = "&id&"")
				
				'Response.write strSQL
				oConn.execute(strSQL)
				end if
				
				
				if request("menu") <> "no" AND session("user") <> "" then
				Response.redirect "signoff.asp?menu=kund"
				else
				%>
				<!--#include file="../inc/regular/header_inc.asp"-->
				<div style="position:absolute; left:250; top:80;">
				<br><br>
				<b>Signoff er oprettet!</b><br>
				Du har nu mulighed for at 
				<a href="signoff.asp?func=opret&menu=no&key=2.013-1803-B015&lto=ravnit" class=vmenu>Oprette et nyt Signoff?</a><br>
				Eller du kan afslutte og <a href="Javascript:window.close()" class=vmenu>lukke dette vindue</a> ned.
				</div><%
				end if
		
			
		end if
	
	case "opret", "red"
	
	
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	dag = day(now)
	md = month(now)
	Aar = year(now)
	
	else
	strSQL = "SELECT id, navn, editor, dato, "_
	&" adresse, postnr, town, kpersnavn, kpersstilling, kperstlf, kpersfax, kpersmobil, "_
	&" kpersemail, itansnavn, itanstlf, itansfax, itansmobil, itansemail, virktype, "_
	&" virktypeandet, pafgift, tohsalgskon, tohsyskon, placering, placeringandet, "_
	&" model, usednew, proveops, levdato, levtid, insdato, instid, kurdato, kurtid, "_ 
	&" kurskonsul, rj45, ethernet, token, IPX_SPX, TCP_IP, appleTalk, novell3, novell4, "_ 
	&" novell5, novell6, nt4, k2, k03, ipadr, subnet, gateway, dns, smtp, ftp, bnavn, pw, "_ 
	&" arb_win98, arb_winME, arb_winNT, arb_win2000, arb_winXP, arb_mac8, arb_mac9, arb_mac10, "_
	&" arb_dos, kopi_farve, kopi_sh, kopi_it, kopi_t1, kopi_t2, kopi_b1_a5, kopi_b1_a4, "_
	&" kopi_b1_a3, kopi_b2_a5, kopi_b2_a4, kopi_b2_a3, kopi_b3_a5, kopi_b3_a4, kopi_b3_a3, "_
	&" faxdriver, faxnr, faxid, faxfarve, fax_papribk_1, fax_papribk_2, fax_papribk_3, "_
	&" fax_papribk_4, fax_it, fax_t1, fax_t2, fax_rap, print_plc5, print_plc6, print_psl2, "_
	&" print_psl3, print_farve, print_sh, print_it, print_t1, print_t2, print_admin, "_
	&" scan_email, scan_efile, scan_ftp, scan_esame, scan_twain, scan_filedownload, "_
	&" scan_topacc, scan_agfa, bemaerk, signoffdato, underskrift, pandet FROM signoff WHERE id=" & id
	
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	instadr = oRec("adresse")
	postnr = oRec("postnr")
	town = oRec("town") 
	kpers = oRec("kpersnavn") 
	kpersstilling = oRec("kpersstilling")
	kperstlf = oRec("kperstlf")
	kpersfax = oRec("kpersfax")
	kpersmobil = oRec("kpersmobil")
	kpersemail = oRec("kpersemail")
	itans = oRec("itansnavn")
	itanstlf = oRec("itanstlf")
	itansfax = oRec("itansfax")
	itansmobil = oRec("itansmobil")
	itansemail = oRec("itansemail")
	
	virktype =	oRec("virktype")
	virktypeandet = oRec("virktypeandet")
	pafgift = oRec("pafgift")
	
	tohsalgskon = oRec("tohsalgskon")
	tohsyskon = oRec("tohsyskon")
	
	placering = oRec("placering")
	placeringandet = oRec("placeringandet")
	model = oRec("model")
	usednew = oRec("usednew")
	proveops = oRec("proveops")
	
	
	levdato = oRec("levdato")
	levdag = day(levdato)
	levmd = month(levdato)
	levaar = year(levdato)
	
	levtid = oRec("levtid")
	
	insdato = oRec("insDato")
	insdag = day(insdato)
	insmd = month(insdato)
	insaar = year(insdato)
	
	instid = oRec("instid")
	
	kurdato = oRec("kurDato")
	kurdag = day(kurdato)
	kurmd = month(kurdato)
	kuraar = year(kurdato)
	
	kurtid = oRec("kurtid")
	kurskonsul = oRec("kurskonsul")
	
	rj45 = oRec("rj45")
	ethernet = oRec("Ethernet")
	token = oRec("Token")
	IPX_SPX = oRec("IPX_SPX")
	TCP_IP = oRec("TCP_IP") 
	appleTalk = oRec("AppleTalk")
	
	novell3 = oRec("Novell3")
	novell4 = oRec("Novell4")
	novell5 = oRec("Novell5")
	novell6 = oRec("Novell6")
	nt4 = oRec("nt4")
	k2 = oRec("k2")
	k03 = oRec("k03")
	
	ipadr1 = left(oRec("ipadr"), 3)
	ipadr2 = mid(oRec("ipadr"),4,3)
	ipadr3 = mid(oRec("ipadr"),8,3)
	ipadr4 = right(oRec("ipadr"), 3)
	subnet = oRec("subnet")
	gateway = oRec("gateway")
	dns = oRec("dns")
	smtp = oRec("smtp")
	ftp = oRec("ftp")
	bnavn = oRec("bnavn")
	pw = oRec("pw")
	
	arb_win98 = oRec("arb_win98") 
	arb_winME = oRec("arb_winME")
	arb_winNT = oRec("arb_winNT")
	arb_win2000 = oRec("arb_win2000")
	arb_winXP = oRec("arb_winXP")
	arb_mac8 = oRec("arb_mac8")
	arb_mac9 = oRec("arb_mac9")
	arb_mac10 = oRec("arb_mac10") 
	arb_dos = oRec("arb_dos")
	
	kopi_farve = oRec("kopi_farve") 
	kopi_sh = oRec("kopi_sh")
	kopi_it = oRec("kopi_it")
	kopi_t1 = oRec("kopi_t1")
	kopi_t2 = oRec("kopi_t2") 
	
	'**** feltnavne
	kopi_b1_a5 = oRec("kopi_b1_a5")
	kopi_b1_a4 = oRec("kopi_b1_a4")
	kopi_b1_a3 = oRec("kopi_b1_a3")
	
	kopi_b2_a5 = oRec("kopi_b2_a5")
	kopi_b2_a4 = oRec("kopi_b2_a4")
	kopi_b2_a3 = oRec("kopi_b2_a3")
	
	kopi_b3_a5 = oRec("kopi_b3_a5")
	kopi_b3_a4 = oRec("kopi_b3_a4")
	kopi_b3_a3 = oRec("kopi_b3_a3")
	 '****
	
	faxdriver = oRec("faxdriver")
	faxnr = oRec("faxnr")
	faxid = oRec("faxid")
	faxfarve = oRec("faxfarve")
	fax_papribk_1 = oRec("fax_papribk_1")
	fax_papribk_2 = oRec("fax_papribk_2")
	fax_papribk_3 = oRec("fax_papribk_3")
	fax_papribk_4 = oRec("fax_papribk_4")
	fax_it = oRec("fax_it")
	fax_t1 = oRec("fax_t1")
	fax_t2 = oRec("fax_t2")
	fax_rap = oRec("fax_rap")
	
	print_plc5 = oRec("print_plc5")
	print_plc6 = oRec("print_plc6") 
	print_psl2 = oRec("print_psl2")
	print_psl3 = oRec("print_psl3")
	print_farve = oRec("print_farve")
	print_sh = oRec("print_sh")
	print_it = oRec("print_it")
	print_t1 = oRec("print_t1")
	print_t2 = oRec("print_t2")
	print_admin = oRec("print_admin")
	
	scan_email = oRec("scan_email") 
	scan_efile = oRec("scan_efile")
	scan_ftp = oRec("scan_ftp")
	scan_esame = oRec("scan_esame")
	scan_twain = oRec("scan_twain")
	scan_filedownload = oRec("scan_filedownload")
	scan_topacc = oRec("scan_topacc")
	scan_agfa = oRec("scan_agfa")
	
	bemaerk = oRec("bemaerk")
	
	signoffdato = oRec("signoffdato")
	dag = day(signoffdato)
	md = month(signoffdato)
	aar = year(signoffdato)
	underskrift = oRec("underskrift")
	pandet = oRec("pandet")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	<%if request("print") <> "j" then%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<%
			if request("menu") <> "no" AND session("user") <> "" then 'no login, kommer direkte fra rdir%>
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<!--#include file="../inc/regular/vmenu.asp"-->
			<%end if%>
	<%
	leftd = 190
	topd = 80
	else
	leftd = 20
	topd = 120%>
	<html>
	<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	<!-------------------------------Sideindhold------------------------------------->
	<%if request("print") <> "j" then
		if func = "red" then%>
		<div style="position:absolute; left:640; top:117; z-index:1000;"><a href="signoff.asp?func=red&print=j&id=<%=id%>" class=vmenu>Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a></div>
		<%end if%>
	<%else%>
	<div style="position:absolute; left:0; top:10; z-index:500;">
	<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width=480 style="padding-left:20;"><img src="http://www.outzource.dk/timeout_xp/wwwroot/ver2_1/inc/upload/ravnit/toshiba.gif" alt="" border="0"></td>
					<td bgcolor="#FFFFFF"><img src="http://www.outzource.dk/timeout_xp/wwwroot/ver2_1/inc/upload/ravnit/Time-out_logo.gif" alt="" border="0"></td>
					<td bgcolor="#FFFFFF"><img src="http://www.outzource.dk/timeout_xp/wwwroot/ver2_1/inc/upload/ravnit/NDK_logo.jpg" alt="" border="0"></td>
				</tr>
			</table>
	</div>
	<%end if%>
	<div id="sindhold" style="position:absolute; left:<%=leftd%>; top:<%=topd%>; visibility:visible;">
	<%if request("print") = "j" then
	bgthis = "#ffffff"
	useborder = 0
	else
	bgthis = "#eff3ff"
	useborder = 1
	end if%>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="<%=bgthis%>">
	<%if request("print") = "j" then%>
	<tr><td colspan=4><br>
	<h3>Sign-off</h3>
		<b>Installationen omfatter</b>
		<ul>
		<li>Kopi
		<li>Fax
		<li>Print
		<li>Scan
		<li>Instruktion
		</ul>
		
		<b>Forudsætninger for installation.</b><br>
		Vedligeholdelse af nedenstående konfiguration påhviler kunden. Ravn IT kan ikke påtage sig ansvar for fejl, der ikke er omfattet af, eller skyldes ændring til nedenstående konfiguration.
		<br><br>
		Nedenstående informationer er vigtige for at Ravn IT kan opfylde forventningerne til det leverede produkt. Det er kundens ansvar, at nedenstående beskrivelse stemmer overens med de virkelige forudsætninger. Ravn IT kan ikke følgende gøres ansvarlig for manglende funktionalitet, som har baggrund i en ufuldstændig beskrivelse.
		<br><br>
		Følgende skal være til rådighed inden installation påbegyndes:
		<ul>
		<li>	Kundens systemadministrator/-systemansvarlig.
		<li>    Originalsoftware til operativsystem.
		<li>	Operativsystem skal være 100% funktionelt på de aktuelle computere.
		<li>	Administratorpassword til server, password til arbejdsstationer.
		<li>	230V edb-stik, netværksstik og telefonstik (fax) max. 2 meter fra TOSHIBA udstyr,
		<li>	med mindre andet er aftalt og anført under bemærkninger.
		<li>    Netværkstilslutning (hub/switch) som understøtter hastigheden for TOSHIBA udstyr.
		</ul>
		Installation er afsluttet ved skriftlig accept af funktionstest.<br><br>&nbsp;
		<%end if%>
	</td></tr>
	
	<form action="signoff.asp?menu=<%=request("menu")%>&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<%if request("print") <> "j" then%>
	<tr bgcolor="#d6dff5">
    	<td valign="top" colspan=4><font class="pageheader">Sign-Off | <%=varbroedkrumme%></font><br>&nbsp;</td>
	</tr>
		<%if dbfunc = "dbred" then%>
		<tr bgcolor="#d6dff5">
			<td colspan="4" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
		</tr>
		<%end if%>
	<%end if%>
	
	<%if request("print") <> "j" then%>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:<%=useborder%>px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=4 class=alt><b>Installationsoplysninger.</b></td>
	</tr>
	<%else%>
	<tr>
		<td colspan=4><b>Installationsoplysninger.</b></td>
	</tr>
	<%end if%>
	<tr>
		<td rowspan=84 style="border-left:<%=useborder%>px #003399 solid; border-top:<%=useborder%>px #003399 solid; padding-left:10;">&nbsp;</td>
		<td width="120" style="border-top:<%=useborder%>px #003399 solid;"><br><b>Firma:</b></td>
		<td style="border-top:<%=useborder%>px #003399 solid;"><br><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td rowspan=84 style="border-right:<%=useborder%>px #003399 solid; border-top:<%=useborder%>px #003399 solid; padding-right:5;">&nbsp;</td>
	</tr>
	<tr>
		<td>Installationsadresse:</td>
		<td><input type="text" name="FM_instadr" value="<%=instadr%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td>Postnr:</td>
		<td><input type="text" name="FM_postnr" value="<%=postnr%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		By:<input type="text" name="FM_town" value="<%=town%>" style="width:140; !border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td><br><b>Kontaktperson:</b></td>
		<td><br><input type="text" name="FM_kpers" value="<%=kpers%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Stilling:</td>
		<td><input type="text" name="FM_kpersstilling" value="<%=kpersstilling%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Telefon:</td>
		<td><input type="text" name="FM_kperstlf" value="<%=kperstlf%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Fax:</td>
		<td><input type="text" name="FM_kpersfax" value="<%=kpersfax%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Mobil:</td>
		<td><input type="text" name="FM_kpersmobil" value="<%=kpersmobil%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >E-mail:</td>
		<td><input type="text" name="FM_kpersemail" value="<%=kpersemail%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td ><br><b>IT-ansvarlig:</b></td>
		<td><br><input type="text" name="FM_itans" value="<%=itans%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Telefon:</td>
		<td><input type="text" name="FM_itanstlf" value="<%=itanstlf%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Fax:</td>
		<td><input type="text" name="FM_itansfax" value="<%=itansfax%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Mobil:</td>
		<td><input type="text" name="FM_itansmobil" value="<%=itansmobil%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >E-mail:</td>
		<td><input type="text" name="FM_itansemail" value="<%=itansemail%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan=2><br><b>Virksomhed:</b></td>
	</tr>
	<tr>
		<td valign=top> Type:</td>
		<td><select name="FM_virktype" style="width:202;">
		<%
		x = 0
		for x = 0 to 7
		
		select case x
		case 0
		namevalue = "Handel"
		case 1
		namevalue = "Grafisk"
		case 2
		namevalue = "Produkton"
		case 3
		namevalue = "Finansielle branche"
		case 4
		namevalue = "Liberale erhverv"
		case 5
		namevalue = "Advokat/Revisor"
		case 6
		namevalue = "Stat/Amt/Kommune"
		case 7
		namevalue = "Andet"
		end select
			if virktype = namevalue then
			selthis = "SELECTED"
			else
			selthis = ""
			end if
		%>
		<option value="<%=namevalue%>" <%=selthis%>><%=namevalue%></option>
		<%
		next
		%>
		</select></td>
	</tr>
	<%if selthis = "SELECTED" then
		virktypeandet = virktypeandet
		else
		virktypeandet = ""
		end if%>
	<tr>
		<td>Andet:</td>
		<td><input type="text" name="FM_virktype_andet" value="<%=virktypeandet%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>	
		
		
	<tr>
		<td valign="top"><br><b>Parkeringsforhold:</b></td>
		<%call checked(pafgift)%>
		<td><br>P-afgift <input type="checkbox" <%=chk%> name="FM_pafgift" value="j"></td>
	</tr>
		<tr><td>Andet: </td><td><input type="text" name="FM_pandet" value="<%=pandet%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan=2><br><b>TOSHIBA kontakt info:</b></td>
	</tr>
	<tr>
		<td >Salgskonsulent:</td>
		<td><input type="text" name="FM_tohsalgskon" value="<%=tohsalgskon%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Systemkonsulent:</td>
		<td><input type="text" name="FM_tohsyskon" value="<%=tohsyskon%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td valign=top><br><b>Maskinens fysiske placering:</b></td>
		<td><br><select name="FM_placering" style="width:202;">
		<%
		x = 0
		for x = 0 to 4
		
		select case x
		case 0
		namevalue = "Stuen"
		case 1
		namevalue = "1. Sal"
		case 2
		namevalue = "2. Sal"
		case 3
		namevalue = "3. Sal"
		case 4
		namevalue = "Andet"
		end select
			if placering = namevalue then
			selthis = "SELECTED"
			else
			selthis = ""
			end if
		%>
		<option value="<%=namevalue%>" <%=selthis%>><%=namevalue%></option>
		<%
		next
		%>
		</select><br>
		<%if selthis = "SELECTED" then
		placeringandet = placeringandet
		else
		placeringandet = ""
		end if%>
	</td></tr>
	<tr>
		<td>Andet:</td><td> <input type="text" name="FM_placering_andet" value="<%=placeringandet%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan=2><br><b>Maskintype:</b></td>
	</tr>
	<tr>
		<td >Model:</td>
		<td><input type="text" name="FM_model" value="<%=model%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td >Ny / Brugt:</td>
		<%call checked(usednew)%>
		<td>Ny<input type="radio" <%=chk%> name="FM_usednew" value="j">
		<%if chk = "CHECKED" then
		chk = ""
		else
		chk = "CHECKED"
		end if%>
		&nbsp;&nbsp;Brugt<input type="radio" <%=chk%> name="FM_usednew" value="n"></td>
	</tr>
	<tr>
		<td >Prøveopstilling:</td>
		<%call checked(proveops)%> 
		<td><input type="checkbox" <%=chk%> name="FM_proveops" value="j"></td>
	</tr>
	<tr>
		<td colspan=2><br><b>Levering:</b></td>
	</tr>
	<tr>
		<td >Dato:</td>
		<td><input type="text" name="FM_levdag" value="<%=levdag%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_levmd" value="<%=levmd%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_levaar" value="<%=levaar%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;dd-mm-åååå</td>
	</tr>
	<tr>
		<td >Tid:</td>
		<td><input type="text" name="FM_levtid" value="<%
		if instr(levtid, ":") > 0 then
		Response.write right(levtid, 8)
		else
		levtid = ""
		Response.write levtid
		end if
		%>" size="8" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;tt:mm</td>
	</tr>
	<tr>
		<td colspan=2><br><b>Installation:</b></td>
	</tr>
	<tr>
		<td >Dato:</td>
		<td><input type="text" name="FM_insdag" value="<%=insdag%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_insmd" value="<%=insmd%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_insaar" value="<%=insaar%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;dd-mm-åååå</td>
	</tr>
	<tr>
		<td >Tid:</td>
		<td><input type="text" name="FM_instid" value="<%
		if instr(instid, ":") > 0 then
		Response.write right(instid, 8)
		else
		instid = ""
		Response.write instid 
		end if
		%>" size="8" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;tt:mm</td>
	</tr>
	<tr>
		<td colspan=2><br><b>Kursus:</b></td>
	</tr>
	<tr>
		<td >Dato:</td>
		<td><input type="text" name="FM_kurdag" value="<%=kurdag%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_kurmd" value="<%=kurmd%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_kuraar" value="<%=kuraar%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;dd-mm-åååå</td>
	</tr>
	<tr>
		<td >Tid:</td>
		<td><input type="text" name="FM_kurtid" value="<%
		if instr(kurtid, ":") > 0 then
		Response.write right(kurtid, 8)
		else
		kurtid = ""
		Response.write kurtid
		end if
		%>" size="8" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;tt:mm</td>
	</tr>
	<tr>
		<td >Konsulent:</td>
		<td><input type="text" name="FM_kurskonsul" value="<%=kurskonsul%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td valign=top><br><b>Netværkstopologi:</b></td>
		<td><br>
		<%call checked(rj45)%> 
		<input type="checkbox" <%=chk%> name="FM_rj45" value="j">&nbsp;Ethernet PDS RJ-45
		&nbsp;&nbsp;&nbsp;
		<%call checked(ethernet)%> 
		<input type="checkbox" <%=chk%> name="FM_Ethernet" value="j">&nbsp;Ethernet BNC Coax
		<%call checked(token)%> 
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_Token" value="j">&nbsp;Token Ring
		</td>
	</tr>
	<tr>
		<td valign=top><br><b>Netværksprotokol:</b></td>
		<td><br>
		<%call checked(IPX_SPX)%> 
		<input type="checkbox" <%=chk%> name="FM_IPX_SPX" value="j">&nbsp;IPX/SPX
		<%call checked(tcp_ip)%> 
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_TCP_IP" value="j">&nbsp;TCP/IP
		<%call checked(appleTalk)%> 
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_AppleTalk" value="j">&nbsp;AppleTalk
		</td>
	</tr>
	<tr>
		<td valign=top><br><b>Server operativsystem:</b></td>
		<td valign=top><br>
		<table><tr>
		<td><%call checked(Novell3)%> 
		<input type="checkbox" <%=chk%> name="FM_Novell3" value="j">&nbsp;Novell 3.xx</td>
		<td><%call checked(Novell4)%> 
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_Novell4" value="j">&nbsp;Novell 4.xx
		</td>
		<td><%call checked(Novell5)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_Novell5" value="j">&nbsp;Novell 5.xx
		</td>
		<td><%call checked(Novell6)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_Novell6" value="j">&nbsp;Novell 6.xx
		</td>
		</tr>
		<tr>
		<td><%call checked(nt4)%>
		<input type="checkbox" <%=chk%> name="FM_nt4" value="j">&nbsp;NT 4.0 
		</td>
		<td><%call checked(k2)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_2000" value="j">&nbsp;Windows 2000 
		</td>
		<td><%call checked(k03)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_2003" value="j">&nbsp;Windows 2003 
		</td>
		<td>&nbsp;</td>
		</tr>
		</table>
		
		</td>
	</tr>
	<tr>
		<td valign=top colspan=2><br><b>IP indstillinger:</b></td>
	</tr>
	<tr>
		<td>IP adresse:</td><td><input type="text" name="FM_ipadr1" value="<%=ipadr1%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_ipadr2" value="<%=ipadr2%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_ipadr3" value="<%=ipadr3%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		 <input type="text" name="FM_ipadr4" value="<%=ipadr4%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		 </td>
	</tr>
	<tr>
		<td>Subnet mask:</td><td><input type="text" name="FM_subnet" value="<%=subnet%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		 </td>
	</tr>
	<tr>
		<td>Default Gateway:</td><td><input type="text" name="FM_gateway" value="<%=gateway%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		</td>
	</tr>
	<tr>
		<td>DNS server:</td><td> <input type="text" name="FM_dns" value="<%=dns%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br></td>
	</tr>
	<tr>
		<td>SMTP server:</td><td> <input type="text" name="FM_smtp" value="<%=smtp%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br></td>
	</tr>
	<tr>
		<td>FTP server:</td><td> <input type="text" name="FM_ftp" value="<%=ftp%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		 </td>
	</tr>
	<tr>
		<td>FTP brugernavn:</td><td><input type="text" name="FM_bnavn" value="<%=bnavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		 </td>
	</tr>
	<tr>
		<td>FTP password:</td><td><input type="text" name="FM_pw" value="<%=pw%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
	</td>
	</tr>
	<tr>
		<td valign=top><br><b>Arbejdsstationer:</b></td>
		<td><br>
		<table><tr>
		<td><%call checked(arb_win98)%>
		<input type="checkbox" <%=chk%> name="FM_arb_win98" value="j">&nbsp;win 95
		</td>
		<td><%call checked(arb_winME)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_winME" value="j">&nbsp;Win 98/ME
		</td>
		<td><%call checked(arb_winNT)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_winNT" value="j">&nbsp;NT
		</td>
		<td><%call checked(arb_win2000)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_win2000" value="j">&nbsp;2000
		</td>
		<td><%call checked(arb_winXP)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_winXP" value="j">&nbsp;XP
		</td>
		</tr>
		<tr>
		<td><%call checked(arb_mac8)%>
		<input type="checkbox" <%=chk%> name="FM_arb_mac8" value="j">&nbsp;MAC 8.x
		</td>
		<td><%call checked(arb_mac9)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_mac9" value="j">&nbsp;MAC 9.x
		</td>
		<td><%call checked(arb_mac10)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_mac10" value="j">&nbsp;MAC 10.x
		</td>
		<td><%call checked(arb_dos)%>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" <%=chk%> name="FM_arb_dos" value="j">&nbsp;Dos
		</td>
		<td>&nbsp;</td>
		</tr></table>
		</td>
	</tr>
	<tr>
		<td valign=top colspan=2><br><b>Kopi opsætning:</b></td>
	</tr>
	<tr>
		<td valign=top>Standard udskrift:</td><td><%call checked(kopi_farve)%>
		<input type="checkbox" <%=chk%> name="FM_kopi_farve" value="j">&nbsp;Farve&nbsp;&nbsp;
		<%call checked(kopi_sh)%>
		<img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><input type="checkbox" <%=chk%> name="FM_kopi_sh" value="j">&nbsp;Sort / hvid</td>
	</tr>
	<tr>
		<td valign=top>Standard udfaldsbakke:</td><td><%call checked(kopi_it)%><input type="checkbox" <%=chk%> name="FM_kopi_it" value="j">&nbsp;Inner Tray
		<%call checked(kopi_t1)%><img src="../ill/blank.gif" width="5" height="1" alt="" border="0"><input type="checkbox" <%=chk%> name="FM_kopi_t1" value="j">&nbsp;Tray 1
		<%call checked(kopi_t2)%><input type="checkbox" <%=chk%> name="FM_kopi_t2" value="j">&nbsp;Tray 2
		</td>
	</tr>
	<tr>
		<td valign=top>Bakke 1:</td><td><%call checked(kopi_b1_a5)%><input type="checkbox" <%=chk%> name="FM_kopi_b1_a5" value="j">&nbsp;A5
		<%call checked(kopi_b1_a4)%><input type="checkbox" <%=chk%> name="FM_kopi_b1_a4" value="j">&nbsp;A4
		<%call checked(kopi_b1_a3)%><input type="checkbox" <%=chk%> name="FM_kopi_b1_a3" value="j">&nbsp;A3
		</td>
	</tr>
	<tr>
		<td valign=top>Bakke 2:</td><td><%call checked(kopi_b2_a5)%><input type="checkbox" <%=chk%> name="FM_kopi_b2_a5" value="j">&nbsp;A5
		<%call checked(kopi_b2_a4)%><input type="checkbox" <%=chk%> name="FM_kopi_b2_a4" value="j">&nbsp;A4
		<%call checked(kopi_b2_a3)%><input type="checkbox" <%=chk%> name="FM_kopi_b2_a3" value="j">&nbsp;A3
		</td>
	</tr>
	<tr>
		<td valign=top>Bakke 3:</td><td><%call checked(kopi_b3_a5)%><input type="checkbox" <%=chk%> name="FM_kopi_b3_a5" value="j">&nbsp;A5
		<%call checked(kopi_b3_a4)%><input type="checkbox" <%=chk%> name="FM_kopi_b3_a4" value="j">&nbsp;A4
		<%call checked(kopi_b3_a3)%><input type="checkbox" <%=chk%> name="FM_kopi_b3_a3" value="j">&nbsp;A3
		</td>
	</tr>
	<tr>
		<td valign=top colspan=2><br><b>FAX opsætning:</b></td>
	</tr>
	<tr>	
		<td>N/W-Fax driver:</td>
		<td><%call checked(faxdriver)%><input type="checkbox" <%=chk%> name="FM_faxdriver" value="j"> </td>
	</tr>
	<tr>	
		<td>Fax nummer:</td>
		<td><input type="text" name="FM_faxnr" value="<%=faxnr%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td>
	</tr>
	<tr><td>	
		Fax ID:</td>
		<td> <input type="text" name="FM_faxid" value="<%=faxid%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td>
	</tr>
	<tr><td>Papirfarve:</td>
		<td> <input type="text" name="FM_faxfarve" value="<%=faxfarve%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td>
	</tr>
	<tr><td>
		Papirbakke:</td>
		<td><%call checked(fax_papribk_1)%><input type="checkbox" <%=chk%> name="FM_fax_papribk_1" value="j">&nbsp;Bakke 1
		&nbsp;&nbsp;&nbsp;<%call checked(fax_papribk_2)%><input type="checkbox" <%=chk%> name="FM_fax_papribk_2" value="j">&nbsp;Bakke 2
		&nbsp;&nbsp;&nbsp;<%call checked(fax_papribk_3)%><input type="checkbox" <%=chk%> name="FM_fax_papribk_3" value="j">&nbsp;Bakke 3
		&nbsp;&nbsp;&nbsp;<%call checked(fax_papribk_4)%><input type="checkbox" <%=chk%> name="FM_fax_papribk_4" value="j">&nbsp;Bakke 4
	</td>
	</tr>
	<tr><td>
		Standard udfaldsbakke:</td>
		<td><%call checked(fax_it)%><input type="checkbox" <%=chk%> name="FM_fax_it" value="j">&nbsp;Inner Tray
		<%call checked(fax_t1)%><input type="checkbox" <%=chk%> name="FM_fax_t1" value="j">&nbsp;Tray 1
		<img src="../ill/blank.gif" width="14" height="1" alt="" border="0"><%call checked(fax_t2)%><input type="checkbox" <%=chk%> name="FM_fax_t2" value="j">&nbsp;Tray 2
	</td>
	</tr>
	<tr><td>
		Bekræftelsesrapport: </td><td><%call checked(fax_rap)%><input type="checkbox" <%=chk%> name="FM_fax_rap" value="j">
	</td>
	</tr>
	
	<tr>
		<td valign=top colspan=2><br><b>Print opsætning:</b></td>
	</tr>
	<tr>	
		<td>Printdriver:</td>
		<td> <%call checked(print_plc5)%><input type="checkbox" <%=chk%> name="FM_print_plc5" value="j">&nbsp;PCL5
		<img src="../ill/blank.gif" width="24" height="1" alt="" border="0"><%call checked(print_plc6)%><input type="checkbox" <%=chk%> name="FM_print_plc6" value="j">&nbsp;PCL6
		<img src="../ill/blank.gif" width="7" height="1" alt="" border="0"><%call checked(print_psl2)%><input type="checkbox" <%=chk%> name="FM_print_psl2" value="j">&nbsp;PSL2
		&nbsp;&nbsp;&nbsp;<%call checked(print_psl3)%><input type="checkbox" <%=chk%> name="FM_print_psl3" value="j">&nbsp;PSL3
	</td>
	</tr>
	<tr><td>
		Standard udskrift:</td>
		<td>
		<%call checked(print_farve)%><input type="checkbox" <%=chk%> name="FM_print_farve" value="j">&nbsp;Farve
		<img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><%call checked(print_sh)%><input type="checkbox" <%=chk%> name="FM_print_sh" value="j">&nbsp;Sort / hvid
		</td>
	</tr>
	<tr><td>
		Standard udfaldsbakke:</td>
		<td><%call checked(print_it)%><input type="checkbox" <%=chk%> name="FM_print_it" value="j">&nbsp;Inner Tray
		<%call checked(print_t1)%><input type="checkbox" <%=chk%> name="FM_print_t1" value="j">&nbsp;Tray 1
		<%call checked(print_t2)%><input type="checkbox" <%=chk%> name="FM_print_t2" value="j">&nbsp;Tray 2
	</td>
	</tr>
	<tr><td>
		Administrators brugernavn:</td><td> <input type="text" size=30 name="FM_print_admin" value="<%=print_admin%>" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
	</td>
	</tr>
	
	<tr>
		<td valign=top colspan=2><br><b>Scan opsætning:</b></td>
	</tr>
	<tr>	
		<td>Scan to e-mail:</td>
		<td> <%call checked(scan_email)%><input type="checkbox" <%=chk%> name="FM_scan_email" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>Scan to e-file:</td>
		<td> <%call checked(scan_efile)%><input type="checkbox" <%=chk%> name="FM_scan_e-file" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>Scan to FTP:</td>
		<td> <%call checked(scan_ftp)%><input type="checkbox" <%=chk%> name="FM_scan_ftp" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>Scan to e-same:</td>
		<td> <%call checked(scan_esame)%><input type="checkbox" <%=chk%> name="FM_scan_e-same" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>Scan to TWAIN driver:</td>
		<td> <%call checked(scan_twain)%><input type="checkbox" <%=chk%> name="FM_scan_twain" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td colspan=2><br><b>Klient software:</b>
	</td>
	</tr>
	<tr>	
		<td>File Downloader:</td>
		<td> <%call checked(scan_filedownload)%><input type="checkbox" <%=chk%> name="FM_scan_filedownload" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>TopAccessDocMon:</td>
		<td> <%call checked(scan_topacc)%><input type="checkbox" <%=chk%> name="FM_scan_topacc" value="j">&nbsp;
		</td>
	</tr>
	<tr>	
		<td>Agfa FontManager:</td>
		<td> <%call checked(scan_agfa)%><input type="checkbox" <%=chk%> name="FM_scan_agfa" value="j">&nbsp;
		</td>
	</tr>
	<tr>
		<td valign=top colspan=2><br><b>Bemærkninger:</b><br>
		<textarea cols="65" rows="6" name="FM_bemaerk" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=bemaerk%></textarea></td>
	</tr>
	<tr>
		<td valign=top colspan=2><br><b>Afsluttet og godkendt:</b></td>
	</tr>
	<%if request("print") <> "j" then%>
	<tr>
		<td>Dato:</td>
		<td><input type="text" name="FM_dag" value="<%=dag%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_md" value="<%=md%>" size="1" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="text" name="FM_aar" value="<%=aar%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;dd-mm-åååå</td>
	</tr>
	<tr><td>
		Kundeunderskrift:</td><td> <input type="text" size=30 name="FM_underskrift" value="<%=underskrift%>" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
	</td>
	</tr>
	<tr>
		<td colspan="4" style="border-left:<%=useborder%>px #003399 solid; border-right:<%=useborder%>px #003399 solid; border-bottom:<%=useborder%>px #003399 solid;" align=center><br><br><input type="image" src="../ill/<%=varSubVal%>.gif"><br><br>&nbsp;</td>
	</tr>
	<%else%>
	<tr><td colspan=2><b>Dato:</b> <%=formatdatetime(signoffdato, 1)%> af <%=underskrift%></td></tr>
	<tr><td colspan=2><br><br>
	<b>Underskrift:</b>&nbsp;_________________________
	</td>
	</tr>
	<tr>
		<td colspan="4" style="border-left:<%=useborder%>px #003399 solid; border-right:<%=useborder%>px #003399 solid; border-bottom:<%=useborder%>px #003399 solid;" align=center><br><br>&nbsp;</td>
	</tr>
	<%end if%>
	</form>
	</table>
	<br><br>
	<br>
	<%if request("print") <> "j" then%>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<%end if%>
	<br>
	<br>
	</div>
	<%case else%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:70; visibility:visible;">
	<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Sign-Off.</b><br>
	Opret og vedligehold Sign-Off's her.</td>
	</tr>
	</table>
	
		

	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><img src="../ill/blank.gif" width="490" height="1" alt="" border="0">
	<a href="signoff.asp?menu=kund&func=opret" class=vmenu>Opret ny <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="520" bgcolor="#ffffff">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=5 valign="top"><img src="../ill/tabel_top.gif" width="504" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=5 class=alt><b>Sign-Off's.</b></td>
	</tr>
	<tr>
	<td width="5" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td width=80><b><a href="signoff.asp?menu=kund&sort=dato">Dato</a></b></td>
	<td height="30" width=160><b><a href="signoff.asp?menu=kund">Firma</a></b></td>
	<td height="30" width=120><b>Kontaktperson</b></td>
	<td height="30" width=80><b>K.pers. Mobil</b></td>
	<td>&nbsp;</td>
	<td width="5" style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	if request("sort") = "dato" then
	sortby = "dato DESC"
	else
	sortby = "navn"
	end if
	
	strSQL = "SELECT id, navn, editor, dato, "_
	&" adresse, postnr, town, kpersnavn, kpersstilling, kperstlf, kpersfax, kpersmobil, "_
	&" kpersemail, itansnavn, itanstlf, itansfax, itansmobil, itansemail, virktype, "_
	&" virktypeandet, pafgift, tohsalgskon, tohsyskon, placering, placeringandet, "_
	&" model, usednew, proveops, levdato, levtid, insdato, instid, kurdato, kurtid, "_ 
	&" kurskonsul, rj45, ethernet, token, IPX_SPX, TCP_IP, appleTalk, novell3, novell4, "_ 
	&" novell5, novell6, nt4, k2, k03, ipadr, subnet, gateway, dns, smtp, ftp, bnavn, pw, "_ 
	&" arb_win98, arb_winME, arb_winNT, arb_win2000, arb_winXP, arb_mac8, arb_mac9, arb_mac10, "_
	&" arb_dos, kopi_farve, kopi_sh, kopi_it, kopi_t1, kopi_t2, kopi_b1_a5, kopi_b1_a4, "_
	&" kopi_b1_a3, kopi_b2_a5, kopi_b2_a4, kopi_b2_a3, kopi_b3_a5, kopi_b3_a4, kopi_b3_a3, "_
	&" faxdriver, faxnr, faxid, faxfarve, fax_papribk_1, fax_papribk_2, fax_papribk_3, "_
	&" fax_papribk_4, fax_it, fax_t1, fax_t2, fax_rap, print_plc5, print_plc6, print_psl2, "_
	&" print_psl3, print_farve, print_sh, print_it, print_t1, print_t2, print_admin, "_
	&" scan_email, scan_efile, scan_ftp, scan_esame, scan_twain, scan_filedownload, "_
	&" scan_topacc, scan_agfa, bemaerk, signoffdato, underskrift FROM signoff ORDER BY "& sortby

	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="7" style="border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=formatdatetime(oRec("dato"), 2)%></td>
		<td height="20"><a href="signoff.asp?menu=kund&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td><%=oRec("kpersnavn")%></td>
		<td><%=oRec("kpersmobil")%></td>
		<td><a href="signoff.asp?menu=kund&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr><td colspan=7 bgcolor="#8caae6" style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td></tr>	
	</table></td>
	</tr></table>
	
	<br><br><br>
	<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
