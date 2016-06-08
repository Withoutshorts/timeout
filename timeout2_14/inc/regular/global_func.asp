
<!--#include file="../xml/login_xml_inc.asp"-->
<!--#include file="../xml/menu_xml_inc.asp"-->
<!--#include file="../xml/global_xml_inc.asp"-->
<!--#include file="../xml/tsa_xml_inc.asp"-->
<!--#include file="../xml/erp_fak_xml_inc.asp"-->





<!--#include file="cls_aktiviteter.asp"-->
<!--#include file="cls_afstem.asp"-->
<!--#include file="cls_afstem_akttyper.asp"-->

<!--#include file="cls_dage.asp"-->

<!--#include file="cls_fak.asp"-->
<!--#include file="cls_filer.asp"-->


<!--#include file="cls_help.asp"-->
<!--#include file="cls_mat.asp"-->
<!--#include file="cls_mat_senlist_mwh1.asp"-->


<!--#include file="cls_medarb.asp"-->
<!--#include file="cls_job.asp"-->
<!--#include file="cls_projektgrp.asp"-->

<!--#include file="cls_stade.asp"-->
<!--#include file="cls_smiley.asp"-->
<!--#include file="cls_stempelur.asp"-->


<!--#include file="cls_timer.asp"-->
<!--#include file="cls_timepriser.asp"-->
<!--#include file="cls_todo.asp"-->

<!--#include file="cls_ugeseddel.asp"-->
<!--#include file="cls_valuta.asp"-->


<!--#include file="global_func2.asp"-->
<!--#include file="global_func3.asp"-->
<!--#include file="global_func4.asp"-->





<%
'*** Er global inc_ inkluderet på den aktuelle side?? 
global_inc = "j"


call basisValutaFN()

public strKundenavnPDFtxt

function kundenavnPDF(strknavn)

             strknavn = replace(strknavn, " ", "_")
             strknavn = replace(strknavn, "&", "_")
             strknavn = replace(strknavn, "<", "_")
             strknavn = replace(strknavn, ">", "_")
             strknavn = replace(strknavn, "/", "_")
             strknavn = replace(strknavn, "!", "_")
             strknavn = replace(strknavn, ":", "_")
             strknavn = replace(strknavn, ".", "_")
             strknavn = replace(strknavn, ",", "_")
             strknavn = replace(strknavn, "¤", "_")
             strknavn = replace(strknavn, "'", "_")
             strknavn = replace(strknavn, "''", "_")
             strknavn = replace(strknavn, "}", "_")
             strknavn = replace(strknavn, "{", "_")
             strknavn = replace(strknavn, "]", "_")
             strknavn = replace(strknavn, "[", "_")
             strknavn = replace(strknavn, "?", "_")
             strknavn = replace(strknavn, "´", "_")
             strknavn = replace(strknavn, "|", "_")
             strknavn = replace(strknavn, "\", "_")
             strknavn = replace(strknavn, "%", "_")
             strknavn = replace(strknavn, "*", "_")
             strknavn = replace(strknavn, "¨", "_")
             strknavn = replace(strknavn, "#", "_")
             strknavn = replace(strknavn, "+", "_")
             strKundenavnPDFtxt = lcase(strknavn)

end function 




public alfanumeriskTxt

function alfanumerisk(txtstr)

             txtstr = replace(txtstr, " ", "")
             txtstr = replace(txtstr, "_", "")
             txtstr = replace(txtstr, "-", "")
             txtstr = replace(txtstr, "&", "")
             txtstr = replace(txtstr, "<", "")
             txtstr = replace(txtstr, ">", "")
             txtstr = replace(txtstr, "/", "")
             txtstr = replace(txtstr, "!", "")
             txtstr = replace(txtstr, ":", "")
             txtstr = replace(txtstr, ".", "")
             txtstr = replace(txtstr, ",", "")
             txtstr = replace(txtstr, "¤", "")
             txtstr = replace(txtstr, "'", "")
             txtstr = replace(txtstr, "''", "")
             txtstr = replace(txtstr, "}", "")
             txtstr = replace(txtstr, "{", "")
             txtstr = replace(txtstr, "]", "")
             txtstr = replace(txtstr, "[", "")
             txtstr = replace(txtstr, "?", "")
             txtstr = replace(txtstr, "´", "")
             txtstr = replace(txtstr, "|", "")
             txtstr = replace(txtstr, "\", "")
             txtstr = replace(txtstr, "%", "")
             txtstr = replace(txtstr, "*", "")
             txtstr = replace(txtstr, "¨", "")
             txtstr = replace(txtstr, "#", "")
             txtstr = replace(txtstr, "+", "")         

             alfanumeriskTxt = txtstr

end function 


'*** Afsluttede uger *****
public afslUgerMedab
redim afslUgerMedab(3000) '2500 '400 '4 år
function afsluger(medarbid, stdato, sldato)
		
        call smileyAfslutSettings()

        if cint(SmiWeekOrMonth) = 0 then
        strPeriode = "WEEK(u.uge, 1)"
        else
        strPeriode = "MONTH(u.uge)"
        end if

		strSQL2 = "SELECT u.status, u.afsluttet, "& strPeriode &" AS periode, YEAR(u.uge) AS aar, "_
		&" u.id, u.mid FROM ugestatus u WHERE mid =  "& medarbid &""_
		&" AND uge BETWEEN '"& stDato &"' AND '"& slDato &"' GROUP BY u.mid, uge"
		
        'Response.write strSQL2
		'Response.end
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
		
			afslUgerMedab(oRec2("mid")) = afslUgerMedab(oRec2("mid")) & "#"& oRec2("periode")&"_"& oRec2("aar") &"#,"
		
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


public SmiWeekOrMonth, SmiantaldageCount, SmiantaldageCountClock, SmiTeamlederCount
function smileyAfslutSettings()


                '** Variable (defaukl værdier)
                SmiWeekOrMonth = 0
                SmiantaldageCount = 1 'mandag
                SmiantaldageCountClock = "12:00:00"
                SmiTeamlederCount = 1
                
        
                strSQL = "SELECT SmiWeekOrMonth, SmiantaldageCount, SmiantaldageCountClock, SmiTeamlederCount FROM licens WHERE id = 1"
                oRec6.open strSQL, oConn, 3
                if not oRec6.EOF then
            
                SmiWeekOrMonth = oRec6("SmiWeekOrMonth")
                SmiantaldageCount = oRec6("SmiantaldageCount")
                
                if oRec6("SmiantaldageCountClock") = 24 then
                SmiantaldageCountClock = "23:59:00"
                else
                SmiantaldageCountClock = oRec6("SmiantaldageCountClock") & ":00:00"
                end if

                SmiTeamlederCount = oRec6("SmiTeamlederCount")
                
                

                end if 
                oRec6.close


end function




public timesimon, timesimh1h2, timesimtp
function timesimon_fn()
    
    timesimon = 0
    timesimh1h2 = 0
    timesimtp = 0
	strSQL6 = "SELECT timesimon, timesimh1h2, timesimtp FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	timesimon = oRec6("timesimon")
    timesimh1h2 = oRec6("timesimh1h2")
    timesimtp = oRec6("timesimtp")
	
	end if
    oRec6.close

end function

public budgetakt
function budgetakt_fn()
    
    budgetakt = 0
    
	strSQL6 = "SELECT budgetakt FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	budgetakt = oRec6("budgetakt")
 	
	end if
    oRec6.close

end function


public fomr_account
function fomr_account_fn()
    
    fomr_account = 0
	strSQL6 = "SELECT fomr_account FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	fomr_account = oRec6("fomr_account")
	
	end if
    oRec6.close

end function


public traveldietexp_maxhours, traveldietexp_on
function traveldietexp_fn()

    traveldietexp_maxhours = 0
    traveldietexp_on = 0
	strSQL6 = "SELECT traveldietexp_maxhours, traveldietexp_on FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	traveldietexp_maxhours = oRec6("traveldietexp_maxhours")
    traveldietexp_on = oRec6("traveldietexp_on")
	
	end if
    oRec6.close

end function


public medarbtypligmedarb
function medarbtypligmedarb_fn()

    medarbtypligmedarb = 0
    strSQL6 = "SELECT medarbtypligmedarb FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	medarbtypligmedarb = oRec6("medarbtypligmedarb")
	
	end if
    oRec6.close

end function



public minimumslageremail
function minimumslageremail_fn()
    
    minimumslageremail = 0
	strSQL6 = "SELECT minimumslageremail FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	minimumslageremail = oRec6("minimumslageremail")
	
	end if
    oRec6.close

end function

    
public visAktlinjerSimpel, visAktlinjerSimpel_datoer, visAktlinjerSimpel_timebudget, visAktlinjerSimpel_realtimer, visAktlinjerSimpel_restimer
public visAktlinjerSimpel_medarbtimepriser, visAktlinjerSimpel_medarbrealtimer, visAktlinjerSimpel_akttype
function visAktSimpel_fn()
    
    visAktlinjerSimpel = 0
    visAktlinjerSimpel_datoer = 0
    visAktlinjerSimpel_timebudget = 0
    visAktlinjerSimpel_realtimer = 0
    visAktlinjerSimpel_restimer = 0
    visAktlinjerSimpel_medarbtimepriser = 0 
    visAktlinjerSimpel_medarbrealtimer = 0
    visAktlinjerSimpel_akttype = 0

	strSQL6 = "SELECT visAktlinjerSimpel, visAktlinjerSimpel_datoer, visAktlinjerSimpel_timebudget, visAktlinjerSimpel_realtimer, visAktlinjerSimpel_restimer, "_
    &" visAktlinjerSimpel_medarbtimepriser, visAktlinjerSimpel_medarbrealtimer, visAktlinjerSimpel_akttype FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	visAktlinjerSimpel = oRec6("visAktlinjerSimpel")
    visAktlinjerSimpel_datoer = oRec6("visAktlinjerSimpel_datoer")
    visAktlinjerSimpel_timebudget = oRec6("visAktlinjerSimpel_timebudget")
    visAktlinjerSimpel_realtimer = oRec6("visAktlinjerSimpel_realtimer")
    visAktlinjerSimpel_restimer = oRec6("visAktlinjerSimpel_restimer")
    visAktlinjerSimpel_medarbtimepriser = oRec6("visAktlinjerSimpel_medarbtimepriser") 
    visAktlinjerSimpel_medarbrealtimer = oRec6("visAktlinjerSimpel_medarbrealtimer")
    visAktlinjerSimpel_akttype = oRec6("visAktlinjerSimpel_akttype")
	
	end if
    oRec6.close

end function

public positiv_aktivering_akt_val
function positiv_aktivering_akt_fn()
    
    positiv_aktivering_akt_val = 0
	strSQL6 = "SELECT positiv_aktivering_akt FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	positiv_aktivering_akt_val = oRec6("positiv_aktivering_akt")
	
	end if
    oRec6.close

end function



public timerround
function timerround_fn()
    
    timerround = 0
	strSQL6 = "SELECT timerround FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	timerround = oRec6("timerround")
	
	end if
    oRec6.close

end function


public smiley_agg, hidesmileyicon
function smiley_agg_fn()
    
    smiley_agg = 0
    hidesmileyicon = 0

	strSQL6 = "SELECT smileyaggressiv, hidesmileyicon FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	smiley_agg = oRec6("smileyaggressiv")
    hidesmileyicon = oRec6("hidesmileyicon")
	
	end if
    oRec6.close

end function


public teamleder_flad
function teamleder_flad_fn()
    
    teamleder_flad = 0
	strSQL6 = "SELECT teamleder_flad FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	teamleder_flad = oRec6("teamleder_flad")
	
	end if
    oRec6.close

end function



public akt_maksbudget_treg, akt_maksforecast_treg
function akt_maksbudget_treg_fn

    akt_maksbudget_treg = 0
    akt_maksforecast_treg = 0
	strSQL6 = "SELECT akt_maksbudget_treg, akt_maksforecast_treg FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	akt_maksbudget_treg = oRec6("akt_maksbudget_treg")
    akt_maksforecast_treg = oRec6("akt_maksforecast_treg")
	
	end if
    oRec6.close



end function




public fomr_mandatoryOn
function fomr_mandatory_fn()
    
    fomr_mandatoryOn = 0
	strSQL6 = "SELECT fomr_mandatory FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	fomr_mandatoryOn = oRec6("fomr_mandatory")
	
	end if
    oRec6.close

end function


public lukaktvdato
function lukaktvdato_fn()
    
    lukaktvdato = 0
	strSQL6 = "SELECT lukaktvdato FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	lukaktvdato = oRec6("lukaktvdato")
	
	end if
    oRec6.close

end function


public showSalgsAnv
function salgsans_fn()
    
    showSalgsAnv = 0
	strSQL6 = "SELECT salgsans FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	showSalgsAnv = oRec6("salgsans")
	
	end if
    oRec6.close

end function


public bdgmtypon_val, bdgmtypon_prgrp
function bdgmtypon_fn()
    
    bdgmtypon_val = 0
    bdgmtypon_prgrp = 0
	strSQL6 = "SELECT bdgmtypon FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	bdgmtypon_val = oRec6("bdgmtypon")
	
	end if
    oRec6.close


    strSQL6 = "SELECT COUNT(id) AS antalIds FROM medarbtyper_grp WHERE id <> 0"

    oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	bdgmtypon_prgrp = oRec6("antalIds")
	
	end if
    oRec6.close


end function

public aktBudgettjkOn, aktBudgettjkOn_afgr, aktBudgettjkOnRegAarSt, aktBudgettjkOnViskunmbgt
function aktBudgettjkOn_fn()
            
    'Response.write "her"
    'Response.Flush

    aktBudgettjkOn = 0
    aktBudgettjkOnRegAarSt = "01-01-2001"
    aktBudgettjkOn_afgr = 0
    aktBudgettjkOnViskunmbgt = 0
    
	strSQL6 = "SELECT forcebudget_onakttreg, forcebudget_onakttreg_afgr, regnskabsaar_start, forcebudget_onakttreg_filt_viskunmbgt FROM licens AS l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	aktBudgettjkOn = oRec6("forcebudget_onakttreg")
    aktBudgettjkOn_afgr = oRec6("forcebudget_onakttreg_afgr")
    aktBudgettjkOnRegAarSt = oRec6("regnskabsaar_start")
    aktBudgettjkOnViskunmbgt = oRec6("forcebudget_onakttreg_filt_viskunmbgt")
    
	
	end if
    oRec6.close

end function

public showuploadimport
function showuploadimport_fn()
    
    showuploadimport = 0
	strSQL6 = "SELECT showupload FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	showuploadimport = oRec6("showupload")
	
	end if
    oRec6.close

end function


public showEasyreg_val
function showEasyreg_fn()
    
    showEasyreg_val = 0
	strSQL6 = "SELECT showeasyreg FROM licens l WHERE id = 1"
	oRec6.open strSQL6, oConn, 3
	If not oRec6.EOF then
	
	showEasyreg_val = oRec6("showeasyreg")
	
	end if
    oRec6.close

end function

public startDatoAar, startDatoMd, startDatoDag, licensklienter, licensstdato
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
    
    userAgent = request.servervariables("HTTP_USER_AGENT")

    if (instr(lcase(userAgent), "iphone") <> 0 OR instr(lcase(userAgent), "iemobile") <> 0 _
    OR instr(lcase(userAgent), "android") <> 0 OR instr(lcase(userAgent), "mobile") <> 0 _ 
    OR inStr(1, userAgent, "iphone", 1) > 0 or inStr(1, userAgent, "windows ce", 1) > 0 or inStr(1, userAgent, "blackberry", 1) > 0 or inStr(1, userAgent, "opera mini", 1) > 0 _
    OR inStr(1, userAgent, "mobile", 1) > 0 or inStr(1, userAgent, "palm", 1) > 0 or inStr(1, userAgent, "portable", 1) > 0) AND instr(lcase(userAgent), "ipad") = 0 then
	'** Iphone **'
    
    browstype_client = "ip"
    else


	if instr(userAgent , "Firefox") <> 0 then
	browstype_client = "mz"
	else
            
            if instr(userAgent , "Chrome") <> 0 then
	        browstype_client = "ch"
	        else

                if instr(userAgent , "Safari") <> 0 then
	            browstype_client = "sf"
	            else
                browstype_client = "ie"
	            end if

            end if

	
	end if

    end if
   
	user_agent_txt = userAgent

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

public stempelurOn, stempelur_hideloginOn, stempelur_ignokomkravOn
function erStempelurOn()
'** Stempelur ***'
	stempelurOn = 0
    stempelur_hideloginOn = 0
    stempelur_ignokomkravOn = 0
	strSQL = "SELECT stempelur, stempelur_hidelogin, stempelur_igno_komkrav FROM licens WHERE id = 1"
	
    'response.write strSQL
    'response.flush
    
    oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	
    stempelurOn = oRec("stempelur") 
    stempelur_hideloginOn = oRec("stempelur_hidelogin")
    stempelur_ignokomkravOn = oRec("stempelur_igno_komkrav")

	end if
	oRec.close 
end function

public timeout_version, toVer, toVerPath, toSubVer, toSubVerPath14, toSubVerPath15
function TimeOutVersion()
    
    '** TimeOut version ***'
	timeout_version = "ver2_1"
	strSQL = "SELECT timeout_version FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	timeout_version = oRec("timeout_version") 
	end if
	oRec.close 

    ''** 3_99 eller 2_1 PATH **'

    if instr(request.servervariables("PATH_TRANSLATED"), "ver2_1") <> 0 then
	toVer = "ver2_1"
	end if

	if instr(request.servervariables("PATH_TRANSLATED"), "ver2_10") <> 0 then
	toVer = "ver2_10"
	end if

    if instr(request.servervariables("PATH_TRANSLATED"), "ver2_14") <> 0 then
	toVer = "ver2_14"
	end if

    if instr(request.servervariables("PATH_TRANSLATED"), "ver3_99") <> 0 then
	toVer = "ver3_99"
	end if

     if instr(request.servervariables("PATH_TRANSLATED"), "to_2015") <> 0 then
	'toVer = "ver3_99"
	end if

    toSubVerPath14 = "../timereg/"
    toSubVerPath15 = "../to_2015/"


    'Response.Write request.servervariables("PATH_TRANSLATED")
	toVerPath = toVer

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
		
		
		
		function xresstopmenu()

       
        resmenupkt = "Ressource Forecast (timebudget)"
       

		%>
			<br>
			<a href="jbpla_w.asp?menu=res" class=rmenu>Ressource Kalender</a>
			&nbsp;&nbsp;|&nbsp;&nbsp;<a href="ressource_belaeg_jbpla.asp" class=rmenu><%=resmenupkt %></a>
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
		
		'if cint(treg0206thisMid) = 0 then
		'	tregLink = "timereg.asp?menu=timereg"
		'else
			'tregLink = "timereg_2006_fs.asp"
			tregLink = "timereg_akt_2006.asp?hideallbut_first=2"
		'end if
		
		call mmenuPkt(1, "102", tregLink,""& global_txt_120 &"",pkt)
		
		if level <= 2 OR level = 6 OR lto = "kejd_pb" OR lto = "kejd_pb2" then
            
            if lto = "kejd_pb" OR lto = "kejd_pb2" then
            call mmenuPkt(2, "101", "ressource_belaeg_jbpla.asp?menu=webblik",""& replace(global_txt_121,"|", "&") &"",pkt)
            else
            call mmenuPkt(2, "101", "webblik_joblisten.asp?menu=webblik&fromvmenu=1",""& replace(global_txt_121,"|", "&") &"",pkt)
            end if
		end if
		
		
		
		'if level <= 2 OR level = 6 then
		    'call mmenuPkt(4, "20", "ressource_belaeg_jbpla.asp?menu=res",""& global_txt_123 &"",pkt)
		'end if
		
		
		if level <= 3 OR level = 6 then
		   call mmenuPkt(5, "09", "kunder.asp?menu=kund&visikkekunder=1",""& global_txt_124 &"",pkt)
		end if
		
		if level <= 3 OR level = 6 then
		    call mmenuPkt(3, "63", "jobs.asp?menu=job&shokselector=1&fromvemenu=j",""& global_txt_122 &"",pkt)
		end if
		
		if level = 1 then
		call mmenuPkt(6, "19", "medarb.asp?menu=medarb",""& global_txt_125 &"",pkt)
        else
        call mmenuPkt(6, "19", "medarb_red.asp?menu=medarb&func=red&id="& session("mid") &"","Rediger profil",pkt)
        end if
		
		if level <= 2 OR level = 6 then
		    call mmenuPkt(7, "33", "joblog_timetotaler.asp",""& global_txt_126 &"",pkt)
		end if
		
		
        if level <= 2 then
		call mmenuPkt(10, "38", "filer.asp?kundeid=0&jobid=0",""& global_txt_127 &"",pkt)
        end if
		
		
		if (level <= 3 OR level = 6) AND lto <> "synergi1" then
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
		call mmenuPkt(1, "12", "crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0","Kalender",pkt)
		call mmenuPkt(2, "09", "kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt","Kontakter",pkt)
		call mmenuPkt(3, "56", "crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist","Aktions historik",pkt)
		call mmenuPkt(4, "99", "crmstat.asp?menu=crm","CRM stat",pkt)
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
		'call mmenuPkt(2, "45", "erp_afstem_md.asp?menu=erp","Afstemning",pkt)
		call mmenuPkt(2, "45", "erp_job_afstem.asp?menu=erp&show=joblog_afstem","Afstemning",pkt)

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
		<table cellspacing=0 cellpadding=0 border=0 id="menutable" style="width:1116px;" bgcolor="#ffffff">
		<tr>
			<td colspan=50 height=25 valign=top><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
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
		     <td id="screenw" colspan=50 valign=top style="border-top:3px #8CAAE6 solid; height:1px;">
               <img src="../ill/blank.gif" width="1" height="1" alt="" border="0"><br /></td>
				</tr>
		<%end if%>
		</table>
		
		<%
		
		end function
		
		
		function mmenuPkt(nb, img, lnk, lnkTxt, valgtPkt)
		
		if cint(nb) = cint(valgtPkt) then
		bgthis = "#ffff99"
		bgr = 0
		else
		bgthis = "#EFF3FF" '"#eff3ff" 
		bgr = 0	
		end if
		
		'if img = "56" then
		'tgt = "_blank"
		'else
		tgt = "_top"
		'end if
		
		tdwdt = (len(lnkTxt) * 8.5)
		%>
		
		
        
		
		<td align=center width="<%=tdwdt %>" id="mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgthis%>; border:<%=bgr %>px orange solid; border-bottom:0px;" onmouseover="bgcolthisMON('mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>');" onmouseout="bgcolthisOFF('mmenupkt_<%=nb%>_<%=left(lnkTxt,3)%>','<%=bgthis%>');">
	    <a href="<%=lnk%>" class="mainmenu" target="<%=tgt%>"><%=lnkTxt%></a>
		</td>
		<td bgcolor="#cccccc" style="width:1px;"><img src="../ill/blank.gif" width="1" height="3" alt="" border="0"></td>
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
                        <td height=30 bgcolor="#Eff3ff" style="border-bottom:1px #999999 solid;" colspan=2><b>Registrerings-log for:</b></td>
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
					
					call helligdage(CheckDay&"/"&md&"/"&ye, 0, lto)
					
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


   
	
	
	
	
	
	
	
		
	
	
		function visfirmalogo(plefther, ptopher, kundeid)
		
		%>
		<div id="firmalogo" style="position:absolute; left:<%=plefther%>; top:<%=ptopher%>; background-color:#ffffff; z-index=10; padding:20px; border:10px silver solid; overflow:hidden;">
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
	
	
	function opretLink_2013(lnkUrl, lnkTxt, lnkAlt, lnkTgt, lnkWdt)

        %>
	 <div style="background-color:forestgreen; padding:5px 5px 5px 5px; width:<%=lnkWdt%>px;"><a href='<%=lnkUrl %>' class='alt' alt="<%=lnkAlt %>" target="<%=lnkTgt %>"><%=lnkTxt %> +</a></div>
	<%
	end function
	
	
        
        
     function AjaxUpdate(table, column, msg)
	
    'Response.Write "her"
    'Response.end
    if table = "job" then 'må ikke opdatere interne

        tjkRisiko = 0
        strSQL = "SELECT risiko FROM job WHERE id = " & request.Form("id")
        oRec8.open strSQL, oConn, 3
        if not oREc8.EOF then

        tjkRisiko = oRec8("risiko")

        end if
        oRec8.close



    end if


    if request.Form("value") <> "" AND table <> "" AND column <> "" AND tjkRisiko >= 0 then
	strSQL = "Update " & table & " set " & column & " = '" & request.Form("value") & "' where id = " & request.Form("id")
	
	'Response.Write strSQL
	'Response.flush
	oConn.execute(strSQL)
	
    
        'strSQL = "Update " & table & " set " & column & " = '" & column+1 & "' where id <> " & request.Form("id") & " AND " & column & " >= " & request.Form("value")
        
        'Response.Write strSQL
	    'oConn.execute(strSQL)
        'Response.end


    Response.Write "timeout notifikation: <br />" & msg & strSQL
	Response.End
	end if
	end function


    function AjaxUpdateTreg(table, column, msg)
	
    'Response.Write "her"
    'Response.flush
    
    dtnow = year(now) &"/"& month(now) &"/"& day(now)

    if request.Form("value") <> "" AND table <> "" AND column <> "" then
	strSQL = "Update " & table & " set " & column & " = '" & request.Form("value") & "', forvalgt_af = '"& session("mid") &"', forvalgt_dt = '"& dtnow &"' where id = " & request.Form("id")
	
	'Response.Write strSQL
	'Response.flush
	oConn.execute(strSQL)
	'Response.Write "timeout notifikation: <br />" & msg
	'Response.End
	end if
	end function


	
	function infoUnisport(uWdt, uTxt)%>
	 <div style="width:<%=uWdt%>px; padding:3px 7px 3px 7px; border:1px #6CAE1C solid; background-color:#DCF5BD;"><%=uTxt %></div>
    <%end function
    
    function infoUnisportAB(uWdt, uTxt, uTop, uLeft)%>
	 <div style="position: absolute; top:<%=uTop%>px; left:<%=uLeft%>px; width:<%=uWdt%>px; padding:3px 7px 3px 7px; border:1px #6CAE1C solid; background-color:#DCF5BD; z-index:90000000;"><%=uTxt %></div>
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




    function  qmarkhelpnote(qmtxt,qmtop,qmleft,qmid,qmwdt) 
        %>
           <div id="qmarkhelptxt_<%=qmid %>" class="qmarkhelptxt" style="position:absolute; top:<%=qmtop%>px; left:<%=qmleft%>px; visibility:hidden; font-weight:normal; background-color:#F7F7F7; border:1px #999999 solid; display:none; padding:10px; width:<%=qmwdt%>px; z-index:4000;">
             <table><tr>
                    <td style="font-size:11px; color:#999999;"><b>Info:</b></td>
                    <td align="right"><span class="closeqmnote" style="color:red;">[X]</span></td>
                </tr>
                 <tr><td colspan="2" style="font-size:11px; color:#999999;"><%=qmtxt %></td></tr>
             </table>

           </div>

        <%
    end function
     
	
	public lastjobnr, nextjobnr, lastTilbudsnr, nextTilbudsnr
    function lastjobnr_fn()


            lastJobnr = 0
            lastTilbudsnr = 0
            strSQLjobnr = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1 "
            oRec5.open strSQLjobnr, oConn, 3
            if not oRec5.EOF then

            lastJobnr = oRec5("jobnr")
            lastTilbudsnr = oRec5("tilbudsnr")

            end if
            oRec5.close

            nextjobnr = lastjobnr + 1
            nextTilbudsnr = nextTilbudsnr + 1

    end function


    function sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

            %>
   <!--slet sidens indhold-->
        <div class="container" style="width:500px;">
            <div class="porlet">
            
            <h3 class="portlet-title">
               <u><%=oskrift %></u>
            </h3>
            
                <div class="portlet-body">
                    <div style="text-align:center;"><%=slttxta%>
                        <%if len(trim(slttxtb)) <> 0 then %>
                        <br /><br />&nbsp;
                        <%=slttxtb%>
                        <%end if %>
                    </div><br />
                    <div style="text-align:center;"><a  class="btn btn-primary btn-sm" role="button" href="<%=slturl%>">&nbsp;Ja&nbsp;</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="Javascript:history.back()"><b>Nej</b></a>
                    </div>
                    <br /><br />
                 </div>

            </div>
        </div>
            <%
    end function
%>
    