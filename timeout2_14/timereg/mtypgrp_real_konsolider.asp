<%response.buffer = true
    
   



if len(trim(request("dothis"))) <> 0 then
dothis = request("dothis")
else
dothis = 1
end if


'** Er det første mnauelle konsolidering = 1, Specifikt job = 2, Auto hver nat = 0/3? (der kun tager de seneste 3 dage)
if len(trim(request("first"))) <> 0 then
first = request("first")
else
first = 0
end if


if len(trim(request("jobid"))) <> 0 then 'kun i forbindelse med first = 2
jobid = request("jobid")
else
jobid = 0
end if

if first = 2 then 
    highend = 0
else
    highend = 0 '3 NÅR ALLE EPI DATABASER ER flyttet
end if




'*** PSEUDO åBNER FOR AT KUNNE LÆSE GLBALFUNC

if first = 0 then 'Auto konsolider via service

        Set oConn = Server.CreateObject("ADODB.Connection")
        Set oRec = Server.CreateObject("ADODB.Recordset")
        Set oRec6 = Server.CreateObject("ADODB.Recordset")
        Set oRec4 = Server.CreateObject("ADODB.Recordset")
        Set oRec5 = Server.CreateObject("ADODB.Recordset")
        Set oRec3 = Server.CreateObject("ADODB.Recordset")
        Set oRec2 = Server.CreateObject("ADODB.Recordset")
        Set oCmd = Server.CreateObject("ADODB.Command")


        'ltothisPseudo = "demo"
        ltothisPseudo = "start_5_56"

         strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_"& ltothisPseudo &";"
         'strConnect = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;
         'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
         oConn.open strConnect

else
        %>
        <!--#include file="../inc/connection/conn_db_inc.asp"-->
        <%
end if



 %>

     

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="inc/isint_func.asp"-->


<%

if first = 0 then 'Auto konsolider via service
oConn.close
end if


'***********************************************************************************
'*********************************** LOOP ******************************************
'***********************************************************************************
l = 0
for l = 0 to highend


if first = 0 then 'Auto konsolider via service
    select case l
    case 0
        ltothis = "epi"
    case 1
        ltothis = "epi_ab"
    case 2
        ltothis = "epi_uk"
    case 3
        ltothis = "epi_no"
    end select


  strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_"& ltothis &";"
 'strConnect = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;
 'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
 oConn.open strConnect


end if


response.Write "LTOthis: " & ltothis & "<br>"
response.flush



 


call fn_medarbtyper()
call bdgmtypon_fn()
call mtyperIGrp_fn(0,1) 
call akttyper2009(2)



   

sqlDatoslut = now






    'first = 1

  select case first
    case 2
    
    %>
    <div style="position:absolute; left:50px; top:50px; border:0px #CCCCCC solid; padding:0px;">
        	<img src="../ill/outzource_logo_200.gif" />
    <h4>Job konsolidering</h4>

    Henter timer for hver måned og indlæser i konsolideringen.<br /><br />
        <h4 style="color:red;">Afvent besked om at <b>"Job er færdig konsolideret"</b> før dette vindue lukkes ned igen.</h4><br /><br />
    <%
   
    end select




        

if cint(first) = 0 then

    tasteDato = dateAdd("d", -2, sqlDatoslut) ' -2 dage
    'tasteDato = dateAdd("d", -130, sqlDatoslut) ' -130
    tasteDatoSQL = year(tasteDato)&"/"&month(tasteDato)&"/"&day(tasteDato)
    tasteDatoSQLend = year(sqlDatoslut)&"/"&month(sqlDatoslut)&"/"&day(sqlDatoslut)

    '**** finder job der er indtastet på igår + 2 dage ****'
    jobSQLkri = "AND (id = 0"
    strSQLtregigar = "SELECT t.tjobnr, j.id FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE t.tastedato BETWEEN '"& tasteDatoSQL & "' AND '"& tasteDatoSQLend &"'  GROUP BY t.tjobnr"
    oRec.open strSQLtregigar, oConn, 3
    while not oRec.EOF 
    
    jobSQLkri = jobSQLkri & " OR id = "& oRec("id")

    oRec.movenext
    wend
    oRec.close 

    jobSQLkri = jobSQLkri & ")"

 end if



'Ignorer job indtastet de sidste 3 dage ***'
select case first
   case 1 'alle job
    jobSQLkri = ""
   case 2 'specifikt jobid

    if instr(jobid, ",") <> 0 then 'multible fra konsolider siden

        jobSQLkri = "AND (id = 0 " 
        jobids = split(jobid, ",") 

        for j = 0 TO UBOUND(jobids)
        jobSQLkri = jobSQLkri & " OR id = " & jobids(j)
        next

        jobSQLkri = jobSQLkri & ")"

    else
    jobSQLkri = "AND (id = "& jobid &")" 
    end if

   case else 'job indenfor de sidste 3 dage '0 eller blank
    jobSQLkri = jobSQLkri
 end select 



    
        '**** FINDER SPECIFIKKE jobnr *****'
    
      if cint(first) = 1 then '1
                
    '''jobSQLkrinr = "AND (jobnr = 15719)"
    '''jobSQLkrinr = "AND (jobnr = 16070 OR jobnr = 15481 OR jobnr = 16243 OR jobnr = 16258 OR jobnr = 16262 OR jobnr = 17186 OR jobnr = 17187 OR jobnr = 17188 OR jobnr = 16875 OR jobnr = 16241 OR jobnr = 15867 OR jobnr = 15964 OR jobnr = 17209 OR jobnr = 17197 OR jobnr = 15687 OR jobnr = 15694 OR jobnr = 16194 OR jobnr = 16062 OR jobnr = 16298 OR jobnr = 17147 OR jobnr = 15604 OR jobnr = 16121 OR jobnr = 15932 OR jobnr = 15558 OR jobnr = 16095 OR jobnr = 16098 OR jobnr = 14860 OR jobnr = 16210 OR jobnr = 15760 OR jobnr = 15902 OR jobnr = 16694 OR jobnr = 16757 OR jobnr = 16322 OR jobnr = 15348 OR jobnr = 16181 OR jobnr = 16386 OR jobnr = 15199 OR jobnr = 16246 OR jobnr = 16342 OR jobnr = 16281 OR jobnr = 16277 OR jobnr = 16279 OR jobnr = 16303)"
    

   
 
        
     'jobSQLkrinr = "AND (jobnr = 15799 OR jobnr = 16430 OR jobnr = 16736 OR jobnr = 17237 OR jobnr = 16733 OR jobnr = 16431 OR jobnr = 17234 OR jobnr = 16996 OR jobnr = 15715 OR jobnr = 17054 OR jobnr = 17516 OR jobnr = 16562 OR jobnr = 15085 OR jobnr = 16071 OR jobnr = 16696 OR jobnr = 16975 OR jobnr = 17307 OR jobnr = 17080 OR jobnr = 16593 OR jobnr = 16652 OR jobnr = 16390 OR jobnr = 16503 OR jobnr = 16727 OR jobnr = 17134 OR jobnr = 17135 OR jobnr = 15483 OR jobnr = 16540 OR jobnr = 17232 OR jobnr = 16263 OR jobnr = 15760 OR jobnr = 16375 OR jobnr = 17140 OR jobnr = 17253 OR jobnr = 14387 OR jobnr = 15215 OR jobnr = 16482 OR jobnr = 16534 OR jobnr = 16539 OR jobnr = 16640 OR jobnr = 16748 OR jobnr = 16799 OR jobnr = 16811 OR jobnr = 16902 OR jobnr = 16941 OR jobnr = 17154 OR jobnr = 17199 OR jobnr = 17220 OR jobnr = 17230 OR jobnr = 17353 OR jobnr = 17396 OR jobnr = 17397 OR jobnr = 17406 OR jobnr = 17446 OR jobnr = 17658 OR jobnr = 15007 OR jobnr = 15235 OR jobnr = 15236 OR jobnr = 16376 OR jobnr = 16377 OR jobnr = 16446 OR jobnr = 17421 OR jobnr = 16541 OR jobnr = 17566 OR jobnr = 17567 OR jobnr = 17407 OR jobnr = 17012 OR jobnr = 16336 OR jobnr = 16471 OR jobnr = 17383 OR jobnr = 16476 OR jobnr = 16550 OR jobnr = 16599 OR jobnr = 16626 OR jobnr = 17088 OR jobnr = 17456 OR jobnr = 16382 OR jobnr = 16685 OR jobnr = 17204 OR jobnr = 17217 OR jobnr = 17469 OR jobnr = 17509 OR jobnr = 17653 OR jobnr = 16378 OR jobnr = 16337 OR jobnr = 17179 OR jobnr = 16489 OR jobnr = 16266 OR jobnr = 16519 OR jobnr = 16741 OR jobnr = 17268 OR jobnr = 17364 OR jobnr = 17413 OR jobnr = 16273 OR jobnr = 15932 OR jobnr = 16971 OR jobnr = 16824 OR jobnr = 16781 OR jobnr = 17192 OR jobnr = 17548 OR jobnr = 16774 OR jobnr = 17240 OR jobnr = 16691 OR jobnr = 16284 OR jobnr = 17398 OR jobnr = 17221 OR jobnr = 16792 OR jobnr = 17108 OR jobnr = 17401 OR jobnr = 15880 OR jobnr = 16488 OR jobnr = 16329 OR jobnr = 16339 OR jobnr = 16425 OR jobnr = 16439 OR jobnr = 16458 OR jobnr = 16508 OR jobnr = 16637 OR jobnr = 16717 OR jobnr = 16763 OR jobnr = 16789 OR jobnr = 16930 OR jobnr = 16931 OR jobnr = 17210 OR jobnr = 17499 OR jobnr = 17501 OR jobnr = 17502 OR jobnr = 16942 OR jobnr = 17079 OR jobnr = 17422 OR jobnr = 16280 OR jobnr = 16014 OR jobnr = 16629 OR jobnr = 17238 OR jobnr = 16545 OR jobnr = 17070 OR jobnr = 16616 OR jobnr = 17189 OR jobnr = 17190 OR jobnr = 17191 OR jobnr = 14981 OR jobnr = 17216 OR jobnr = 15113 OR jobnr = 15834 OR jobnr = 16196 OR jobnr = 16243 OR jobnr = 16477 OR jobnr = 16479 OR jobnr = 16480 OR jobnr = 16779 OR jobnr = 16780 OR jobnr = 16968 OR jobnr = 16969 OR jobnr = 17125 OR jobnr = 17186 OR jobnr = 17187 OR jobnr = 17188 OR jobnr = 17442 OR jobnr = 16246 OR jobnr = 17297 OR jobnr = 17006 OR jobnr = 17115 OR jobnr = 16463 OR jobnr = 16497 OR jobnr = 16716 OR jobnr = 17261 OR jobnr = 16198 OR jobnr = 16305 OR jobnr = 16442 OR jobnr = 17266 OR jobnr = 17447 OR jobnr = 17448 OR jobnr = 16054 OR jobnr = 16089 OR jobnr = 16869 OR jobnr = 17180 OR jobnr = 16955 OR jobnr = 16186 OR jobnr = 16383 OR jobnr = 17027 OR jobnr = 17101 OR jobnr = 17385 OR jobnr = 16563 OR jobnr = 16323 OR jobnr = 16107 OR jobnr = 16496 OR jobnr = 16542 OR jobnr = 16702 OR jobnr = 17256 OR jobnr = 17486 OR jobnr = 16688 OR jobnr = 13037 OR jobnr = 16136 OR jobnr = 16470 OR jobnr = 16473 OR jobnr = 16535 OR jobnr = 17584 OR jobnr = 16965 OR jobnr = 16947 OR jobnr = 16333 OR jobnr = 16647 OR jobnr = 17487 OR jobnr = 16543 OR jobnr = 17252 OR jobnr = 17067 OR jobnr = 16821 OR jobnr = 16084 OR jobnr = 17308 OR jobnr = 16521 OR jobnr = 15986 OR jobnr = 16009 OR jobnr = 16432 OR jobnr = 16398 OR jobnr = 14596 OR jobnr = 16281 OR jobnr = 16282 OR jobnr = 17214 OR jobnr = 16649 OR jobnr = 17106 OR jobnr = 17107 OR jobnr = 16357 OR jobnr = 16192 OR jobnr = 15888 OR jobnr = 16537 OR jobnr = 15719 OR jobnr = 16528 OR jobnr = 16700 OR jobnr = 16970 OR jobnr = 17306 OR jobnr = 16240 OR jobnr = 16330 OR jobnr = 16462 OR jobnr = 16728 OR jobnr = 16729 OR jobnr = 17049 OR jobnr = 17409 OR jobnr = 15781 OR jobnr = 16628 OR jobnr = 16759 OR jobnr = 16174 OR jobnr = 16424 OR jobnr = 16612 OR jobnr = 16944 OR jobnr = 17147 OR jobnr = 17303 OR jobnr = 17410 OR jobnr = 16478 OR jobnr = 17185 OR jobnr = 16659 OR jobnr = 16953 OR jobnr = 16743 OR jobnr = 16261 OR jobnr = 16904 OR jobnr = 16632 OR jobnr = 16324 OR jobnr = 15696 OR jobnr = 17053 OR jobnr = 16894 OR jobnr = 17201)"
     'jobnr = 3
     'jobSQLkrinr = "AND (jobnr = 100 OR jobnr = 100000 OR jobnr = 100001 OR jobnr = 100002 OR jobnr = 100003 OR jobnr = 100004 OR jobnr = 100005 OR jobnr = 100007 OR jobnr = 100011 OR jobnr = 100012 OR jobnr = 100013 OR jobnr = 100014 OR jobnr = 100018 OR jobnr = 10365 OR jobnr =10401 OR jobnr =10404 OR jobnr =10406 OR jobnr =10408 OR jobnr =10411 OR jobnr =10413 OR jobnr =10416 OR jobnr =10418 OR jobnr =10430 OR jobnr =10431 OR jobnr =10433 OR jobnr =10434 OR jobnr =10437 OR jobnr =10439 OR jobnr =10446 OR jobnr =10450 OR jobnr =10451 OR jobnr =10452 OR jobnr =10453 OR jobnr =10454 OR jobnr =10455 OR jobnr =10457 OR jobnr =10466 OR jobnr =10467 OR jobnr =10469 OR jobnr =10472 OR jobnr =10474 OR jobnr =10480 OR jobnr =10483 OR jobnr =10488 OR jobnr =10490 OR jobnr =10496 OR jobnr =10498 OR jobnr =10505 OR jobnr =10513 OR jobnr =10521 OR jobnr =10523 OR jobnr =10526 OR jobnr =10527 OR jobnr =10528 OR jobnr =10530 OR jobnr =10537 OR jobnr =10538 OR jobnr =10542 OR jobnr =10549 OR jobnr =10551 OR jobnr =10552 OR jobnr =10553 OR jobnr =10561 OR jobnr =10562 OR jobnr =10563 OR jobnr =10564 OR jobnr =10570 OR jobnr =10571 OR jobnr =10573 OR jobnr =10581 OR jobnr =10584 OR jobnr =10589 OR jobnr =10593 OR jobnr =10594 OR jobnr =10595 OR jobnr =10599 OR jobnr =10603 OR jobnr =10604 OR jobnr =10605 OR jobnr =10607 OR jobnr =10608 OR jobnr =10609 OR jobnr =10610 OR jobnr =10613 OR jobnr =10615 OR jobnr =10616 OR jobnr =10617 OR jobnr =10620 OR jobnr =10621 OR jobnr =10623 OR jobnr =10625 OR jobnr =10627 OR jobnr =10628 OR jobnr =10629 OR jobnr =10630 OR jobnr =10631 OR jobnr =10632 OR jobnr =10633 OR jobnr =10651 OR jobnr =10652 OR jobnr =10654 OR jobnr =10655 OR jobnr =10657 OR jobnr =10658 OR jobnr =10662 OR jobnr =10663 OR jobnr =10664 OR jobnr =10666 OR jobnr =10667 OR jobnr =10669 OR jobnr =10679 OR jobnr =10684 OR jobnr =10686 OR jobnr =10690 OR jobnr =10695 OR jobnr =10697 OR jobnr =10702 OR jobnr =10703 OR jobnr =10704 OR jobnr =10705 OR jobnr =10707 OR jobnr =10708 OR jobnr =10711 OR jobnr =10712 OR jobnr =10715 OR jobnr =10717 OR jobnr =10719 OR jobnr =10720 OR jobnr =10721 OR jobnr =10723 OR jobnr =10725 OR jobnr =10727 OR jobnr =10729 OR jobnr =10730 OR jobnr =10732 OR jobnr =10735 OR jobnr =10736 OR jobnr =10740 OR jobnr =10742 OR jobnr =10743 OR jobnr =10744 OR jobnr =10745 OR jobnr =10746 OR jobnr =10748 OR jobnr =10749 OR jobnr =10750 OR jobnr =10752 OR jobnr =10753 OR jobnr =10756 OR jobnr =10761 OR jobnr =10764 OR jobnr =10778 OR jobnr =10779 OR jobnr =10780 OR jobnr =10781 OR jobnr =10782 OR jobnr =10783 OR jobnr =10786 OR jobnr =10790 OR jobnr =10793 OR jobnr =10794 OR jobnr =10795 OR jobnr =10798 OR jobnr =10799 OR jobnr =10800 OR jobnr =10801 OR jobnr =10803 OR jobnr =10805 OR jobnr =10806 OR jobnr =10807 OR jobnr =10812 OR jobnr =10813 OR jobnr =10815 OR jobnr =10816 OR jobnr =10822 OR jobnr =10823 OR jobnr =10824 OR jobnr =10826 OR jobnr =10827 OR jobnr =10828 OR jobnr =10830 OR jobnr =10831 OR jobnr =10832 OR jobnr =10834 OR jobnr =10838 OR jobnr =10852 OR jobnr =10855 OR jobnr =10875 OR jobnr =10886 OR jobnr =10893 OR jobnr =10897 OR jobnr =10899 OR jobnr =10901 OR jobnr =10909 OR jobnr =10914 OR jobnr =10916 OR jobnr =10918 OR jobnr =10930 OR jobnr =10933 OR jobnr =10939 OR jobnr =10941 OR jobnr =10947 OR jobnr =10950 OR jobnr =10951 OR jobnr =10957 OR jobnr =10958 OR jobnr =10959 OR jobnr =10960 OR jobnr =10963 OR jobnr =10968 OR jobnr =10972 OR jobnr =10973 OR jobnr =10974 OR jobnr =10978 OR jobnr =10997 OR jobnr =10999 OR jobnr =11000 OR jobnr =110000 OR jobnr = 11002 OR jobnr =11003 OR jobnr =11004 OR jobnr =11005 OR jobnr =11006 OR jobnr =11008 OR jobnr =11010 OR jobnr =11011 OR jobnr =11013 OR jobnr =11014 OR jobnr =11015 OR jobnr =11018 OR jobnr =11020 OR jobnr =11024 OR jobnr =11026 OR jobnr =11029 OR jobnr =11030 OR jobnr =11031 OR jobnr =11049 OR jobnr =11055 OR jobnr =11056 OR jobnr =11058 OR jobnr =11074 OR jobnr =11078 OR jobnr =11079 OR jobnr =11086 OR jobnr =11090 OR jobnr =11092 OR jobnr =11097 OR jobnr =11099 OR jobnr =11105 OR jobnr =11107 OR jobnr =11114 OR jobnr =11121 OR jobnr =11122 OR jobnr =11123 OR jobnr =11124 OR jobnr =11125 OR jobnr =11126 OR jobnr =11127 OR jobnr =11129 OR jobnr =12040 OR jobnr =12064 OR jobnr =13012 OR jobnr =13018 OR jobnr =13032 OR jobnr =13037 OR jobnr =13053 OR jobnr =13056 OR jobnr =13058 OR jobnr =13060 OR jobnr =13062 OR jobnr = 13065 OR jobnr =13073 OR jobnr =13080 OR jobnr =13082 OR jobnr =13083 OR jobnr =13086 OR jobnr =13097 OR jobnr =13101 OR jobnr =13109 OR jobnr =13113 OR jobnr =13121 OR jobnr =13147 OR jobnr =13153 OR jobnr =13165 OR jobnr =13174 OR jobnr =13177 OR jobnr = 13179 OR jobnr =13180 OR jobnr =13181 OR jobnr =13189 OR jobnr =13190 OR jobnr =13194 OR jobnr =13198 OR jobnr =13209 OR jobnr =13212 OR jobnr =13213 OR jobnr =13218 OR jobnr =13220 OR jobnr =13222 OR jobnr =13224 OR jobnr =13226 OR jobnr =13230 OR jobnr =13237 OR jobnr =13255 OR jobnr =14036 OR jobnr =14044 OR jobnr =14046 OR jobnr =14051 OR jobnr =14059 OR jobnr =14062 OR jobnr =14063 OR jobnr =14073 OR jobnr =14080 OR jobnr =14082 OR jobnr =14088 OR jobnr =14090 OR jobnr =14164 OR jobnr =14196 OR jobnr =14197 OR jobnr =14202 OR jobnr =14204 OR jobnr =14205 OR jobnr =14209 OR jobnr =14218 OR jobnr =14245 OR jobnr =14246 OR jobnr =14247 OR jobnr =14248 OR jobnr =14249 OR jobnr =14265 OR jobnr =14266 OR jobnr =14269 OR jobnr =14271 OR jobnr =14272 OR jobnr =14276 OR jobnr =14277 OR jobnr =14287 OR jobnr =14288 OR jobnr =14319 OR jobnr =14330 OR jobnr =14345 OR jobnr =14370 OR jobnr =14381 OR jobnr =14388 OR jobnr = 14389 OR jobnr =14394 OR jobnr =14397 OR jobnr =14407 OR jobnr =14430 OR jobnr =14441 OR jobnr =14449 OR jobnr =14451 OR jobnr =14455 OR jobnr =14461 OR jobnr =14479 OR jobnr =14500 OR jobnr =14512 OR jobnr =14526 OR jobnr =14531 OR jobnr =14578 OR jobnr =14580 OR jobnr =14592 OR jobnr =14598 OR jobnr =14600 OR jobnr =14725 OR jobnr =14726 OR jobnr =14727 OR jobnr = 14781 OR jobnr =14785 OR jobnr =14815 OR jobnr =14870 OR jobnr =14872 OR jobnr =14920 OR jobnr =14921 OR jobnr =14924 OR jobnr =14959 OR jobnr =15025 OR jobnr =15090 OR jobnr =15120 OR jobnr =15177 OR jobnr =15180 OR jobnr =15208 OR jobnr =15331 OR jobnr =15394 OR jobnr =15491 OR jobnr =15562 OR jobnr =15744 OR jobnr =15885 OR jobnr =15932 OR jobnr =16129 OR jobnr =16169 OR jobnr =16378 OR jobnr =888 OR jobnr = 16550 OR jobnr =16608 OR jobnr =16763 OR jobnr =16764 OR jobnr =16767 OR jobnr =17097 OR jobnr =17682 OR jobnr = 200002 OR jobnr = 200003 OR jobnr = 200004 OR jobnr = 200009 OR jobnr = 200010 OR jobnr = 200011 OR jobnr = 200013 OR jobnr = 200019 OR jobnr = 200020 OR jobnr = 200021 OR jobnr = 200041 OR jobnr = 300007 OR jobnr = 300012 OR jobnr = 300014 OR jobnr = 300016 OR jobnr = 300019 OR jobnr = 300020 OR jobnr = 400001 OR jobnr = 400002 OR jobnr = 400003 OR jobnr = 400004 OR jobnr = 400006 OR jobnr = 5 OR jobnr =500001 OR jobnr = 500003 OR jobnr = 500004 OR jobnr = 500005 OR jobnr = 500006 OR jobnr = 500007 OR jobnr = 500008 OR jobnr = 600001 OR jobnr = 600002 OR jobnr = 600003 OR jobnr = 600004 OR jobnr = 600005 OR jobnr = 7 OR jobnr =700005 OR jobnr = 700025 OR jobnr = 700117 OR jobnr = 800008 OR jobnr = 800019 OR jobnr = 800022 OR jobnr = 800023 OR jobnr = 800024 OR jobnr = 800107 OR jobnr = 800141 OR jobnr = 800142 OR jobnr = 800143 OR jobnr = 910042)"
      'jobSQLkrinr = "AND (jobnr = 10941 OR jobnr = 10963 OR jobnr = 10933 OR jobnr = 10950 OR jobnr = 10957 OR jobnr = 10916 OR jobnr = 10958 OR jobnr = 10960 OR jobnr = 10930 OR jobnr = 10918 OR jobnr = 11003 OR jobnr = 10951 OR jobnr = 10973 OR jobnr = 10947 OR jobnr = 10939 OR jobnr = 11008 OR jobnr = 10999 OR jobnr = 10972 OR jobnr = 11020 OR jobnr = 11011 OR jobnr = 11026 OR jobnr = 11024 OR jobnr = 11004 OR jobnr = 11005 OR jobnr = 11010 OR jobnr = 11014 OR jobnr = 11029 OR jobnr = 11018 OR jobnr = 11030 OR jobnr = 11031 OR jobnr = 11000 OR jobnr = 10968 OR jobnr = 11055 OR jobnr = 11092 OR jobnr = 11090 OR jobnr = 11126 OR jobnr = 11123 OR jobnr = 11125 OR jobnr = 11129 OR jobnr = 11127 OR jobnr = 11124 OR jobnr = 11122 OR jobnr = 11006 OR jobnr = 11079 OR jobnr = 11074 OR jobnr = 11107 OR jobnr = 11056 OR jobnr = 11058 OR jobnr = 11105 OR jobnr = 10974 OR jobnr = 10959 OR jobnr = 11015 OR jobnr = 11078 OR jobnr = 11097 OR jobnr = 11114 OR jobnr = 10978 OR jobnr = 11013 OR jobnr = 11086 OR jobnr = 11099 OR jobnr = 11049 OR jobnr = 13012 OR jobnr = 13018 OR jobnr = 13058 OR jobnr = 13056 OR jobnr = 13065 OR jobnr = 13060 OR jobnr = 13053 OR jobnr = 13037 OR jobnr = 13086 OR jobnr = 13082 OR jobnr = 13083 OR jobnr = 13032 OR jobnr = 13097 OR jobnr = 13109 OR jobnr = 13113 OR jobnr = 13121 OR jobnr = 13165 OR jobnr = 13147 OR jobnr = 10997 OR jobnr = 13218 OR jobnr = 12064 OR jobnr = 13213 OR jobnr = 13224 OR jobnr = 13080 OR jobnr = 11121 OR jobnr = 13209 OR jobnr = 13179 OR jobnr = 13180 OR jobnr = 13226 OR jobnr = 13198 OR jobnr = 13181 OR jobnr = 13230 OR jobnr = 12040 OR jobnr = 13220 OR jobnr = 13194 OR jobnr = 13189 OR jobnr = 14090 OR jobnr = 13190 OR jobnr = 14036 OR jobnr = 14073 OR jobnr = 13174 OR jobnr = 14051 OR jobnr = 13255 OR jobnr = 14059 OR jobnr = 14080 OR jobnr = 13101 OR jobnr = 14046 OR jobnr = 11002 OR jobnr = 13222 OR jobnr = 13177 OR jobnr = 13062 OR jobnr = 14082 OR jobnr = 14063 OR jobnr = 14062 OR jobnr = 13153 OR jobnr = 14088 OR jobnr = 14044 OR jobnr = 13212 OR jobnr = 13073 OR jobnr = 13237 OR jobnr = 14164 OR jobnr = 14205 OR jobnr = 14204 OR jobnr = 14202 OR jobnr = 14218 OR jobnr = 14209 OR jobnr = 14197 OR jobnr = 14196 OR jobnr = 14247 OR jobnr = 14245 OR jobnr = 14248 OR jobnr = 14246 OR jobnr = 14319 OR jobnr = 14249 OR jobnr = 14277 OR jobnr = 14271 OR jobnr = 14265 OR jobnr = 14288 OR jobnr = 14266 OR jobnr = 14272 OR jobnr = 14269 OR jobnr = 14287 OR jobnr = 14330 OR jobnr = 14345 OR jobnr = 14388 OR jobnr = 14389 OR jobnr = 14407 OR jobnr = 14394 OR jobnr = 14397 OR jobnr = 14381 OR jobnr = 14430 OR jobnr = 14441 OR jobnr = 14455 OR jobnr = 14461 OR jobnr = 14451 OR jobnr = 14449 OR jobnr = 14526 OR jobnr = 14531 OR jobnr = 14512 OR jobnr = 14580 OR jobnr = 400003 OR jobnr = 600001 OR jobnr = 300007 OR jobnr = 110000 OR jobnr = 100001 OR jobnr = 200002 OR jobnr = 300014 OR jobnr = 100005 OR jobnr = 600004 OR jobnr = 100003 OR jobnr = 100002 OR jobnr = 600002 OR jobnr = 600003 OR jobnr = 100000 OR jobnr = 200003 OR jobnr = 100004 OR jobnr = 500001 OR jobnr = 300012 OR jobnr = 400002 OR jobnr = 400001 OR jobnr = 600005 OR jobnr = 400004 OR jobnr = 14592 OR jobnr = 200004 OR jobnr = 14781 OR jobnr = 14785 OR jobnr = 400006 OR jobnr = 14600 OR jobnr = 200009 OR jobnr = 100007 OR jobnr = 300016 OR jobnr = 14921 OR jobnr = 14920 OR jobnr = 14924 OR jobnr = 200010 OR jobnr = 15025 OR jobnr = 14370 OR jobnr = 14815 OR jobnr = 14598 OR jobnr = 14959 OR jobnr = 14479 OR jobnr = 15090 OR jobnr = 14870 OR jobnr = 14727 OR jobnr = 14872 OR jobnr = 300019 OR jobnr = 300020 OR jobnr = 14276 OR jobnr = 14726 OR jobnr = 14725 OR jobnr = 200013 OR jobnr = 15120 OR jobnr = 14578 OR jobnr = 200011 OR jobnr = 15180 OR jobnr = 15208 OR jobnr = 15177 OR jobnr = 14500 OR jobnr = 500003 OR jobnr = 500005 OR jobnr = 500004 OR jobnr = 500006 OR jobnr = 500007 OR jobnr = 15331 OR jobnr = 15394 OR jobnr = 700005 OR jobnr = 15491 OR jobnr = 100011 OR jobnr = 500008 OR jobnr = 100012 OR jobnr = 800008 OR jobnr = 15562 OR jobnr = 200020 OR jobnr = 200019 OR jobnr = 100013 OR jobnr = 100014 OR jobnr = 800022 OR jobnr = 800019 OR jobnr = 800023 OR jobnr = 800024 OR jobnr = 200021 OR jobnr = 15885 OR jobnr = 700025 OR jobnr = 15932 OR jobnr = 16169 OR jobnr = 16129 OR jobnr = 16378 OR jobnr = 16608 OR jobnr = 16550 OR jobnr = 16763 OR jobnr = 16767 OR jobnr = 16764 OR jobnr = 100018 OR jobnr = 910042 OR jobnr = 800107 OR jobnr = 17097 OR jobnr = 15744 OR jobnr = 200041 OR jobnr = 800142 OR jobnr = 800143 OR jobnr = 800141 OR jobnr = 17682 OR jobnr = 700117)"
      'jobSQLkrinr = "AND (jobnr = 10892 OR jobnr = 11022 OR jobnr = 11039 OR jobnr = 11155 OR jobnr = 12015 OR jobnr = 12064 OR jobnr = 13029 OR jobnr = 13037 OR jobnr = 13074 OR jobnr = 13078 OR jobnr = 13108)"  
       'jobSQLkrinr = "AND (jobnr = 11155 OR jobnr = 12064)"   
         'jobSQLkrinr = "AND (jobnr = 17149 OR jobnr = 17341 OR jobnr = 17342 OR jobnr = 17355 OR jobnr = 17370 OR jobnr = 17373 OR jobnr = 17381 OR jobnr = 17292 OR jobnr = 17468 OR jobnr = 17484 OR jobnr = 17494 OR jobnr = 17280)"   

        'jobSQLkrinr = "AND (jobnr = 10892 OR jobnr = 13078 OR jobnr = 13074 OR jobnr = 12064 OR jobnr = 14477 OR jobnr = 14688 OR jobnr = 14231 OR jobnr = 14754 OR jobnr = 14878 OR jobnr = 14937 OR jobnr = 15007 OR jobnr = 15053 OR jobnr = 15044 OR jobnr = 15008 OR jobnr = 15113 OR jobnr = 15125 OR jobnr = 15139 OR jobnr = 14486 OR jobnr = 14600 OR jobnr = 15149 OR jobnr = 15180 OR jobnr = 14500 OR jobnr = 15238 OR jobnr = 15282 OR jobnr = 15300 OR jobnr = 15318 OR jobnr = 15343 OR jobnr = 14949 OR jobnr = 15199 OR jobnr = 15351 OR jobnr = 15484 OR jobnr = 15416 OR jobnr = 15480 OR jobnr = 15524 OR jobnr = 15227 OR jobnr = 15568 OR jobnr = 15187 OR jobnr = 15580 OR jobnr = 15594 OR jobnr = 15607 OR jobnr = 15612 OR jobnr = 15610 OR jobnr = 15088 OR jobnr = 15687 OR jobnr = 15481 OR jobnr = 15637 OR jobnr = 15392 OR jobnr = 15658 OR jobnr = 15694 OR jobnr = 15699 OR jobnr = 11039 OR jobnr = 15208 OR jobnr = 15394 OR jobnr = 15743 OR jobnr = 15741 OR jobnr = 15750 OR jobnr = 14300 OR jobnr = 15952 OR jobnr = 15814 OR jobnr = 15004 OR jobnr = 15844 OR jobnr = 15834 OR jobnr = 15861 OR jobnr = 15938 OR jobnr = 15873 OR jobnr = 15880 OR jobnr = 15959 OR jobnr = 15896 OR jobnr = 15904 OR jobnr = 15975 OR jobnr = 15969 OR jobnr = 15493 OR jobnr = 15986 OR jobnr = 15987 OR jobnr = 15995)"
         
        'jobSQLkrinr = "AND (jobnr = 16659 OR jobnr = 16667 OR jobnr = 16592 OR jobnr = 16554 OR jobnr = 16724 OR jobnr = 16752 OR jobnr = 16757 OR jobnr = 16800 OR jobnr = 16788 OR jobnr = 16623 "_
        '&" OR jobnr = 16766 OR jobnr = 16771 OR jobnr = 16817 OR jobnr = 16825 OR jobnr = 16861 OR jobnr = 16915 OR jobnr = 16919 OR jobnr = 16924 OR jobnr = 16940 OR jobnr = 16952 OR jobnr = 16961)"

        'jobnr = 17485 OR jobnr = 11039 OR jobnr = 16292 OR jobnr = 16273 OR jobnr = 16281 OR jobnr = 15299 OR jobnr = 16288 OR jobnr = 15799 OR jobnr = 16261 OR jobnr = 16282 OR jobnr = 15261 OR jobnr = 14387 OR jobnr = 17336 OR jobnr = 16277 OR jobnr = 14860 OR jobnr = 15235 OR jobnr = 16266 OR jobnr = 15964 OR jobnr = 16289 OR jobnr = 15715 OR jobnr = 16386 OR jobnr = 15534 OR jobnr = 12015 OR jobnr = 15970 OR jobnr = 15574 OR jobnr = 16267 OR jobnr = 15462 OR jobnr = 17280 OR jobnr = 16284 OR jobnr = 16259 OR jobnr = 17484 OR jobnr = 16279 OR jobnr = 16919 OR jobnr = 15483 OR jobnr = 16214 OR jobnr = 14981 OR jobnr = 15515 OR jobnr = 15173 OR jobnr = 15649 OR jobnr = 15274 OR jobnr = 15082 OR jobnr = 16554 OR jobnr = 14890 OR jobnr = 15823 OR jobnr = 17355 OR jobnr = 16294 OR jobnr = 15395 OR jobnr = 16275 OR
        'jobSQLkrinr = "AND (jobnr = 15125 OR jobnr = 16061 OR jobnr = 14617 OR jobnr = 17530 OR jobnr = 17483 OR jobnr = 15330 OR jobnr = 16567 OR jobnr = 15605 OR jobnr = 14911 OR jobnr = 15079 OR jobnr = 14596 OR jobnr = 16262 OR jobnr = 16757 OR jobnr = 16861 OR jobnr = 15580 OR jobnr = 16531 OR jobnr = 16724 OR jobnr = 17149 OR jobnr = 16118 OR jobnr = 16409 OR jobnr = 16258 OR jobnr = 16393 OR jobnr = 16592 OR jobnr = 16800 OR jobnr = 16396 OR jobnr = 17370 OR jobnr = 16924 OR jobnr = 17341 OR jobnr = 17342 OR jobnr = 16168 OR jobnr = 16210)"


         'jobSQLkrinr = "AND (jobnr = 17001 OR jobnr = 17280 OR jobnr = 17254 OR jobnr = 17292 OR jobnr = 17235 OR jobnr = 17295 OR jobnr = 17284 OR jobnr = 17149 OR jobnr = 17100 OR jobnr = 17341 OR jobnr = 17342 OR jobnr = 17355 OR jobnr = 17370 OR jobnr = 17373 OR jobnr = 17381 OR jobnr = 17556 OR jobnr = 17468 OR jobnr = 17484 OR jobnr = 17518 OR jobnr = 17528 OR jobnr = 17494 OR jobnr = 17590)"

        'jobSQLkrinr = "AND (jobnr = 17201 OR jobnr = 16503 OR jobnr = 16507 OR jobnr = 16535 OR jobnr = 16550 OR jobnr = 16558 OR jobnr = 16612 OR jobnr = 16620 OR jobnr = 16647 OR jobnr = 16685 OR jobnr = 16716 OR jobnr = 16727 OR jobnr = 16774 OR jobnr = 16779 OR jobnr = 16816 OR jobnr = 16883 OR jobnr = 16882 OR jobnr = 16909 OR jobnr = 16938 OR jobnr = 16944 OR jobnr = 16953 OR jobnr = 16955 OR jobnr = 16960 OR jobnr = 16967 OR jobnr = 16968 OR jobnr = 16969 OR jobnr = 17035 OR jobnr = 17060 OR jobnr = 17067 OR jobnr = 17232 OR jobnr = 17134 OR jobnr = 17135 OR jobnr = 17139 OR jobnr = 17147 OR jobnr = 17213 OR jobnr = 17185 OR jobnr = 17186 OR jobnr = 17187 OR jobnr = 17188 OR jobnr = 17189 OR jobnr = 17190 OR jobnr = 17191 OR jobnr = 17200 OR jobnr = 17203 OR jobnr = 17204 OR jobnr = 17207 OR jobnr = 17214 OR jobnr = 17216 OR jobnr = 17221 OR jobnr = 17231 OR jobnr = 17240 OR jobnr = 17253 OR jobnr = 17261 OR jobnr = 17283 OR jobnr = 17297 OR jobnr = 17301 OR jobnr = 17303 OR jobnr = 17306 OR jobnr = 17338 OR jobnr = 17339 OR jobnr = 17352 OR jobnr = 17363 OR jobnr = 17397 OR jobnr = 17401 OR jobnr = 17406 OR jobnr = 17407 OR jobnr = 17411 OR jobnr = 17422 OR jobnr = 17423 OR jobnr = 17425 OR jobnr = 17428 OR jobnr = 17438 OR jobnr = 17446 OR jobnr = 17448 OR jobnr = 17475 OR jobnr = 17481 OR jobnr = 17484 OR jobnr = 17485)"

        '14500 - 14999
        'jobSQLkrinr = "AND (jobnr = 14616 OR jobnr = 14928 OR jobnr = 14939 OR jobnr = 14573 OR jobnr = 14623 OR jobnr = 14534 OR jobnr = 14568 OR jobnr = 14504 OR jobnr = 14500 OR jobnr = 14502 OR jobnr = 14505 OR jobnr = 14514 OR jobnr = 14510 OR jobnr = 14509 OR jobnr = 14516 OR jobnr = 14518 OR jobnr = 14520 OR jobnr = 14521 OR jobnr = 14522 OR jobnr = 14523 OR jobnr = 14524 OR jobnr = 14525 OR jobnr = 14527 OR jobnr = 14528 OR jobnr = 14529 OR jobnr = 14530 OR jobnr = 14533 OR jobnr = 14539 OR jobnr = 14537 OR jobnr = 14538 OR jobnr = 14540 OR jobnr = 14541 OR jobnr = 14542 OR jobnr = 14544 OR jobnr = 14545 OR jobnr = 14547 OR jobnr = 14548 OR jobnr = 14549 OR jobnr = 14550 OR jobnr = 14551 OR jobnr = 14552 OR jobnr = 14556 OR jobnr = 14557 OR jobnr = 14558 OR jobnr = 14559 OR jobnr = 14560 OR jobnr = 14561 OR jobnr = 14563 OR jobnr = 14565 OR jobnr = 14566 OR jobnr = 14585 OR jobnr = 14591 OR jobnr = 14567 OR jobnr = 14663 OR jobnr = 14571 OR jobnr = 14578 OR jobnr = 14579 OR jobnr = 14582 OR jobnr = 14583 OR jobnr = 14586 OR jobnr = 14587 OR jobnr = 14589 OR jobnr = 14782 OR jobnr = 14593 OR jobnr = 14594 OR jobnr = 14595 OR jobnr = 14597 OR jobnr = 14599 OR jobnr = 14602 OR jobnr = 14603 OR jobnr = 14604 OR jobnr = 14605 OR jobnr = 14606 OR jobnr = 14609 OR jobnr = 14611 OR jobnr = 14613 OR jobnr = 14612 OR jobnr = 14614 OR jobnr = 14619 OR jobnr = 14620 OR jobnr = 14622 OR jobnr = 14624 OR jobnr = 14625 OR jobnr = 14626 OR jobnr = 14627 OR jobnr = 14628 OR jobnr = 14629 OR jobnr = 14630 OR jobnr = 14631 OR jobnr = 14633 OR jobnr = 14657 OR jobnr = 14658 OR jobnr = 14660 OR jobnr = 14661 OR jobnr = 14662 OR jobnr = 14664 OR jobnr = 14665 OR jobnr = 14666 OR jobnr = 14668 OR jobnr = 14681 OR jobnr = 14685 OR jobnr = 14686 OR jobnr = 14687 OR jobnr = 14691 OR jobnr = 14692 OR jobnr = 14693 OR jobnr = 14694 OR jobnr = 14695 OR jobnr = 14696 OR jobnr = 14697 OR jobnr = 14698 OR jobnr = 14700 OR jobnr = 14713 OR jobnr = 14830 OR jobnr = 14710 OR jobnr = 14716 OR jobnr = 14712 OR jobnr = 14714 OR jobnr = 14717 OR jobnr = 14721 OR jobnr = 14722 OR jobnr = 14723 OR jobnr = 14725 OR jobnr = 14726 OR jobnr = 14727 OR jobnr = 14728 OR jobnr = 14729 OR jobnr = 14731 OR jobnr = 14733 OR jobnr = 14734 OR jobnr = 14735 OR jobnr = 14737 OR jobnr = 14738 OR jobnr = 14739 OR jobnr = 14740 OR jobnr = 14741 OR jobnr = 14742 OR jobnr = 14743 OR jobnr = 14744 OR jobnr = 14745 OR jobnr = 14749 OR jobnr = 14750 OR jobnr = 14751 OR jobnr = 14752 OR jobnr = 14753 OR jobnr = 14756 OR jobnr = 14757 OR jobnr = 14760 OR jobnr = 14761 OR jobnr = 14763 OR jobnr = 14764 OR jobnr = 14770 OR jobnr = 14772 OR jobnr = 14774 OR jobnr = 14775 OR jobnr = 14777 OR jobnr = 14778 OR jobnr = 14779 OR jobnr = 14790 OR jobnr = 14784 OR jobnr = 14788 OR jobnr = 14789 OR jobnr = 14791 OR jobnr = 14796 OR jobnr = 14814 OR jobnr = 14813 OR jobnr = 14797 OR jobnr = 14798 OR jobnr = 14799 OR jobnr = 14800 OR jobnr = 14801 OR jobnr = 14802 OR jobnr = 14803 OR jobnr = 14805 OR jobnr = 14806 OR jobnr = 14807 OR jobnr = 14808 OR jobnr = 14810 OR jobnr = 14812 OR jobnr = 14815 OR jobnr = 14820 OR jobnr = 14821 OR jobnr = 14822 OR jobnr = 14823 OR jobnr = 14824 OR jobnr = 14827 OR jobnr = 14828 OR jobnr = 14829 OR jobnr = 14831 OR jobnr = 14832 OR jobnr = 14833 OR jobnr = 14869 OR jobnr = 14835 OR jobnr = 14836 OR jobnr = 14837 OR jobnr = 14839 OR jobnr = 14840 OR jobnr = 14841 OR jobnr = 14843 OR jobnr = 14844 OR jobnr = 14848 OR jobnr = 14849 OR jobnr = 14851 OR jobnr = 14852 OR jobnr = 14853 OR jobnr = 14854 OR jobnr = 14855 OR jobnr = 14857 OR jobnr = 14871 OR jobnr = 14863 OR jobnr = 14868 OR jobnr = 14867 OR jobnr = 14872 OR jobnr = 14874 OR jobnr = 14875 OR jobnr = 14879 OR jobnr = 14881 OR jobnr = 14882 OR jobnr = 14883 OR jobnr = 14884 OR jobnr = 14887 OR jobnr = 14891 OR jobnr = 14892 OR jobnr = 14894 OR jobnr = 14895 OR jobnr = 14896 OR jobnr = 14918 OR jobnr = 14898 OR jobnr = 14900 OR jobnr = 14901 OR jobnr = 14902 OR jobnr = 14904 OR jobnr = 14905 OR jobnr = 14912 OR jobnr = 14914 OR jobnr = 14927 OR jobnr = 14925 OR jobnr = 14936 OR jobnr = 14995 OR jobnr = 14948 OR jobnr = 14952 OR jobnr = 14955 OR jobnr = 14957 OR jobnr = 14960 OR jobnr = 14962 OR jobnr = 14964 OR jobnr = 14966 OR jobnr = 14967 OR jobnr = 14969 OR jobnr = 14970 OR jobnr = 14972 OR jobnr = 14973 OR jobnr = 14997 OR jobnr = 14996 OR jobnr = 14576)" 

        '15000 - 15299
        'jobSQLkrinr = "AND (jobnr = 15215 OR jobnr = 15184 OR jobnr = 15068 OR jobnr = 15104 OR jobnr = 15061 OR jobnr = 15047 OR jobnr = 15049 OR jobnr = 15050 OR jobnr = 15051 OR jobnr = 15057 OR jobnr = 15058 OR jobnr = 15063 OR jobnr = 15067 OR jobnr = 15081 OR jobnr = 15071 OR jobnr = 15073 OR jobnr = 15074 OR jobnr = 15075 OR jobnr = 15076 OR jobnr = 15077 OR jobnr = 15083 OR jobnr = 15111 OR jobnr = 15084 OR jobnr = 15086 OR jobnr = 15092 OR jobnr = 15273 OR jobnr = 15298 OR jobnr = 15132 OR jobnr = 15091 OR jobnr = 15094 OR jobnr = 15097 OR jobnr = 15099 OR jobnr = 15107 OR jobnr = 15112 OR jobnr = 15115 OR jobnr = 15172 OR jobnr = 15123 OR jobnr = 15117 OR jobnr = 15118 OR jobnr = 15121 OR jobnr = 15128 OR jobnr = 15141 OR jobnr = 15127 OR jobnr = 15161 OR jobnr = 15133 OR jobnr = 15136 OR jobnr = 15137 OR jobnr = 15143 OR jobnr = 15145 OR jobnr = 15147 OR jobnr = 15150 OR jobnr = 15153 OR jobnr = 15154 OR jobnr = 15156 OR jobnr = 15157 OR jobnr = 15158 OR jobnr = 15159 OR jobnr = 15165 OR jobnr = 15166 OR jobnr = 15167 OR jobnr = 15169 OR jobnr = 15175 OR jobnr = 15178 OR jobnr = 15179 OR jobnr = 15181 OR jobnr = 15182 OR jobnr = 15185 OR jobnr = 15191 OR jobnr = 15192 OR jobnr = 15193 OR jobnr = 15195 OR jobnr = 15196 OR jobnr = 15198 OR jobnr = 15200 OR jobnr = 15201 OR jobnr = 15204 OR jobnr = 15205 OR jobnr = 15207 OR jobnr = 15209 OR jobnr = 15210 OR jobnr = 15219 OR jobnr = 15212 OR jobnr = 15214 OR jobnr = 15217 OR jobnr = 15221 OR jobnr = 15223 OR jobnr = 15226 OR jobnr = 15229 OR jobnr = 15232 OR jobnr = 15236 OR jobnr = 15239 OR jobnr = 15240 OR jobnr = 15245 OR jobnr = 15255 OR jobnr = 15260 OR jobnr = 15270 OR jobnr = 15271 OR jobnr = 15272 OR jobnr = 15281 OR jobnr = 15283 OR jobnr = 15284 OR jobnr = 15285 OR jobnr = 15287 OR jobnr = 15293 OR jobnr = 15288 OR jobnr = 15289 OR jobnr = 15290 OR jobnr = 15294 OR jobnr = 15295 OR jobnr = 15296 OR jobnr = 15297)"


           'jobSQLkrinr = jobSQLkrinr &" AND (jobnr = 15180 OR jobnr = 15236 OR jobnr = 15283 OR jobnr = 15384 OR jobnr = 15658 OR jobnr = 15744 OR jobnr = 15896 OR jobnr = 15991 OR jobnr = 15992 OR jobnr = 15995 OR jobnr = 16017 OR jobnr = 16089 OR jobnr = 16121 OR jobnr = 16175 OR jobnr = 16292 OR jobnr = 16304 OR jobnr = 16376 OR jobnr = 16487 OR jobnr = 16509 OR jobnr = 16537 OR jobnr = 16558 OR jobnr = 16590 OR jobnr = 16597 OR jobnr = 16620 OR jobnr = 16638 OR jobnr = 16792 OR jobnr = 16794 OR jobnr = 16904 OR jobnr = 16909 OR jobnr = 16938 OR jobnr = 17028 OR jobnr = 17038 OR jobnr = 17080 OR jobnr = 17107 OR jobnr = 17200 OR jobnr = 17213 OR jobnr = 17216 OR jobnr = 17253 OR jobnr = 17301 OR jobnr = 17393 OR jobnr = 17438 OR jobnr = 17441 OR jobnr = 17481 OR jobnr = 17515 OR jobnr = 17547 OR jobnr = 17549 OR jobnr = 17564 OR jobnr = 17580 OR jobnr = 17678 OR jobnr = 17720 OR jobnr = 17725 OR jobnr = 17746 OR jobnr = 17753 OR jobnr = 17757 OR jobnr = 17761 OR jobnr = 17769 OR jobnr = 17770 OR jobnr = 17777 OR jobnr = 17779 OR jobnr = 17792 OR jobnr = 17793 OR jobnr = 17853 OR jobnr = 17875 OR jobnr = 17876)"

            jobSQLkrinr = jobSQLkrinr &" AND (jobnr = 17438 OR jobnr = 17515 OR jobnr = 17553 OR jobnr = 17568 OR jobnr = 17627 OR jobnr = 17643 OR jobnr = 17646 OR jobnr = 17677 OR jobnr = 17717 OR jobnr = 17725 OR jobnr = 17728 OR jobnr = 17732 OR jobnr = 17733 OR jobnr = 17746 OR jobnr = 17754 OR jobnr = 17755 OR jobnr = 17757 OR jobnr = 17760 OR jobnr = 17761 OR jobnr = 17769 OR jobnr = 17770 OR jobnr = 17779 OR jobnr = 17785 OR jobnr = 17789 OR jobnr = 17792 OR jobnr = 17843 OR jobnr = 17849 OR jobnr = 17853 OR jobnr = 17872 OR jobnr = 17875 OR jobnr = 17876 OR jobnr = 17906 OR jobnr = 17922 OR jobnr = 17923 OR jobnr = 17930 OR jobnr = 17941 OR jobnr = 17958 OR jobnr = 17964 OR jobnr = 17967 OR jobnr = 17971 OR jobnr = 17973 OR jobnr = 17975 OR jobnr = 17993 OR jobnr = 17998 OR jobnr = 18060 OR jobnr = 18064 OR jobnr = 18076 OR jobnr = 18084 OR jobnr = 18090 OR jobnr = 18095 OR jobnr = 18103 OR jobnr = 18107 OR jobnr = 18111 OR jobnr = 18135 OR jobnr = 18137 OR jobnr = 18140 OR jobnr = 18145 OR jobnr = 18150 OR jobnr = 18152 OR jobnr = 18155 OR jobnr = 18160 OR jobnr = 18168 OR jobnr = 18169 OR jobnr = 18171 OR jobnr = 18172 OR jobnr = 18174 OR jobnr = 18183 OR jobnr = 18204 OR jobnr = 18205 OR jobnr = 18213 OR jobnr = 18217 OR jobnr = 18218 OR jobnr = 18224 OR jobnr = 18228 OR jobnr = 18240 OR jobnr = 18241 OR jobnr = 18248 OR jobnr = 18249 OR jobnr = 18250 OR jobnr = 18262 OR jobnr = 18301 OR jobnr = 18307 OR jobnr = 18309 OR jobnr = 18317 OR jobnr = 18318 OR jobnr = 18321 OR jobnr = 18332 OR jobnr = 18340 OR jobnr = 18398 OR jobnr = 18403 OR jobnr = 18404 OR jobnr = 18405 OR jobnr = 18469 OR jobnr = 18476 OR jobnr = 18489)"
    
    
           jobSQLkri = "AND ( id = 0 " ''Overrruler
           strSQLjoids = "SELECT id FROM job WHERE id <> 0 "& jobSQLkrinr & " ORDER BY jobnr"
            oRec.open strSQLjoids, oConn, 3
            while not oRec.EOF 
            
            jobSQLkri = jobSQLkri & " OR id = " & oRec("id")

            oRec.movenext
            wend
            oRec.close
            jobSQLkri = jobSQLkri & ")"

           bdgmtypon_val = 1 'overrruler
          
      '*******************************************

     end if



 
bdgmtypon_val = 1 'overrruler medarbbydget slået til

'sqlDatostart = year(sqlDatostart) &"/"& month(sqlDatostart) &"/"& day(sqlDatostart)

sqlDatoslut = year(sqlDatoslut) &"/"& month(sqlDatoslut) &"/"& day(sqlDatoslut)
dddato = now

Response.write "Dato: "& dddato & "<br><br>"

'** Første gang eller 12 md bagud ***'
'** Bruges IKKE job startdato bruges - undtagleser
'if cint(first) <> 1 AND first <> 2 then
'dddato_start = dateAdd("m", -6, dddato) '6 md
'dddato_start = "1/"& month(dddato_start) &"/"& year(dddato_start)
'else
dddato_start = "01-06-2010" 'li startdato BASIS bare for at være sikker på der er en dato
'end if

dddatoSQL = year(dddato)&"/"&month(dddato)&"/"&day(dddato)


    'Response.write "dddato:"& dddato &"<br>bdgmtypon_val:; "& bdgmtypon_val & "<br>"
    'Response.flush

if cint(bdgmtypon_val) = 1 then 
            
            'select case dothis
            'case 1 
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id <> 0 AND (jobstatus = 1 OR jobstatus = 0 OR jobstatus = 2 OR jobstatus = 4) "& jobSQLkri &" ORDER BY id"
            'case 2
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id > 0 AND jobstatus = 0 AND jobstartdato BETWEEN '2014-01-01' AND '2014-12-31' "& jobSQLkri &" ORDER BY id" 
            'case 3
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id > 4170 AND jobstatus = 0 AND jobstartdato BETWEEN '2013-01-01' AND '2013-12-31' "& jobSQLkri &" ORDER BY id LIMIT 50" 
            'case 4
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id > 3428 AND jobstatus = 0 AND jobstartdato BETWEEN '2012-01-01' AND '2012-12-31' "& jobSQLkri &" ORDER BY id LIMIT 60" 
            'case 5
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id > 2326 AND jobstatus = 0 AND jobstartdato BETWEEN '2011-01-01' AND '2011-12-31' "& jobSQLkri &" ORDER BY id LIMIT 60"
            'case 6
            'strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id > 1438 AND jobstatus = 0 AND jobstartdato BETWEEN '2010-01-01' AND '2010-12-31' "& jobSQLkri &" ORDER BY id"
            'end select
            
                'if first = 1 then
                'jobSQLkri = "AND (jobnr BETWEEN 15996 AND 15999)"
                'end if

             strSQLj = "SELECT id, jobnavn, jobnr, jobstartdato FROM job WHERE id <> 0 AND (jobstatus = 1 OR jobstatus = 0 OR jobstatus = 2 OR jobstatus = 4) "& jobSQLkri &" ORDER BY id LIMIT 200"
            'AND (jobnr > 14632)
            'AND (jobnr BETWEEN 15374 AND 15399)
            'AND (jobnr BETWEEN 15487 AND 15499)
            'AND (jobnr BETWEEN 15577 AND 15599)
            'AND id > 4777
            'AND (jobnr BETWEEN 15884 AND 16800)
        'AND (jobnr BETWEEN 17600 AND 17800)




            lastJid = 0
            j = 0
            Response.write "<br><br><br><b>Henter job:<br> "& strSQLj & "</b><br>"
            Response.flush

            'if cint(first) = 0 then
            'response.end
            'end if

            oRec.open strSQLj, oConn, 3
            while not oRec.EOF 
            j = j + 1
            
            
            if cint(first) = 2 OR cint(first) = 1 then '2: specifikt job = SLETTER ALT eller 1: første kons på alle
            'Nulstiller interval ved næste job    
            if lastJid <> oRec("id") then
               
                strSQLdelKons = "DELETE FROM timer_konsolideret_tot WHERE jobid = "& oRec("id")
                response.write "<br>Tømmer: "& strSQLdelKons & "<hr>"
                oConn.execute(strSQLdelKons)

            end if
            end if



            m = 0

            Response.write oRec("id") &" - "& oRec("jobnavn") & " ("& oRec("jobnr") &")<br><br>"

            strAktpajobKri = "taktivitetid = 0"

            strSQLaktpajob = "SELECT id, navn, tid FROM aktiviteter LEFT JOIN timer ON (taktivitetid = id) WHERE job = " & oRec("id") &" AND ("& replace(aty_sql_realhours, "tfaktim", "fakturerbar") & ") GROUP BY id" 
          
            'if session("mid") = 1 then
            'response.write strSQLaktpajob & "<br>"
            'end if
           
            oRec6.open strSQLaktpajob, oConn, 3
            while not oRec6.EOF 

            if isNull(oRec6("tid")) <> true then
            strAktpajobKri = strAktpajobKri & " OR taktivitetid = " & oRec6("id")
            end if

            oRec6.movenext
            wend
            oRec6.close
            
            'response.write "strAktpajobKri: "& strAktpajobKri
            'response.write "<hr>"


            strAktpajobKri = replace(strAktpajobKri, "taktivitetid = 0 OR", "")
            
            '** Brug jobstdato
            '** JobstDao - måneder
            dddato_start = dateAdd("m", -3, dddato) '6 md Standard hver nat

            if cint(first) = 2 then 'specifikt job
            dddato_start = dateAdd("m", -12, oRec("jobstartdato")) 
            end if
            
            if cint(first) = 1 then 'alle job første gang
            dddato_start = dateAdd("m", -2, oRec("jobstartdato")) 
            end if        

            dddato_start = year(dddato_start) &"/"& month(dddato_start) &"/1"
            dddato_start = dddato_start '"1/9/2007"    '2010 job

           '*** Første gang  / Ellers 12 md bagud
            periodeInterval = dateDiff("m", dddato_start, dddato, 2,2)
            
            tb_ext = "" '"_2014_1" '"" '"_mt"

            for m = 0 TO periodeInterval
            
            for t = 1 to UBOUND(mtypgrpids)


           
            

            thisMtypTimerKost = 0
            thisMtypTimerBelob = 0
            timerReal = 0
            if len(trim(mtypnavn(t))) <> 0 AND strAktpajobKri <> "taktivitetid = 0" then

         
            
            '*** Hvis der kun skal kigges tilbgae til jobstartsato **
            '** Mindre datediff, hurtigere load
            'sqlDatostart = year(oRec("jobstartdato")) &"/"& month(oRec("jobstartdato")) &"/"& day(oRec("jobstartdato"))

            sqlDatostart = dateAdd("m", m, dddato_start)        
            sqlDatoslut = dateAdd("m", 1, sqlDatostart)
            sqlDatoslut = dateAdd("d", -1, sqlDatoslut)        

            sqlDatostart = year(sqlDatostart) &"/"& month(sqlDatostart) &"/1"
            sqlDatoslut = year(sqlDatoslut) & "/"& month(sqlDatoslut) & "/"& day(sqlDatoslut)


           'Response.write "strAktpajobKri: " & strAktpajobKri & "<br>"

            strSQLmtypbgt = "SELECT SUM(timer*timepris*(kurs/100)) AS mtyptimerbelob, SUM(timer*kostpris*(kurs/100)) AS mtyptimerkost, SUM(timer) AS timerreal FROM timer WHERE tjobnr = "& oRec("jobnr") &" AND ("& medarbimType(t) &") AND ("& strAktpajobKri &") "

             'if cint(realfakpertot) <> 0 then 
             strSQLmtypbgt = strSQLmtypbgt & " AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
             'else
            'strSQLmtypbgt = strSQLmtypbgt & " AND tdato BETWEEN '2001-01-01' AND '2044-01-01'"
             'end if
           
            strSQLmtypbgt = strSQLmtypbgt & " GROUP BY tjobnr"

            
                
               'if session("mid") = 1 then
               'Response.write " <br># "& strSQLmtypbgt  &" #<br><br> " 
                if first <> 2 then
                Response.write "SELECT TIMER: "& sqlDatostart &"' AND '"& sqlDatoslut &"' Jobnr: " & oRec("jobnr") & "<br><br>"
                else
                Response.write "."
                end if
                Response.flush           
    'end if

            'Response.write " <br># "& strSQLmtypbgt  &" #<br><br> " 
            'response.flush


            oRec6.open strSQLmtypbgt, oConn, 3
            if not oRec6.EOF then
            
            thisMtypTimerKost = replace(oRec6("mtyptimerkost"), ",", ".")
            thisMtypTimerBelob = replace(oRec6("mtyptimerbelob"), ",", ".")
            timerReal = replace(oRec6("timerreal"), ",", ".") 

            end if
            oRec6.close
            


            '**** Real timer pr. medarbtype konsolideret **'

            'strSQljobf = "SELECT jobid, id FROM timer_konsolideret_tot"& tb_ext &" WHERE jobid = "& oRec("id") & " AND mtype = "& mtypid(t)  &" AND mtgid = "& mtypgrpid(t) & " AND (year(dato) = '"& year(sqlDatostart) & "' AND month(dato) = '"& month(sqlDatostart) & "')"
            
            ''Response.write "findes reg på job pr. md<br>" & strSQljobf & "<br><br>"

            'thisId = 0
            'oRec6.open strSQljobf , oConn, 3
            'if not oRec6.EOF then
            'thisId = oRec6("id")

            'end if
            'oRec6.close
            


            
            'if thisId <> 0 then 'update
            'strSQlupdRealtKons = "UPDATE timer_konsolideret_tot"& tb_ext &" SET timer = "& timerReal &", belob = "& thisMtypTimerBelob &", kost = "& thisMtypTimerKost &" WHERE jobid = "& oRec("id") & " AND mtype = "& mtypid(t) &"  AND mtgid = "& mtypgrpid(t) & " AND (year(dato) = '"& year(sqlDatostart) & "' AND month(dato) = '"& month(sqlDatostart) & "')"
                
            'Response.write "<br>"& strSQlupdRealtKons & "<br><hr>"
            'Response.flush
            'oConn.execute(strSQlupdRealtKons)
    
            'else
            
            if timerReal > 0 then
            
                if cint(first) = 0 then 'Nat kørsel seneste 2 dage, sletter kun for den peridoe der blvier konsolidere
                'Nulstiller interval ved næste job    
                if lastJid <> oRec("id") then
               
                    strSQLdelKons2 = "DELETE FROM timer_konsolideret_tot WHERE jobid = "& oRec("id") & " AND mtype = "& mtypid(t) &" AND mtgid = "& mtypgrpid(t) &" AND dato = '"& sqlDatostart &"'"
                    response.write "<br>Tømmer specifik: "& strSQLdelKons2 & "<br>"
                    oConn.execute(strSQLdelKons2)

                end if

            end if




            strSQlupdRealtKons = "INSERT INTO timer_konsolideret_tot"& tb_ext &" (jobid, mtype, mtgid, timer, belob, kost, dato) VALUES ("& oRec("id") &", "& mtypid(t) &", "& mtypgrpid(t) &", "& timerReal &", "& thisMtypTimerBelob &", "& thisMtypTimerKost &", '"& sqlDatostart &"')"   
            
            Response.write "<br>mtypeId: " & mtypid(t) & " medarbType: "& mtypnavn(t) 
            Response.write "<br>"& strSQlupdRealtKons & "<br>"
            Response.flush
            oConn.execute(strSQlupdRealtKons) 
                   
            end if

    
            'end if
            

            
            

            end if
            next

            next 'måneder m


        lastJid = oRec("id")

    oRec.movenext
    wend
    oRec.close

    end if

    oConn.close

    next' L ltos


    
    '************* resustalt txt ****
    select case first
    case 2
    
    %>

    <br /><br /><H4>Job er færdig konsolideret.</H4><br><br>
    <a href="#" onclick="Javascript:window.close()" class="red">[Luk vindue]</a> 
        <br /><br /><br />
        &nbsp;
      </div>

    <%
    case else
    Response.write "<br><br>Alle records er indlæst! Job: "& j
    end select

%>
<!--#include file="../inc/regular/footer_inc.asp"-->