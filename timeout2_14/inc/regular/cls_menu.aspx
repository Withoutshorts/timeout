
<%
function tsamainmenu(pkt)
		
		call treg0206use(session("mid"))
		%>
		
		<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible; z-index:9000000;">
	
		<div style="position:relative;">
		
		
		
		<%
		call mmenuTableSt()
		
		%>
		<td><ul class='sf-menu'>
		<%
		
		
		'*** Timereg menu **'
		
		'if cint(treg0206thisMid) = 0 then
		'	tregLink = "timereg.asp?menu=timereg"
		'else
			'tregLink = "timereg_2006_fs.asp"
			tregLink = "timereg_akt_2006.asp"
		'end if
		
		call mmenuPktLi(1, "102", tregLink,""& global_txt_120 &"",pkt)
		
		%>
		<ul>
		<!--<li><a href="timereg_akt_2006.asp?showakt=1" class=rmenu><%=tsa_txt_116 %></a></li>-->
	
	    <%if ((level < 8 AND lto <> "bowe") OR (level < 7)) then %>
	    <li><a href="materialer_indtast.asp?id=0&fromsdsk=0&aftid=0" class=rmenu target="_blank"><%=tsa_txt_191 %></a></li>
	    <%end if %>
    	    
	       
	        <li><a href="joblog_timetotaler.asp" class=rmenu target="_blank">Statistik // <%=tsa_txt_119 %></a></li>
	        <li><a href="joblog.asp" class=rmenu target="_blank">Statistik // <%=tsa_txt_118 %></a></li>
	       
    	
	        <%if level = 1 then %>
		    <li><a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie, Afspad. & Sygdom's kalender</a></li>
		    <%else %>
		    <li><a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie & Afspad. kalender</a></li>
		    <%end if %>
    	
	    </ul>
	    </li><!--level 1-->
	
		<%
		
		if level <= 2 OR level = 6 then
            call mmenuPktLi(2, "101", "webblik_joblisten.asp?menu=webblik",""& replace(global_txt_121,"|", "&") &"",pkt)
		end if
		
		
		
		'if level <= 2 OR level = 6 then
		    'call mmenuPkt(4, "20", "ressource_belaeg_jbpla.asp?menu=res",""& global_txt_123 &"",pkt)
		'end if
		
		
		if level <= 3 OR level = 6 then
		   call mmenuPktLi(5, "09", "kunder.asp?menu=kund&visikkekunder=1",""& global_txt_124 &"",pkt)
		end if
		
		
		'**** Job menu ***'
		if level <= 3 OR level = 6 then
		    call mmenuPktLi(3, "63", "jobs.asp?menu=job&shokselector=1&fromvemenu=j",""& global_txt_122 &"",pkt)
		    
		    %>
		    <ul>
		     <li><a href='jobs.asp?menu=job&fromvemenu=j' class='rmenu'>Joboversigt</a></li>
		    
		    <li><a href='jobs.asp?menu=job&func=opret&id=0&int=1' class='rmenu'>Opret nyt job</a></li>
		    <%if level <= 2 OR level = 6 then %>
		    <li><a href='job_print.asp?menu=job&kid=0&id=0' class='rmenu'>Print job</a></li>
		    <li><a href='projektgrupper.asp?menu=job' class='rmenu'>Projektgrupper</a></li>
		    <li><a href='akt_gruppe.asp?menu=job&func=favorit' class='rmenu'>Stam-aktivitets grupper (Job skabeloner)</a></li>
		    <!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href='aktiv_multi.asp?menu=job' class='rmenu'>Multi opdater akt.</a>-->
    		
    		<li><a href='filer.asp?kundeid=0&jobid=0' class='rmenu'><%=global_txt_127 %></a></li>
		   
		    <li><a href='pipeline.asp?menu=job&FM_kunde=0&FM_progrupper=10' class='rmenu'>Pipeline</a></li>
		    <li><a href='webblik_joblisten21.asp' class='rmenu'>Jobplanner (Gantt årsoversigt)</a></li>
		    <%end if %>
		    
		    </ul>
		
		<%end if
		
		
		call mmenuPktLi(6, "19", "medarb.asp?menu=medarb",""& global_txt_125 &"",pkt)
		
		if level <= 2 OR level = 6 then
		    call mmenuPktLi(7, "33", "joblog_timetotaler.asp",""& global_txt_126 &"",pkt)
		end if
		
		
		'call mmenuPktLi(10, "38", "filer.asp?kundeid=0&jobid=0",""& global_txt_127 &"",pkt)
		
		
		if level <= 3 OR level = 6 then
            call mmenuPktLi(11, "30", "materialer.asp?menu=mat",""& global_txt_128 &"",pkt)
		end if
		
		%>
		
		
		
		
		</ul></td><%
		
		call mmenuTableEnd(1)%>
		
		
		</div>
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
		     <td id="screenw" colspan=32 valign=top style="border-top:3px #8CAAE6 solid; height:1px;">
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
		
		
		function mmenuPktLi(nb, img, lnk, lnkTxt, valgtPkt)
		
		if cint(nb) = cint(valgtPkt) then
		cls = "current"
		else
		cls = ""
		end if
		
		
		tgt = "_top"
		
		
		%>
		
		
        
		
		<li class="<%=cls %>"><a href="<%=lnk%>" class="mainmenu" target="<%=tgt%>"><%=lnkTxt%></a>
		<%
		end function
		
		%>
