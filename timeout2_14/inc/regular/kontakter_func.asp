<%
public stamdiv_vzb, stamdiv_dsp, logodiv_vzb, logodiv_dspk, logdiv_vzb, klogdiv_dsp
		public submLeft, submTop
		
		function kundemenu(showdiv)
		
		select case showdiv
		case "seraft"
		stamdiv_vzb = "hidden"
		stamdiv_dsp = "none"
		stamdiv_col = "#003399"
		stamdiv_top = 118
		
		kpersdiv_col = "#003399"
		kpersdiv_top = 118
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "hidden"
		klogdiv_dsp = "none"
		klogdiv_col = "#003399"
		klogdiv_top = 118
		
		logodiv_vzb = "hidden"
		logodiv_dsp = "none"
		logodiv_col = "#003399"
		logodiv_top = 118
		
		serviceaft_col = "limegreen"
		serviceaftdiv_top = 120
		
		case "kpers"
		stamdiv_vzb = "hidden"
		stamdiv_dsp = "none"
		stamdiv_col = "#003399"
		stamdiv_top = 118
		
		kpersdiv_col = "limegreen"
		kpersdiv_top = 120
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "hidden"
		klogdiv_dsp = "none"
		klogdiv_col = "#003399"
		klogdiv_top = 118
		
		logodiv_vzb = "hidden"
		logodiv_dsp = "none"
		logodiv_col = "#003399"
		logodiv_top = 118
		
		serviceaft_col = "#003399"
		serviceaftdiv_top = 118
		
		case "logo"
		stamdiv_vzb = "hidden"
		stamdiv_dsp = "none"
		stamdiv_col = "#003399"
		stamdiv_top = 118
		
		kpersdiv_col = "#003399"
		kpersdiv_top = 118
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "hidden"
		klogdiv_dsp = "none"
		klogdiv_col = "#003399"
		klogdiv_top = 118
		
		logodiv_vzb = "visible"
		logodiv_dsp = ""
		logodiv_col = "limegreen"
		logodiv_top = 120
		
		
		serviceaft_col = "#003399"
		serviceaftdiv_top = 118
	
		
		case "klog"
		stamdiv_vzb = "hidden"
		stamdiv_dsp = "none"
		stamdiv_col = "#003399"
		stamdiv_top = 118
		
		kpersdiv_col = "#003399"
		kpersdiv_top = 118
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "visible"
		klogdiv_dsp = ""
		klogdiv_col = "limegreen"
		klogdiv_top = 120
		
		logodiv_vzb = "hidden"
		logodiv_dsp = "none"
		logodiv_col = "#003399"
		logodiv_top = 118
		
		serviceaft_col = "#003399"
		serviceaftdiv_top = 118
	
		
		case "onthefly"
		
		stamdiv_vzb = "visible"
		stamdiv_dsp = ""
		stamdiv_col = "limegreen"
		stamdiv_top = 40
		
		kpersdiv_col = "#003399"
		kpersdiv_top = 118
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "hidden"
		klogdiv_dsp = "none"
		klogdiv_col = "#003399"
		klogdiv_top = 118
		
		logodiv_vzb = "hidden"
		logodiv_dsp = "none"
		logodiv_col = "#003399"
		logodiv_top = 118
		
		serviceaft_col = "#003399"
		serviceaftdiv_top = 118
		
	
		
		
		case else
		stamdiv_vzb = "visible"
		stamdiv_dsp = ""
		stamdiv_col = "limegreen"
		stamdiv_top = 120
		
		kpersdiv_col = "#003399"
		kpersdiv_top = 118
		
		dokdiv_vzb = "hidden"
		dokdiv_dsp = "none"
		dokdiv_col = "#003399"
		dokdiv_top = 118
		
		klogdiv_vzb = "hidden"
		klogdiv_dsp = "none"
		klogdiv_col = "#003399"
		klogdiv_top = 118
		
		logodiv_vzb = "hidden"
		logodiv_dsp = "none"
		logodiv_col = "#003399"
		logodiv_top = 118
		
		serviceaft_col = "#003399"
		serviceaftdiv_top = 118
		
	
		
		end select
		'****
	
	
	oimg = "ikon_kunder_48.png"
	oleft = 20
	otop = 70
	owdt = 300
	oskrift = "Kontakter"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
		
		%>
		
		
		
		
	<div id=kmenu style="position:relative; left:20px; top:100px;">
	<table cellspacing="1" cellpadding="2" border="0">
		
		<tr bgcolor="#EFF3FF">
		
		<%if intKid <> 0 then%>
			<td style="width:80px; height:22px;" align=center><a href="kunder.asp?menu=<%=menu%>&func=red&id=<%=intKid%>&FM_soeg=<%=thiskri%>" class='vmenu'>Stamdata</a></td>
		<%end if %>
	
	
	
	<%if intKid <> 0 then%>
	
		
			<td style="width:140px;" align=center><a href="kontaktpers.asp?menu=<%=menu%>&kid=<%=intKid%>&FM_soeg=<%=thiskri%>" class='vmenu'>Kontaktpers. / filialer</a></td>
		
	<%end if%>
	        
	        
	       
	
	<%if intKid <> 0 then%>
	
		
			<td style="width:80px;" align=center><a href="javascript:popUp('filer.asp?kundeid=<%=intKid%>&jobid=0&nomenu=1', '900', '500','50', '50')" target="_self" class='vmenu'>Filarkiv</a>
			<!--<a href="#" onClick="showdok()" class='alt'>Dokumenter</a>--></td>
		
	<%end if%>
	
	
	<%if intKid <> 0 then%>
	
		
			<td style="width:180px;" align=center><a href="kunder.asp?menu=<%=menu%>&func=red&id=<%=intKid%>&showdiv=klog&FM_soeg=<%=thiskri%>" class='vmenu'>Faktura, Lev. bet. og besk.</a></td>
		
		
	<%end if%>
	
	
	
	<%if intKid <> 0 then%>
	
	
		
			<td style="width:80px;" align=center><a href="kunder.asp?menu=<%=menu%>&func=red&id=<%=intKid%>&showdiv=logo&FM_soeg=<%=thiskri%>" class='vmenu'>Logoer</a></td>
		
	<%end if%>
	
	
	<%if intKid <> 0 then%>
		<%if menu = "crm" then%>
		
				<td style="width:100px;" align=center>
				<!--<a href="javascript:NewWin_large('../inc/regular/kundelogview.asp?useKid=<%=intKid%>')" target="_self" class=vmenu>Se aktions- og kunde log</a>&nbsp;-->
				<a href="crmhistorik.asp?menu=crm&func=hist&id=<%=intKid%>&selpkt=hist" class=vmenu target="_blank">Aktions historik</a>
				</td>
			
			
		<%end if%>
	<%end if%>
		
	
	<%if menu <> "crm" then%>
		<%if intKid <> 0 then%>
		
		
		
			<td style="width:80px;" align=center><a href="serviceaft_osigt.asp?menu=<%=menu%>&id=<%=intKid%>&FM_soeg=<%=thiskri%>" class='vmenu'>Aftaler</a></td><!--onClick="showSer()"-->
	
		
		<%end if%>
	
	<%else%>
	
	
	<%end if%>
	
	</tr>
	</table>
	</div>
	
	
	
	<!--
	<span id=note name=note style="width:150px; position:absolute; left:840px; top:210px; visibility:visible; display:; z-index:250;">
	<%if intKid <> 0 then%>
	
	<%else%>
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
		<tr>
			<td style="background-color:#ffffe1; border:1px red solid; padding:15px;">
			<img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
			Kontaktpersoner, aftaler, filer mm. kan først tilknyttes efter oprettelse af stamdata.</td>
		</tr>
		</table>
	<%end if%>
	</span>
	-->
	
	
	<%
	end function
	
	
	function kundtopmenu()
		%><br>
		<a href='kunder.asp?menu=<%=menu%>&visikkekunder=1' class='rmenu'>Kontakter</a>
		<%if menu <> "crm" then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='serviceaft_osigt.asp?menu=<%=menu%>&id=0&func=osigtall' class='rmenu'>Aftaleoversigt</a>
		<%end if %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href="kontaktpers.asp?menu=<%=menu%>&func=list" class='rmenu'>Kontaktpersoner </a>
		
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href='infobase.asp?menu=kund&id=0' class='rmenu'>Infobase</a>-->
		<br>&nbsp;
		<%
	end function
	
	
	
	
	%>