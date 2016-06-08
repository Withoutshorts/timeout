	<%
	function erptopmenu()
		%>
		<br>
		<a href='erp_tilfakturering.asp' target="_top" class='rmenu'>Til fakturering</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='erp_opr_faktura_fs.asp?reset=0' class='rmenu' target="_top">Opret / Rediger faktura</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href='erp_fakturaer_find.asp' class='rmenu' target="_blank">Søg i fakturalinjer</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='erp_fakhist.asp' class='rmenu' target="_top">Faktura historik (søg i fakturaer)</a>
		<br>&nbsp;
		<%
	end function
	
	
	function erptopmenu2()
		%>
		<br>
		<!--<a href='erp_afstem_md.asp?menu=erp' class='rmenu' target="_top">Månedsafstemning</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->
		<a href='erp_job_afstem.asp?menu=erp&show=joblog_afstem' class='rmenu' target="_top">Afstemning real./faktureret (job og medarb.)</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href='erp_serviceaft_saldo.asp?menu=erp' class='rmenu' target="_top">Afstemning aftaler detail</a>
		<br>&nbsp;
		<%
	end function
	
	function erptopmenu3()
		%>
		<br>
		<a href='kontoplan.asp?menu=erp' class='rmenu' target="_top">Kontoplan</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='posteringer.asp?menu=erp&id=0' class='rmenu' target="_top">Kassekladde (0)</a>
	    &nbsp;&nbsp;|&nbsp;&nbsp;<a href='posteringer.asp?menu=erp&id=0&kontonr=-1&kid=0' class='rmenu' target="_top">I-Moms (-1)</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='posteringer.asp?menu=erp&id=0&kontonr=-2&kid=0' class='rmenu' target="_top">U-Moms (-2)</a>
		<%if level = 1 then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='resultatop.asp?menu=erp' class='rmenu' target="_top">Resultatopgørelse</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='momskoder.asp?menu=erp' class='rmenu' target="_top">Momskoder</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='nogletalskoder.asp?menu=erp' class='rmenu' target="_top">Nøgletalskoder</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='momsafsluttet.asp?menu=erp' class='rmenu' target="_top">Momsperiode (afsluttet dato)</a>
		<%end if %>
		<br>&nbsp;
		<%
	end function
	%> 
	
	