	<%
	function webbliktopmenu()

      
        resmenupkt = "Ressource Forecast (timebudget)"
      
        %>
        <br />
        <%

        if level <= 2 then 
		%>
		<a href='webblik_joblisten.asp?menu=webblik' class='rmenu'>Igangværende job (liste)</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='webblik_joblisten21.asp?menu=webblik' class='rmenu'>Jobplanner (gantt)</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;
		<%end if %>


		<%if cint(erpOnOff) = 1 then%> 
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href='erp_tilfakturering.asp?menu=erp' class='rmenu'>Job til fakturering.. (ERP)</a>-->
		<%end if %>
		

         <% if level <= 7 then  %>
		    <a href="ressource_belaeg_jbpla.asp?menu=webblik" class=rmenu><%=resmenupkt %></a>
            <%end if %>

        <% if level <= 2 then  %>
		<!--	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="jbpla_w.asp?menu=webblik" class=rmenu>Ressource Kalender</a> -->
			<!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href="ressource_belaeg.asp" class=rmenu>Ressource Belægning</a>
			&nbsp;&nbsp;|&nbsp;&nbsp;<a href="jbpla_k.asp" class=rmenu>Job Belastning</a><br>&nbsp;-->
		
		
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;<a href='webblik_todo.asp' class='rmenu'>ToDo's</a>-->
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='webblik_milepale.asp?menu=webblik' class='rmenu'>Milepæle</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='pipeline.asp?menu=webblik&FM_kunde=0&FM_progrupper=10' class='rmenu'>Pipeline</a>
		
        <%end if %>

		<%if level = 1 then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=webblik' class='rmenu'>Ferie, Afspad. & Sygdom's kalender</a>
		<%else %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=webblik' class='rmenu'>Ferie & Afspad. kalender</a>
		<%end if %>
		<br>&nbsp;
		<%
	end function


   

	%>