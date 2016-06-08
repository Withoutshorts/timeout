
<%function helpandguides(lto, vis)

if vis <> 1 then
dsp = "none"
wzb = "hidden"
else
dsp = ""
wzb = "visible"
end if %>
 


<div id="helpandfaq" style="display:<%=dsp%>; position:absolute; left:90px; visibility:<%=wzb%>; padding:20px; width:80%; background-color:#FFFFFF;">
<br />
<table width=100% cellpadding=0 cellspacing=0 border=0>
    <tr>
        <td colspan=2><b>Hjælp, guides og manualer til TimeOut. PDF versioner</b><br /><br />&nbsp;</td>
    </tr>

    <%select case lto
     case "tec"%>
    <tr>
        <td >
        <a href="http://www.eniga.dk/it-service/vejledninger/til-medarbejdere/tidsregistrering/" target="_blank" >Vejledning til TEC TimeOut</a>
        </td>
    </tr>
      <%case "esn"%>
    <tr>
        <td >
        <a href="http://www.eniga.dk/it-service/vejledninger/til-medarbejdere/tidsregistrering/" target="_blank" >Vejledning til ESN TimeOut</a>
        </td>
    </tr>
        
     <%case else%>
    <tr>
        <td  style="width:80%;"><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_kom_godt_igang._rev_100420_1.pdf" target="_blank" >Kom Godt igang [PDF]</a></td>
        <td rowspan=7  valign=top>
        
        
        <b>Kontakt TimeOut Support:</b><br />
        Man-Tor 09.00 - 15.00<br />
        Fredag 09.00 - 14.00<br /><br />
        Tel. +45 26 84 20 00<br />
        Email: <a href="mailto:support@outzource.dk" >support@outzource.dk</a>
        <br /><br />
        
        <b>Online Support:</b><br />
        <A href="https://www.islonline.net/start/ISLLightClient"  target="_blank">
            OutZourCE Online Remote Desktop Support.
        </A>
        </td>
        
        </td>
    </tr>
     <tr>
        <td >
        <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_indtasttimer_rev_20130102.pdf" target="_blank" >Timeregistrering - generel [PDF]</a>
        </td>
        
        
    </tr>
     <tr>
        <td >
        <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_job_oprettelse_rev100915.pdf" target="_blank" >Joboprettelse [PDF]</a>
        </td>
        
        
    </tr>
      <tr>
        <td >
        <%if lto <> "immenso" AND lto <> "execon" then%>
        <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_fakturering_rev3.pdf" target="_blank" >Fakturering [PDF]</a>
        <%else %>
        <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_fakturering_rev_091220-4_ex.pdf" target="_blank" >Fakturering [PDF]</a>
        <%end if %>
        </td>
        
        
    </tr>
     <tr>
        <td >
        <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/help_and_faq/TimeOut_udgiftsbilag_rev1.pdf" target="_blank" >Udgiftsbilag [PDF]</a>
        </td>
        
        
    </tr>
    <%end select %>
   
   
    
   
</table>

</div>

<center>
<br /><br />
© 2002 - <%=year(now) %> OutZourCE og OutZourCE -underleverandører. Alle rettigheder forbeholdes.
</center>

<br /><br />


<%end function %>