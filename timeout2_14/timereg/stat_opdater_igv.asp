<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<script>
    function popUp(URL, width, height, left, top) {
        window.open(URL, 'navn', 'left=' + left + ',top=' + top + ',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
    }
</script>

<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else

     
     level = session("rettigheder")

     
     if len(trim(request("usemrn"))) <> 0 then
     usemrn = request("usemrn")
     else
     usemrn = session("mid")
     end if

     func = request("func")

     if len(trim(request("mthuse"))) <> 0 then
     mthuse = request("mthuse") 
     else
     mthuse = now
     end if

     mthuse_minus = dateAdd("m", -1, mthuse)
     mthuse_plus = dateAdd("m", 1, mthuse)

     select case func
     case "opdaterdb"

     call stadeopdater()

     'Response.write "her"
     'Response.end

     'Response.Write("<script language=""JavaScript"">window.opener.location.href = 'stat_opdater_igv.asp';</script>")
     'Response.flush
     'Response.Write("<script language=""JavaScript"">window.location.href('stat_opdater_igv.asp?func=opdater&usemrn="+ usemrn +"');</script>")
     
     Response.redirect "stat_opdater_igv.asp?func=opdater&usemrn="& usemrn &"&mthuse="& mthuse

     case "opdater"

     %>
     <!--#include file="../inc/regular/header_hvd_inc.asp"-->
     <%
     '**** Stade indmelding ****'
     call stadeindm(usemrn, 2, mthuse)



 

    call ekportogprint_fn()
               
   


     case "opdateralle"

       %>
     <!--#include file="../inc/regular/header_hvd_inc.asp"-->
     <%
     '**** Stade indmelding ****'
     call stadeindm(0, 3, mthuse)

       call ekportogprint_fn()

    case else

	%>


   
    



      <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
     <script src="inc/stat_opdater_igv_jav.js"></script>
	
	
	



    

   
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(7)%>
	</div>
	
	
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	
		'call stattopmenu()
	

	%>
	</div>
-->


    <%call menu_2014()%>

    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:200px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" /><br />&nbsp;
  
	</td></tr>
    <tr><td colspan=2>  <div id="load_cdown">Forventet loadtid: 4-7 sekunder...</div></td></tr>
    </table>

	</div>
    
    

   
	<%
	
	Response.Flush 
    
    %>

	
   



 <%



    
   


    tTop = 102
	tLeft = 20
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth)


    
    'oimg = "ikon_timereg_48.png"
	oleft = 0
	otop = 0
	owdt = 500
	oskrift = "Stade-indmeldinger, sen. 3 måneder"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
   


     %>
<br />&nbsp;
   
      <table border=0 cellspacing=1 cellpadding=0 width="100%">
	           
               <tr>
               <td><%if level = 1 then %>
               <a href="stat_opdater_igv.asp?func=opdateralle" class=vmenu target="_blank">Vis alle job (uanset jobansvarlig)</a>
               <%end if %>
              
               
               <table cellspacing=0 cellpadding=3 width="100%">

               <%
               dtNow = mthuse
               aJob = 0
               for dt = 0 to 2
               
               dtUse = dateAdd("m", -(dt), dtNow)

               yNow = year(dtUse)
               mNow = month(dtUse)

               aJob = 0
               %>
               <tr><td colspan=3 style="padding-left:10px;"><br /><br />&nbsp;</td></tr>
               <tr bgcolor="#CCCCCC"><td colspan=3><br /><b><%=monthname(mNow) %> - <%= yNow %></b></td></tr>
               <tr><td>Medarbejder</td><td>Antal job hvor jobansvarlig</td><td>Stade-indmelding afsluttet?</td></tr>
               <%

               strSQLm = "SELECT mid, COUNT(j.id) AS antaljob FROM medarbejdere AS m "_
               &" LEFT JOIN job AS j ON (jobans1 = m.mid AND (jobstatus = 1 OR jobstatus = 3) AND risiko >= 0) WHERE mid <> 0 AND mansat = 1 AND (jobstatus = 1 OR jobstatus = 3) AND risiko >= 0 GROUP BY mid ORDER BY mnavn"

               'Response.write strSQLm
               'Response.flush

               oRec2.open strSQLm, oConn, 3
               while not oRec2.EOF

               if isNULL(oRec2("antaljob")) <> true then
               aJob = trim(oRec2("antaljob"))
               else
               aJob = 0
               end if

               if aJob >= 1 then

                 call meStamdata(oRec2("mid"))
                 %>
                    <tr><td style="border-bottom:1px #CCCCCC solid;"><%=meTxt%></td><td style="border-bottom:1px #CCCCCC solid;">
                    <%if dt = 0 then %>
                    <a href="stat_opdater_igv.asp?menu=stat&usemrn=<%=oRec2("mid") %>&func=opdater" target="_blank" class=vmenu><%=oRec2("antaljob") %></a>
                    <%else %>
                    <%=oRec2("antaljob") %>
                    <%end if %>
                    </td>
                 <%

               strSQLigvst = "SELECT medid, maaned, aar, dato FROM job_igv_status "_
               &"  WHERE medid = "& oRec2("mid") & " AND aar = "& yNow & " AND maaned = "& mNow 

               afsl = 0
               oRec.open strSQLigvst, oConn, 3
               While not oRec.EOF 


               %>
            <td style="border-bottom:1px #CCCCCC solid;">Ja: <%=oRec("dato") %></td></tr>
               <%

               afsl = 1

               oRec.movenext
               wend
               oRec.close


               if afsl = 0 then
               %>
               <td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td></tr>
               <%
               end if


               end if

               oRec2.movenext
               wend
               oRec2.close
               

               next
               %>

               
            

               </table>
               <br /><br /><br /><br /><br />&nbsp;
	            </td></tr>
	            </table>

                <br /><br /><br /><br /><br />&nbsp;

    
    </div> <!-- table div -->



  
	
    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	
	
	

<%

end select
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
