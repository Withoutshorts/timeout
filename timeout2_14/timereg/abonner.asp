<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
    if len(session("user")) = 0 then

    %>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
    

    func = request("func")

    
    select case func


    case "abo"

      '**** Abonner ****'
            

            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

            oConn_admin.open strConnect_admin
            
            editor = session("user")

            'Response.write request("FM_abo_mid") & "<br>"

            aMedid = split(request("FM_abo_mid"), ", ") 'session("mid")
            aMedTyp = request("FM_abo_mtyp")
            
            if trim(aMedTyp) = "-1" then 'kan kun abonnere på enten eller
            aProgrp = request("FM_abo_progrp")
            else
            aProgrp = "-1"
            end if 

            lto = lto
            
            
            
            
            if request("FM_abo") = "1" then 'aboner
            
           for m = 0 to UBOUND(aMedid)
            

            'Response.write aMedid(m)

            dtNowSQL = year(now) &"/"& month(now) &"/"& day(now)
          
            call meStamdata(aMedid(m))

         

            'if meBrugergruppe = 3 then 'Admin
            rapporttype = 0 'Viser for alle medarbejdere
            'else
            'rapporttype = 1 'Viser kun tal for sig sel
            'end if

                '** Findes den i forvejen ==> gør intet ***'
                strSQLAbo = "SELECT medid FROM rapport_abo WHERE medid = "& aMedid(m) & " AND lto = '"& lto & "'"
                oRec_admin.open strSQLAbo, oConn_admin, 3
                if not oRec_admin.EOF then

                else '**Insert

                 strSQLAbo = "INSERT INTO rapport_abo (dato, editor, lto, navn, email, medid, rapporttype, akttyper, medarbejdertyper, projektgrupper) "_
                 &" VALUES ('"& dtNowSQL &"','"& editor &"', '"& lto &"', '"& meNavn &"', '"& meEmail &"', "& aMedid(m) &", "& rapporttype &", '#0#', '"& aMedTyp &"', '"& aProgrp &"' )"
                 oConn_admin.execute(strSQLAbo)

                end if
                oRec_admin.close
          
            
            next


            else '**slet
               

            

            end if
            
            
          
            oConn_admin.close

            'Response.end


            Response.redirect "abonner.asp"

        case "abo_slet"


            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

            oConn_admin.open strConnect_admin
            
            aMedid = request("FM_abo_mid") 'session("mid")
            lto = lto


         strSQLAbo = "DELETE FROM rapport_abo WHERE medid = "& aMedid & " AND lto = '"& lto & "'"
               'Response.write strSQLAbo
                oConn_admin.execute(strSQLAbo)


         Response.redirect "abonner.asp"
   
case else
	

        %>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call stattopmenu()
	'end if
	%>
	</div>
-->

<%

    call menu_2014()

tLeft = 20
tWdth = 1004

tTop = 120
oskrift = "Abonnér"

    call filterheader_2013(100,20,1004,oskrift)
'call tableDiv(tTop,tLeft,tWdth)


            aboCHK = ""
         

%>

             Abonnér på uge rapporten. Rapporten indeholder bl.a normtid, realiseret tid og fravær mm.<br />
            Valgte medarbejder ønsker at abonnére på uge-rapporten. Udsendes tirsdag morgen.<br /><br />
            Vælg <b>medarbejder</b>, samt hvilke <b>medarbejdertyper <u>eller</u> projektgrupper</b> de skal abonnere på.<br />&nbsp;


        <FORM method="post" action="abonner.asp?func=abo">
        <table cellspacing=0 cellpadding=0 border=0 width=100%>
       
        <tr>     
	            <td>Vælg modtager<br /><b>Medarbejdere:</b> (kun dem med email adr.)<br />
        <%strSQLm = "SELECT mnavn, mid, email FROM medarbejdere WHERE mansat = 1 AND email <> '' ORDER BY mnavn"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:230px; font-size:11px;" size="10" multiple name="FM_abo_mid">
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("mid")%>"><%=oRec2("mnavn") %></option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


        <br />
         <input type="hidden" name="FM_abo" value="1" /> 
        </td>


             <td>Vælg hvilke medarbejdere der skal med i rapporten. <br />Enten medarbejdertyper eller projektgrupper<br /><b>A) Medarbejdertyper:</b><br />
        <%strSQLm = "SELECT id, type FROM medarbejdertyper WHERE id <> 0 ORDER BY type"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:230px; font-size:11px;" size="10" multiple name="FM_abo_mtyp">
            <option value="-2" SELECTED>Kun sig selv</option>
              <option value="-1">Ingen</option>
            <option value="0">Alle</option>
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("id")%>"><%=oRec2("type") %> (id:<%=oRec2("id") %>)</option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


       
        </td>
                 <td><br /><b>B) Projektgrupper:</b><br />
        <%strSQLm = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:230px; font-size:11px;" size="10" multiple name="FM_abo_progrp">
            <option value="-1" SELECTED>Ingen</option>
            <option value="0">Alle</option>
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("id")%>"><%=oRec2("navn") %> (id:<%=oRec2("id") %>)</option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


       
        </td>




        </tr>
        <tr>
	    <td align=right valign=bottom colspan="3">


            <input id="Submit5" type="submit" value="Tilføj >>" />

          
	</td></tr>
</table>
</FORM>


        <%  

if request.servervariables("PATH_TRANSLATED") <> "c:\www\timeout_xp\wwwroot\ver2_1\timereg\abonner.asp" then





            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")
            oConn_admin.open strConnect_admin

         %>

  <b>Abonenter:</b> (rapporttype)<br />
   <span style="color:#999999;">
  0: Standard<br />
  1: Tilpasset</span><br /><br />

<table width="100%">
    <tr><td><b>Medarbejder</b></td><td><b>Rapporttype</b></td><td>Medarbejdertyper tilvalgt (0=alle, -1=ingen, -2=sig selv)</td><td>Projekgrupper tilvalgt (0=alle)</td><td>&nbsp;</td></tr>

 

          
            <% '** subscribers ***'
            strSQLAbo = "SELECT navn, medid, rapporttype, medarbejdertyper, projektgrupper FROM rapport_abo WHERE lto = '"& lto & "' ORDER BY navn "
            oRec_admin.open strSQLAbo, oConn_admin, 3
            while not oRec_admin.EOF
            
            %>
            <tr>
            <td><%=oRec_admin("navn")%></td><td><%=oRec_admin("rapporttype") %></td>
            <td><%=oRec_admin("medarbejdertyper") %></td>        
            <td><%=oRec_admin("projektgrupper") %></td>
                <td><a href="abonner.asp?func=abo_slet&FM_abo_mid=<%=oRec_admin("medid") %>" class="slet">Slet</a></td>
                </tr>
    <%


            oRec_admin.movenext
            wend
           
              
           
            oRec_admin.close
            oConn_admin.close

    
    %>
    </table>
    <%
    
    end if

    %>
      <!-- filter header sLut -->
	</td></tr></table>
	</div>
    

    <%

    end select
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
