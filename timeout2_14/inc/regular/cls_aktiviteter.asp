<%

'*** Akttyper ***'
 	public aty_fakbar, aty_real, aty_pre, aty_enh, aty_on, aty_tfval, aty_medpafak, aty_pre_dg, aty_pre_prg
	function akttyper2009Prop(fakturerbartype)
	
	aty_pre = 0
	aty_fakbar = 0
	aty_real = 0
	aty_medpafak = 0
	
	'Response.Write fakturerbartype
	'Response.flush
	
	
	 strSQL4 = "SELECT aty_id, aty_on, aty_label, aty_desc, aty_on_realhours, "_
	 &" aty_on_invoice, aty_on_invoiceble, aty_pre, aty_enh, et_navn, aty_on_workhours, aty_pre_dg, aty_pre_prg FROM akt_typer "_
	 &" LEFT JOIN enheds_typer ON (et_id = aty_enh) "_
	 &" WHERE aty_id = "& fakturerbartype 
	 
	'Response.Write strSQL4
	'Response.flush
	 
	 oRec4.open strSQL4, oConn, 3
	 
	
	 if not oRec4.EOF then
	
	  if oRec4("aty_on_invoiceble") = 1 then
	  aty_fakbar = 1
	  else
	  aty_fakbar = 0
	  end if
	  
	  if oRec4("aty_on_invoice") = 1 then
	  aty_medpafak = 1
	  else
	  aty_medpafak = 0
	  end if
	  
	  if oRec4("aty_on_realhours") = 1 then
	  aty_real = 1
	  else
	  aty_real = 0
	  end if
	  
	  if oRec4("aty_pre") <> 0 then
	  aty_pre = oRec4("aty_pre")
	  else
	  aty_pre = 0
	  end if


      if len(trim(oRec4("aty_pre_dg"))) <> 0 then
      aty_pre_dg = oRec4("aty_pre_dg")
      else  
      aty_pre_dg = "1,2,3,4,5"
      end if

       if len(trim(oRec4("aty_pre_prg"))) <> 0 then
      aty_pre_prg = oRec4("aty_pre_prg")
      else  
      aty_pre_prg = "10"
      end if
    
       

	  
	  aty_on = oRec4("aty_on")
	  
	  aty_enh = oRec4("et_navn")
	  
	  
	  select case oRec4("aty_on_workhours")
	  case 0
	  aty_tfval = ""
	  case 1 '** fradrag
	  aty_tfval = "(-)"
	  case 2 '** tillæg
	  aty_tfval = "(+)"
	  end select
	  
	  'Response.Write "aty_enh" & aty_enh
	  
	 end if 
	 oRec4.close
	
	end function



    public aty_sql_real, aty_sql_fak, aty_sql_fak_on, aty_options, aty_sql_frawhours, aty_sql_sel
	public aty_sql_ikfakbar, aty_sql_fakbar, aty_sql_onfak, aty_sql_realhours, aty_sql_realHoursFakbar
	public aktiveTyper, aty_sql_realHoursIkFakbar, aty_sql_tilwhours, aty_sql_onfaknotReal, aty_sql_hide_on_treg, aty_sql_active
    public aty_sql_frawhours2, aty_sql_tilwhours2
	 
	function akttyper2009(dothis)
	
	'Response.Write "akttype_sel::"& akttype_sel
	 
	 aty_sql_fakbar = "fakturerbar = 0"
	 aty_sql_ikfakbar = "fakturerbar = 0"
	 aty_sql_onfak =  "aktiviteter.fakturerbar = 0"
	 
	 aty_sql_realhours = "tfaktim = 0"
	 aty_sql_realHoursFakbar = "t.tfaktim = 0"
	 aty_sql_realHoursIkFakbar = "t.tfaktim = 0"
	 
	 aty_sql_frawhours = "t.tfaktim = 0"
	 aty_sql_tilwhours = "t.tfaktim = 0"

     aty_sql_frawhours2 = "t.tfaktim = 0"
	 aty_sql_tilwhours2 = "t.tfaktim = 0"


	 aty_sql_sel = "t.tfaktim = 0"

     aty_sql_active = "tfaktim = 0"

     aty_sql_hide_on_treg = " a.fakturerbar <> 0"
	 
     aty_sql_onfaknotReal = "aktiviteter.fakturerbar = 0"
	 aty_lastsort = 0
     aty_options = ""
	 
	 '** vis kun admin typer på internejob
	 if dothis = 1 then
	  
	  
	  call licKid()
	  '** finder kid på valgte job ***'
	  thisKid = 0
	  strSQLjid = "SELECT j.jobknr FROM job j "_
	  &" WHERE j.id = "& jobid 
	  
	  oRec4.open strSQLjid, oConn, 3
	  if not oRec4.EOF then
	  
	  thisKid = oRec4("jobknr")
	  
	  end if
	  oRec4.close
	  
	 if (cint(licensindehaverKid) <> cint(thisKid)) then 
	 'Vis kun admintyper på interne
	 AstrSQLwhKri = " AND ((aty_id BETWEEN 1 AND 6) OR (aty_id BETWEEN 50 AND 61) OR (aty_id BETWEEN 90 AND 100) OR aty_id = 30 OR aty_id = 25)"
	 else
	 AstrSQLwhKri = ""
	 end if
	
	end if
	 
	 'aty_on_invoiceble
	 '*** Typer der skal med på timeregsiden og medregnes i dagligt timeforbrug ***'
	 
	 strSQL4 = "SELECT aty_id, aty_label, aty_desc, aty_on_realhours, aty_on_invoice, "_
	 &" aty_on_invoice, aty_on_invoiceble, aty_sort, aty_on_workhours, aty_hide_on_treg, aty_enh FROM akt_typer "_
	 &" WHERE aty_on = 1 "& AstrSQLwhKri &" ORDER BY aty_sort"
	
	 
	 'Response.Write strSQL4
	 'Response.flush
	 
	  oRec4.open strSQL4, oConn, 3
	 
	 xi = 0
	 while not oRec4.EOF
	 
	 if dothis = 1 then
	  
    

      if oRec4("aty_hide_on_treg") = 1 then
	  hideontreg = "S"
	  else
	  hideontreg = "-"
	  end if

	  if oRec4("aty_on_realhours") = 1 then
	  meditimeregn = "M"
	  else
	  meditimeregn = "-"
	  end if
	  
	  if oRec4("aty_on_invoice") = 1 then
	  medpafak = "E"
	  else
	  medpafak = "-"
	  end if
	  
	  if oRec4("aty_on_invoiceble") = 1 then
	  fakturerbar = "Z"
	  else
	  fakturerbar = "-"
	  end if
	  
	  if oRec4("aty_on_workhours") = 1 then
	  fradrag = "F"
	  else
	  fradrag = "-"
	  end if
	  
	    call akttyper(oRec4("aty_id"), 1)
	    
	    if cint(strFakturerbart) = cint(oRec4("aty_id")) then
	    aktCHK = "SELECTED"
	    else
	    aktCHK = ""
	    end if
	    
	    aty_options = aty_options & "<option value='"& oRec4("aty_id")&"' "&aktCHK&">"& akttypenavn 


        '** Salgs aktvitiet ***'
        select case lto
        case "tec", "esn"

        case else

            if oRec4("aty_id") = 6 then 'Salg
            aty_options = aty_options & " - Vises på timereg. siden på tilbud"
            end if

        end select
	    
	    if level = 1 then
	    aty_options = aty_options & " ("& meditimeregn & " "& medpafak &" "&fakturerbar&" "&  fradrag &" "& hideontreg &")"
	    end if
	    
	    aty_options = aty_options & "</option>"
	    
	 end if
	 
	 
	 if dothis = 2 then
	      
	      '** Aktiviteter tabel ***'  
	      '** Fakturerbare / ik fakturebare
	      if oRec4("aty_on_invoiceble") = 1 then
	      aty_sql_fakbar = aty_sql_fakbar & " OR fakturerbar = "& oRec4("aty_id")
	      else
	      aty_sql_fakbar = aty_sql_fakbar 
	      end if
	      
	      if oRec4("aty_on_invoiceble") = 2 then
	      aty_sql_ikfakbar = aty_sql_ikfakbar & " OR fakturerbar = "& oRec4("aty_id")
	      else
	      aty_sql_ikfakbar = aty_sql_ikfakbar 
	      end if
	        
	      
	      '**** Alle aktive typer ****'
          aty_sql_active = aty_sql_active &" OR tfaktim = "& oRec4("aty_id")
	      
	      '** Timer tabel ***'
	        
	        '*** Tælles med i timereg **'
	        if oRec4("aty_on_realhours") = 1 then
	        aty_sql_realhours = aty_sql_realhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realhours = aty_sql_realhours
	        end if
	        
	        
	        '*** Fakturerbare timer **'
	        if oRec4("aty_on_invoiceble") = 1 then
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar
	        end if
	        
	         '*** Ikke Fakturerbare timer **'
	        if oRec4("aty_on_invoiceble") = 2 then
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar
	        end if
	        
	        
	          
	        '**** Fradrag fra løntimer ***'
	        if oRec4("aty_on_workhours") = 1 AND oRec4("aty_enh") = 0 then 'kun enh = timer
	        aty_sql_frawhours = aty_sql_frawhours &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_frawhours = aty_sql_frawhours
	        end if
	        
	        '**** Tillæg fra løntimer ***'
	        if oRec4("aty_on_workhours") = 2 AND oRec4("aty_enh") = 0 then
	        aty_sql_tilwhours = aty_sql_tilwhours &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_tilwhours = aty_sql_tilwhours
	        end if



              '**** Fradrag fra løntimer ***'
	        if oRec4("aty_on_workhours") = 1 AND oRec4("aty_enh") = 2 then 'kun enh = enheder
	        aty_sql_frawhours2 = aty_sql_frawhours2 &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_frawhours2 = aty_sql_frawhours2
	        end if
	        
	        '**** Tillæg fra løntimer ***'
	        if oRec4("aty_on_workhours") = 2 AND oRec4("aty_enh") = 2 then
	        aty_sql_tilwhours2 = aty_sql_tilwhours2 &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_tilwhours2 = aty_sql_tilwhours2
	        end if

	        
	        
            '**** Med på faktura ***'
	          if oRec4("aty_on_invoice") = 1 then
              aty_sql_onfak = aty_sql_onfak & " OR aktiviteter.fakturerbar = "& oRec4("aty_id")
              else
              aty_sql_onfak = aty_sql_onfak 
              end if

             '**** Med på faktura, men ikke med i dagligt timergnskab ***'
	          if oRec4("aty_on_invoice") = 1 AND oRec4("aty_on_realhours") = 0 then
              aty_sql_onfaknotReal = aty_sql_onfaknotReal & " OR aktiviteter.fakturerbar = "& oRec4("aty_id")
              else
              aty_sql_onfaknotReal = aty_sql_onfaknotReal
              end if
	        
	        aktiveTyper = aktiveTyper & "#"&oRec4("aty_id") & "#, "

            '*** Skjult på timereg siden **'
            if oRec4("aty_hide_on_treg") = 1 then
            aty_sql_hide_on_treg = aty_sql_hide_on_treg & " AND a.fakturerbar <> "& oRec4("aty_id")
            else
            aty_sql_hide_on_treg = aty_sql_hide_on_treg
            end if
	        
	  end if
	 
	 
	 if dothis = 3 then
	 
	 if xi = 0 then
	 
	 %>
	 <b>Aktive aktivitets typer:</b> (Konfiguration, se kontrolpanel)<br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr bgcolor="#EFf3ff">
			    <td valign=bottom><b>Id</b></td>
			    <td valign=bottom><b>Navn</b></td>
			    <td valign=bottom><b>Fakturerbar</b><br />
			    Ikke fakturerbar ell. administrativ</td>
			    <td valign=bottom><b>Medregnes i dagligt timeregnskab *)</b><br />
			    (Dvs. indgår denne type i de dagligt realiserede timer der udgør balancen mellem normeret tid pr. uge og realiseret tid)</td>
			    <%
                if session("stempelur") <> 0 then %>
			    <td valign=bottom><b>Tillæg/Fradrag +/- </b><br />
			    I fht. løntimer</td>
			    <%end if %>
			</tr>
	 <%end if
	 %>
			<tr>
			<td valign=top><%=oRec4("aty_id")%></td>
			<td>
			    <%call akttyper(oRec4("aty_id"), 1) %>
	            <%=akttypenavn%>
			</td>
			<td>
			<%select case oRec4("aty_on_invoiceble")
			case 1 %>
			Fakturerbar
			<%case 2 %>
			Ikke Fakt. bar
			<%case else %>
			Administrativ
			<%end select %></td>
			<td>
			<%if oRec4("aty_on_realhours") = 1 then %>
			Ja
			<%end if %>
			</td>
			<%

            if session("stempelur") <> 0 then %>
			        <td>
			        <%select case oRec4("aty_on_workhours") 
			        case 1 %>
			        Fradrag
			        <% 
			        case 2
			        %>
			        Tillæg
			        <%
			        case else
			        end select %>
			        </td>
			        <%end if %>
			        </tr>
			
			        <%
	 
	         end if
	 
	 
	 
	 if dothis = 4 then
	 
	  if oRec4("aty_on_invoice") = 1 then
	  aty_sql_onfak = aty_sql_onfak & " OR aktiviteter.fakturerbar = "& oRec4("aty_id")
	  else
	  aty_sql_onfak = aty_sql_onfak 
	  end if
	 
	 
	 end if
	 
	 '******* Vælg akttyper joblog + bal_real_norm (medarbafstemning)
	  if dothis = 5 then
	  

	  if xi = 0 then  %>
	  
	  <!-- start tæller pga af trim felter -->
	  <input name="FM_akttype" type="hidden" value="#-99#" />
	  
	   <table cellspacing=0 cellpadding=1 border=0 width=100%>
	  
	        <%if thisfile = "bal_real_norm_2007.asp" then %>

               


			            <tr>
			                <td><input id="chkalle_0" type="checkbox" class="akt_afst" /></td>
			                <td><b>Afstemning</b>
			    
			                </td>
			                 <!-- <td><b>Time Regnsk.</b></td> -->
			            </tr>
			
			
			
			            <!-- faaste felter og sumtotaler -->
			            <%	if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then
                        akttypeCHK = "CHECKED"
                        else
                        akttypeCHK = ""
                        end if %>
            
                        <%if cint(stempelurOn) = 1 then%>
			            <tr>
			            <td  valign=top>
			
			
			
                        <input id="FM_akttype_id_0_0" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-5#" class="akt_afst" /></td>
			            <td >
			            Komme/Gå tid (lønt.)
			            </td>
			
			
			
			            </tr>
			            <%end if %>


			<tr bgcolor="#EFF3ff">
			<td  valign=top>
			
			<%
			if instr(akttype_sel, "#-1#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_1" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-1#" class="akt_afst" /></td>
			<td >
			Vis sum-totaler på alle aktivitetstyper 
			</td>
			</tr>
			
			<tr>
			<td >
			
			<%
			if instr(akttype_sel, "#-2#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_2" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-2#" class="akt_afst" /></td>
			<td >
			Ressource timer
			</td>
			</tr>
			
			<tr bgcolor="#EFF3ff">
			<td >
			
			<%
			if instr(akttype_sel, "#-3#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_3" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-3#" class="akt_afst" /></td>
			<td >
			Faktureret timer
			</td>
			</tr>
			
			<tr>
			<td >
			
			<%
			if instr(akttype_sel, "#-4#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_4" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-4#" class="akt_afst" /></td>
			<td >
			Mat. frb. / Udlæg
			</td>
			</tr>
		

           <tr><td colspan="2" style="padding-top:20px;">Vis kun saldo:</td></tr>

            <%if cint(stempelurOn) = 1 then%>

            <%
			if instr(akttype_sel, "#-10#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>

			<tr>
			<td  valign=top bgcolor="#DCF5BD">
			
			
			
            <input id="FM_akttype_id_0_10" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-10#" class="akt_afst_saldo" /></td>
			<td >
			Saldo Komme/Gå tid (lønt.) / Normtid 
			</td>
			
			
			
			</tr>
			<%end if %>

            <%
			if instr(akttype_sel, "#-20#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
           
			<tr>
			<td  valign=top bgcolor="lightpink">
			
			
			
            <input id="FM_akttype_id_0_20" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-20#" class="akt_afst_saldo" /></td>
			<td >
			Saldo Realiseret / Normtid 
			</td>
			
			
			
			</tr>
				<%end if %>
			
	 <%
	 v = 7 '5
	
	    end if 'thisfile = "bal_real_norm_2007.asp"  
	  
	   if cint(left(oRec4("aty_sort"), 1)) <> cint(aty_lastsort) then ' AND cint(left(oRec4("aty_sort"), 1)) <> 2 then
	   %>
	   </table>
        <input id="antal_v_<%=aty_lastsort%>" type="hidden" value="<%=v %>" />
	    <%v = 0%>

                

	  </td>

    
                

	  <td valign=top style="padding:3px; width:125px; border:1px #D6DFf5 solid;">
	   
	  <table cellspacing=0 cellpadding=1 border=0 width=100%>
	   <%
			    select case left(oRec4("aty_sort"), 1)
			    'case 0
			    'straktgrpnavn = "Afstemning"
			    case 1
			    straktgrpnavn = "Udspecificering<br> atk. typer"
			    aktcls = "akt_udspec"
			    case 2
			    straktgrpnavn = "Udspecificering<br>akt. typer<br> (kat. 2)"
			    aktcls = "akt_flex"
			    case 3
			    straktgrpnavn = "Ferie"
			    aktcls = "akt_ferie"
			    case 4
			    straktgrpnavn = "Overarbejde / Afspadsering"
			    aktcls = "akt_overarb"
			    case 5
			    straktgrpnavn = "Sygdom / Fravær"
			    aktcls = "akt_syg"
			    case else
			    straktgrpnavn = ""
			    end select  
                
                
                if (left(oRec4("aty_sort"), 1) <> 5) OR (level = 1) then 'sygdom kun admin %>
			<tr>
			    <td><input id="chkalle_<%=left(oRec4("aty_sort"), 1) %>" type="checkbox" class=<%=aktcls %> /></td>
			    <td><b><%=straktgrpnavn %></b></td>
			    <!-- <td><b>Time Regnsk.</b></td> -->
			</tr>

	 <%     end if 
     end if %>
	 
	  <%select case right(xi, 1)
	  case 2,4,6,8,0
	  bgthis = "#EFF3ff"
	  case else
	  bgthis = "#ffffff"
	  end select 
	  
	  
	  if instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	  akttypeCHK = "CHECKED"
	  else
	  akttypeCHK = ""
	  end if
	  %>
	        
	        <%
	        '**** Sygdom og sundhed kun admin ***'

	        if (oRec4("aty_id") = 20 OR oRec4("aty_id") = 21 OR oRec4("aty_id") = 22 OR oRec4("aty_id") = 23 OR oRec4("aty_id") = 24 OR oRec4("aty_id") = 8 OR oRec4("aty_id") = 81) AND level <> 1 then 'AND thisfile = "bal_real_norm_2007.asp" then 
	        hdflt = 1
	        else
	        hdflt = 0
	        end if 
	        
	        if hdflt <> 1 then%>

          
                 <%if oRec4("aty_sort") = "2,5" OR (oRec4("aty_sort") = "2,6" AND korrOversktiftIsWrt <> 1) then %>
            <tr>
                <td colspan="2"><br /><br />Korrektion (overført) i forb. med lønperiode</td>
                 </tr>
                <%
                    korrOversktiftIsWrt = 1
                    end if %>
           
	        
			<tr bgcolor="<%=bgthis%>">
			<td >
                <input id="FM_akttype_id_<%=left(oRec4("aty_sort"), 1)%>_<%=v%>" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#<%=oRec4("aty_id")%>#" class=<%=aktcls %> /></td>
			<td >
			
			<%call akttyper(oRec4("aty_id"), 1) %>
            <%=akttypenavn%>
			</td>
			</tr>
			
			<%end if
	 
	 aty_lastsort = left(oRec4("aty_sort"), 1)
	 v = v + 1
	 end if
	 
	 '*** slut vælg akt typer **'
	 
	 
	 
	 
	 if dothis = 6 then
	 
	 aktiveTyper = aktiveTyper & "#"&oRec4("aty_id") & "#, "
	 
	 end if
	 
	 
	 '*** Medarbejder afstemning / Joblog **'
	 if dothis = 7 then
	 
	        
	        
	        '*** Tælles med i timereg **'
	        if oRec4("aty_on_realhours") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realhours = aty_sql_realhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realhours = aty_sql_realhours
	        end if
	        
	        
	        
	         
	        '*** Fakrurerbare **'
	        if oRec4("aty_on_invoiceble") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar 
	        end if
	        

            '*** Ikke Fakturerbare **'
	        if oRec4("aty_on_invoiceble") = 2 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar
	        end if
	        
	        
	        '**** Fradrag fra løntimer ***'
	        if oRec4("aty_on_workhours") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_frawhours = aty_sql_frawhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_frawhours = aty_sql_frawhours
	        end if
	        
	        '**** Tillæg fra løntimer ***'
	        if oRec4("aty_on_workhours") = 2 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_tilwhours = aty_sql_tilwhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_tilwhours = aty_sql_tilwhours
	        end if
	        
	         '*** Aktive typer **'
	        if instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_sel = aty_sql_sel &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_sel = aty_sql_sel
	        end if
	        
	 end if
	 
	
	 
	 'aty_sql_real = aty_sql_real & " AND "
	 
	 
	 
	 xi = xi + 1
	 oRec4.movenext
	 wend
	 oRec4.close
	 
	 if dothis = 5 then
	 %>
	 <input id="antal_v_<%=aty_lastsort%>" type="hidden" value="<%=v %>" />
	 <%
	 end if
	 
	 if dothis = 3 then
	 %>
	 <tr>
	 <td colspan=5><br /><b>*)</b> Hvis man angiver den normerede uge som 37 timer excl. frokost skal typen <b>"frokost"</b> sættes til <b>ikke</b> at tælle med i det daglige timeregnskab (de realiserede timer pr. dag). I dette tilfælde sættes den til "Administrativ".<br /><br />
	 Typen <b>"Afspadsering"</b> sættes normalt til Ja, så den tæller med i den realiserede tid pr. uge, således at man ikke går i minus på den løbende fleks saldo ved at afspadsere. normalt kan man kun afspadsere optjent overarbejde. Afspadsering bliver automatisk altid modregnet typen "Overarbejde" <br /><br />
	 I det tilfælde hvor man har optjent et plus på norm/real tid saldoen bruges typen <b>"Fleks"</b> til at "nulle" saldoen. Derfor skal typen <b>"Fleks"</b> normalt ikke medregnes i de realiserede timer pr. dag.
	 </td>
	 </tr>
	 <%end if
	
	 
	end function

            
public totBelDiasbled
sub lastFaseSumSub

         if cint(budgetakt) = 0 then
         lastFaseSumSubColspanA = 3
         lastFaseSumSubColspanB = 4      
         else
         lastFaseSumSubColspanA = 1
         lastFaseSumSubColspanB = 5
         end if



        '**lastFaseSum **'
	    strAktListe = strAktListe &"<tr bgcolor=""#FFFFFF""><td colspan="& lastFaseSumSubColspanB &" style='padding-left:5px; padding-bottom:10px;'><b>"& replace(lastFase, "_", " ") &"</b> ialt: (kun viste akt.)</td>"_
	    &"<td style='padding-bottom:10px; text-align:right;'><span style='width:40px; font-size:9px; font-family:arial; border:0px #FFc0CB solid; padding:2px;' id='sltimer_"&lcase(lastFase)&"'><b>"& formatnumber(lastFaseTimer, 2) &"</b></span></td>"_
	    &"<td colspan="& lastFaseSumSubColspanA &"><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"<td style='padding-bottom:10px; white-space:nowrap; text-align:right;'><span style='width:60px; font-size:9px; font-family:arial; border:0px #FFc0CB solid; padding:2px;' id='slsum_"&lcase(lastFase)&"'><b>"& formatnumber(lastFaseSum, 2) &"</b></span></td>"_
	    &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"

        
	    strAktListe = strAktListe &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"
        

        strAktListe = strAktListe &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"</tr><input id='fa_"&lcase(lastFase)&"' value='"&fa&"' type=""hidden"" />"_ 
	    &"<input id='fatot_"&lcase(lastFase)&"' value='"& faTot &"' type=""hidden"" />"_
	    &"<input id='fatot_val_"&faTot&"' value='"&formatnumber(lastFaseSum, 2) &"' type=""hidden"" />"_
	    &"<input id='fatottimer_val_"&faTot&"' value='"&formatnumber(lastFaseTimer, 2)&"' type=""hidden"" />"
	
	     fa = 0
        lastFaseTimer = 0
        lastFaseSum = 0    
	    faTot = faTot + 1
	    
end sub


sub nyFaseSub

    if cint(budgetakt) = 0 then
         cpsnaNyfsSub = 8
    else

         cpsnaNyfsSub = 7

    end if

    strAktListe = strAktListe & "<tr bgcolor=""#eff3ff""><td colspan=2 style='padding:5px 5px 5px 5px;' align=right>fase: "_
    &"<input id='"&lcase(trim(thisFase))&"' name='' class=""faseoskrift_navn"" value='"&replace(thisFase, "_", " ")&"' type=""text"" style='width:195px; font-size:9px; font-family:arial;' /></td>"_
	&"<td><select name=""faseoskrift"" class=""faseoskrift"" id='"&lcase(trim(thisFase))&"' style='font-size:9px; font-family:arial;'>"_
	&"<option value=""1"">Vælg..</option>"_
	&"<option value=""1"">Aktiv</option>"_
	&"<option value=""0"">Lukket</option>"_
	&"<option value=""2"">Passiv</option>"_
	&"</select></td>"_
	&"<td colspan='"& cpsnaNyfsSub &"'>&nbsp;</td><td align=right><input id='sl_"&lcase(trim(thisFase))&"' class=""faseoskrift_slet"" type=""checkbox"" value=""1"" /></td><td>&nbsp;</td></tr>" 

end sub



public straktOptionlist 
function aktlisteOptions(jobid, status, vis, aktid)

    if len(trim(jobid)) <> 0 AND jobid <> 0 then
         jobid = jobid
        else 
         jobid = -1
    end if 

    if len(trim(aktid)) <> 0 then
         aktid = aktid
    else
         aktid = -1
    end if

         strSQLaktid = ""
         if aktid <> -1 then
         strSQLaktid = " OR a.id = "& aktid
         end if

    strSQL = "SELECT a.id, a.navn, fase, job"_ 	
    &" FROM aktiviteter a "_
	&" WHERE (job = "& jobid &" AND aktstatus = 1) "& strSQLaktid &" ORDER BY a.fase, a.navn" 
	
	'Response.Write strSQL
	'response.Flush
    'Response.end 
	c = 0
    lastFase = ""
	straktOptionlist = "<option value='0'>Vælg aktivitet</option>"
    straktOptionlist = straktOptionlist & "<option value='0'>Ingen (uspec.)</option>"
	oRec6.open strSQL, oConn, 3
    while not oRec6.EOF

         if lastFase <> oRec6("fase") AND len(trim(oRec6("fase"))) <> 0 then
         straktOptionlist = straktOptionlist & "<option value='0' DISABLED></option>"
         straktOptionlist = straktOptionlist & "<option value='0' DISABLED>"& oRec6("fase") &"</option>"
         end if

         if cdbl(aktid) = cdbl(oRec6("id")) then
         aktSEL = "SELECTED"
         else
         aktSEL = ""
         end if

         straktOptionlist = straktOptionlist & "<option value='"& oRec6("id") &"' "& aktSEL &">"& oRec6("navn") &"</option>"

         lastFase = oRec6("fase")

    oRec6.movenext
    wend
    oRec6.close


end function





public lastFase, lastFaseSum, thisFase, lastFaseTimer, fa, faTot 
public strAktListe, aktbudgetsamlet, akttfaktimtildelt
function hentaktiviterListe(jobid, func, vispasluk, sort)

                        '*** Finder jobnr *********
                        jobnrThis = 0
                        strSQLjobnr = "SELECT jobnr FROM job WHERE id = " & jobid
                        oRec8.open strSQLjobnr, oConn, 3

	                    if not oRec8.EOF then
                
                        jobnrThis = oRec8("jobnr")

                        end if
                        oRec8.close


      call timesimon_fn()
      call budgetakt_fn() 

      if cint(budgetakt) = 1 then 
         '** Vis kontonr eller. Stk. på budget på aktivitetslinjer
         '** HUSK OGSÅ job_jav_aktliste.js om den skal regne timer pbg. af budget
         
           
            'case "xintranet - local", "oko"
            resOskrift = "Ress. beløb"
            totBelDiasbled = "DISABLED"
            akttotprisName = "FM_akttotpris"  '** --> ÆNDRES TIL BLANK og de 2 XX slettes (nedenfor)
            stkTimOskrift = "Pr. time"
            'kontoClas = "kontocls_"& oRec6("kontoid")

                    
                     strKontoplan = ""
                     strSQLkontoplan = "SELECT id, kontonr FROM kontoplan"

                    oRec6.open strSQLkontoplan, oConn, 3
                    while not oRec6.EOF

                     strKontoplan = strKontoplan & "<option value="& oRec6("id") &">"& oRec6("kontonr") &"</option>" 


                    oRec6.movenext
                    wend
                    oRec6.close


        else

            resOskrift = "Ress. tim."
            totBelDiasbled = ""
            akttotprisName = "FM_akttotpris"
            stkTimOskrift = "Pr. stk./time"
            'kontoClas = ""
          

         end if


   

    
    samletverdi = 0
    a = 0
	orderBy = "a.fase, a.sortorder, a.navn"
	
	
	lastFase = ""
	lastFaseSum = 0
    kontototalsTYP = ""
    isKontoWrt = "0"
	
	'if vispasluk = 1 then
	vispaslukKri = " AND aktstatus <> - 1"
	'else
	'vispaslukKri = " AND aktstatus = 1"
	'end if
	
	strSQL = "SELECT a.id, a.navn, job, beskrivelse, fakturerbar, aktstartdato, "_
	&" aktslutdato, budgettimer, aktstatus, aktbudget, fomr.navn AS fomr, antalstk, tidslaas, "_
	&" a.fase, a.sortorder, a.bgr, easyreg, a.aktbudgetsum, aty_desc, aty_id, k.kontonr, k.id AS kontoid, aktkonto "_
	&" FROM aktiviteter a LEFT JOIN fomr ON (fomr.id = a.fomr) "_
    &" LEFT JOIN akt_typer aty ON (aty_id = a.fakturerbar) "_
    &" LEFT JOIN kontoplan k ON (k.id = a.aktkonto) "_
	&" WHERE job = "& jobid &" "& vispaslukKri &" ORDER BY "& orderBy 
	
	'Response.Write strSQL
	'Response.end 
	c = 0
	fa = 0
	faTot = 0
	oRec6.open strSQL, oConn, 3
	while not oRec6.EOF 
	
                
                        '*** Forecast timer / BELØB ***'

                        '*** Kun finansår?? ***'
                         resTimerThis = 0

                        select case lto
                        case "oko", "intranet - local"

                       
                            strSQLFC = "SELECT COALESCE(SUM(rmd.timer*6timepris), 0) AS restimer FROM ressourcer_md AS rmd "_
                            &" LEFT JOIN timepriser AS tp ON (tp.aktid = rmd.aktid AND tp.medarbid = rmd.medid) "_
                            &" WHERE rmd.aktid = "& oRec6("id") &" AND tp.jobid = "& jobid &" AND rmd.jobid = "& jobid &" GROUP BY rmd.aktid"
            	            oRec8.open strSQLFC, oConn, 3

	                        if not oRec8.EOF then
                
                                resTimerThis = formatnumber(oRec8("restimer"), 0)

                            end if
                            oRec8.close

                        case else
                       
                            strSQLFC = "SELECT COALESCE(SUM(timer), 0) AS restimer FROM ressourcer_md WHERE aktid = "& oRec6("id") &" GROUP BY aktid"
            	            oRec8.open strSQLFC, oConn, 3

	                        if not oRec8.EOF then
                
                                resTimerThis = formatnumber(oRec8("restimer"), 0)

                            end if
                            oRec8.close

                        end select



            

    'Response.write "hre:" & strAktListe
    'Response.end
	
    if not IsNull(oRec6("fase")) then
    thisFase = oRec6("fase")
    else
    thisFase = ""
    end if

   
	
	if lcase(trim(lastFase)) <> lcase(trim(thisFase)) AND c <> 0 AND len(trim(thisFase)) <> 0 then
	
	   '**lastFaseSum **'
	   call lastFaseSumSub
	     
       call nyFaseSub

    end if
	
    'strAktListe = strAktListe & strSQLFC

	if c = 0 then
	
	
          

          

            strAktListe = strAktListe &"<table cellspacing=""0"" cellpadding=""0"" width=""100%"" border=""0"" id=""incidentlist"">"_
            &"<tr><td><b>Navn</b></td>"_
            &"<td><b>Fase</b></td>"_
            &"<td><b>Status</b></td>"_
            &"<td><b>Type</b></td>"

            if cint(budgetakt) = 1 then 
            strAktListe = strAktListe &"<td><b>Konto</b></td>"
            end if

            strAktListe = strAktListe &"<td><b>Timer</b></td>"
           
            if cint(budgetakt) = 0 then 
            strAktListe = strAktListe &"<td><b>Stk.</b></td>"
            end if
           
            if cint(budgetakt) = 0 then
            strAktListe = strAktListe &"<td><b>Grundlag</b></td>"
            end if
            
            strAktListe = strAktListe &"<td><b>"& stkTimOskrift &"</b></td>"_
            &"<td style=""padding-left:10px;""><b>Pris i alt "& basisValISO &"</b></td>"_
            &"<td>&nbsp;</td>"_
            &"<td align=right><b>"& resOskrift  &"</b></td>"_
            &"<td align=right><b>Slet</b><br> <input id=""sl_"" class=""faseoskrift_slet"" type=""checkbox"" value=""1"" /></td>"_
            &"<td>&nbsp;</td>"_
            &"<tr>"
    
                if len(trim(thisFase)) <> 0 then
    
                call nyFaseSub
    
	            end if
	
	end if

    


    if vispasluk = 1 OR (vispasluk <> 1 AND oRec6("aktstatus") = 1) then
	
	select case right(c, 1)
    case 0,2,4,6,8
    bgcolor = "#EFf3FF"
    case else
    bgcolor = "#FFFFFF"
    end select
    
    bgrSEL0 = "SELECTED"
    bgrSEL1 = ""
    bgrSEL2 = ""
    
    '** budget grundlag **'
    select case oRec6("bgr")
    case 0 '** Intent angivet
    bgrSEL0 = "SELECTED"
    'aktsumtot = oRec6("aktbudget")
    case 1 '** timer
    bgrSEL1 = "SELECTED"
    'aktsumtot = oRec6("aktbudget") * budgettimer 
    case 2 '** Stk
    bgrSEL2 = "SELECTED"
    'aktsumtot = oRec6("aktbudget") * oRec6("antalstk")
    end select
    

    aktTypeThis = oRec6("fakturerbar")
    
    
    strAktListe = strAktListe & "<tr bgcolor="&bgcolor&">"_
    &"<td style=""width:175px;"">"
    
    strAktListe = strAktListe & "<input type=""hidden"" name=""SortOrder"" class=""SortOrder"" value='"&oRec6("sortorder")&"' />"
	strAktListe = strAktListe & "<input type=""hidden"" name=""rowId"" value='"&oRec6("id")&"' />"
        

    
    if oRec6("fakturerbar") <> 5 then
    strAktListe = strAktListe &"<input type=""text"" name=""FM_aktnavn"" value='"& oRec6("navn") &"' style='width:150px; font-size:9px; font-family:arial;' />"
    else
    strAktListe = strAktListe &"<input type=""hidden"" name=""FM_aktnavn"" value='"& oRec6("navn") &"'/>"
    strAktListe = strAktListe &"<input type=""text"" name=""disFM_aktnavn"" DISABLED value='"& oRec6("navn") &"' style='width:150px; font-size:9px; font-family:arial;' />"
    end if
    strAktListe = strAktListe &"<input type=""hidden"" name=""FM_aktnavn"" value='#' />"
	
    
    'if len(trim(oRec6("fase"))) <> 0 then
    'strAktListe = strAktListe & " ("& oRec6("fase") &")</td>"
    'end if
    
    strAktListe = strAktListe &"<td>"_
    &"<input class='aktFase_"& trim(lcase(thisFase)) &"' id='FM_aktfas_"& oRec6("id") &"' type=""text"" name=""FM_aktfase"" value='"& lcase(trim(replace(thisFase, "_", " "))) &"' style='width:75px; font-size:9px; font-family:arial;'  />"_
    &"</td>"

    strAktListe = strAktListe &"<input class='aktFase_"& trim(lcase(thisFase)) &"' id=""FM_aktfas_h_"& oRec6("id") &""" type=""hidden"" name="""" value='"& lcase(trim(thisFase)) &"' />"
    
	strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktfase"" value='#' />"
	strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktid"" value='"& oRec6("id") &"' />"
	
	strAktListe = strAktListe &"<td>"
	 
	select case oRec6("aktstatus")
	case 1
	stCHK0 = ""
	stCHK1 = "SELECTED"
	stCHK2 = ""
	'selbgcol = "#DCF5BD"
	case 2
	stCHK0 = ""
	stCHK1 = ""
	stCHK2 = "SELECTED"
	'selbgcol = "#cccccc"
	case 0
	stCHK0 = "SELECTED"
	stCHK1 = ""
	stCHK2 = ""
	'selbgcol = "Crimson"
	end select
	
	strAktListe = strAktListe &"<select name=""FM_aktstatus"" id='af_"&lcase(trim(thisFase))&"_"&fa&"' style='background-color:"&selbgcol&"; font-family:arial; width:40px; font-size:9px;'>"_
	&"<option value=""1"" "&stCHK1&">Aktiv</option>"_
	&"<option value=""0"" "&stCHK0&">Lukket</option>"_
	&"<option value=""2"" "&stCHK2&">Passiv</option>"_
	&"</select>"_
	&"</td>"

    call akttyper(oRec6("aty_id"), 1)
    

    if cint(timesimon) = 1 then
        fcLink = "../to_2015/timbudgetsim.asp?FM_fy="&year(now)&"&FM_visrealprdato=1-1-"&year(now)&"&FM_sog="& jobnrThis '& "&jobid="& jobid &"&func=forecast"
    else
        fcLink = "ressource_belaeg_jbpla.asp?FM_sog="& jobnrThis
    end if

    if cdbl(oRec6("budgettimer")) < cdbl(resTimerThis) then 'budgettimer
    resBgColor = "crimson"
    resAColor = "#ffffff"
    else
    resBgColor = ""
    resAColor = "#5582d2"
    end if

    strAktListe = strAktListe &"<td>"& left(akttypenavn, 7) &".</td>"

    if cint(budgetakt) = 1 then
    strAktListe = strAktListe &"<td><select name='FM_aktkonto' id='FM_aktkonto_"& oRec6("id") &"' class='FM_aktkonto FM_aktkonto_"& oRec6("kontoid") &"' style='font-family:arial; width:40px; font-size:9px;'><option value="& oRec6("kontoid") &">"& oRec6("kontonr") &"</option>"& strKontoplan &"</select></td>"
   
         '*** Opretter kontototal felter
         if instr(isKontoWrt, "#"& oRec6("kontoid") &"#") = 0 then
         kontototalsTYP = kontototalsTYP & "<input id='FM_kontot_"& oRec6("kontoid") &"' class='FM_kontot' type='hidden' value='0'>"
         kontototalsTYP = kontototalsTYP & "<input id='FM_kontotp_"& oRec6("kontoid") &"' class='FM_kontotp' type='hidden' value='0'>"
         kontototalsTYP = kontototalsTYP & "<input id='FM_kontosum_"& oRec6("kontoid") &"' class='FM_kontosum' type='hidden' value='0'>"  
         isKontoWrt = isKontoWrt & ",#"& oRec6("kontoid") &"#"
         end if
         
   else

           strAktListe = strAktListe &"<input type='hidden' name='FM_aktkonto' id='FM_aktkonto_"& oRec6("id") &"' value='"& oRec6("aktkonto") &"'>"
         
    end if
	
	'strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktstatus"" value='"& oRec6("aktstatus") &"' />"
	if cint(budgetakt) = 1 then

         if cint(aktTypeThis) = 90 then 'SUM linje E1 OKO
            kontoTClas = "kontoTclsSum" '_"& oRec6("kontoid")
            kontoTpClas = "kontoTpclsSum"
            kontoSumClas = "kontoSumclsSum"
            totBelDiasbledlinje = "" 'DISABLED
            totBelDiasbledlinjeVal = 1
            akttotprisNameLinje = "FM_akttotpris"
            akttotprisNameLinjeBGCol = "#999999"
            akttotprismxlnght = 0

         else
            kontoTClas = "kontoTclslinje"
            kontoTpClas = "kontoTpclslinje"
            kontoSumClas = "kontoSumclslinje"
            totBelDiasbledlinje = ""
            totBelDiasbledlinjeVal = 0
            akttotprisNameLinje = "FM_akttotpris"
            akttotprisNameLinjeBGCol = ""
            akttotprismxlnght = 30
         end if

     else

         kontoTClas = ""
         kontoTpClas = ""
         kontoSumClas = ""
         totBelDiasbledlinje = ""
         totBelDiasbledlinjeVal = 0
         akttotprisNameLinje = "FM_akttotpris"
         akttotprisNameLinjeBGCol = ""
         akttotprismxlnght = 30

    end if
	
    strAktListe = strAktListe & "<input type=""hidden"" id=""budgetakt_"& oRec6("id") &""" name="""" value='"&totBelDiasbledlinjeVal&"' />"
	strAktListe = strAktListe &"<td><input class='jq_akttimerstk jq_akttimer "& kontoTClas &"' id='FM_akttim_"& oRec6("id") &"' name=""FM_akttimer"" type=""text"" value='"& formatnumber(oRec6("budgettimer"), 2) &"' style='width:40px; text-align:right; font-size:9px; font-family:arial;'/></td>"
    
    if cint(budgetakt) = 0 then
    strAktListe = strAktListe &"<td><input class='jq_akttimerstk' id='FM_aktstk_"& oRec6("id") &"' name=""FM_aktantalstk"" type=""text"" value='"& formatnumber(oRec6("antalstk"), 2) &"' style='width:40px; text-align:right; font-size:9px; font-family:arial;' /></td>"
    else
    strAktListe = strAktListe &"<input class='jq_akttimerstk' id='FM_aktstk_"& oRec6("id") &"' name=""FM_aktantalstk"" type=""hidden"" value=""0"" />"
    end if

    strAktListe = strAktListe &"<input name='af_timer_"&oRec6("id")&"' id='af_timer_"&lcase(trim(thisFase))&"_"&fa&"' type=""hidden"" value='"& formatnumber(oRec6("budgettimer"), 2) &"' />"_
    &"<input name='af_sum_"& oRec6("id") &"' id='af_sum_"&lcase(trim(thisFase))&"_"&fa&"' type=""hidden"" value='"& formatnumber(oRec6("aktbudgetsum"), 2) &"' />"_
    &"<input name='FM_sum_aid_fa_"&oRec6("id")&"' id='FM_sum_aid_fa_"&oRec6("id")&"' type=""hidden"" value='"&fa&"' />"
    
    if cint(budgetakt) = 0 then
    strAktListe = strAktListe &"<td><select class='bgr' id='FM_aktbgr_"& oRec6("id") &"' name=""FM_aktbgr"" style='width:45px; font-size:9px; font-family:arial;'>"_
    &"<option value=0 "& bgrSEL0 &">Ingen</option>"_
    &"<option value=1 "& bgrSEL1 &">Timer</option>"_
    &"<option value=2 "& bgrSEL2 &">Stk.</option>"_
    &"</select></td>"
    else
    strAktListe = strAktListe &"<input type='hidden' id='FM_aktbgr_"& oRec6("id") &"' value='1' name='FM_aktbgr'/>"
    end if


    strAktListe = strAktListe &"<td><input class='jq_akttimerstk "& kontoTpClas &"' id='FM_aktpri_"& oRec6("id") &"' type=""text"" name=""FM_aktpris"" value="& formatnumber(oRec6("aktbudget"), 2) &" style='width:50px; text-align:right; text-align:right; font-size:9px; font-family:arial;' /></td>"_
    &"<td class=lille style='white-space:nowrap;'> = <input class='jq_akttotal "& kontoSumClas &"' id='"& akttotprisNameLinje &"_"&oRec6("id")&"' name="& akttotprisNameLinje &" type=""text"" value="&formatnumber(oRec6("aktbudgetsum"), 2)&" maxlenght='"& akttotprismxlnght &"' style='width:60px; text-align:right; font-size:9px; font-family:arial; background-color:"& akttotprisNameLinjeBGCol &";' "& totBelDiasbledlinje &" /></td>"
    
    if totBelDiasbledlinje = "DISABLED" then
    strAktListe = strAktListe &"<input id='FM_akttotpris_"&oRec6("id")&"' type=""hidden"" name=""FM_akttotpris"" value='"& formatnumber(oRec6("aktbudgetsum"), 2) &"' />"
    end if

    strAktListe = strAktListe &"<td align=right><a href='aktiv.asp?menu=job&func=red&id="&oRec6("id")&"&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir=job2&nomenu=1' target=""_blank"" style=""font-size:9px; color:#6CAE1C;"">"_
    &"[R]</a></td>"_
    &"<td align=right style=""background-color:"& resBgColor &";""><a href="& fcLink &" target=""_blank"" style=""color:"& resAColor &"; font-size:9px; font-family:arial;"">"& formatnumber(resTimerThis, 0) &"</a>"_
    &"</td>"_
    &"<td align=right><input name='FM_slet_aid_"&c&"' id='af_sl_"&lcase(trim(thisFase))&"_"&fa&"' type=""checkbox"" value=""1"" />"_
    &"<td align=right><img src=""../ill/pile_drag.gif"" alt='Klik, træk og sorter aktivitet' border=""0"" class='drag' /></td>"_
    &"<input name=""FM_slet_aid"" type=""hidden"" value='"&c&"' />"_
    &"</tr>"
    'businesspeople
    end if 'aktstatus (før sum beregning) 

    lastFaseTimer = lastFaseTimer + oRec6("budgettimer")
    lastFaseSum = lastFaseSum + oRec6("aktbudgetsum") 
    
    'if len(trim(oRec6("fase"))) <> 0 then
    'lastFase = lcase(trim(oRec6("fase")))
    'else
    lastFase = thisFase
    'end if
    
    aktbudgetsamlet = aktbudgetsamlet + oRec6("aktbudgetsum") 
    akttfaktimtildelt = akttfaktimtildelt + oRec6("budgettimer")  
    
    fa = fa + 1
    c = c + 1

    
	
	
	oRec6.movenext
	wend
	oRec6.close
	
	'thisFase = lastFase
	
	
	call lastFaseSumSub
	    
	'akttfaktimtildelt  
    '*** Totaler ***'
    if cint(budgetakt) = 1 then
    response.write kontototalsTYP
    end if

    strAktListe = strAktListe & "<input id=""jq_akttottimer"" value='"& formatnumber(akttfaktimtildelt, 2) &"' type=""hidden"" />"
    strAktListe = strAktListe & "<input id=""jq_akttotsum"" value='"& formatnumber(aktbudgetsamlet, 2) &"' type=""hidden"" />"
    
    strAktListe = strAktListe & "<input id=""fatot_ialt"" value='"&faTot-1&"' type=""hidden"" />" 
    
    strAktListe = strAktListe & "</table>"
    '</td></tr>
    
   
    
    
end function   




'sub lastFaseTr
'    strAktListe = strAktListe & "<tr bgcolor=""#8CAAE6""><td colspan=9 style='padding-left:5px;'><b>"& thisFase &"</b></td></tr>" 
'end sub

'sub LastFaseTrSum

'end sub



function opdateraktliste(jobid, aktids, aktnavn, akttimer, aktantalstk, aktfaser, aktbgr, aktpris, aktstatus, akttotpris, aktslet, aktslet_aids, aktkonto)

        
    'Response.Write aktnavn & "<br>"    
    'Response.Write aktstatus
	'Response.end
	
	aktnavn = split(aktnavn, ", #,")
	akttimer = split(akttimer, ", ")
	
    aktkonto = split(aktkonto, ", ") 
         
    aktantalstk = split(aktantalstk, ", ")
	aktpris = split(aktpris, ", ")
	aktids = split(aktids, ",")
	aktstatus = split(aktstatus, ", ")
	
	aktfaser =  split(aktfaser, ", #,")
	aktbgr = split(aktbgr, ",")
	
	'Response.Write request("FM_akt_totpris")
	'Response.end
	
	'akttotpris = split(akttotpris, ", #")
	akttotpris = split(akttotpris, ", ")
	
	
	
	aktslet = split(aktslet, ", ")
	
	'Response.Write aktSlet
	'Response.end
	
	
	
	for t = 0 to UBOUND(aktids)
	err = 0
	
	    akttimer(t) = replace(akttimer(t), ".", "")
	    akttimer(t) = replace(akttimer(t), ",", ".")
	    
	    if len(trim(akttimer(t))) <> 0 then
	    akttimer(t) = akttimer(t)
	    else
	    akttimer(t) = 0
	    end if
	    
	    call erDetInt(akttimer(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktantalstk(t) = replace(aktantalstk(t), ".", "")
	    aktantalstk(t) = replace(aktantalstk(t), ",", ".")
	    
	    if len(trim(aktantalstk(t))) <> 0 then
	    aktantalstk(t) = aktantalstk(t)
	    else
	    aktantalstk(t) = 0
	    end if
	    
	    call erDetInt(aktantalstk(t))
	        
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktpris(t) = replace(aktpris(t), ".", "")
	    aktpris(t) = replace(aktpris(t), ",", ".")
	    
	    if len(trim(aktpris(t))) <> 0 then
	    aktpris(t) = aktpris(t)
	    else
	    aktpris(t) = 0
	    end if
	    
	    call erDetInt(aktpris(t))
	      
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	      
	      'akttotpris(t) = replace(akttotpris(t), "#", "")
	      akttotpris(t) = replace(akttotpris(t), ".", "")
	      akttotpris(t) = replace(akttotpris(t), ",", ".")
	      
	     
	      'Response.Write "err:" & err & "<br>"
	      'Response.Write "int:" & isInt &"<br>"
	      
	      'Response.flush
	    aktFaser(t) = trim(aktFaser(t))
		call illChar(aktFaser(t))
		aktFaser(t) = vTxt
		

        aktkonto(t) = aktkonto(t) 
	    
	    aktNavn(t) = trim(aktNavn(t))
	    aktNavn(t) = replace(aktNavn(t), "'", "")
	    
	    if cint(err) = 0 then
	    
	    aktSletval = 0
	    aktSletval = request("FM_slet_aid_"& aktSlet(t) &"")
		
		if aktSletval <> "1" then
		
		strSQL = "UPDATE aktiviteter SET navn = '"& aktNavn(t) &"', aktstatus = " & aktstatus(t) & ", "_
		&" budgettimer = "& akttimer(t) &", aktbudget = "& aktpris(t) &", antalstk = "& aktantalstk(t) &", "
        
        if len(trim(aktFaser(t))) <> 0 then
        strSQL = strSQL &" fase = '"& aktFaser(t) &"', "
        else
        strSQL = strSQL &" fase = Null, "
        end if

        if len(trim(aktkonto(t))) <> 0 then
        strSQL = strSQL &" aktkonto = "& aktkonto(t) &", "
        else
        strSQL = strSQL &" aktkonto = 0, "
        end if
		
        strSQL = strSQL &" bgr = "& aktBgr(t) &", aktbudgetsum = "& akttotpris(t) &""_
		&" WHERE id = "& aktids(t)
		
		'Response.write strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)

            '*** Opdaterer timereg. tabellen ***'
            oConn.execute("UPDATE timer SET "_
            & " TAktivitetNavn = '"& aktNavn(t) &"'"_
            & " WHERE TAktivitetId = "& aktids(t) & "")


		else
		    
		    call delakt(aktids(t))
		
		end if
		
		'** Sync job ***'
		'if len(trim(request("opdjobv"))) <> 0 then 
		'opdjobv = 1		
		'else
		'opdjobv = 0
		'end if		
		
		
		'if cint(opdjobv) = 1 then
        '    call syncJob(jobid)
        'end if
		
		end if
	
	next




end function



function delakt(id)
	
	strSQL = "SELECT id, navn FROM aktiviteter WHERE id = "& id &"" 
	oRec5.open strSQL, oConn, 3
	while not oRec5.EOF 
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("akt", oRec5("id"), 0, oRec5("navn"), session("mid"), session("user"))
		
	oRec5.movenext
	wend
	oRec5.close
		
	
	
	oConn.execute("DELETE FROM aktiviteter WHERE id = "& id &"")
	oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& id &"")
	
end function






function tilknytstamakt(a, intAktfavgp, strAktFase, opretAlleAktiGrp)
		'Response.write "her a "& a &", og grp:" & intAktfavgp & "<br>"
		'Response.flush
		                ause = trim(intAktfavgp) ' a ' + 1
		                
		                '** fase må ikke indeholde mellemrum **'
		                strAktFase = trim(strAktFase)
		                call illChar(strAktFase)
		                strAktFase = trim(vTxt)
		                
						if intAktfavgp <> 0 then
						
                      
                        varjobIdUse = varjobId
                        
						
						strSQL = "select id, navn, fakturerbar, beskrivelse, budgettimer, fomr, faktor, "_
						&" aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
						&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
				        &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, "_
                        &" brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto "_
						&" FROM aktiviteter WHERE aktFavorit = "& intAktfavgp &" AND job = 0"
						

                        'response.write strSQL &"<br>j:"& j &"<br>andreJob: "& andreJob
                        'response.flush

						    oRec2.open strSQL, oConn, 3
							while not oRec2.EOF
							
							aktid = oRec2("id")
                
                           	aktNavn = trim(request("FM_stakt_navn_"& ause &"_"& oRec2("id"))) 'oRec2("navn")
						    aktNavn = replace(aktNavn, "'", "")

         					aktFakbar = oRec2("fakturerbar")
							aktFomr = oRec2("fomr")
							aktFaktor = replace(oRec2("faktor"), ",", ".")
							aktstatus = oRec2("aktstatus") 'request("FM_stakt_status_"& ause &"_"& oRec2("id")) 
							tidslaas = oRec2("tidslaas")
							tidslaas_st = oRec2("tidslaas_st")
							tidslaas_sl = oRec2("tidslaas_sl")
							
							tidslaas_man = oRec2("tidslaas_man")
							tidslaas_tir = oRec2("tidslaas_tir")
							tidslaas_ons = oRec2("tidslaas_ons")
							tidslaas_tor = oRec2("tidslaas_tor")
							tidslaas_fre = oRec2("tidslaas_fre")
							tidslaas_lor = oRec2("tidslaas_lor")
							tidslaas_son = oRec2("tidslaas_son")
							
							'** behold eller overskriv fase **'
							 if cint(opretAlleAktiGrp) = 0 then
                
							    if len(trim(request("FM_stakt_fase_"& ause &"_"& oRec2("id")))) <> 0 then
							    strFase = request("FM_stakt_fase_"& ause &"_"& oRec2("id")) 'replace(oRec2("fase"), "'", "")
							    else
							    strFase = ""
							    end if
							
                            else

                                if isNull(oRec2("fase")) <> true then 
                                strFase = replace(oRec2("fase"), "'", "")
                                else
                                strFase = ""
                                end if

                            end if
							
                            
                            call illChar(strFase)
		                    strFase = vTxt
                            strFase = trim(strFase)

							if len(trim(oRec2("beskrivelse"))) <> 0 then
							
							'beskrivelse = replace(oRec2("beskrivelse"), "'", "''")
							'call htmlparseCSV(beskrivelse)
							'beskrivelse = htmlparseCSV
							
							beskrivelse = replace(oRec2("beskrivelse"), "'", "''")
							
							beskrivelse = replace(beskrivelse, "<span", "")
							beskrivelse = replace(beskrivelse, "</span>", "")
							beskrivelse = replace(beskrivelse, "<div", "")
							beskrivelse = replace(beskrivelse, "</div>", "")
							beskrivelse = replace(beskrivelse, "<table", "")
							beskrivelse = replace(beskrivelse, "</table>", "")
							beskrivelse = replace(beskrivelse, "<tr", "")
							beskrivelse = replace(beskrivelse, "</tr>", "")
							beskrivelse = replace(beskrivelse, "<td", "")
							beskrivelse = replace(beskrivelse, "</td>", "")
							
							''beskrivelse = replace(beskrivelse, "<p>", "")
							''beskrivelse = replace(beskrivelse, "</p>", "")
							
							else
							
							beskrivelse = ""
							
							end if
							
							'antalstk = replace(oRec2("antalstk"), ",", ".")
							
							
							
							
								if len(tidslaas) <> 0 then
								tidslaas = tidslaas
								else
								tidslaas = 0
								end if
								
								
								if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
								tidslaas_st = left(formatdatetime(tidslaas_st, 3), 5)
								tidslaas_sl = left(formatdatetime(tidslaas_sl, 3), 5)
								else
								tidslaas_st = "07:00:00"
								tidslaas_sl = "23:30:00"
								end if
	
							
							
							if len(aktFaktor) <> 0 then
							aktFaktor = aktFaktor
							else
							aktFaktor = 0
							end if
							
						
							
							sortorder = oRec2("sortorder")
							
							'*************************************************
							'** Henter fra forkalkulation // eller fra job ***
							'*************************************************
                          
							
							

							'aktBudgettimer = replace(request("FM_stakt_timer_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							aktBudgettimer = replace(oRec2("budgettimer"), ",", ".")
							
							antalstk = oRec2("antalstk") 'replace(request("FM_stakt_stk_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							antalstk = replace(antalstk, ",", ".")
							
							bgr = oRec2("bgr") 'request("FM_stakt_bgr_"& ause &"_"& oRec2("id") &"") 'oRec2("bgr")
							
							aktBudget = replace(oRec2("aktbudget"), ",", ".")
							
							aktbudgetsum = oRec2("aktbudgetsum") 'replace(request("FM_stakt_totaktsum_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							aktbudgetsum = replace(aktbudgetsum, ",", ".")					
                



                    
                            call erDetInt(antalstk)
							if isInt > 0 OR len(trim(antalstk)) = 0 then
							antalstk = 0
							else
							antalstk = antalstk
							end if

                            

                            call erDetInt(aktBudgettimer)
							if isInt > 0 OR len(trim(aktBudgettimer)) = 0 then
							aktBudgettimer = 0
							else
							aktBudgettimer = aktBudgettimer
							end if           


							call erDetInt(aktBudget)
							if isInt > 0 OR len(trim(aktBudget)) = 0 then
							aktBudget = 0
							else
							aktBudget = aktBudget
							end if
							
							
							
							call erDetInt(aktbudgetsum)
							if isInt > 0 OR len(trim(aktbudgetsum)) = 0 then
							aktbudgetsum = 0
							else
							aktbudgetsum = aktbudgetsum
							end if 
         
         
         
         
         
         
         
                           easyreg = oRec2("easyreg")


                            brug_fasttp = oRec2("brug_fasttp")
                            brug_fastkp = oRec2("brug_fastkp")
                            fasttp = replace(oRec2("fasttp"), ",", ".")
                            fasttp_val = oRec2("fasttp_val")
                            fastkp = replace(oRec2("fastkp"), ",", ".")
                            fastkp_val = oRec2("fastkp_val")
                         
                            avarenr = oRec2("avarenr")

                            kostpristarif = oRec2("kostpristarif")

                            aktkonto = oRec2("aktkonto")
						    
						    '*** Tilføj fravalgt ***'
						    'Response.write "chk: "& request("FM_stakt_tilfoj_"& ause &"_"& oRec2("id") &"") & "("& ause &"- "& oRec2("id") &")<br>"
						    'Response.flush
						    if request("FM_stakt_tilfoj_"& ause &"_"& oRec2("id") &"") = "1" OR cint(opretAlleAktiGrp) = 1 then


                          
                                    aj = 1 '** Hvad bliver opdateret? Aktivitet eller job 1 = job, 2 = akt.
                                    aid = aktid '** Aktid
                                    nedfod = pgrp_arvefode '*** ignorer fød/nedarv -1 = ignorer, 0 Nedarv,  1 = fød job
                                    'fm_alle = -1 '*** bruges kun ved oprettelse af enkelt aktivitet. Skal den følge job's projektgrupper eller beholde sinde egne. 1 = Behold, 0/"" = Nedarv fra job, - 1 ignorer (ved jobopr / rediger job)
                                    jid = varjobId 'jobid
                                

                            



                                call tilfojProGrp(func,aj,jid,aid,nedfod,strProjektgr1,strProjektgr2,strProjektgr3,strProjektgr4,strProjektgr5,strProjektgr6,strProjektgr7,strProjektgr8,strProjektgr9,strProjektgr10,firstLoop)
                           
                            

						    
							strSQLins = "INSERT INTO aktiviteter "_
							&" (navn, dato, editor, job, fakturerbar, "_
							&" projektgruppe1, projektgruppe2, projektgruppe3, "_
							&" projektgruppe4, projektgruppe5, projektgruppe6, "_
							&" projektgruppe7, projektgruppe8, projektgruppe9, "_
							&" projektgruppe10, aktstartdato, aktslutdato, "_
							&" budgettimer, fomr, faktor, aktBudget, aktstatus, tidslaas, "_
							&" tidslaas_st, tidslaas_sl, beskrivelse, antalstk, "_
							&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
				            &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, brug_fasttp, "_
                            &" brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto "_
				            &" ) VALUES "_
							&"('"& aktNavn &"', "_
							&"'"& strDato &"', "_ 
							&"'"& strEditor &"', "_
							&""& varjobId  &", "_ 
							&""& aktFakbar &", "_
							&""& strProjektgr1 &", "_ 
							&""& strProjektgr2 &", "_ 
							&""& strProjektgr3 &", "_ 
							&""& strProjektgr4 &", "_ 
							&""& strProjektgr5 &", "_
							&""& strProjektgr6 &", "_ 
							&""& strProjektgr7 &", "_ 
							&""& strProjektgr8 &", "_ 
							&""& strProjektgr9 &", "_ 
							&""& strProjektgr10 &", "_     
							&"'"& startDato &"', "_ 
							&"'"& slutDato &"', "_
							&""& aktBudgettimer & ", "& aktFomr &", "_
							&""& aktFaktor &", "& aktBudget &", "& aktstatus &", "_
							&""&tidslaas&", '"&tidslaas_st&"', '"&tidslaas_sl&"', '"& beskrivelse &"', "& antalstk &","_
							&""& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
				            &""& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son &", "

                            if len(trim(strFase)) <> 0 then
				            strSQLins = strSQLins &" '"& strFase &"', "
                            else
                            strSQLins = strSQLins &" NULL , "
                            end if
                            
                            strSQLins = strSQLins & sortorder &", "& bgr &", "& aktbudgetsum &", "& easyreg &", "& brug_fasttp &","& brug_fastkp &","& fasttp &","& fasttp_val &","& fastkp &","& fastkp_val &", '"& avarenr &"', '"& kostpristarif &"', "& aktkonto &""_
				            &")"
							
							
							'Response.write strSQLins & "<br><br>"
							'Response.flush
							oConn.execute(strSQLins)
							
							'*** Henter det netop oprettede akt-id ***
							strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
							oRec3.open strSQLid, oConn, 3
							if not oRec3.EOF then
							useNewAktid = oRec3("id")
							end if
							oRec3.close
							
							if len(useNewAktid) <> 0 then
							useNewAktid = useNewAktid
							else
							useNewAktid = 0
							end if
							
                            '**** Ved opret aktiviteter på job ****'
							'**** Overfører de timepriser hver enkelt stamaktivitet er født med, for hver enkelt medarbejder ****
							'**** Eller nedarv fra job 
							'**** 1: Behold medarbejdertimepriser på aktiviteter   
							'**** 0: Nedarv fra job
                            
                            tpAlt = 6
                            tpPris = 0

                            'if instr(lcase(lto), "epi") then
                            'tifojikke 
							if (request("FM_timepriser") = 1 OR oRec2("fakturerbar") = 5) then '** 5: Kørsel
								
								strSQLfindtp = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE aktid = "& aktid
								oRec3.open strSQLfindtp, oConn, 3 
								while not oRec3.EOF 
									
                                    if oRec2("fakturerbar") = 5 then '** Altid 3,56
                                    tpAlt = 6
                                    tpPris = 0
                                            
                                            strSQLtp = "SELECT kmpris FROM licens WHERE id = 1"
                                            oRec6.open strSQLtp, oConn, 3
                                            if not oRec6.EOF then
                                            tpPris = replace(oRec6("kmpris"), ",",".")
                                            end if
                                            oRec6.close

                                    else
                                    tpAlt = replace(oRec3("timeprisalt"), ",", ".")
                                    tpPris = replace(oRec3("6timepris"), ",", ".")
                                    end if 

									strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES "_
									&" ("& varjobId &", "& useNewAktid &", "& oRec3("medarbid") &", "_
									&" "& tpAlt &", "& tpPris &")"
									oConn.execute(strSQLtp)
									
									'Response.write strSQLtp
									'Response.flush
								oRec3.movenext
								wend
								oRec3.close  
								
										
							end if
							
							   
                               
                               
                               
                                '**** Forretningsområder tildeles (kan overskrives af tilvalg på job hvis sync akt. = 1) ****
                                
                                
                                '*** nulstiller ***
                                'strSQLdela = "DELETE FROM fomr_rel WHERE for_aktid = " & useNewAktid
                                'oConn.execute(strSQLdela) 

                                strSQLa = "SELECT for_fomr, for_faktor FROM fomr_rel WHERE for_aktid = " & aktid
                                'Response.Write strSQLa & "<br>"
                                'Response.end
                                oRec3.open strSQLa, oConn, 3
                                while not oRec3.EOF 

                                    strSQLfomrai = "INSERT INTO fomr_rel "_
                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                    &" VALUES ("& oRec3("for_fomr") &", "& varJobId &","& useNewAktid &", "& oRec3("for_faktor") &")"

                                    'Response.Write strSQLfomrai & "<br>"
                                    'Response.flush
                                    oConn.execute(strSQLfomrai)

                                oRec3.movenext
                                wend
                                oRec3.close

                               
                               'Response.end

							end if '** tilføj fravalgt
							
					    oRec2.movenext
					    wend
						oRec2.close

                    
                        
                    

						end if 'intAktfavgp


                        'response.end
		
		end function




    public akttypenavn
	function akttyper(akttype, visning)
	
	if len(trim(akttype)) <> 0 then
	akttype = akttype
	else
	akttype = 0
	end if
	
	    select case akttype 
	    case 1 
	    akttypenavn = global_txt_129 '"Fakturerbar"
	    case 5 '2
	    akttypenavn = global_txt_130 '"Km"
	    case 2 '0
	    akttypenavn = global_txt_131 '"Ikke fakturerbar"
	    case 6
	    akttypenavn = replace(global_txt_132, "|", "&") '"Salg & NewBizz"
	    case 7
	    akttypenavn = global_txt_147 '"Flex brugt"
	    case 8
	    akttypenavn = global_txt_148 '"Læge / Massage / Fysioterapi"
	    case 9
	    akttypenavn = global_txt_150 '"Pause"
	    case 10
	    akttypenavn = global_txt_133 '"Frokost / Pause"
	    case 11
	    akttypenavn = global_txt_134 '"Ferie planlagt"
	    case 12
	    akttypenavn = global_txt_136 '"Ferie fridage optjent"
        case 13
	    akttypenavn = global_txt_137 '"Ferie fridage brugt"
	    case 14
	    akttypenavn = global_txt_135 '"Ferie afholdt"
	    case 15
	    akttypenavn = global_txt_143 '"Ferie optjent"
	    case 16
	    akttypenavn = global_txt_156 '"Ferie udbetalt"
	    case 17
	    akttypenavn = global_txt_149 '"Ferie fridage udbetalt"
	    case 18
	    akttypenavn = global_txt_164 '"Ferie Fridage Planlagt"
	    case 19
	    akttypenavn = global_txt_165 '"Ferie u. Løn"
	    case 20
	    akttypenavn = global_txt_138 '"Syg"
	    case 21
	    akttypenavn = global_txt_139 '"Barn syg"
        case 22
	    akttypenavn = global_txt_171 '"Barsel"

        case 23
	    akttypenavn = global_txt_172 '"Omsorgsdage"
        case 24
	    akttypenavn = global_txt_173 '"Seniortimer"
        case 25
	    akttypenavn = global_txt_174 '"1 maj timer"

        case 26
	    akttypenavn = global_txt_188 '"Aldersreduktion planlagt"


        case 27
	    akttypenavn = global_txt_185 '"Aldersreduktion optjent"
        case 28
	    akttypenavn = global_txt_186 '"Aldersreduktion brugt"
        case 29
	    akttypenavn = global_txt_187 '"Aldersreduktion udbetalt"

	    case 30
	    akttypenavn = global_txt_144 '"Afspad. optjent"
        case 31
	    akttypenavn = global_txt_145 '"Afspad. brugt"
	    case 32
	    akttypenavn = global_txt_146 '"Afspad. udbetalt"
	    case 33
	    akttypenavn = global_txt_159 '"Afspad. Ø udbetalt"
	    case 50
	    akttypenavn = global_txt_166 '"Dag"
	    case 51
	    akttypenavn = global_txt_151 '"Nat"
	    case 52
	    akttypenavn = global_txt_152 '"Weekend"
	    case 53
	    akttypenavn = global_txt_153 '"Afspad. optjent"
	    case 54
	    akttypenavn = global_txt_154 '"Weekend Nat"
	    case 55
	    akttypenavn = global_txt_155 '"Weekend Aften"
	    case 60
	    akttypenavn = global_txt_157 '"Ad-hoc"
	    case 61
	    akttypenavn = global_txt_158 '"Stk. Antal"
	    case 81
	    akttypenavn = global_txt_160 '"Læge"
	    case 90
	    akttypenavn = global_txt_161 '"E1"
	    case 91
	    akttypenavn = global_txt_162 '"E2"

        case 111
	    akttypenavn = global_txt_175 '"Ferie overført"
	    case 112
	    akttypenavn = global_txt_176 '"Ferie optjent u. løn"

         case 113
	    akttypenavn = global_txt_177 '"Korrektion Komme 7Gå"

        case 114
	    akttypenavn = global_txt_178 '"Korrektion Real"

        case 115
	    akttypenavn = global_txt_179 '"Tjenestefri"

        case 120
	    akttypenavn = global_txt_189 '"Omsorgsdag 2 planlagt"

        case 121
	    akttypenavn = global_txt_190 '"Omsorgsdag 10 planlagt"

        case 122
	    akttypenavn = global_txt_191 '"Omsorgsdag K planlagt"

        case 123
	    akttypenavn = global_txt_192 '"Ulempe1706Udb"

        case 124
	    akttypenavn = global_txt_193 '"UlempeWUdb"
        case 125
	    akttypenavn = global_txt_194 '"Rejse"
       
        case 92
	    akttypenavn = global_txt_195 '"E3"

	    case else
	    akttypenavn = "-"
	    end select
	
	
	
	end function



         

	  
    
%>