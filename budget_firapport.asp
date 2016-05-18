<%'response.buffer = true %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%


         '**** Søgekriterier AJAX ***************************************
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_tilfojper"

                   budgetid = request("budgetid") '4

                   strSQLins = "INSERT INTO budget_job_per (budgetid, beskrivelse, budget_app) VALUES ("& budgetid &", 'ww', 0)"
                   oConn.execute(strSQLins)


                   'response.write strSQLins
                   'response.end

        case "FN_valutakurs"

                   valutaid = request("valutaid")

                   call valutaKurs(valutaid)

                   if dblKurs <> 100 then
                   response.write dblKurs/100
                   else
                   response.write 1
                   end if
                   'response.end


        end select
        Response.end
        end if 'JQ AJAX Slut 
        '****************************************************************


    %>

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/errors/error_inc.asp"-->
    <!--#include file="../inc/regular/global_func.asp"-->
    <!--#include file="inc/budget_firapport_inc.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->

    <!-- <script src="../inc/jquery/budget_firapport_jav.js" type="text/javascript"></script> -->
    <script src="js/budget_firapport_jav.js" type="text/javascript"></script>
    

    <%

    media = request("media")
    func = request("func")
    
    if len(trim(request("id"))) <> 0 then
          id = request("id")
          else
          id = 0
          end if   
    
 %>



<%
  if media <> "export" then    
        
    call menu_2014 
    
 end if%>

<div id="wrapper">
  <div class="content">


      <%select case func

      
        case "slet" 


          'response.write request("FM_budgetid_del")
          'response.flush

          if len(trim(request("FM_budgetid_del"))) <> 0 then
          budgetids = split(request("FM_budgetid_del"), ", ")
          else
          budgetids = 0
          end if

          for s = 0 TO UBOUND(budgetids)

            if len(trim(budgetids(s))) <> 0 then
            budgetids(s) = budgetids(s)
            else
            budgetids(s) = 0
            end if

            strSQLDelA = "DELETE FROM budget_job WHERE id =" & budgetids(s)
            oConn.execute(strSQLDelA)

            strSQLDelB = "DELETE FROM budget_job_per WHERE budgetid =" & budgetids(s)
            oConn.execute(strSQLDelB)

            strSQLDelC = "DELETE FROM budget_job_exp WHERE budgetid =" & budgetids(s)
            oConn.execute(strSQLDelC)

            strSQLDelD = "DELETE FROM budget_job_exp_rel WHERE budgetid =" & budgetids(s)
            oConn.execute(strSQLDelD)


          next

          'response.end

          response.redirect "budget_firapport.asp"


       case "dbopr", "dbred"


       

          if len(trim(request("FM_budgetnavn"))) then
          budgetnavn = replace(request("FM_budgetnavn"), "'", "''")
          else
          budgetnavn = ""
          end if

          jobid = request("FM_kunde")

          status = request("FM_status")

          if len(trim(request("FM_refno"))) <> 0 then
          refno = request("FM_refno")
          else
          refno = ""
          end if

          if len(trim(request("FM_prodno"))) <> 0 then
          prodno = request("FM_prodno")
          else
          prodno = ""
          end if

          if len(trim(request("FM_totbudget"))) <> 0 then
          totbudget = request("FM_totbudget")

          totbudget = replace(totbudget, ".", "")
          totbudget = replace(totbudget, ",", ".")

          else
          totbudget = 0
          end if

          if len(trim(request("FM_extfound"))) <> 0 then
          extfound = request("FM_extfound")

          extfound = replace(extfound, ".", "")
          extfound = replace(extfound, ",", ".")

          else
          extfound = 0
          end if

          if len(trim(request("FM_stdato"))) <> 0 then
          stdato = year(request("FM_stdato")) &"-"& month(request("FM_stdato")) &"-"& day(request("FM_stdato"))
          else
          stdato = "2002-01-01"
          end if
          
          if len(trim(request("FM_sldato"))) <> 0 then
          sldato = year(request("FM_sldato")) &"-"& month(request("FM_sldato")) &"-"& day(request("FM_sldato"))
          else
          sldato = "2044-01-01"
          end if

          budget_view = request("FM_budget_view")

          if len(trim(request("FM_valuta_0"))) <> 0 then
          b_valuta = request("FM_valuta_0")
          else
          b_valuta = 1
          end if

          if len(trim(request("FM_valuta_0_kurs"))) <> 0 then
          b_valuta_kurs = replace(request("FM_valuta_0_kurs"), ",", ".")
          else
          b_valuta_kurs = 100
          end if
          
          

          dato = year(now) & "-" & month(now) & "-" & day(now)
          editor = session("user")

          '**** ERRORs *********************
          if len(trim(budgetnavn)) = 0 then
          %>
            
                <div class="container">
                  <div class="portlet">
                      <div style="padding:40px;">
                      <div class="row">
                           <h4>Error</h4>
                 
                Required field <b>Budget Name</b> is missing.
            
            <br /><br />

            <a href="Javascript:history.back()"><< Back</a>
            </div>
                        
                          </div>
                      </div>
                    </div>
          <%
          response.end
          end if 





          '***** Indlæser Budget ******************
          lastId = 0
          if func = "dbopr" then
          strSQLins = "INSERT INTO budget_job (dato, editor, budgetnavn, jobid, budgetstatus, refno, prodno, budget_tot, budget_extra_fo, stdato, sldato, budget_view, b_valuta, b_valuta_kurs) "_
          &" VALUES ('"& dato &"', '"& editor &"', '"& budgetnavn &"', "& jobid &", "& status &", '"& refno &"', '"& prodno &"', "& totbudget &", "& extfound &", "_
          &" '"& stdato &"', '"& sldato & "', "& budget_view &", "& b_valuta &","& b_valuta_kurs &")"
          oConn.execute(strSQLins)
                

                        '*** LAST ID ****
                        strSQLlastid = "SELECT id FROM budget_job WHERE id <> 0 ORDER BY id DESC LIMIT 1"
            
                        'response.write strSQLlastid
                        'response.end            
            
                        oRec.open strSQLlastid, oConn, 3
                        if not oRec.EOF then
                        lastId = oRec("id")
                        end if
                        oRec.close


          else
          strSQLupd = "UPDATE budget_job SET dato = '"& dato &"', editor = '"& editor &"', budgetnavn = '"& budgetnavn &"', jobid = "& jobid &", "_
          &" budgetstatus = "& status &", refno = '"& refno &"', prodno = '"& prodno &"', budget_tot = "& totbudget &", budget_extra_fo = "& extfound &", "_
          &" stdato = '"& stdato &"', sldato = '"& sldato &"', budget_view = "& budget_view &", b_valuta = "& b_valuta &", b_valuta_kurs = "& b_valuta_kurs &" WHERE id = "& id
          
          'response.write strSQLupd
          'response.flush
          
          oConn.execute(strSQLupd)
          
          end if


          if func = "dbopr" then
          id = lastId
          else
          id = id
          end if

           

          '**** Perioder *****************************
          datoThis = year(now) &"-"& month(now) &"-"& day(now)
          
          if len(trim(request("FM_current_periodeid"))) <> 0 then 
          current_periodeid = request("FM_current_periodeid")
          else
          current_periodeid = 0
          end if

         'response.write "-------------------------------"& request("FM_percum_id")
         'response.write "<br>-------------------------------"& request("FM_percum_desc")
         'response.write "<br>-------------------------------"& request("FM_percum_stdate")
         'response.Flush

          perid = split(request("FM_percum_id"), ", ")
          percum_desc = split(request("FM_percum_desc"), ", ##, ")
          percum_budget_app = split(request("FM_percum_budget_app"), ", ##, ")
          percum_kurs = split(request("FM_percum_kurs"), ", ##, ")
          percum_budget_stdato = split(request("FM_percum_stdate"), ", ##, ")
          percum_budget_sldato = split(request("FM_percum_enddate"), ", ##, ")


        
          'current_periodeid_new = split("1,0,0,0,0", ",")
          'response.write "-------------------------------"& request("FM_del_per") & "<br>"
          'response.flush
          
          perDelete = split(request("FM_del_per"), "##, ")

           for p = 0 TO UBOUND(perid)

                perDelete(p) = replace(perDelete(p), "##", "")
                perDelete(p) = replace(perDelete(p), ",", "")
                perDelete(p) = trim(perDelete(p))

              if perDelete(p) = "1" then
            
                strSQldel = "DELETE FROM budget_job_per WHERE id = "& perid(p)
                oConn.execute(strSQldel)

              else

              percum_desc(p) = replace(percum_desc(p), ", ##", "")
              
              '********* Budget ***************
              percum_budget_app(p) = replace(percum_budget_app(p), ", ##", "")
              percum_budget_app(p) = replace(percum_budget_app(p), ".", "")
              percum_budget_app(p) = replace(percum_budget_app(p), ",", ".")

              if len(trim(percum_budget_app(p))) = 0 then
              percum_budget_app(p) = 0
              else
              percum_budget_app(p) = trim(percum_budget_app(p))
              end if

              '****** KURS *****
              percum_kurs(p) = replace(percum_kurs(p), ", ##", "")
              percum_kurs(p) = replace(percum_kurs(p), ".", "")
              percum_kurs(p) = replace(percum_kurs(p), ",", ".")

              if len(trim(percum_kurs(p))) = 0 then
              percum_kurs(p) = 0
              else
              percum_kurs(p) = trim(percum_kurs(p))
              end if

             

           

               'response.write "---------------------------------"& percum_budget_stdato(p) & "<br>"
               'response.Flush
            '************ DATOER ****************
             percum_budget_stdato(p) = replace(percum_budget_stdato(p), ", ##", "")
             percum_budget_stdato(p) = replace(percum_budget_stdato(p), ".", "-")
             if len(trim(percum_budget_stdato(p))) <> 0 AND isDate(percum_budget_stdato(p)) = true then
             percum_budget_stdato(p) = year(percum_budget_stdato(p)) & "-" & month(percum_budget_stdato(p)) & "-" & day(percum_budget_stdato(p))
             else
             percum_budget_stdato(p) = "2002-01-01"
             end if


             percum_budget_sldato(p) = replace(percum_budget_sldato(p), ", ##", "")
             percum_budget_sldato(p) = replace(percum_budget_sldato(p), ".", "-")
             if len(trim(percum_budget_sldato(p))) <> 0 AND isDate(percum_budget_sldato(p)) = true then
             percum_budget_sldato(p) = year(percum_budget_sldato(p)) & "-" & month(percum_budget_sldato(p)) & "-" & day(percum_budget_sldato(p))
             else
             percum_budget_sldato(p) = "2002-01-01"
             end if

                if cint(current_periodeid) = cint(perid(p)) then ' OR func = "dbopr" AND p = 0 then
                currentperiod = current_periodeid
                else
                currentperiod = 0
                end if

                  if perid(p) = 0 then

                    if len(trim(percum_desc(p))) <> 0 then '** KAN ikke indsætte tomme linjer **'
                    strSQLins = "INSERT INTO budget_job_per (budgetid, beskrivelse, budget_app, stdato, sldato, currentperiod, kurs) VALUES "_
                    &" ("& id &", '"& percum_desc(p) &"', "& percum_budget_app(p) &", '"& percum_budget_stdato(p) &"', '"& percum_budget_sldato(p) &"', "& currentperiod &", "& percum_kurs(p) &")"
                    oConn.execute(strSQLins)
          
          
                        if func = "dbopr" AND p = 0 then '** Sætter første current
                        strSQLfirstLastper = "SELECT id FROM budget_job_per WHERE id <> 0 ORDER BY id DESC"
                        oRec5.open strSQLfirstLastper, oConn, 3
                        if not oRec5.EOF then

                    
                              strSQLupdFirstLastPer = "UPDATE budget_job_per SET currentperiod = "& oRec5("id") &" WHERE id = " & oRec5("id") 
                              oConn.execute(strSQLupdFirstLastPer)               


                        end if
                        oRec5.close
                    

                        end if
          
          
                    end if




                  else

                
                     '*** VED DISABLED ****'
                      percum_desc(p) = replace(percum_desc(p), "##, ", "")
                      percum_desc(p) = replace(percum_desc(p), "##", "")
                      percum_budget_stdato(p) = replace(percum_budget_stdato(p), "##, ", "")
                      percum_budget_stdato(p) = replace(percum_budget_stdato(p), "##", "")
                      percum_budget_sldato(p) = replace(percum_budget_sldato(p), "##, ", "")
                      percum_budget_sldato(p) = replace(percum_budget_sldato(p), "##", "")
                      percum_budget_app(p) = replace(percum_budget_app(p), "##. ", "")
                      percum_budget_app(p) = replace(percum_budget_app(p), "##", "")
                      percum_kurs(p) = replace(percum_kurs(p), "##. ", "")
                      percum_kurs(p) = replace(percum_kurs(p), "##", "")
         

                      if len(trim(percum_desc(p))) <> 0 then '** KAN ikke indsætte tomme linjer ELLER lukkede perider (DISABLED) **'

                      strSQLupd = "UPDATE budget_job_per SET beskrivelse = '"& percum_desc(p) &"', budget_app = "& percum_budget_app(p) &", stdato = '"& percum_budget_stdato(p) &"', sldato = '"& percum_budget_sldato(p) &"', currentperiod = "& currentperiod &", kurs = "& percum_kurs(p) &" WHERE id = "& perid(p)
                      
                      'response.write "<br>-----------------------------------------"& strSQLupd
                      'response.flush
                      oConn.execute(strSQLupd)
                      end if

                  end if

              end if

          next

          'response.end



          '**** Activities / Expenditures *************************


          'response.write "--------------------------"& request("FM_aktfas_budget")
          'response.write "<br>--------------------------"& request("FM_aktfase_konto")
          'response.flush

          aktid = split(request("FM_aktid"), ", ")
          aktnavn = split(request("FM_aktnavn"), ", ##, ")
          aktfas_konto = split(request("FM_aktfase_konto"), ", ##, ")
          aktfase = split(request("FM_aktfase"), ", ##, ")
          aktfas_budget = split(request("FM_aktfas_budget"), ", ##, ")
          aktfas_current = split(request("FM_aktfas_current"), ", ##, ")


          'response.write "<br>--------------------------"& request("FM_delete")
          'response.end


          aktDelete = split(request("FM_delete"), "##, ")

         
          for a = 0 TO UBOUND(aktnavn)

                aktDelete(a) = replace(aktDelete(a), "##", "")
                aktDelete(a) = replace(aktDelete(a), ",", "")
                aktDelete(a) = trim(aktDelete(a))


            'response.write "<br>---------------------------------------------"& aktid(a) & ": "& aktDelete(a) & "<br>"
            'response.flush 

              if aktDelete(a) = "1" then
            
                strSQldel = "DELETE FROM budget_job_exp WHERE id = "& aktid(a)
                oConn.execute(strSQldel)

                 strSQLdel = "DELETE FROM budget_job_exp_rel WHERE aktid = " & aktid(a)
                 oConn.execute(strSQLdel)

              else

              aktfase(a) = replace(aktfase(a), ", ##", "")
              aktfas_konto(a) = replace(aktfas_konto (a), ", ##", "")
              aktnavn(a) = replace(aktnavn(a), ", ##", "")
              aktfas_current(a) = replace(aktfas_current(a), ", ##", "")

              
                if len(trim(aktfas_konto(a))) <> 0 then
                aktfas_konto(a) = aktfas_konto(a)
                else
                aktfas_konto(a) = 0
                end if

          
              if len(trim(aktfas_current(a))) <> 0 then
              aktfas_current(a) = replace(aktfas_current(a), ".", "")
              aktfas_current(a) = replace(aktfas_current(a), ",", ".")
              else
              aktfas_current(a) = 0 
              end if


              
              aktfas_budget(a) = replace(aktfas_budget(a), ", ##", "")
              if len(trim(aktfas_budget(a))) <> 0 then
              aktfas_budget(a) = replace(aktfas_budget(a), ".", "")
              aktfas_budget(a) = replace(aktfas_budget(a), ",", ".")
              else
              aktfas_budget(a) = 0
              end if


                  if aktid(a) = 0 AND len(trim(aktnavn(a))) <> 0 then

                  strSQLins = "INSERT INTO budget_job_exp (budgetid, periodeid, fase, aktnavn, budget_app, konto) VALUES ("& id &", 1, '"& aktfase(a) &"', '"& aktnavn(a) &"',  "& aktfas_budget(a) &", "& aktfas_konto(a) &")"
                  'response.write strSQLins
                  'response.flush
              
                  oConn.execute(strSQLins)

                  else

                      strSQLupd = "UPDATE budget_job_exp SET fase = '"& aktfase(a) &"', aktnavn = '"& aktnavn(a) &"', budget_app = "& aktfas_budget(a) &", konto = "& aktfas_konto(a) &" WHERE id = "& aktid(a)
                      'response.write strSQLupd & "<br>"
                      'response.flush
                      oConn.execute(strSQLupd)

                  end if


                  '*** Expend current '****
                if len(trim(aktfas_current(a))) <> "" AND aktfas_current(a) <> "-9999" then

                   

                        strSQlexprelfindes = "SELECT id FROM budget_job_exp_rel WHERE aktid = "& aktid(a) & " AND budgetid = "& id & " AND periodeid = "& current_periodeid
                
                        oRec.open strSQlexprelfindes, oConn, 3
                        if not oRec.EOF then

                             if aktfas_current(a) = 0 then 'delete

                                 strSQLdel = "DELETE FROM budget_job_exp_rel  WHERE id = " & oRec("id")
                                 oConn.execute(strSQLdel)
                             
          
                             else


                                 strSQLupdexp = "UPDATE budget_job_exp_rel SET budgetid = "& id &", aktid = "& aktid(a) &", dato = '"& datoThis &"', belob = "& aktfas_current(a) &" WHERE id = " & oRec("id")
                                 'response.write strSQLupdexp
                                 'response.flush 
                                 oConn.execute(strSQLupdexp)
                                'periodeid = "& current_periodeid &"

                            end if

                        else

                             if aktfas_current(a) <> 0 then

                             strSQLinsexp = "INSERT INTO budget_job_exp_rel SET budgetid = "& id &", aktid = "& aktid(a) &", dato = '"& datoThis &"', periodeid = "& current_periodeid &", belob = "& aktfas_current(a) 
                             'response.write strSQLinsexp
                             'response.flush 
                             oConn.execute(strSQLinsexp)

                            end if

                        end if
                        oRec.close

                

                end if 'len trim aktfas_current(a)

                


              end if

          next

          'response.Redirect "budget_firapport.asp?func=red&id="& id
          'response.end

          rdir = request("rdir")
          select case rdir
          case "1"
          Response.Redirect "budget_firapport.asp?func=red&id="&id
          case else
          Response.Redirect "budget_firapport.asp"
          end select


       case "opr", "red"



          budget_viewSEL1 = ""
          budget_viewSEL2 = ""
          budget_viewSEL4 = ""
          budget_viewSEL12 = ""



          if func = "red" then
          dbfunc = "dbred"

          strSQL = "SELECT dato, editor, budgetnavn, jobid, budgetstatus, refno, prodno, budget_tot, budget_extra_fo, stdato, sldato, budget_view, b_valuta, b_valuta_kurs FROM budget_job WHERE id = "& id
          
           'response.write strSQL
           'response.flush
              
          oRec.open strSQL, oConn, 3
          if not oRec.EOF then

          dato = oRec("dato")
          editor = oRec("editor")
          budgetnavn = oRec("budgetnavn")
          jobid = oRec("jobid") 
          budgetstatus = oRec("budgetstatus")
          refno = oRec("refno") 
          prodno = oRec("prodno")
          budget_tot = formatnumber(oRec("budget_tot"), 2)
          budget_extra_fo = formatnumber(oRec("budget_extra_fo"), 2)
          stdato = oRec("stdato")
          sldato = oRec("sldato")
          budget_view = oRec("budget_view")
          valuta = oRec("b_valuta")
          valuta_kurs = oRec("b_valuta_kurs") 

          end if
          oRec.close

           


          else
          dbfunc = "dbopr"

          budgetnavn = ""
          kundeid = 0
          budgetstatus = 1

          budget_view = 4
          jobid = 0

          valuta = 2
          
          valutaKurs(valuta)
          valuta_kurs = dblKurs/100 '00


          if day(now) < 10 then
          dayN = "0"& day(now)
          else
          dayN = day(now)
          end if

          if month(now) < 10 then
          monthN = "0"& month(now)
          else
          monthN = month(now)
          end if

          stdato = dayN &"-"& monthN &"-"& year(now) 
          sldato = dateAdd("m", 12, stdato) 

          end if


          select case budget_view
          case "1"
          budget_viewSEL1 = "SELECTED"
          case "2"
          budget_viewSEL2 = "SELECTED"
          case "4"
          budget_viewSEL4 = "SELECTED"
          case "12"
          budget_viewSEL12 = "SELECTED"
          end select


          
            strSQLacc = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
		    &" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
		    &" ORDER BY k.kontonr, k.navn"
		    oRec.open strSQLacc, oConn, 3 
            if func <> "red" then
		    strKontoOptions = "<option value=0>Choose Account..</option>"
            end if
		  
            while not oRec.EOF

            strKontoOptions = strKontoOptions & "<option value="& oRec("kontonr") &">"& oRec("navn") &" ("& oRec("kontonr") &")</option>"

            oRec.movenext
            wend
            oRec.close

                                                           

          %>

         
              
                <div class="container">
                  <div class="portlet">


                        <form method="post" id="budget" action="budget_firapport.asp?func=<%=dbfunc %>">
                            <input type="hidden" name="rdir" id="rdir" value="0" /> 
                            <input type="hidden" name="id" id="budgetid" value="<%=id %>" />
                            <input type="hidden" name="fn" id="fn" value="<%=func %>" /> 
                            <input type="hidden" id="fajl" value=".." />
                    <h3 class="portlet-title">
                      <u>Redigér Budget</u>
                    </h3>
                    <!-- Opdater/Annuller knapper -->
                    <div style="margin-top:-15px;margin-bottom:15px;">
                      <button type="submit" class="btn btn-secondary btn-sm pull-right" id="sbm_upd0"><b>Update & Exit</b></button>
                      <!--<button type="button" class="btn btn-danger btn-sm pull-right" style="margin-right:10px;" /><b>Cancel</b></button>-->
                      <div class="clearfix"></div>
                    </div>
                    <div class="portlet-body">
                        <!-- NAVN / SORTERING / ID -->
                        <section>
                            <div class="well well-white">
                                <div class="row">
                                    <div class="col-lg-2">
                                        <h4 class="panel-title-well">Basedata</h4>
                                    </div>


                                    <!-- Budget title -->
                                    <div class="col-lg-12">

                                        <div class="row">
                                            
                                           
                                            
                                            <div class="col-lg-6 pad-t5">Budget title: <span style="color:red;">*</span>
                                                <input type="text" name="FM_budgetnavn" class="form-control input-small" value="<%=budgetnavn %>" />
                                            </div>
                                               
                                            

                                             <div class="col-lg-4 pad-t5">Customer / Job:

                                                 <%strSQLkunder = "SELECT jobnavn, jobnr, id, k.kkundenavn, k.kid FROM job "_
                                                     &" LEFT JOIN kunder AS k ON (kid = jobknr) WHERE (id <> 0 AND jobstatus = 1) OR id = "& jobid &" ORDER BY kkundenavn, jobnavn" 
                                                     
                                                     'response.write strSQLkunder
                                                     'response.flush%>

                                                <select name="FM_kunde" class="form-control input-small" />
                                                <option value="0">Choose..</option>

                                                 <%
                                                   
                                                    lastk = ""
                                                    oRec.open strSQLkunder, oConn, 3
                                                    while not oRec.EOF 


                                                     if lastk <> oRec("kkundenavn") AND lastk <> "" then
                                                     %>
                                                    <option value="0"></option>
                                                     <%
                                                     end if

                                                     if cint(jobid) = cint(oRec("id")) then
                                                     kSel = "SELECTED"
                                                     else
                                                     kSel = ""
                                                     end if

                                                     %><option value="<%=oRec("id") %>" <%=kSel %>><%=oRec("kkundenavn") %> - <%=oRec("jobnavn") %></option><%


                                                    lastk = oRec("kkundenavn")
                                                    oRec.movenext
                                                    wend
                                                    oRec.close

                                                      %>

                                              
                                               
                                                </select>
                                            </div>
                                             
                                            <div class="col-lg-2 pad-r30 pad-t5">Status:
                                                <select name="FM_status" class="form-control input-small">

                                                    <%
                                                       
                                                        bgtStatusSEL0 = ""
                                                        bgtStatusSEL1 = "" 
                                                        bgtStatusSEL2 = ""

                                                       select case budgetstatus
                                                       case 1
                                                        bgtStatusSEL1 = "SELECTED"
                                                       case 2
                                                        bgtStatusSEL2 = "SELECTED"
                                                        case 0
                                                        bgtStatusSEL0 = "SELECTED"
                                                        case else
                                                        bgtStatusSEL1 = "SELECTED"
                                                        end select  
                                                        
                                                       %>

                                                    <option value="1" <%=bgtStatusSEL1%>>Active</option>
                                                    <option value="0" <%=bgtStatusSEL0%>>Closed</option>
                                                    <option value="2" <%=bgtStatusSEL2%>>Passive</option>
                                                </select>
                                            </div>
                                             
                                            
                                        </div>

                                   
                                        <div class="row">
                                            <div class="col-lg-3 pad-t5">Reference no.:   
                                                <input type="text" name="FM_refno" class="form-control input-small" value="<%=refno%>" />
                                            </div>

                                            <div class="col-lg-3 pad-t5">Product no.:
                                          
                                                <input type="text" name="FM_prodno" class="form-control input-small" value="<%=prodno %>" />
                                            </div>


                                             <div class="col-lg-2 pad-t5">Total budget:
                                                <input type="text" name="FM_totbudget" class="form-control input-small" value="<%=budget_tot %>" />
                                            </div>

                                             <div class="col-lg-2 pad-t5">Extra founding:
                                                    <input type="text" name="FM_extfound" class="form-control input-small" value="<%=budget_extra_fo %>" />
                                                </div>

                                              <div class="col-lg-2 pad-r30 pad-t5">Currency:

                                                  <%call valutakoder(0, valuta, 1) %>

                                                </div>
                                            
                                        </div>


                                       
                                             <div class="row">
                                               

                                               


                                                   

                                                <div class="col-lg-2 pad-t5">Start date:
                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control input-small" name="FM_stdato" id="FM_stdato" value="<%=stdato %>" />
                                                          <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>
                                                </div>

                                            
                                                  <div class="col-lg-2 pad-t5">End date:
                                                    <div class='input-group date' id='datepicker_sldato'>
                                                    <input type="text" class="form-control input-small" name="FM_sldato" id="FM_sldato" value="<%=sldato %>" />
                                                       <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>
                                                </div>

                                                    <div class="col-lg-2"><!--Months calc-->

                                                        <input type="hidden" name="" id="months_diff" value="12" />

                                                    </div>

                                                 <div class="col-lg-4 pad-t5">Budget view:
                                                    <select name="FM_budget_view" id="FM_budget_view" class="form-control input-small" />
                                                     <option value="12" <%=budget_viewSEl12 %>>Monthly</option>
                                                     <option value="4" <%=budget_viewSEl4 %>>Quarter</option>
                                                     <option value="2" <%=budget_viewSEl2 %>>½Y</option>
                                                     <option value="1" <%=budget_viewSEl1 %>>FY</option>
                                                    </select>
                                                </div>

                                                  <div class="col-lg-2 pad-t5 pad-r30">Rate:
                                                       <input type="text" class="form-control input-small" name="FM_valuta_0_kurs" id="FM_valuta_0_kurs" value="<%=valuta_kurs %>" />
                                                  </div>
                                            
                                            </div>

                                     </div>

                                    
                                   
                                  
                                </div>
                            </div>
                        </section>


                        <!-- PERIODS -->
                        <br />
                        <section>
                             <div class="well well-white">
                                <div class="row">
                                    <div class="col-lg-2">
                                        <h4 class="panel-title-well">Periods</h4>
                                    </div>

                                  

                                    <!-- Period lines -->
                                  
                                               <div class="row">
                                               

                                                <div class="col-lg-12 pad-l30 pad-r50">


                                                  <table class="table table-striped table-bordered" id="xtable-2">
                                                <thead>
                                                  <tr>
                                                    <th style="width: 5%">Current</th>
                                                    <th style="width: 20%">Description</th>
                                                    <th style="width: 10%">St. Date</th>
                                                    <th style="width: 10%">End Date</th>
                                                
                                                    <th style="width: 12%">Budget App.</th>
                                                    <th style="width: 7%">Rate</th>
                                                    <th style="width: 12%"><span style="font-size:9px; font-weight:lighter;">Expenditures</span><br />Prev. period</th>
                                                    <th style="width: 12%"><span style="font-size:9px; font-weight:lighter;">Expenditures</span><br />Cur. period</th>
                                                    <th style="width: 12%"><span style="font-size:9px; font-weight:lighter;">Expenditures</span><br />Commulative</th>
                                                    <th style="width: 12%">Rem. budget</th>

                                                    <th>Delete</th>
                                                   
                                                    
                                                  </tr>
                                                </thead>
                                                  <tbody>

                                                      <%
                                                      budget_appTot = 0
                                                      exp_prevTot = 0
                                                      exp_currTot = 0
                                                      exp_cummTot = 0
                                                      exp_remaTot = 0
                                                      currentperiodId = 0   
                                                      currentPeriodTXT = ""  
                                                      'antalPerMKurs = 0
                                                      'kursTot = 0
                                                      
                                                      
                                                          
                                                          
                                                      strSQLbgtper = "SELECT id, beskrivelse, budget_app, stdato, sldato, currentperiod, kurs FROM budget_job_per WHERE budgetid = "& id & " ORDER BY id"
                                                      p = 0
                                                      oRec3.open strSQLbgtper, oConn, 3
                                                      while not oRec3.EOF
                                                          
                                                          stdato = oRec3("stdato")
                                                          sldato = oRec3("sldato")


                                                              if cDate(stdato) <> "01-01-2002" then
                                                              stdato = stdato
                                                              else
                                                              stdato = ""
                                                              end if

                                                              if cDate(sldato) <> "01-01-2002" then
                                                              sldato = sldato
                                                              else
                                                              sldato = ""
                                                              end if


                                                               if cdbl(oRec3("currentperiod")) = cdbl(oRec3("id")) then
                                                               periodActiveColor = "yellowgreen" '"#5cb85c" '
                                                               bdpx = 2
                                                               currentperiodCHK = "CHECKED"
                                                               currentperiodId = oRec3("id")
                                                               currentPeriodTXT = left(oRec3("beskrivelse"), 20) '& "<br>" & stdato & " - " & sldato
                                                               else
                                                               currentperiodCHK = ""
                                                               bdpx = 0
                                                               periodActiveColor = ""
                                                               end if


                                                            'if cdbl(currentperiodId) = 0 AND func = "red" then
                                                            'periodDisabled = "DISABLED"
                                                            'else
                                                            'periodDisabled = ""
                                                            'end if 

                                                          

                                                            if oRec3("kurs") <> 0 AND len(trim(oRec3("kurs")) <> 0) then
                                                            'kursTot = (kursTot/1 + oRec3("kurs")/1)
                                                            kurs = formatnumber(oRec3("kurs"),2)
                                                            'antalPerMKurs = antalPerMKurs + 1 
                                                            else
                                                            kurs = ""
                                                            end if
                                                      %>

                                                     
                                                      <tr>
                                                          <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="hidden" class="rw_<%=oRec3("id")%>" name="FM_percum_id" value="<%=oRec3("id") %>" <%=periodDisabled %> />
                                                              <input type="radio" class="rdbtn_per" id="rw_<%=oRec3("id")%>" name="FM_current_periodeid" value="<%=oRec3("id") %>" <%=currentperiodCHK %> />
                                                           </td>
                                                           <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;">
                                                               <input type="text" name="FM_percum_desc" class="form-control input-small rw_<%=oRec3("id")%>" value="<%=oRec3("beskrivelse")%>" <%=periodDisabled %> />
                                                               <input type="hidden" name="FM_percum_desc" value="##" />
                                                           </td>
                                                          <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="text" name="FM_percum_stdate" class="form-control input-small rw_<%=oRec3("id")%>" value="<%=stdato%>" <%=periodDisabled %> />
                                                              <input type="hidden" name="FM_percum_stdate" value="##" />
                                                          </td>
                                                          <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="text" name="FM_percum_enddate" class="form-control input-small rw_<%=oRec3("id")%>" value="<%=sldato %>" <%=periodDisabled %> />
                                                              <input type="hidden" name="FM_percum_enddate" value="##" />
                                                          </td>
                                                          <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="text" name="FM_percum_budget_app" style="text-align:right;" class="form-control input-small rw_<%=oRec3("id")%>" value="<%=formatnumber(oRec3("budget_app"), 2) %>" <%=periodDisabled %> />
                                                               <input type="hidden" name="FM_percum_budget_app" value="##" />
                                                          </td>

                                                           <td style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="text" name="FM_percum_kurs" class="form-control input-small" value="<%=kurs%>" />
                                                               <input type="hidden" name="FM_percum_kurs" value="##" /></td>

                                                          <%
                                                              expend_prev_per = 0
                                                              strSQlprv = "SELECT COALESCE(SUM(br.belob), 0) AS expend_prev FROM budget_job_exp_rel AS br WHERE (br.budgetid = "& id &" AND br.periodeid < "& oRec3("id") &") GROUP BY br.budgetid"
                                                              oRec4.open strSQlprv, oConn, 3
                                                              if not oRec4.EOF then 
                                                              
                                                                expend_prev_per = oRec4("expend_prev")

                                                              end if 
                                                              oRec4.close    


                                                              expend_cur_per = 0
                                                              strSQlprv = "SELECT COALESCE(SUM(br.belob), 0) AS expend_prev FROM budget_job_exp_rel AS br WHERE (br.budgetid = "& id &"  AND br.periodeid = "& oRec3("id") &") GROUP BY br.budgetid, br.periodeid"
                                                              oRec4.open strSQlprv, oConn, 3
                                                              if not oRec4.EOF then 
                                                              
                                                                expend_cur_per = oRec4("expend_prev")

                                                              end if 
                                                              oRec4.close  
                                                               

                                                      
                                                            
                                                               expend_cum_per = expend_prev_per + (expend_cur_per)

                                                               budget_app_per = oRec3("budget_app") 
                                                               budget_app_rem = (budget_app_per - expend_cum_per)


                                                              expend_cum_perTot = expend_cum_perTot + (expend_cum_per)
                                                              expend_prev_perTot = expend_prev_perTot + (expend_prev_per)
                                                              expend_cur_perTot = expend_cur_perTot + (expend_cur_per) 
                                                              budget_appTot = budget_appTot + (oRec3("budget_app"))
                                                              budget_app_remTot = budget_app_remTot + (budget_app_rem) 

                                                               
                                                              if expend_prev_per <> 0 then
                                                              expend_prev_per = formatnumber(expend_prev_per, 2)
                                                              else
                                                              expend_prev_per = ""
                                                              end if


                                                               if budget_app_rem <> 0 then
                                                               budget_app_rem = formatnumber(budget_app_rem, 2)
                                                               else
                                                               budget_app_rem = ""
                                                               end if

                                                               if expend_cur_per <> 0 then
                                                               expend_cur_per = formatnumber(expend_cur_per, 2)
                                                               else
                                                               expend_cur_per = ""
                                                               end if

                                                               if expend_cum_per <> 0 then
                                                               expend_cum_per = formatnumber(expend_cum_per, 2)
                                                               else
                                                               expend_cum_per = ""
                                                               end if

                                                               
                                                                 

                                                              
                                                          %>
                                                              <td style="text-align:right; border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><%=expend_prev_per%></td>
                                                          <td style="text-align:right; border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><%=expend_cur_per %></td>
                                                          <td style="text-align:right; border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><%=expend_cum_per%></td>
                                                          <td style="text-align:right; border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><%=budget_app_rem %></td>
                                                          <td align="center" style="border-bottom:<%=bdpx%>px <%=periodActiveColor%> solid;"><input type="checkbox" name="FM_del_per" value="1" class="rw_<%=oRec3("id")%>" <%=periodDisabled %> />
                                                              <input type="hidden" name="FM_del_per" value="##" />
                                                          </td>

                                                      </tr>

                                                      <%
                                                          
                                                    
                                                      p = p + 1

                                                      oRec3.movenext 
                                                      wend 
                                                      oRec3.close



                                                            '********************************************************************************
                                                            '*** HENT Expenditures, find periode. Find vægtning på kurs og beregn gennemsnit.
                                                            '********************************************************************************
                                                            if func <> "opr" then


                                                            dim thisPeriodBel, thisPeriodRate
                                                            redim thisPeriodBel(30), thisPeriodRate(30)

                                                           
                                                            antalperiodeids = 0
                                                            lastperiodeid = 0
                                                            totalBelob = 0
                                                            strSQLavgrate = "SELECT e.periodeid, e.belob, p.kurs FROM budget_job_exp_rel AS e "_
                                                            &" LEFT JOIN budget_job_per AS p ON (p.id = e.periodeid) WHERE e.budgetid = " & id & " ORDER BY e.periodeid"
                                                            
                                                            'response.write strSQLavgrate & "<br><br>"
                                                          
                                                            oRec5.open strSQLavgrate, oConn, 3
                                                            while not oRec5.EOF 
                                                          
                                                            
                                                                if cdbl(lastperiodeid) <> cdbl(oRec5("periodeid")) then
                                                                antalperiodeids = antalperiodeids + 1
                                                                end if 
                                                            
                                                            'response.write "LIBNJE: "& oRec5("belob") & " -- "
                                                            thisPeriodRate(antalperiodeids) = oRec5("kurs")
                                                            thisPeriodBel(antalperiodeids) = thisPeriodBel(antalperiodeids) + oRec5("belob")
                                                            

                                                            'response.Write antalperiodeids & ": "& thisPeriodBel(antalperiodeids) & "<br>"
                                                            totalBelob = totalBelob + oRec5("belob")
                                                            lastperiodeid = oRec5("periodeid")
                                                            oRec5.movenext
                                                            wend 
                                                            oRec5.close 



                                                            for b = 1 to antalperiodeids

                                                             if totalBelob <> 0 then
                                                             vaegtning = thisPeriodBel(b) / totalBelob  
                                                             end if

                                                            'response.Write "Per " & b & ":"& thisPeriodBel(b) & " totalBelob : "& totalBelob  &" Rate: "& thisPeriodRate(b) &" vaegtning: "& vaegtning &"<br>"
                                                            avgRate = avgRate + (thisPeriodRate(b) * vaegtning)


                                                            next

                                                            
                                                            avgRate = formatnumber(avgRate, 2)


                                                            else

                                                            avgRate = 0

                                                            end if
                                                            

                                                            'response.Write "avgRate: " & avgRate

                                                            '********************************************************************************



                                                    
                                                    if func = "Xopr" then

                                                      e = 3
                                                    

                                                       for p = 0 TO e


                                                          
                                                      %>      

                                                       <tr>
                                                           
                                                           <td>
                                                               <input type="hidden" name="FM_percum_id" value="0" />
                                                                
                                                              
                                                            </td>
                                                           <td><input type="text" name="FM_percum_desc" class="form-control input-small" value="" />
                                                                <input type="hidden" name="FM_percum_desc" value="##" />
                                                           </td>
                                                          <td><input type="text" name="FM_percum_stdate" class="form-control input-small" value="" />
                                                              <input type="hidden" name="FM_percum_stdate" value="##" />
                                                          </td>
                                                          <td><input type="text" name="FM_percum_enddate" class="form-control input-small" value="" />
                                                              <input type="hidden" name="FM_percum_enddate" value="##" />
                                                          </td>
                                                          <td><input type="text" name="FM_percum_budget_app" class="form-control input-small" value="" />
                                                               <input type="hidden" name="FM_percum_budget_app" value="##" />
                                                          </td>
                                                           <td><input type="text" name="FM_percum_kurs" class="form-control input-small" value="" />
                                                               <input type="hidden" name="FM_percum_kurs" value="##" /></td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                           <td>&nbsp;</td>
                                                           <td align="center"><input type="checkbox" name="FM_del_per" value="1" />
                                                              <input type="hidden" name="FM_del_per" value="##" />
                                                          </td>

                                                      </tr>
                                                      <%next %>


                                                      <%end if 
                                                          
                                                      if antalPerMKurs <> 0 then
                                                      kursTotAvg = (kursTot/antalPerMKurs)
                                                      else
                                                      kursTotAvg = kursTot
                                                      end if
                                                          
                                                          
                                                          
                                                      if func = "red" then%>


                                                      <tr>
                                                          
                                                           <td>Total:</td>
                                                          <td>&nbsp;</td>
                                                           <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                           <td style="text-align:right;"><%=formatnumber(budget_appTot, 2) %></td>
                                                           <td style="text-align:right; white-space:nowrap;"><span style="font-size:8px;">Avg.</span> <%=avgRate %></td>
                                                             <td style="text-align:right;"><%=formatnumber(expend_prev_perTot, 2) %></td>
                                                           <td style="text-align:right;"><%=formatnumber(expend_cur_perTot, 2) %></td>
                                                          <td style="text-align:right;"><%=formatnumber(expend_cum_perTot, 2) %></td>
                                                           <td style="text-align:right;"><%=formatnumber(budget_app_remTot, 2) %></td>
                                                      
                                                         
                                                          <td>&nbsp;</td>

                                                           <tr>
                                                                <td>&nbsp;</td>
                                                                <td>&nbsp;</td>
                                                                <td>Currency:</td>
                                                                <td>
                                                                <select class="form-control input-small" / disabled>
                                                   
                                                                 <option>DKK</option>
                                                                 <option>USD</option>
                                                                 <option>EUR</option>
                                                                </select></td>
                                                                <td style="text-align:right;"><%=formatnumber(budget_appTot*avgRate, 2) %></td>
                                                                <td>&nbsp;</td>
                                                                 <td style="text-align:right;"><%=formatnumber(expend_prev_perTot*avgRate, 2) %></td>
                                                           <td style="text-align:right;"><%=formatnumber(expend_cur_perTot*avgRate, 2) %></td>
                                                          <td style="text-align:right;"><%=formatnumber(expend_cum_perTot*avgRate, 2) %></td>
                                                           <td style="text-align:right;"><%=formatnumber(budget_app_remTot*avgRate, 2) %></td>
                                                                <td>&nbsp;</td>
                                                         
                                                      </tr>


                                                      <tr>
                                                          <td colspan="11" align="right">
                                                            <span id="sp_add_period">Add Period +</span>

                                                        
                                                      </td>
                                                      </tr>

                                                    <%end if %>

                                                     
                                           
                                                    <%for b = 0 to 20 %>
                                                    
                                                        <tr class="tr_nyjobper tr_nyjobper_<%=b%> ">
                                                           <td class="tr_nyjobper tr_nyjobper_<%=b%> ">
                                                                <input type="hidden" name="FM_percum_id" value="0" />
                                                                 
                                                              
                                                           </td>
                                                           <td class="tr_nyjobper tr_nyjobper_<%=b%> ">
                                                               <input type="text" name="FM_percum_desc" class="form-control input-small" value="" />
                                                               <input type="hidden" name="FM_percum_desc" value="##" />
                                                           </td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> "><input type="text" name="FM_percum_stdate" class="form-control input-small" value="" />
                                                              <input type="hidden" name="FM_percum_stdate" value="##" />
                                                          </td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> "><input type="text" name="FM_percum_enddate" class="form-control input-small" value="" />
                                                              <input type="hidden" name="FM_percum_enddate" value="##" />
                                                          </td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> "><input type="text" name="FM_percum_budget_app" class="form-control input-small" value="" />
                                                               <input type="hidden" name="FM_percum_budget_app" value="##" />
                                                          </td>
                                                           <td class="tr_nyjobper tr_nyjobper_<%=b%> "><input type="text" name="FM_percum_kurs" class="form-control input-small percum_kurs" value="" />
                                                               <input type="hidden" name="FM_percum_kurs" value="##" />
                                                          </td>

                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> "></td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> ">&nbsp;</td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> ">&nbsp;</td>
                                                          <td class="tr_nyjobper tr_nyjobper_<%=b%> ">&nbsp;</td>
                                                           <td class="tr_nyjobper tr_nyjobper_<%=b%> " align="center"><input type="checkbox" name="FM_del_per" value="1" />
                                                              <input type="hidden" name="FM_del_per" value="##" />
                                                          </td>

                                                      </tr>

                                                      <%next %>

                                                      </tbody>
                                                    </table>

                                               

                                        </div>
                                    </div>
                                   
                                  
                                </div>
                            </div>

                        </section>
                        <!-- </form>

                            <form method="post" action="budget_firapport.asp?func=opdlist&id=<%=id %>">
                            -->

                        <!-- Expenditures -->
                        
                        <section>

                            <br /><br /><br /> 
                             <!-- Opdater/Annuller knapper -->
                                    <div style="margin-top:-15px;margin-bottom:15px;">
                                      <button type="submit" class="btn btn-secondary btn-sm pull-right" id="sbm_upd1"><b>Update</b></button>
                                 
                                      <div class="clearfix"></div>
                                    </div>

                         
                             <div class="well well-white">
                                <div class="row">
                                    <div class="col-lg-2">
                                        <h4 class="panel-title-well">Expenditures</h4>
                                    </div>

                                   



                                                  <div class="row">
                                               

                                                <div class="col-lg-12 pad-l30 pad-r50">


                                                  <table class="table table-striped table-bordered" id="xtable-1">
                                                <thead>
                                                  <tr>
                                                    
                                                    <th style="width: 30%">Phase/Activity</th>
                                                    
                                                    <th style="width: 15%">Account</th>
                                                    <th style="width: 10%">Budget App.</th>
                                                    <th style="width: 10%">Prev. period</th>
                                                    <th style="width: 10%">Cur. period<br />
                                                        <span style="font-size:9px; color:yellowgreen; font-weight:lighter;"><%=currentPeriodTXT %></span>
                                                    </th>
                                                    <th style="width: 10%">Commulative</th>
                                                    <th style="width: 10%">Rem. budget</th>
                                                    <th>Delete</th>
                                                    
                                                  </tr>
                                                </thead>
                                                  <tbody>

                                                      <%
                                                      if func = "red" then
                                                      strSQLakt = "SELECT b.id, b.budgetid, b.budget_app, b.periodeid, b.fase, b.aktnavn, b.konto, COALESCE(SUM(br.belob), 0) AS expend_current FROM budget_job_exp AS b "_
                                                      &" LEFT JOIN budget_job_exp_rel AS br ON (br.budgetid = "& id &" AND br.aktid = b.id AND br.periodeid = "& currentperiodId &")"_
                                                      &" WHERE b.budgetid = "& id & " GROUP BY b.id ORDER BY b.fase, b.aktnavn"
                                                      else
                                                      strSQLakt = "SELECT a.id, a.navn AS aktnavn, a.fase FROM akt_gruppe AS ag "_
                                                      &" LEFT JOIN aktiviteter AS a ON (job = 0 AND a.aktfavorit = ag.id) WHERE ag.forvalgt = 1 AND ag.skabelontype = 1 ORDER BY a.fase, a.sortorder, a.navn" '& jobid
                                                      end if

                                                      lastFase = ""
                                                      f = 0
                                                      fs = 0

                                                     'response.write strSQLakt 
                                                     'response.flush
                                                      expend_prevTot = 0
                                                      expend_currentTot = 0
                                                      expend_cumTot = 0
                                                      expend_rembudTot = 0 
                                                      aktfaseOptions = "<option value=''>None (blank)</option>"  

                                                      oRec.open strSQLakt, oConn, 3
                                                      while not oRec.EOF 

                                                      '*** Phase ****'
                                                       if lastFase <> oRec("fase") OR f = 0 then
                                                          fs = fs + 1

                                                          if isNull(oRec("fase")) <> true then
                                                          fsThis = oRec("fase")
                                                          else
                                                          fsThis = ""
                                                          end if

                                                          %>
                                                         <thead>
                                                          <tr>
                                                              <td><input type="text" name="FM_fase" class="form-control input-small fstxt" id="fs_<%=fs%>" value="<%=replace(fsThis, "_"," ") %>" /></td>
                                                              
                                                               <td>
                                                               <select name="fsFM_konto" class="form-control input-small fsacc" id="fs_acc_<%=fs %>">
                                                                
                                                               <%if func = "red" then 
                                                                   
                                                                           if oRec("konto") = "0" then
                                                                           %>
                                                                           <option value="0">Choose Account..</option>
                                                                           <%
                                                                           else

                                                                                strSQLaccThis = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
		                                                                        &" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
		                                                                        &" WHERE kontonr = " & oRec("konto")
		                                                                        oRec3.open strSQLaccThis, oConn, 3 

                                                                                if not oRec3.EOF then
                                                                       
                                                                                strKontoNavnNr = oRec3("navn") & " ("& oRec3("kontonr") &")"
                                                                                'strKonto = oRec3("konto")

                                                                                end if 
                                                                                oRec3.close

                                                                                %>
                                                                                <option value="<%=oRec("konto")%>"><%=strKontoNavnNr%></option>
                                                                                <%


                                                                           end if

                                                                   end if %>

                                                               <%=strKontoOptions %>
                                                               
                                                            </select></td>
                                                              <td>&nbsp;</td>
                                                              <td>&nbsp;</td>
                                                              <td>&nbsp;</td>
                                                              <td>&nbsp;</td>
                                                              <td>&nbsp;</td>
                                                              <td>&nbsp;</td>
                                                          </tr>
                                                            </thead>
                                                          <%
                                                       end if


                                                     if func = "red" then



                                                              thisAktid = oRec("id")
                                                              '*** PREV
                                                              expend_prev = 0
                                                              expend_cum = 0
                                                              rem_budget = 0
                                                              budget_app = oRec("budget_app")

                                                              strSQlprv = "SELECT COALESCE(SUM(br.belob), 0) AS expend_prev FROM budget_job_exp_rel AS br WHERE (br.budgetid = "& id &" AND br.aktid = "& thisAktid &" AND br.periodeid < "& currentperiodId &") GROUP BY br.aktid"
                                                              
                                                              'response.write strSQlprv

                                                              oRec4.open strSQlprv, oConn, 3
                                                              if not oRec4.EOF then 
                                                              
                                                                expend_prev = oRec4("expend_prev")

                                                              end if 
                                                              oRec4.close

                                                              if expend_prev <> 0 then
                                                              expend_prevTot = expend_prevTot + expend_prev
                                                              expend_prev = formatnumber(expend_prev, 2)

                                                              expend_cum = expend_prev

                                                              else
                                                              expend_prev = ""
                                                              end if



                                                              '*** CURRENT
                                                              expend_current = 0 
                                                              if oRec("expend_current") <> 0 then
                                                              expend_cum = expend_cum + oRec("expend_current")
                                                              expend_cum = formatnumber(expend_cum, 2)

                                                              expend_current = formatnumber(oRec("expend_current"), 2)

                                                              expend_currentTot = expend_currentTot + oRec("expend_current")

                                                              else
                                                              expend_current = ""
                                                              end if


                                                              '*REMAINING BUDGET
                                                              rem_budget = budget_app - (expend_cum)
                                                              expend_rembudTot = expend_rembudTot + (rem_budget)

                                                              '** CUM **
                                                              if expend_cum <> 0 then
                                                              expend_cum = expend_cum 
                                                              expend_cumTot = expend_cumTot + (expend_cum)
                                                              else
                                                              expend_cum = ""
                                                              end if

                                                              '** Budget APP 
                                                              if budget_app <> 0 then
                                                              budget_app = formatnumber(budget_app, 2)
                                                              else
                                                              budget_app = ""
                                                              end if

                                                              '** Remaining budget 
                                                              if rem_budget <> 0 then
                                                              rem_budget = formatnumber(rem_budget, 2)
                                                              else
                                                              rem_budget = ""
                                                              end if


                                                              expend_acc = oRec("konto")


                                                     else
                                                          thisAktid = 0
                                                          expend_prev = ""
                                                          expend_current = ""
                                                          expend_cum = ""
                                                          budget_app = ""
                                                          expend_acc = 0
                                                     end if

                                                      if len(trim(oRec("fase"))) <> 0 AND instr(aktfaseOptions, oRec("fase")) = 0 then
                                                       aktfaseOptions = aktfaseOptions & "<option value='"& oRec("fase") &"'>"& replace(oRec("fase"), "_", " ") &"</option>" 
                                                      end if

                                                      %>


                                                     
                                                    <tr>
                                                        <td style="padding-left:40px;">
                                                        <input type="hidden" name="FM_aktid" value="<%=thisAktid%>" />
                                                        <!--<input type="text" class="afs_acc_<%=fs%>" name="FM_aktfase_konto" value="<%=expend_acc %>" />-->
                                                        


                                                        <input type="hidden" class="afs_<%=fs%>" name="FM_aktfase" value="<%=oRec("fase")%>" />
                                                        <input type="hidden" name="FM_aktfase" value="##" />
                                                            
                                                            <input type="text" name="FM_aktnavn" class="form-control input-small" value="<%=oRec("aktnavn") %>" />
                                                            <input type="hidden" name="FM_aktnavn" value="##" />
                                                        </td>
                                                        

                                                         <td><select class="afs_acc_<%=fs%> form-control input-small" name="FM_aktfase_konto">

                                                             <!-- value="<%=expend_acc %>" -->
                                                                
                                                               <%if func = "red" then 
                                                                   
                                                                           if oRec("konto") = "0" then
                                                                           %>
                                                                           <option value="0">Choose Account..</option>
                                                                           <%
                                                                           else

                                                                                strSQLaccThis = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
		                                                                        &" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
		                                                                        &" WHERE kontonr = " & oRec("konto")
		                                                                        oRec3.open strSQLaccThis, oConn, 3 

                                                                                if not oRec3.EOF then
                                                                       
                                                                                strKontoNavnNr = oRec3("navn") & " ("& oRec3("kontonr") &")"
                                                                                'strKonto = oRec3("konto")

                                                                                end if 
                                                                                oRec3.close

                                                                                %>
                                                                                <option value="<%=oRec("konto")%>"><%=strKontoNavnNr%></option>
                                                                                <%


                                                                           end if

                                                                   end if %>

                                                               <%=strKontoOptions %>
                                                               
                                                            </select>
                                                             <input type="hidden" name="FM_aktfase_konto" value="##" />
                                                         </td>


                                                        <td><input type="text" name="FM_aktfas_budget" style="text-align:right;" class="form-control input-small" value="<%=budget_app %>" />
                                                            <input type="hidden" name="FM_aktfas_budget" value="##" />
                                                        </td>
                                                        <td style="text-align:right;"><%=expend_prev %></td>
                                                        <td><input type="text" name="FM_aktfas_current" style="text-align:right;" class="form-control input-small aktcurrent" value="<%=expend_current %>" />
                                                            <input type="hidden" name="FM_aktfas_current" value="##" />
                                                        </td>
                                                        <td style="text-align:right;"><%=expend_cum %></td>
                                                        <td style="text-align:right;"><%=rem_budget %></td>
                                                        <td align="center"><input type="checkbox" name="FM_delete" value="1" />
                                                            <input type="hidden" name="FM_delete" value="##" />
                                                        </td>
                                                       
                                                    </tr>
                                                   
                                                    <%
                                                        f = f + 1
                                                        lastFase = oRec("fase")
                                                     oRec.movenext   
                                                     wend
                                                    oRec.close %>



                                                     
                                                     

                                                      <tr>
                                                          <td>Total:</td>
                                                          <td></td>
                                                          <td></td>
                                                          
                                                          <td style="text-align:right;"><%=formatnumber(expend_prevTot, 2) %></td>
                                                          <td style="text-align:right;"><%=formatnumber(expend_currentTot, 2) %></td>
                                                          <td style="text-align:right;"><%=formatnumber(expend_cumTot, 2) %></td>
                                                          <td style="text-align:right;"><%=formatnumber(expend_rembudTot, 2) %></td>
                                                          <td></td>
                                                      </tr>

                                                       <tr>
                                                          <td colspan="8" align="right"><span id="sp_add_activity">Add Activity +</span></td>

                                                      </tr>


                                                       <tr class="tr_nyakt">
                                                            <td style="padding-left:40px;"> <input type="hidden" name="FM_aktid" value="0" />
                                                           Phase: <select name="FM_aktfase" class="form-control input-small"><%=aktfaseOptions %></select>
                                                           <input type="hidden" name="FM_aktfase" value="##" /></td>
                                                           <td colspan="7">&nbsp;</td>
                                                        </tr>
                                                        
                                                        <tr class="tr_nyakt">
                                                        <td style="padding-left:40px;"><input type="text" name="FM_aktnavn" class="form-control input-small" value="" />
                                                            <input type="hidden" name="FM_aktnavn" value="##" />

                                                             <input type="hidden" name="FM_aktfase_konto" value="0" />
                                                        <input type="hidden" name="FM_aktfase_konto" value="##" />

                                                        </td>
                                                        <td>&nbsp;</td>
                                                        <td><input type="text" name="FM_aktfas_budget" style="text-align:right;" class="form-control input-small" value="" />
                                                            <input type="hidden" name="FM_aktfas_budget" value="##" />
                                                        </td>
                                                        <td style="text-align:right;">&nbsp;</td>
                                                        <td><input type="text" name="FM_aktfas_current" style="text-align:right;" class="form-control input-small aktcurrent" value="" />
                                                            <input type="hidden" name="FM_aktfas_current" value="##" />
                                                        </td>
                                                        <td style="text-align:right;">&nbsp;</td>
                                                        <td style="text-align:right;">&nbsp;</td>
                                                        <td align="center"><input type="checkbox" name="FM_delete" value="1" />
                                                            <input type="hidden" name="FM_delete" value="##" />
                                                        </td>
                                                       
                                                    </tr>


                                                </tbody>
                                                <tfoot>
                                                  <tr>
                                                   <th>Phase/Activity</th>
                                                    
                                                    <th>Account</th>
                                                    <th>Budget App.</th>
                                                    <th>Prev. period</th>
                                                    <th>Cur. period</th>
                                                    <th>Commulative</th>
                                                    <th>Rem. budget</th>
                                                    <th>Delete</th>
                                                  </tr>
                                                </tfoot>

                                                     

                                              </table>

                                          

                                          </div>

                                     


                                </div>
                                   
                                  
                                     

                                </div>

                             </div>

                                    <br />
                                  <div style="margin-top:-15px;margin-bottom:15px;">
                                      <button type="submit" class="btn btn-secondary btn-sm pull-right" id="sbm_upd2"><b>Update</b></button>
                                 
                                      <div class="clearfix"></div>
                                    </div>

                        </section>
                        
                      

                       

                    </div> <!-- /.portlet-body -->
                      </form>
                     
                  </div> <!-- /.portlet -->
                </div> <!-- /.container -->

          <%
       case else 'list 
              
              
     if media <> "export" then%>

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Budgetter</u>
        </h3>
        <!--
          <form class="mainnav-form pull-right" role="search" style="margin-top:-15px;margin-bottom:15px;">
          <input type="text" class="form-control input-md mainnav-search-query" placeholder="Search">
          <button class="btn btn-sm mainnav-form-btn"><i class="fa fa-search"></i></button>
        </form>
        -->

          <form action="budget_firapport.asp?media=export" method="post" target="_blank">
              <input type="submit" class="btn btn-sm btn-secondary pull-right" value="Export as .CSV file" style="margin-top:-15px;margin-bottom:15px;">
              <br /><br />

          </form>
        
          <form action="budget_firapport.asp?func=opr" method="post">


        <input type="submit" class="btn btn-sm btn-success pull-right" value="Create" style="margin-top:-15px;margin-bottom:15px;">
        


            </form>
        <div class="portlet-body">

           <form action="budget_firapport.asp?func=slet" method="post">
          <table class="table table-striped table-bordered" id="xtable-1">
            <thead>
              <tr>
                <th class="text-center" style="width: 5%">Id</th>
                <th style="width: 15%">Budget title</th>
                <th style="width: 15%">Customer - Job</th>
                <th style="width: 15%">Period</th>
                <th style="width: 10%">Ref. No.</th>
                <th style="width: 12%">Prod. No.</th>
                <th style="width: 10%">Budget</th>
                <th style="width: 10%">Extra fo.</th>
                  <th>Delete</th>
                
              </tr>
            </thead>
              <tbody>


                  <%
                  end if
                      
                          
                      
                  strSQLlist = "SELECT bj.id AS id, budgetnavn, kkundenavn, jobnavn, budget_tot, budget_extra_fo, refno, prodno, stdato, sldato "_
                  &" FROM budget_job AS bj "_
                  &" LEFT JOIN job AS j ON (j.id = bj.jobid) "_
                  &" LEFT JOIN kunder AS k ON (kid = j.jobknr) WHERE bj.id <> 0 AND budgetstatus = 1 ORDER BY budgetnavn" 

                      'response.write strSQLlist
                      'response.Flush

                       oRec.open strSQLlist, oConn, 3
                        while not oRec.EOF 

                       
                       if media <> "export" then
                        %>

                  <tr>
                   
                        <td class="text-center"><%=oRec("id") %></td>
                        <td><a href="budget_firapport.asp?func=red&id=<%=oRec("id") %>"><%=oRec("budgetnavn") %></a></td>
                        <td><%=left(oRec("kkundenavn"), 20) %> - <%=left(oRec("jobnavn"), 20) %></td>
                        <td><%=oRec("stdato") %> - <%=oRec("sldato") %></td>
                        <td><%=oRec("refno") %></td>
                        <td><%=oRec("prodno") %></td>
                        <td><%=oRec("budget_tot") %></td>
                        <td><%=oRec("budget_extra_fo") %></td>
                          <td align=center><input type="checkbox" name="FM_budgetid_del" value="<%=oRec("id") %>" />
                          
                          </td>
                    
                    </tr>
                      

                        <%
                        end if 'media 


                        ekspTxt = ekspTxt & oRec("budgetnavn") &";"& oRec("stdato") &";"& oRec("sldato") &";"& oRec("kkundenavn") &";"& oRec("jobnavn") & ";" & oRec("refno") &";"& oRec("prodno") &";"& oRec("budget_tot") &";"& oRec("budget_extra_fo") & ";;;;;;xx99123sy#z"
                                

                                   if media = "export" then 'expences
                                
                                        strSQLExp = "SELECT id, fase, aktnavn, budget_app, konto FROM budget_job_exp WHERE budgetid = "& oRec("id") &"  ORDER BY fase, aktnavn"

                                          oRec3.open strSQLExp, oConn, 3
                                          while not oRec3.EOF 
                                
                                            ekspTxt = ekspTxt & oRec("budgetnavn") &";"& oRec("stdato") &";"& oRec("sldato") &";"& oRec("kkundenavn") &";"& oRec("jobnavn") & ";" & oRec("refno") &";"& oRec("prodno") &";"& oRec("budget_tot") &";"& oRec("budget_extra_fo") &";"& oRec3("konto") &";"& oRec3("fase") &";"& oRec3("aktnavn") &";"& oRec3("budget_app") &";;xx99123sy#z"
                        
                                    
                                                          strSQLExp_rel = "SELECT budgetid, aktid, belob, dato FROM budget_job_exp_rel WHERE "& oRec("id") &" AND aktid = "& oRec3("id") & " ORDER BY dato"

                                                          'response.write strSQLExp_rel
                                                           'response.Flush

                                                          oRec4.open strSQLExp_rel, oConn, 3
                                                          while not oRec4.EOF 

                                                                  ekspTxt = ekspTxt & oRec("budgetnavn") &";"& oRec("stdato") &";"& oRec("sldato") &";"& oRec("kkundenavn") &";"& oRec("jobnavn") & ";" & oRec("refno") &";"& oRec("prodno") &";"& oRec("budget_tot") &";"& oRec("budget_extra_fo") &";"& oRec3("konto") &";"& oRec3("fase") &";"& oRec3("aktnavn") &";"& oRec3("budget_app") &";"& oRec4("dato") &";"& oRec4("belob") &"xx99123sy#z"


                                                         oRec4.movenext
                                                         wend
                                                         oRec4.close

                                    
                                         oRec3.movenext
                                         wend
                                         oRec3.close

                                   end if

                        oRec.movenext
                        wend
                        oRec.close
                            
                            
                            
                         
                        if media <> "export" then
                            %>
                     

                
              
                        </tbody>
                        <tfoot>
                          <tr>
             
                            <th class="text-center">Id</th>
                            <th>Budget title</th>
                            <th>Customer</th>
                            <th>Period</th>
                            <th>Ref. No.</th>
                            <th>Prod. No.</th>
                            <th>Budget</th>
                            <th>Extra fo.</th>
                            <th></th>
                
                          </tr>
                        </tfoot>
                      </table>

                              <input type="submit" class="btn btn-danger btn-sm pull-right" value="Delete" style="margin-top:-15px;margin-bottom:15px;">

                        </form>

                    </div> <!-- /.portlet-body -->
                  </div> <!-- /.portlet -->
                </div> <!-- /.container -->




                <% end if 'media
      
            '******************* Eksport **************************' 
                if media = "export" then


    
                    call TimeOutVersion()
    


                  

	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	                'datointerval = request("datointerval")
	
	                lto = session("lto")
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\budget_firapport.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\budget_firapportexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\budget_firapportexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\budget_firapportexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\budget_firapportexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "budget_firapportexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
                              
                                strOskrifter = "Budget;St.date;Enddate;Customer;Project;Ref. No.;Prod. No.;Budget tot.;Budget Extra Fo.;Account;Phase;Act.name;Budget app.;Expend Date;Expend Amount;"
                                strOskrifter = strOskrifter & strExportOskriftDage

                               
				
				                'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="position:absolute; top:-100px; left:60px; padding:20px; background-color:#ffffff;">
	                                <table border=0 cellspacing=1 cellpadding=0 width="200">
	                                <tr><td valign=top style="padding:5px;">
	                                <img src="../ill/outzource_logo_200.gif" />
	                                </td>
	                                </tr>
	                                <tr>
	                                <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	                                <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Your CSV. file is ready >></a>
	                                </td></tr>
	                                </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if 'media%>
          
          
          
              
          
          
      %>



    <%end select %>



  </div> <!-- .content -->
</div> <!-- /#wrapper -->











</body>
</html>