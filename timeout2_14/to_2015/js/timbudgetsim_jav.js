






$(window).load(function () {

    //visSubtotJobmed();

});


    $(document).ready(function () {

        
        $("#load").hide(1000);
        

        
        $("#sorterjobnr").click(function () {

            $("#sorter").val(1); 
            $("#sogmainlist").submit();

        });

        $("#sorterjobnavn").click(function () {

            $("#sorter").val(0);
            $("#sogmainlist").submit();

        });



        
        $("#sp_luk_fc").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });


        $(".sp_jid").mouseover(function () {

            $(this).css('cursor', 'pointer');


        });



        $(".sp_jid").click(function () {

           var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(3, idlngt)




            if ($(".tr_" + idtrim).css('display') == "none") {

                $(".tr_" + idtrim).css("display", "");
                $(".tr_" + idtrim).css("visibility", "visible");
                $(".tr_" + idtrim).show(2000);

                str = idtrim
                var n = str.indexOf("_");
                if (n == -1) {

                    $("#sp_" + idtrim).html('<b>[-]&nbsp;</b>');
                    $("#sm_" + idtrim).html('<b>[-]&nbsp;</b>');
                }



            } else {

                $(".tr_" + idtrim).css("display", "none");
                $(".tr_" + idtrim).css("visibility", "hidden");
                $(".tr_" + idtrim).hide(300);

                str = idtrim
                var n = str.indexOf("_");

                if (n == -1) {
                    $("#sp_" + idtrim).html('<b>[+]&nbsp;</b>');
                    $("#sm_" + idtrim).html('<b>[+]&nbsp;</b>');
                }

            }


        });


        $(".fp_jid").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });


        $(".sp_p").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });

        $("#a_afdm").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });

        $("#a_jobakt").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });
        
   
       

        $("#minit").keyup(function () {  
           
            disVisAllejob();
        });


        //$("#allejob").click(function () {


        //if ($('#allejob').is(':checked') != true) {
        //    $("#jq_jobid").removeAttr("disabled");
        //} else {
        //    $("#jq_jobid").attr("disabled", true);
        //}
        
        //});

        
        function disVisAllejob() {

            //alert("her")

            var thisVal = $("#minit").val()

            //alert(thisVal.length)

            if (thisVal.length > 0) {
                $("#allejob").removeAttr("disabled");
                $("#viskunemedarbfcreal").attr("disabled", true);
                $("#viskunemedarbfcreal").prop("checked", false);
                $("#progrpid").attr("disabled", true);

            } else {
                    //if ($("#progrpid").val() != 0) {
                        $("#allejob").removeAttr("disabled");
                //} else {
                //       $("#allejob").prop("checked", false);
                //      $("#allejob").attr("disabled", true);
                //  }
                $("#viskunemedarbfcreal").removeAttr("disabled");
                $("#progrpid").removeAttr("disabled");
            }

        }

        disVisAllejob();


        $("#progrpid").change(function () {

            if ($("#progrpid").val() != 0) {
                $("#allejob").removeAttr("disabled");
            } else {
                $("#allejob").attr("disabled", true);
            }

        });
        

        $(".sp_p").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(5, idlngt)

            
           

            if ($(".afd_p_" + idtrim).css('display') == "none") {


                //$(".afd_p").css("display", "none");
                //$(".afd_p").css("visibility", "hidden");
                $(".afd_p").hide(100);
                $(".td_p").css("background-color", "");

                $(".afd_p_" + idtrim).css("display", "");
                $(".afd_p_" + idtrim).css("visibility", "visible");
                $("#td_p_" + idtrim).css("background-color", "#FFFFE1");
                $("#an_p_" + idtrim).html("[-]");
                $(".afd_p_" + idtrim).show(100);

               


            } else {

                $(".afd_p_" + idtrim).css("display", "none");
                $(".afd_p_" + idtrim).css("visibility", "hidden");
                $("#td_p_" + idtrim).css("background-color", "");
                $("#an_p_" + idtrim).html("[+]");
                $(".afd_p_" + idtrim).hide(100);

               

            }


        });


        $("#a_afdm").click(function () {

            
                //$("#d_jobakt").hide(100);
              
                $("#dv_afdm").css("display", "");
                $("#dv_afdm").css("visibility", "visible");
                $("#dv_afdm").show(100);

             


         
        
        });


        $("#a_jobakt").click(function () {

         
            $("#dv_afdm").hide(100);


                $("#d_jobakt").css("display", "");
                $("#d_jobakt").css("visibility", "visible");
                $("#d_jobakt").show(100);


        });

        $("#sp_luk_fc").click(function () {


            $("#dv_afdm").hide(100);


            $("#d_jobakt").css("display", "");
            $("#d_jobakt").css("visibility", "visible");
            $("#d_jobakt").show(100);


        });


     
    

        $(".fp_jid").click(function () {


            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(3, idlngt)

           
           

            if ($(".tr_" + idtrim).css('display') == "none") {

                $(".tr_" + idtrim).css("display", "");
                $(".tr_" + idtrim).css("visibility", "visible");
                $("#tr_job_" + idtrim).css("background-color", "#FFFFE1");
                $(".tr_" + idtrim).show(2000);

                str = idtrim
                var n = str.indexOf("_");
                if (n == -1) {
                    
                    $("#an_" + idtrim).html('<b>[-]&nbsp;</b>');
                    $("#sp_" + idtrim).html('<b>[-]&nbsp;</b>');
                    $("#sm_" + idtrim).html('<b>[-]&nbsp;</b>');
                }

          

            } else {

                $(".tr_" + idtrim).css("display", "none");
                $(".tr_" + idtrim).css("visibility", "hidden");
                $("#tr_job_" + idtrim).css("background-color", "");
                $(".tr_" + idtrim).hide(300);

                str = idtrim
                var n = str.indexOf("_");
            
                if (n == -1) {
                    $("#an_" + idtrim).html('<b>[+]&nbsp;</b>');
                    $("#sp_" + idtrim).html('<b>[+]&nbsp;</b>');
                    $("#sm_" + idtrim).html('<b>[+]&nbsp;</b>');
                }

            }
      


        });


        // farveskift ved tildel budgettimer //
        $(".jobakt_budgettimer_FY, .jobakt_budgettimer_job").keyup(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

          
            budgettimer = $("#FM_timerbudget_FY0_" + idtrim).val().replace(",", ".")
            realtimer = $("#realtimer_jobakt_" + idtrim).val().replace(",", ".")
            
            thisval = Math.round((budgettimer / 1 - realtimer / 1) * 100) / 100

            if (thisval < 0) {
                $("#h2t_jobakt_" + idtrim).css("background-color", "lightpink")
            } else {
                $("#h2t_jobakt_" + idtrim).css("background-color", "")
            }

            thisval = String(thisval).replace(".", ",")

            $("#h2t_jobakt_" + idtrim).val(thisval)



            budgettimer1 = $("#FM_timerbudget_FY1_" + idtrim).val().replace(",", ".")
            budgettimer2 = $("#FM_timerbudget_FY2_" + idtrim).val().replace(",", ".")

            budgettimerTotFY = (budgettimer / 1 + budgettimer1 / 1 + budgettimer2 / 1)
            jobbudgettimer = $("#budgettimer_jobakt_" + idtrim).val().replace(",", ".")


            if ((jobbudgettimer/1 - budgettimerTotFY/1) < 0) {
                $("#budgettimer_jobakt_" + idtrim).css("background-color", "lightpink")
            } else {
                $("#budgettimer_jobakt_" + idtrim).css("background-color", "")
            }


          


            //alert(budgettimer)

            //budgettimerColor(h1totjobakt, h2totjobakt, budgettimer, idtrim);

        });



        ////////////////////////////////////////////////////////////////////////////////////
        // Beregn H1H2 MEDARBEJDER FORECAST///
        $(".mh1, .mh2, mh1t").keyup(function () {

           
            //medarbejder total på job / akt h1 + h2
            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(15, idlngt)


            var arr = idtrim.split('_');

            jobid = $("#h1_jobid_" + idtrim).val()
            aktid = $("#h1_aktid_" + idtrim).val()

            //////////////////////////////////////////////////////////////////////////////////
            //// AKT --> Joblinje

            
        

           
            // H1 //
            h1jobTot = 0;

            if (aktid != 0) { //Akt sumer til job

              

               


                $(".mh1h_akt_" + jobid + "_" + arr[2]).each(function () {

                    


                    h1job = $(this).val().replace(".", "")
                    h1job = h1job.replace(",", ".")

                    if (h1job == "NaN" || (h1job == "-Infinity") || h1job == "") {
                        h1job = 0
                    }

                    h1jobTot = (h1jobTot / 1 + h1job / 1)


                });





                //alert("#mh1h_jobaktmid_" + jobid + "_0_" + "_" + arr[2])

                h1jobTot = Math.round(h1jobTot)
                h1jobTot = h1jobTot
                h1jobTot = String(h1jobTot).replace(".", ",")
                $("#mh1h_jobaktmid_" + jobid + "_0_" + arr[2]).val(h1jobTot)

                //$("#h1t_jobakt_" + jobid + "_0").val(h1jobTot) //h1t_jobakt


            }

           

            // H2 //
            h2jobTot = 0;
            /*
            if (aktid != 0) { //Akt sumer til job

                $(".mh2h_akt_" + jobid + "_" + arr[2]).each(function () {



                    h2job = $(this).val().replace(".", "")
                    h2job = h2job.replace(",", ".")

                    if (h2job == "NaN" || (h2job == "-Infinity") || h2job == "") {
                        h2job = 0
                    }

                    h2jobTot = (h2jobTot / 1 + h2job / 1)

                   
                });

 
                h2jobTot = Math.round(h2jobTot)
                h2jobTot = h2jobTot
                h2jobTot = String(h2jobTot).replace(".", ",")
                $("#mh2h_jobaktmid_" + jobid + "_0_" + arr[2]).val(h2jobTot)
               
            }
            */

            
          

            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////



            beregnh1h2(idtrim);

           
          
        });


        
      

        

        /////////////////////////////////////////////////////////////////////////////////////////



        function beregnh1h2(idtrim) {


            

            h1 = $("#mh1h_jobaktmid_" + idtrim).val()
            var h2 = 0; //$("#mh2h_jobaktmid_" + idtrim).val()

           

            ///H1 - Real Saldo
            forecastTimer = $("#mh1h_jobaktmid_" + idtrim).val().replace(",", ".") / 1
            forecastTpris = $("#mh1t_jobaktmid_" + idtrim).val().replace(",", ".") / 1

            realTimer = $("#mreal_jobaktmid_" + idtrim).val().replace(",", ".") / 1

            if (forecastTimer == "NaN" || (forecastTimer == "-Infinity") || forecastTimer == "") {
                forecastTimer = 0
            }

            if (forecastTpris == "NaN" || (forecastTpris == "-Infinity") || forecastTpris == "") {
                forecastTpris = 0
            }

            if (realTimer == "NaN" || (realTimer == "-Infinity") || realTimer == "") {
                realTimer = 0
            }

            timerSaldo = (forecastTimer - realTimer)
            forecastBelob = (forecastTimer * forecastTpris)
            //alert(timerSaldo)

           
          
            h1 = h1.replace(",", ".")
            //h2 = h2.replace(",", ".")

        
            if (h1 == "NaN" || (h1 == "-Infinity") || h1 == "") {
                h1 = 0
            }

            //if (h2 == "NaN" || (h2 == "-Infinity") || h2 == "") {
            //    h2 = 0
            //}

           
            h1h2 = Math.round(h1 / 1 + h2 / 1)
            h1h2 = String(h1h2).replace(".",",")
            //alert("HER 7")
           

            if (timerSaldo < 0) {
                $("#mh12h_jobaktmid_" + idtrim).css("background-color", "lightpink")

            } else {
                $("#mh12h_jobaktmid_" + idtrim).css("background-color", "#Eff3ff")

            }


            
            timerSaldo = String(timerSaldo).replace(".", ",")
            $("#mh12h_jobaktmid_" + idtrim).val(timerSaldo) // h1h2
            $("#mh12tp_jobaktmid_" + idtrim).val(forecastBelob) // h1h2
            $("#mh12tps_jobaktmid_" + idtrim).html(forecastBelob + " DKK") // h1h2

            jobid = $("#h1_jobid_" + idtrim).val()
            aktid = $("#h1_aktid_" + idtrim).val()

            h1totjobakt = 0
            h2totjobakt = 0

            h1jobakt = 0
            h2jobakt = 0

         

            //var arr = idtrim.split('_');
            thisP = $("#mh1h_jobaktafd_" + idtrim).val()
            //alert(thisP)



            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // TOTAL PR. JOB/  AKT LINJE
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            /////////////////////////////////////////////
            // Timer
            /////////////////////////////////////////////
            $(".mh1h_jobaktmid_" + jobid + "_" + aktid + "_" + thisP).each(function () {

               
             
                //alert("mh1h_jobaktmid_" + jobid + "_" + aktid + "_" + thisP + "// this:" + this)

                h1jobakt = $(this).val()
                h1TPjobakt = $("#mh1t_jobaktmid_" + jobid + "_" + aktid + "_" + thisP).val()
               
                if (h1jobakt == "NaN" || (h1jobakt == "-Infinity") || h1jobakt == "") {
                    h1jobakt = 0
                    
                } else {
                    h1jobakt = h1jobakt.replace(".", "")
                    h1jobakt = h1jobakt.replace(",", ".")
                }

                h1totjobakt = h1totjobakt / 1 + (h1jobakt / 1)

                
            });


            h1totjobakt = String(h1totjobakt).replace(".", ",")
            //alert("#afd_jobakt_" + jobid + "_" + aktid + "_" + thisP)
            
            $("#afd_jobakt_" + jobid + "_" + aktid + "_" + thisP).val(h1totjobakt)
            $("#fcjobaktgt_" + jobid + "_" + aktid).val(0)
           
            


            /////////////////////////////////////////////
            ////Belob
            /////////////////////////////////////////////
            h1totBeljobakt = 0;
            $(".mbel_" + jobid + "_" + aktid).each(function () {

                //alert("mh1h_jobaktmid_" + jobid + "_" + aktid + "_" + thisP + "// this:" + this)

                h1TPjobakt = $(this).val()

                //alert(h1TPjobakt)
              
                if (h1TPjobakt == "NaN" || (h1TPjobakt == "-Infinity") || h1TPjobakt == "") {
                    h1TPjobakt = 0

                } else {
                    h1TPjobakt = h1TPjobakt.replace(".", "")
                    h1TPjobakt = h1TPjobakt.replace(",", ".")
                }

                h1totBeljobakt = h1totBeljobakt / 1 + (h1TPjobakt / 1)

            });

            

            h1totBeljobakt = String(h1totBeljobakt).replace(".", ",")
            //alert(h1totBeljobakt)
            $("#afd_jobaktbel_" + jobid + "_" + aktid + "_" + thisP).val(h1totBeljobakt)
            $("#afd_jobaktbels_" + jobid + "_" + aktid + "_" + thisP).html(h1totBeljobakt + " DKK")

            /////////////////////////////////////////////
            //// Belob GT på Job/AKt
            /////////////////////////////////////////////
            h1totBeljobaktGt = 0;

            $(".afd_jobaktbel_" + jobid + "_" + aktid).each(function () {

                //alert("j:" + jobid + " a: " + aktid)

                h1Beljobakt = $(this).val()

                //h1Beljobakt = h1Beljobakt.replace(" DKK", "")

                //alert(h1Beljobakt)


                if (h1Beljobakt == "NaN" || (h1Beljobakt == "-Infinity") || h1Beljobakt == "") {
                    h1Beljobakt = 0

                } else {
                    h1Beljobakt = h1Beljobakt.replace(".", "")
                    h1Beljobakt = h1Beljobakt.replace(",", ".")
                }

                h1totBeljobaktGt = h1totBeljobaktGt / 1 + (h1Beljobakt / 1)

            });



            h1totBeljobaktGt = String(h1totBeljobaktGt).replace(".", ",")
            //alert(h1totBeljobaktGt)
            $("#fcjobaktBelgt_" + jobid + "_" + aktid).val(h1totBeljobaktGt)
            $("#fcjobaktBelgts_" + jobid + "_" + aktid).html(h1totBeljobaktGt + " DKK")
            $("#fcjobaktBelgts_" + jobid + "_" + aktid).css("border-bottom", "1px darkred dashed")
            //////// END Belob GT
          
           

           

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Total på projektgruppe LODRET (sum af aktiviteter)
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            afDGt(thisP);

            visallejob = $("#visallejob").val();
            viskunminit = $("#viskunminit").val();

            // Total på job SUMMER aktiviteter lodret
            if (visallejob == 1) {
                //jobfcGt();
            }
          

            if (visallejob != 1 && viskunminit != 1) {
                beregnFCprProgrpTilTotPaJob(jobid,aktid);
            }

           

            /* $(".mh2h_jobaktmid_" + jobid + "_" + aktid).each(function () {

                

                h2jobakt = $(this).val().replace(".", "")
                h2jobakt = h2jobakt.replace(",", ".")
                if (h2jobakt == "NaN" || (h2jobakt == "-Infinity") || h2jobakt == "") {
                    h2jobakt = 0
                }


               
                h2totjobakt = h2totjobakt + (h2jobakt / 1)

            }); */


           

            
           
            $("#h1h_jobaktmid_" + jobid + "_" + aktid).val(h1totjobakt)
            //$("#h2h_jobaktmid_" + jobid + "_" + aktid).val(h2totjobakt)

            h1totjobaktTotLinje = h1totjobakt

            String(h1totjobaktTotLinje).replace(".", ",")
            $("#afd_jobakt_" + jobid + "_" + aktid).val(h1totjobaktTotLinje)
            //$("#afd_jobaktbel_" + jobid + "_" + aktid).val('600')

   



            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // GT på Job / AKT 
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            h1totjobaktGT = 0;
            $(".mh1h_jobaktmid_" + jobid +"_0").each(function () {

                h1jobakt = $(this).val().replace(".", "")

             
                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(14, idlngt)

                var arr = idtrim.split('_');

                //alert(this.id + " aar1: " + arr[1])

                if (arr[1] != 0) {
                //h1h_jobaktmid_<%=jobid%>_<%=aktid %>

                
                    if (h1jobakt == "NaN" || (h1jobakt == "-Infinity") || h1jobakt == "") {
                        h1jobakt = 0
                    } else {
                        h1jobakt = h1jobakt.replace(",", ".")
                    }


                //alert(thisid +" "+ h1jobakt)

                h1totjobaktGT = h1totjobaktGT + (h1jobakt / 1)

                }


            });


            h1totjobaktGT = Math.round(h1totjobaktGT)
            h1totjobaktGT = h1totjobaktGT 
            h1totjobaktGT = String(h1totjobaktGT ).replace(".", ",")


            //// Total beløb lodret på job / akt
            beljobaktGT = 0;
            $(".fcjobaktBelgt").each(function () {

                

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(14, idlngt)

                var arr = idtrim.split('_');

                //alert(this.id + " aar1: " + arr[1])
                jobaktBel = $(this).val()
                //jobaktBel = jobaktBel.replace(",", ".")
                if (arr[1] != 0) {

                   
                    if (jobaktBel == "NaN" || (jobaktBel == "-Infinity") || jobaktBel == "") {
                        jobaktBel = 0
                    } else {
                        jobaktBel = jobaktBel.replace(".", "")
                        jobaktBel = jobaktBel.replace(",", ".")
                        
                    }


                    
                    beljobaktGT = beljobaktGT/1 + (jobaktBel / 1)
                    //alert("B: " + jobaktBel + ": " + beljobaktGT)

                }

            });
            //alert(jobid +"BGT: " + beljobaktGT)


            //alert(jobid)
            beljobaktGT = Math.round(beljobaktGT)
            beljobaktGT = beljobaktGT
            beljobaktGT = String(beljobaktGT).replace(".", ",")
            $("#fcjobaktBelgts_" + jobid + "_0").html(beljobaktGT + " DKK")
            $("#fcjobaktBelgt_" + jobid + "_0").val(beljobaktGT)


            // SALDO PÅ job
            fcjobaktgtAkkVal = 0
            jobaktBudget = 0
            jobaktBudget = $("#jobaktBudget_" + jobid + "_0").val()

            

            if (jobaktBudget == "NaN" || (jobaktBudget == "-Infinity") || jobaktBudget == "") {
                jobaktBudget = 0
            } else {
                jobaktBudget = jobaktBudget.replace(".", "")
                jobaktBudget = jobaktBudget.replace(",", ".")
            }


            fcjobaktgtAkkVal = $("#fcjobaktBelgt_" + jobid + "_0").val()
            if (fcjobaktgtAkkVal == "NaN" || (fcjobaktgtAkkVal == "-Infinity") || fcjobaktgtAkkVal == "") {
                fcjobaktgtAkkVal = 0
            } else {
                fcjobaktgtAkkVal = fcjobaktgtAkkVal.replace(".", "")
                fcjobaktgtAkkVal = fcjobaktgtAkkVal.replace(",", ".")
            }

          
            fcjobaktgtAkkSaldo = 0
            fcjobaktgtAkkSaldo = (jobaktBudget / 1 - fcjobaktgtAkkVal / 1)

            jobaktBgtTimValGtGt = 0;
            jobaktBgtTimValGt = $("#jobaktT_" + jobid + "_0").val()
            if (jobaktBgtTimValGt == "NaN" || (jobaktBgtTimValGt == "-Infinity") || jobaktBgtTimValGt == "") {
                jobaktBgtTimValGt = 0
            } else {
                jobaktBgtTimValGt = jobaktBgtTimValGt.replace(".", "")
                jobaktBgtTimValGt = jobaktBgtTimValGt.replace(",", ".")
            }

            //alert("HER: " + jobid + " arr[0] ")

            //alert(jobaktBudget + " - " + fcjobaktgtAkkVal)
            //alert(fcjobaktgtAkkSaldo)
            //$("#fcjobaktBelgt_" + jobid + "_" + aktid).html('200 DKK')
            //$("#saldoBelgt_" + jobid + "_0").val(fcjobaktgtAkkSaldo)

            
            jobaktFCGt = 0;
            jobaktFCGt = $("#fcjobaktgt_" + jobid + "_0").val()
            if (jobaktFCGt == "NaN" || (jobaktFCGt == "-Infinity") || jobaktFCGt == "") {
                jobaktFCGt = 0
            } else {
                jobaktFCGt = jobaktFCGt.replace(".", "")
                jobaktFCGt = jobaktFCGt.replace(",", ".")
            }
            
            jobaktBgtTimValGt = (jobaktBgtTimValGt / 1 - jobaktFCGt / 1)

            $("#saldoTgts_" + jobid + "_0").html(jobaktBgtTimValGt + " t.")
            $("#saldoBelgts_" + jobid + "_0").html(fcjobaktgtAkkSaldo + " DKK")
            

            //H2
            h2totjobaktGT = 0;
            /*
            $(".mh2h_jobaktmid_" + jobid + "_0").each(function () {

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(14, idlngt)

                var arr = idtrim.split('_');

                if (arr[1] != 0) {

                    h2jobakt = $(this).val().replace(".", "")
                    h2jobakt = h2jobakt.replace(",", ".")
                    if (h2jobakt == "NaN" || (h2jobakt == "-Infinity") || h2jobakt == "") {
                        h2jobakt = 0
                    }



                    h2totjobaktGT = h2totjobaktGT + (h2jobakt / 1)

                }

            });

            h2totjobaktGT = Math.round(h2totjobaktGT)
            h2totjobaktGT = h2totjobaktGT 
            h2totjobaktGT = String(h2totjobaktGT).replace(".", ",")
            */




            /////////////////////////////////////////////////////////////
            /// Jobniveau
            ////////////////////////////////////////////////////////////
            $("#h1h_jobaktmid_" + jobid + "_0").val(h1totjobaktGT)
            //$("#h1t_jobaktmid_" + jobid + "_0").val(1000)

            h2totjobaktGT = 0;
            //$("#h2h_jobaktmid_" + jobid + "_0").val(h2totjobaktGT)
         





            // JOB og AKT, Afd og MEDARB Sammentæller totaler ///
          
          


            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // AFDELINGER
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            phigh = $("#phigh").val()

            var arr = idtrim.split('_');
            m_arr = arr[2]  
            //mh1h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>

            medarbTot(aktid, jobid, m_arr)

            afdelingTot(aktid, jobid, phigh);
           



        }


        //////////////////////////////////////////////////////////////////////////////////////







        function afdelingTot(aktid, jobid, phigh) {


            for (p = 0; p <= phigh; p++) {

                fs_totTot = 0
                $(".fs_" + jobid + "_" + aktid + "_" + p).each(function () {



                    fs_tot = $(this).val().replace(".", "")
                    fs_tot = fs_tot.replace(",", ".")


                    if (fs_tot == "NaN" || (fs_tot == "-Infinity") || fs_tot == "") {
                        fs_tot = 0
                    }

                    fs_totTot = (fs_totTot / 1 + fs_tot / 1)

                    //alert(fs_totTot)


                });

                fs_totTot = String(fs_totTot).replace(".", ",")

                $("#afd_jobakt_" + jobid + "_" + aktid + "_" + p).val(fs_totTot)
                //$("#afd_jobaktbel_" + jobid + "_" + aktid + "_" + p).val('200')

            }


        }


        function medarbTot(aktid, jobid, m) {

            //mh2h_akt_"& jobid & "_"& antalm(m,1)

           //alert(m)

                mFc_totTot = 0
                $(".mh1h_akt_" + jobid + "_" + m).each(function () {


                

                    mFc_tot = $(this).val().replace(".", "")
                    mFc_tot = mFc_tot.replace(",", ".")


                    if (mFc_tot == "NaN" || (mFc_tot == "-Infinity") || mFc_tot == "") {
                        mFc_tot = 0
                    }

                    mFc_totTot = (mFc_totTot / 1 + mFc_tot / 1)

                   


                });

            //alert(mFc_totTot)
                mFc_totTotGT = 0;
                mFc_totTotGT = $("#totmedarbh1GT_" + m).val().replace(".", "")
                mFc_totTotGT = mFc_totTotGT.replace(",", ".")

                mFc_totTotGTGT = (mFc_totTotGT / 1 + mFc_totTot / 1)

                //alert("her")

                totmedarbn1 = 0;
                totmedarbn1 = $("#totmedarbn1_" + m).val().replace(".", "")
                totmedarbn1 = totmedarbn1.replace(",", ".")

                if (mFc_totTotGTGT > totmedarbn1) {
                   $("#totmedarbn1_" + m).css("background-color", "lightpink")
                } else {
                   $("#totmedarbn1_" + m).css("background-color", "")
                }

                mFc_totTot = String(mFc_totTot).replace(".", ",")
                mFc_totTotGTGT = String(mFc_totTotGTGT).replace(".", ",")

                $("#totmedarbh1_"+ m).val(mFc_totTot)
                $("#totmedarbh1GTGT_" + m).val(mFc_totTotGTGT)



            


        }













        //////////////////////////////////////////////////////////////////////////////////////
        // SAMMENTÆLLING AKT --> JOB
        //////////////////////////////////////////////////////////////////////////////////////



        //////////////////////////////////////////////////////////////////////////////////
        //// Timepriser JOB
        $(".jobakt_budgettimep_job").keyup(function () {

            //alert("her")

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');

            //alert(idtrim)

            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // TPJOB //
            jobbudgettp = 0;
            if (arr[1] != 0) { //Akt sumer til job
                i = 0
                $(".jobakt_budgettimep_" + arr[0]).each(function () {



                    tottp1 = $(this).val().replace(".", "")
                    tottp1 = tottp1.replace(",", ".")

                    

                    if (tottp1 == "NaN" || (tottp1 == "-Infinity") || tottp1 == "") {
                        tottp1 = 0
                    }

                    jobbudgettp = (jobbudgettp / 1 + tottp1 / 1)

                    if (tottp1 != 0) {
                        i = i + 1
                    }



                });

                if (i == 0) {
                    i = 1
                }

                
                jobbudgettp = jobbudgettp / i
                jobbudgettp = Math.round(jobbudgettp)
                jobbudgettp = String(jobbudgettp).replace(".", ",")

                //alert("jer" + jobbudgettp)

                $("#budgettimep_jobakt_" + arr[0] + "_0").val(jobbudgettp)
                

            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            //totalalleh1h2(1)

        });


















        //////////////////////////////////////////////////////////////////////////////////
        //// JOBTIMER
        $(".jobakt_budgettimer_job").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY0 //
            jobbudgetTot = 0;
            if (arr[1] != 0) { //Akt sumer til job

                $(".jobakt_budgettimer_" + arr[0]).each(function () {



                    jobbudget = $(this).val().replace(".", "")
                    jobbudget = jobbudget.replace(",", ".")

                    //alert(jobbudget)

                    if (jobbudget == "NaN" || (jobbudget == "-Infinity") || jobbudget == "") {
                        jobbudget = 0
                    }

                    jobbudgetTot = (jobbudgetTot / 1 + jobbudget / 1)


                });



                //jobbudgetTot = Math.round((jobbudgetTot) / 100) * 100
                jobbudgetTot = jobbudgetTot
                jobbudgetTot = String(jobbudgetTot).replace(".", ",")
                $("#budgettimer_jobakt_" + arr[0] + "_0").val(jobbudgetTot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            //totalalleh1h2(1)

        });




      



        //////////////////////////////////////////////////////////////////////////////////
        //// FY 0 JOB
        $(".jobakt_budgettimer_FY").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY0 //
            totFy0Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
               
                $(".jobakt_fy0_" + arr[0]).each(function () {



                    totFy0 = $(this).val().replace(".", "")
                    totFy0 = totFy0.replace(",", ".")

                    //alert(totFy0)

                    if (totFy0 == "NaN" || (totFy0 == "-Infinity") || totFy0 == "") {
                        totFy0 = 0
                    }

                    totFy0Tot = (totFy0Tot / 1 + totFy0 / 1)

                 
                });

              

                //totFy0Tot = Math.round((totFy0Tot) / 100) * 100
                totFy0Tot = totFy0Tot 
                totFy0Tot = String(totFy0Tot).replace(".", ",")
                $("#FM_timerbudget_FY0_" + arr[0] + "_0").val(totFy0Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            //totalalleh1h2(1)

        });


        //////////////////////////////////////////////////////////////////////////////////
        //// FY 1 JOB
        $(".jobakt_budgettimer_FY1").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY1 //
            totFy1Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
               
                $(".jobakt_fy1_" + arr[0]).each(function () {



                    totFy1 = $(this).val().replace(".", "")
                    totFy1 = totFy1.replace(",", ".")

                    //alert(totFy1)

                    if (totFy1 == "NaN" || (totFy1 == "-Infinity") || totFy1 == "") {
                        totFy1 = 0
                    }

                    totFy1Tot = (totFy1Tot / 1 + totFy1 / 1)

                   
                });

              

                //totFy1Tot = Math.round((totFy1Tot) / 100) * 100
                totFy1Tot = totFy1Tot
                totFy1Tot = String(totFy1Tot).replace(".", ",")
                $("#FM_timerbudget_FY1_" + arr[0] + "_0").val(totFy1Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            //totalalleh1h2(1)

        });




        //////////////////////////////////////////////////////////////////////////////////
        //// FY 2 JOB
        $(".jobakt_budgettimer_FY2").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY1 //
            totFy2Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
                $(".jobakt_fy2_" + arr[0]).each(function () {



                    totFy2 = $(this).val().replace(".", "")
                    totFy2 = totFy2.replace(",", ".")

                    //alert(totFy2)

                    if (totFy2 == "NaN" || (totFy2 == "-Infinity") || totFy2 == "") {
                        totFy2 = 0
                    }

                    totFy2Tot = (totFy2Tot / 1 + totFy2 / 1)

                   
                });

               

                //totFy2Tot = Math.round((totFy2Tot) / 100) * 100
                totFy2Tot = totFy2Tot 
                totFy2Tot = String(totFy2Tot).replace(".", ",")
                $("#FM_timerbudget_FY2_" + arr[0] + "_0").val(totFy2Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            //totalalleh1h2(1)

        });



        



        function beregntpH1(idtrim) {


            //alert("HER")

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // TP1 //
            tottp1Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
                i = 0
                $(".jobakt_tph1_" + arr[0]).each(function () {



                    tottp1 = $(this).val().replace(".", "")
                    tottp1 = tottp1.replace(",", ".")

                    //alert(tottp1)

                    if (tottp1 == "NaN" || (tottp1 == "-Infinity") || tottp1 == "") {
                        tottp1 = 0
                    }

                    tottp1Tot = (tottp1Tot / 1 + tottp1 / 1)

                    if (tottp1 != 0) {
                        i = i + 1
                    }



                });

                if (i == 0) {
                    i = 1
                }


                tottp1Tot = tottp1Tot / i
                tottp1Tot = Math.round(tottp1Tot)
                tottp1Tot = String(tottp1Tot).replace(".", ",")
                $("#fctimepris_h1_jobakt_" + arr[0] + "_0").val(tottp1Tot)


            }


        }






        //////////////////////////////////////////////////////////////////////////////////
        //// Timepriser H1 & H2 JOB
        $(".jobakt_tph1").keyup (function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(21, idlngt)
           

            beregntpH1(idtrim);


          
            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

               //totalalleh1h2(1)

        });



        //////////////////////////////////////////////////////////////////////////////////
        //// Timepriser H1 & H2 JOB
        $(".jobakt_tph2").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(21, idlngt)

            var arr = idtrim.split('_');

            // jobakt_tph2_<%=jobid %>



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // TP2 //
            tottp2Tot = 0;


            /*
            if (arr[1] != 0) { //Akt sumer til job
                i = 0
                $(".jobakt_tph2_" + arr[0]).each(function () {



                    tottp2 = $(this).val().replace(".", "")
                    tottp2 = tottp2.replace(",", ".")

                    //alert(tottp2)

                    if (tottp2 == "NaN" || (tottp2 == "-Infinity") || tottp2 == "") {
                        tottp2 = 0
                    }

                    tottp2Tot = (tottp2Tot / 1 + tottp2 / 1)

                    if (tottp2 != 0) {
                        i = i + 1
                    }



                });

                if (i == 0) {
                    i = 1
                }

                
                tottp2Tot = tottp2Tot / i
                tottp2Tot = Math.round(tottp2Tot)
                tottp2Tot = String(tottp2Tot).replace(".", ",")
                $("#fctimepris_h2_jobakt_" + arr[0] + "_0").val(tottp2Tot)

            }
            */

            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////



            //totalalleh1h2(1)

        });


     



        // SLUT sammentælling akt->job
        //////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////











        function beregnbudget(jobid, aktid){
        
          ////////////////////////////////////////////////////////////////////////////////
            // Budget //
            h1timer = 0;
            h1timer = $("#h1t_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")
            fctimepris = 0;
            fctimepris = $("#fctimepris_h1_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            h2timer = 0;
            h2timer = $("#h2t_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")
            fctimeprish2 = 0;
            fctimeprish2 = 0; //$("#fctimepris_h2_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            //alert(budgettimer)


            realtimer = 0;
            realtimer = $("#realtimer_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            realtimep = 0;
            realtimep = $("#realtimep_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            //alert(realtimer + "*" + realtimep + " = " + (realtimer / 1 * realtimep / 1) + " + " + h2timer / 1 + "*" + fctimeprish2 / 1 + "=" + (h2timer / 1 * fctimeprish2 / 1))

            budgeth1 = Math.round((h1timer / 1 * fctimepris / 1) / 100) * 100
            budgeth2 = Math.round(((realtimer / 1 * realtimep / 1) + (h2timer / 1 * fctimeprish2 / 1)) / 100) * 100

            //alert(budgeth2 / 1)

            budgeth1h2 = (budgeth1 / 1 + budgeth2 / 1)

            budgeth1 = String(budgeth1).replace(".", ",")
            budgeth2 = String(budgeth2).replace(".", ",")
            budgeth1h2 = String(budgeth1h2).replace(".", ",")
            $("#budget_h1_jobakt_" + jobid + "_" + aktid).val(budgeth1)
            //$("#budget_h2_jobakt_" + jobid + "_" + aktid).val(budgeth2)

            if (realtimer == 0) {
                $("#budget_jobakt_" + jobid + "_" + aktid).val(budgeth1)
            } else {

                $("#budget_jobakt_" + jobid + "_" + aktid).val(budgeth2)
            }

            //$("#budgetgt").css('background-color','#999999')
            $("#budgetgt").val('Opdater!')

            //alert("her")
            //////////////////////////////////////////////////////////////////////////////////////
        }

       


        function totalalleh1h2(loadberegn) {


            mhigh = $("#mhigh").val()
            xhigh = $("#xhigh").val()
            phigh = $("#phigh").val()

            //alert(phigh + " M:" + mhigh)
            //xhigh = 4

            // JOB og AKT Sammentæller totaler ///
            for (i = 0; i <= xhigh; i++) {

                

                jobid = $("#FM_jobid_" + i).val()
                aktid = $("#FM_aktid_" + i).val()

                //alert("her: " + jobid + "aktid: " + aktid)
                budgettimer = 0;
                realtimer = 0;
                alert(jobid + "_" + aktid + " budgettimer: " + budgettimer)
               

                h1 = $("#h1h_jobaktmid_" + jobid + "_" + aktid).val()
                $("#h1t_jobakt_" + jobid + "_" + aktid).val(h1)

                h1 = h1.replace(".", "")
                h1 = h1.replace(",", ".")

              

                h2 = $("#h2h_jobaktmid_" + jobid + "_" + aktid).val()
                $("#h2t_jobakt_" + jobid + "_" + aktid).val(h2)
                h2 = h2.replace(".", "")
                h2 = h2.replace(",", ".")

                //h2 = 0

                realtimer = $("#realtimer_jobakt_" + jobid + "_" + aktid).val().replace(".", "")
                realtimer = realtimer.replace(",",".")
                //realtimer = 0

                

                // budgettimer_jobakt_
                // FY0 budgetTimer
                budgettimer = $("#FM_timerbudget_FY0_" + jobid + "_" + aktid).val().replace(",", ".")
                budgettimer = budgettimer.replace(",",".")
                
                //budgettimer = 0
           


                if (h1 == "NaN" || h1 == "-Infinity" || h1 == "") {
                    h1 = 0
                }

                if (h2 == "NaN" || h2 == "-Infinity" || h2 == "") {
                    h2 = 0
                }

                if (realtimer == "NaN" || realtimer == "-Infinity" || realtimer == "") {
                    realtimer = 0
                }

                if (budgettimer == "NaN" || budgettimer == "-Infinity" || budgettimer == "") {
                    budgettimer = 0
                }

                //alert(h1 + "# " + h2 + "# " + budgettimer + "# " + realtimer)

                h1h2total = Math.round((h1/1 + h2/1) * 100) / 100
                h1h2total = String(h1h2total).replace(".", ",")
                $("#h12t_jobakt_" + jobid + "_" + aktid).val(h1h2total)
                

               
                
                budgettimerColor(h1, h2, budgettimer, jobid+"_"+aktid);


                if (loadberegn == 1) {
                
                    beregnbudget(jobid, aktid);



                        }


           
           


                            // MEDARBEJDERE / NORM
                            for (l = 0; l <= mhigh; l++) {

                                //alert(l)

                                mh1 = $("#totmedarbh1_" + l).val().replace(".", "")
                                mh1 = mh1.replace(",", ".")
                                mh2 = 0; //$("#totmedarbh2_" + l).val().replace(".", "")
                                mh2 = mh2.replace(",", ".")

                                mn1 = $("#totmedarbn1_" + l).val().replace(".", "")
                                mn1 = mn1.replace(",", ".")
                                mn2 = $("#totmedarbn2_" + l).val().replace(".", "")
                                mn2 = mn2.replace(",", ".")
                                //alert()
                

                                if (mh1 == "NaN" || (mh1 == "-Infinity") || mh1 == "") {
                                    mh1 = 0
                                }

                                if (mh2 == "NaN" || (mh2 == "-Infinity") || mh2 == "") {
                                    mh2 = 0
                                }

                                if (mn1 == "NaN" || (mn1 == "-Infinity") || mn1 == "") {
                                    mn1 = 0
                                }

                                if (mn2 == "NaN" || (mn2 == "-Infinity") || mn2 == "") {
                                    mn2 = 0
                                }


                                if ((mh1/1) > mn1) {
                                    $("#totmedarbh1_" + l).css("background-color", "lightpink")
                                } else {

                                    if (mh1 = "0,00") {
                                        $("#totmedarbh1_" + l).css("background-color", "")
                                    } else {
                                        $("#totmedarbh1_" + l).css("background-color", "greenyellow")
                                    }
                                }

                                /*
                                if ((mh2/1) > mn2) {
                                    $("#totmedarbh2_" + l).css("background-color", "lightpink")
                                } else {
                                    if (mh1 = "0,00") {
                                        $("#totmedarbh2_" + l).css("background-color", "")
                                    } else {
                                        $("#totmedarbh2_" + l).css("background-color", "greenyellow")
                                    }
                                }
                                */



                            }


            } //loadberegn
           
            

        
        
        } /// All H1H2 totaler func





      



        function budgettimerColor(h1, h2, budgettimer, jobidaktid){

         

            if ((h1 / 1 + h2 / 1) > budgettimer && budgettimer != 0) {
                $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "lightpink")
            } else {

                if (budgettimer == 0) {
                    $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "")
                } else {
                    $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "")
                }

            }


        }


        //////////////////  Ved LOAD /////////////////////////////////////
        
        



        
        realjobaktGT = 0;

        function realjobaktGT_fn() {
        $(".realjobaktgt").each(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(15, idlngt)

            //alert("HH ; " + idtrim)

            var arr = idtrim.split('_');

            jobid = arr[0]
            aktid = arr[1]

            //// REAL 
            realjobakt = $(this).val().replace(".", "")
            realjobakt = realjobakt.replace(",", ".")

            if (realjobakt == "NaN" || (realjobakt == "-Infinity") || realjobakt == "") {
                realjobakt = 0
            }

            realjobaktGT = (realjobaktGT / 1 + realjobakt / 1)
            realjobakt = String(realjobakt).replace(".", ",")


            if (realjobakt != 0) {
                $("#realjobaktgt_" + arr[0] + "_" + arr[1]).val(realjobakt)
            } else {
                $("#realjobaktgt_" + arr[0] + "_" + arr[1]).val('')
            }

           
        });


        }
       
        //realjobaktGT_fn();




        ///////// SUMMER FORECAST TOTAL PÅ JOB PR ProJEKTGRUPPE VANDRET ////
        
        visallejob = $("#visallejob").val();
        viskunminit = $("#viskunminit").val();
        if (visallejob != 1 && viskunminit != 1) {
            //beregnFCprProgrpTilTotPaJob(0, 0);
        }
        


        function beregnFCprProgrpTilTotPaJob(jobid,aktid) {

            //alert("beregnFCprProgrpTilTotPaJob")

            $("#beregnTotFC" + jobid + "_0").val('0')
            $("#fcjobaktgt_" + jobid + "_0").val('0')
            $("#fcjobaktgt_" + jobid + "_" + aktid).val('0')
            //$("#fcjobaktBelgt_" + jobid + "_" + aktid).html('0')
            
            //lastaktid = 0;
            fcjobakt = 0;
            fcjobaktgt = 0;
            fcjobaktgtAkk = 0;
           
            //$("#fcjobaktgt_" + jobid + "_" + aktid).val('0')

            $(".afd_jobakt").each(function () {

                

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(11, idlngt)



                var arr = idtrim.split('_');

                jobid = arr[0]
                aktid = arr[1]
                progrpid = arr[2]


               
                    fcjobakt = $("#afd_jobakt_" + jobid + "_" + aktid + "_" + progrpid).val()
                    //fcjobaktbel = $("#afd_jobaktbel_" + jobid + "_" + aktid + "_" + progrpid).val()


                    if (fcjobakt == "NaN" || (fcjobakt == "-Infinity") || fcjobakt == "") {
                        fcjobakt = 0
                    } else {
                        fcjobakt = fcjobakt.replace(".", "")
                        fcjobakt = fcjobakt.replace(",", ".")
                    }


                    fcjobaktgt = $("#fcjobaktgt_" + jobid + "_" + aktid).val()
             

                if (fcjobaktgt == "NaN" || (fcjobaktgt == "-Infinity") || fcjobaktgt == "") {
                    fcjobaktgt = 0
                } else {
                    fcjobaktgt = fcjobaktgt.replace(".", "")
                    fcjobaktgt = fcjobaktgt.replace(",", ".")
                }


                
                fcjobaktgtAkk = (fcjobaktgt / 1 + fcjobakt / 1)

              
                if (fcjobaktgtAkk != 0) {

                    
                    fcjobaktgtAkkVal = fcjobaktgtAkk
                    String(fcjobaktgtAkk).replace(".", ",")
                    $("#fcjobaktgt_" + jobid + "_" + aktid).val(fcjobaktgtAkk)


                  

                }


               

            });


        }

       
       
        // afDGt FØR 
        // Total på projektgruppe LODRET (summer aktiviteter)
        function afDGt(progrpid) {
            fcAfdGT = 0;

            //alert("afDGt")
           
            lastjobid = 0

            $(".afd_jobakt_" + progrpid).each(function () {

              
                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(11, idlngt)

                var arr = idtrim.split('_');

                jobid = arr[0]
                aktid = arr[1]


                if (lastjobid != jobid && lastjobid > 0) {

                    fcAfdGT = String(fcAfdGT).replace(".", ",")
                    if (fcAfdGT != 0) {
                        $("#afd_jobakt_" + lastjobid + "_0_" + progrpid).val(fcAfdGT)
                    } else {
                        $("#afd_jobakt_" + lastjobid + "_0_" + progrpid).val('')
                    }

                    fcAfdGT = 0;

                }


                

                //alert(jobid +"#"+ aktid)

                if (aktid != 0) {

                    fcjobakt = 0;
                    //// fc 
                    fcjobakt = $("#afd_jobakt_" + jobid + "_" + aktid + "_" + progrpid).val().replace(".", "")
                    fcjobakt = fcjobakt.replace(",", ".")

                    //alert(fcjobakt)
                    //alert("#fcjobaktgt_" + $("#fcjobaktgt_" + idtrim).val() + " id: " + idtrim + "  // r:" + fcjobakt + "// gt:" + fcjobaktGT)

                    if (fcjobakt == "NaN" || (fcjobakt == "-Infinity") || fcjobakt == "") {
                        fcjobakt = 0
                    }

                    fcAfdGT = (fcAfdGT / 1 + fcjobakt / 1)
                }


                lastjobid = jobid

            });


            fcAfdGT = String(fcAfdGT).replace(".", ",")
            if (fcAfdGT != 0) {
                $("#afd_jobakt_" + jobid + "_0_" + progrpid).val(fcAfdGT)
                $("#afd_jobaktbel_" + jobid + "_0_" + progrpid).val('300')
            } else {
                $("#afd_jobakt_" + jobid + "_0_" + progrpid).val('')
                $("#afd_jobaktbel_" + jobid + "_0_" + progrpid).val('')
            }
           
        }




        /////////////////////////////////////////////////////////////////////////////////////
        // DENNE BRUGES IKKE PT
        // Forecast TOTAL PR job LODRET summer aktiviteter til JOB //
        fcjobaktGT = 0;
        function jobfcGt() {

            //alert("jobfcGt")

            $(".fcjobgt").each(function () {

                fcjobaktGT = 0;

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(13, idlngt)

                var arr = idtrim.split('_');

                jobid = arr[0]




                $(".fcjobgt_" + jobid).each(function () {

                    var thisid = this.id
                    var idlngt = thisid.length
                    var idtrim = thisid.slice(13, idlngt)

                    var arr2 = idtrim.split('_');

                    //jobid = arr[0]
                    aktid = arr2[1]

                    //alert("aktid: " + aktid)

                    if (aktid != 0) {

                        fcjobakt = 0;
                        //// fc 
                        fcjobakt = $("#fcjobaktgt_" + idtrim).val().replace(".", "")
                        fcjobakt = fcjobakt.replace(",", ".")

                        //alert("#fcjobaktgt_" + $("#fcjobaktgt_" + idtrim).val() + " id: " + idtrim + "  // r:" + fcjobakt + "// gt:" + fcjobaktGT)

                        if (fcjobakt == "NaN" || (fcjobakt == "-Infinity") || fcjobakt == "") {
                            fcjobakt = 0
                        }

                        fcjobaktGT = (fcjobaktGT / 1 + fcjobakt / 1)
                    }


                });//EaCH

                fcjobaktGT = String(fcjobaktGT).replace(".", ",")
                if (fcjobaktGT != 0) {
                    $("#fcjobaktgt_" + arr[0] + "_0").val(fcjobaktGT)
                } else {
                    $("#fcjobaktgt_" + arr[0] + "_0").val('')
                }

                //$("#fcjobaktBelgt_" + arr[0] + "_0").html('200 DKK')
                //$("#saldoBelgt_" + arr[0] + "_0").html('-900 DKK')

            });//EaCH

        }



        // REAL //
        realjobaktGT = 0;
        function jobrealGt() {

            //alert("jobrealGt")

            $(".realjobgt").each(function () {

                realjobaktGT = 0;

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(13, idlngt)

                var arr = idtrim.split('_');

                jobid = arr[0]




                $(".realjobgt_" + jobid).each(function () {

                    var thisid = this.id
                    var idlngt = thisid.length
                    var idtrim = thisid.slice(13, idlngt)

                    var arr2 = idtrim.split('_');

                    //jobid = arr[0]
                    aktid = arr2[1]

                    //alert("aktid: " + aktid)

                    if (aktid != 0) {

                        realjobakt = 0;
                        //// REAL 
                        realjobakt = $("#realjobaktgt_" + idtrim).val().replace(".", "")
                        realjobakt = realjobakt.replace(",", ".")

                        //alert("#realjobaktgt_" + $("#realjobaktgt_" + idtrim).val() + " id: " + idtrim + "  // r:" + realjobakt + "// gt:" + realjobaktGT)

                        if (realjobakt == "NaN" || (realjobakt == "-Infinity") || realjobakt == "") {
                            realjobakt = 0
                        }

                        realjobaktGT = (realjobaktGT / 1 + realjobakt / 1)
                    }


                });//EaCH

                realjobaktGT = String(realjobaktGT).replace(".", ",")
                if (realjobaktGT != 0) {
                    $("#realjobaktgt_" + arr[0] + "_0").val(realjobaktGT)
                } else {
                    $("#realjobaktgt_" + arr[0] + "_0").val('')
                }

            });//EaCH

        }


        //jobrealGt();
         

        
       
      

    });
      
 




       
      