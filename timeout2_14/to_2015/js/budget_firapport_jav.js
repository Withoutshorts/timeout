






$(window).load(function () {

    //visSubtotJobmed();

});


$(document).ready(function () {

       

        //alert("her XX")

        //$(".tr_nyjobper").css("visibility", "hidden")
        //$(".tr_nyjobper").css("display", "none")

        //$(".tr_nyakt").css("visibility", "hidden")
        //$(".tr_nyakt").css("display", "none")

        fn = $("#fn").val()

        if (fn == "opr") {
            showPeriods();
        } else {
            $(".tr_nyjobper").css("visibility", "hidden")
            $(".tr_nyjobper").css("display", "none")
        }


        $('#example1').datepicker({
            format: "dd/mm/yyyy"
        });

        

        $(".fstxt").keyup(function () {

          
            
            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(3, idlngt)

            thisVal = $("#fs_" + idtrim).val();

            //alert(thisVal)

            $(".afs_" + idtrim).val(thisVal);
           
        });


        $(".fsacc").change(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(7, idlngt)

            thisVal = $("#fs_acc_" + idtrim).val();

            //alert(thisVal + " fs: " + idtrim)

            $(".afs_acc_" + idtrim).val(thisVal);

        });


        
        
   

        $('#FM_stdato, #datepicker_stdato').datepicker({
           
        })

        $('#FM_sldato, #datepicker_sldato').datepicker({
          
        })


        $("#FM_stdato, #FM_sldato").keyup(function () {
            
           
            var stDate = $("#FM_stdato").val()
            //$("#FM_stdato").val()
            stDay = stDate.slice(0, 2)
            stMth = stDate.slice(3, 5)
            stYear = stDate.slice(6, 10)

            //alert(stDay + "." + stMth + "." + stYear)

            var slDate = $("#FM_sldato").val()
            //$("#FM_stdato").val()
            slDay = slDate.slice(0, 2)
            slMth = slDate.slice(3, 5)
            slYear = slDate.slice(6, 10)

        
           monthDiff(
           new Date(stYear, stMth, stDay), // November 4th, 2008
           new Date(slYear, slMth, slDay)  // March 12th, 2010
            );


           fn = $("#fn").val()

           if (fn == "opr") {
               showPeriods();
           }

          

        });



        function monthDiff(d1, d2) {
            var months;
            months = (d2.getFullYear() - d1.getFullYear()) * 12;
            months -= d1.getMonth() + 1;
            months += d2.getMonth();

            $("#months_diff").val(months)
            //return months <= 0 ? 0 : months;

            
        }

       



        $("#FM_valuta_0").change(function () {

           

            thisVal = $("#FM_valuta_0").val()

           

            $.post("?valutaid=" + thisVal, { control: "FN_valutakurs", AjaxUpdateField: "true" }, function (data) {

                
                //$("#div_timereg_0").html(data);
                $("#FM_valuta_0_kurs").val(data);


                fn = $("#fn").val()

                if (fn == "opr") {
                    kurs = $("#FM_valuta_0_kurs").val()

                    $(".percum_kurs").val(kurs)
                }

            });


        });



        $("#sbm_upd0").click(function () {

            $("#rdir").val('0');

        });

        $("#sbm_upd1").click(function () {

            $("#rdir").val('1');

        });

        $("#sbm_upd2").click(function () {

            $("#rdir").val('1');

        });
       



        $("#sp_add_period").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });

        $("#sp_add_activity").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });


        $(".rdbtn_per").click(function () {

            //alert(this.id)

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(3, idlngt)

           if ($(this.id).attr('checked')) {
               $(".rw_" + idtrim).prop("disabled", true);
            } else {
                $(".rw_" + idtrim).prop("disabled", false);
            }
            
            $(".aktcurrent").val('-9999')

            $("#budget").trigger("submit")

        });


        function showPeriods(){


            budget_view = $("#FM_budget_view").val()
            months_diff = $("#months_diff").val()


            if (months_diff <= 0) {
                loops = 0
            }

            if (months_diff > 0 && months_diff <= 12) {
                loops = 1
            }

            if (months_diff > 12 && months_diff <= 24) {
                loops = 2
            }


            if (months_diff > 24 && months_diff <= 36) {
                loops = 3
            }

            if (months_diff > 36 && months_diff <= 48) {
                loops = 4
            }

            if (months_diff > 48) {
                loops = 1
            }

            antalrows = (loops * budget_view)

            //alert(antalrows)

            $(".tr_nyjobper").css("visibility", "hidden")
            $(".tr_nyjobper").css("display", "none")

            fn = $("#fn").val()
           
            if (fn == "opr") {
                kurs = $("#FM_valuta_0_kurs").val()
               
                $(".percum_kurs").val(kurs)
            }
            

            for (rw = 0; rw < antalrows; rw++) {

                            
                if ($(".tr_nyjobper_" + rw).css('display') == "none") {

                    $(".tr_nyjobper_" + rw).css("visibility", "visible")
                    $(".tr_nyjobper_" + rw).css("display", "")


                } else {

                    $(".tr_nyjobper_" + rw).css("visibility", "hidden")
                    $(".tr_nyjobper_" + rw).css("display", "none")

                

                }

            }

        }

        
      
        
        $("#FM_valuta_0_kurs").keyup(function () {
            fn = $("#fn").val()

            if (fn == "opr") {
                kurs = $("#FM_valuta_0_kurs").val()

                $(".percum_kurs").val(kurs)
            }
        });

        
        $("#FM_budget_view").change(function () {

            fn = $("#fn").val()

            if (fn == "opr") {
                showPeriods();
            }


        });

       

        

        $("#sp_add_period").click(function () {


            if ($(".tr_nyjobper_0").css('display') == "none") {

                $(".tr_nyjobper_0").css("visibility", "visible")
                $(".tr_nyjobper_0").css("display", "")

                $("#sp_add_period_0").css("color", "#CCCCCC")

            } else {

                $(".tr_nyjobper_0").css("visibility", "hidden")
                $(".tr_nyjobper_0").css("display", "none")

                $("#sp_add_period").css("color", "#999999")


            }
            

            
            //budgetid = $("#budgetid").val()

           
            //$.post("?budgetid=" + budgetid, { control: "FN_tilfojper", AjaxUpdateField: "true" }, function (data) {
                //$("#fajl").val(data);


                //alert("..")
                //$("#div_jobid").html(data);

            //});
        });



        $("#sp_add_activity").click(function () {


          

            if ($(".tr_nyakt").css('display') == "none") {

                $(".tr_nyakt").css("visibility", "visible")
                $(".tr_nyakt").css("display", "")

                $("#sp_add_activity").css("color", "#CCCCCC")

            } else {

                $(".tr_nyakt").css("visibility", "hidden")
                $(".tr_nyakt").css("display", "none")

                $("#sp_add_activity").css("color", "#999999")


            }



        });
     
        
   

    });
      
 




       
      