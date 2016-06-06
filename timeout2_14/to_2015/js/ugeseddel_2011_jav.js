





$(document).ready(function() {





    $(".status_all").click(function () {

        

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(11, thisvallngt)
        thisval = thisvaltrim

        //alert("Her:" + thisval)

        if ($("#status_all_" + thisval).is(':checked') == true) {

            $(".status_chk_" + thisval).attr("checked", "checked");

            } else {

            $(".status_chk_" + thisval).removeAttr("checked");
           
        }


    });

 








    $("#FM_visallemedarb").click(function () {

       
        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {
       
        $("#filterkri").submit();
    });

    

   

    $(".FM_job").keyup(function () {

       

        $(".dv_akt").hide();

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

     
        
        sogjobogkunde(thisval);

    });

    //$(".FM_job").blur()(function () {

    
    //    $(".dv_job").hide();
    //});

    $(".FM_akt").keyup(function () {

       

        $(".dv_job").hide();

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

     
        sogakt(thisval);

    });



    function sogakt(thisval) {

        //alert(thisval)

        jq_newfilterval = $("#FM_akt_" + thisval).val()
        jq_jobid = $("#FM_jobid_" + thisval).val()
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt_" + thisval).val()
        jq_pa = $("#FM_pa").val()

        //alert(jq_jobid + ":: medid: " + jq_medid + "#" + jq_newfilterval.length + " jq_aktid: " + jq_aktid + "## jq_pa" + jq_pa)

        if (jq_newfilterval.length > 0) {

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_akt_" + thisval).html(data);

            $("#dv_akt_" + thisval).css("display", "");
            $("#dv_akt_" + thisval).css("visibility", "visible");
            $("#dv_akt_" + thisval).show(100);



            $('.luk_aktsog').bind('mouseover', function () {

                $(this).css("cursor", "pointer");



            });

       

            $(".luk_aktsog").bind('click', function () {


                $(".dv_akt").hide();

            });


            $('.chbox_akt').bind('mouseover', function () {

                $(this).css("cursor", "pointer");

            });


            $(".chbox_akt").bind('click', function () {

                

                var thisaktid = this.id
                var thisvallngt = thisaktid.length
                var thisvaltrim = thisaktid.slice(10, thisvallngt)

                thisaktid_val = thisvaltrim
                thisAktnavn = $("#hiddn_akt_" + thisaktid_val).val()
                //thisAktidUse = $("#chbox_akt_" + thisaktid_val).val()
                thisAktidUse = thisaktid_val

                //salert("thisAktidUse: " + thisAktidUse + "thisaktid_val: " + thisaktid_val)

                $("#FM_akt_" + thisval).val(thisAktnavn)
                $("#FM_aktid_" + thisval).val(thisAktidUse)
                $(".dv_akt").hide();

            });


     

        });

        }

            

    }



  
    function sogjobogkunde(thisval) {

        

        jq_newfilterval = $("#FM_job_" + thisval).val()
        jq_medid = $("#FM_medid").val()
       
        //alert(jq_newfilterval.length)

        if (jq_newfilterval.length > 0) {

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_job_" + thisval).html(data);

            $("#dv_job_" + thisval).css("display", "");
            $("#dv_job_" + thisval).css("visibility", "visible");
            $("#dv_job_" + thisval).show(100);
            

            $('.luk_jobsog').bind('mouseover', function () {

                $(this).css("cursor", "pointer");

            });


            $('.chbox_job').bind('mouseover', function () {

                $(this).css("cursor", "pointer");

            });


            

                $(".luk_jobsog").bind('click', function () {

              
                    $(".dv_job").hide();

                });


                $(".luk_jobsog").bind('click', function () {


                    $(".dv_job").hide();

                });

                $(".chbox_job").bind('click', function () {


                   
                    var thisjobid = this.id
                    var thisvallngt = thisjobid.length
                    var thisvaltrim = thisjobid.slice(10, thisvallngt)
                    
                    thisJobid = thisvaltrim
                    thisJobnavn = $("#hiddn_job_" + thisJobid).val()

                    $("#FM_job_" + thisval).val(thisJobnavn)
                    $("#FM_jobid_" + thisval).val(thisJobid)

                    $(".dv_job").hide();

                });


        
                
        });

        }
}

});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});