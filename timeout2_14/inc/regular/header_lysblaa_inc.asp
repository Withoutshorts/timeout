<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">-->
<html xmlns="http://www.w3.org/1999/xhtml" lang="da">
<head>
	<title>TimeOut - Tid, Overblik & Fakturering</title>
	<!--<LINK REL="SHORTCUT ICON" HREF="https://outzource.dk/favicon.ico">-->	
	

            
    <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css' />

     
   
    <link href="../inc/menu/css/chronograph.css" rel="stylesheet" type="text/css" />
   
    
    <!--
    <LINK rel="stylesheet" type="text/css" href="../inc/style/chronograph.css" />
    -->    
        
   <!--

    <script>
  less = {
    env: "development",
    async: false,
    fileAsync: false,
    poll: 1000,
    functions: {},
    dumpLineNumbers: "comments",
    relativeUrls: false,
    rootpath: ""
  };
</script>
       -->
       

       
       
    

     <style type="text/css">

  h3.menuh3 {
  font-size: 16px;
  font-weight:200;
 
} 

</style>
       
   
 



	<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css" />

    <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>

	<script src="../inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.timer.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.corner.js" type="text/javascript"></script>

    

 
    

  



     <!-- MENU og CSS Martin -->

     <!--
     <script src="../inc/js/less.js" type="text/javascript"></script>
     -->
 
    <script src="../inc/menu/js/modernizr.js" type="text/javascript"></script>
     <script src="../inc/menu/js/classie.js" type="text/javascript"></script>
     <script src="../inc/menu/js/menu.js" type="text/javascript"></script>
        



      <!--  Added by Lei 17-06-2013-->
    <!-- <link rel="stylesheet" href="../inc/jquery/jquery-ui-1.10.3.custom/css/ui-lightness/jquery-ui-1.10.3.custom.min.css" /> -->     
    <script src="../inc/jquery/jquery-1.10.1.min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery-ui-1.10.3.custom/js/jquery-ui-1.10.3.custom.min.js" type="text/javascript"></script>

    <style>
        #sortable { list-style-type: none; margin: 0; padding: 0; width: 60%; }
        #sortable li { margin: 0 3px 3px 3px; padding: 0.4em; padding-left: 1.5em; font-size: 1.4em; height: 18px; }
        #sortable li span { position: absolute; margin-left: -1.3em; }
    </style>



    <%'if thisfile <> "erp_opr_faktura" AND thisfile <> "erp_fak.asp" AND thisfile <> "erp_opr_faktura_kontojob" then   
    '** I ERP delen (FRAMESET) voirker Jquery IKKE når denne del er tilføjet. ***'%>
    <script type="text/javascript">

        /*
        Restore globally scoped jQuery variables to the first version loaded
        (the newer version)
        */
        jq1101 = jQuery.noConflict(true);

    </script>
    <script src="../inc/jquery/dragSort.jquery.js" type="text/javascript"></script>
     <%'end if %>
    
    <!--End of added by Lei 17-06-2013-->

   



    <script type="text/javascript">



     
      
        $(window).load(function () {


          

            /*

            
            //scwdt = document.body.offsetWidth
            //window.innerWidth
            //alert(scwdt)


            var winW = 1116 //winH = 460;
            if (document.body && document.body.offsetWidth) {
                winW = document.body.offsetWidth;
                winH = document.body.offsetHeight;
            }
            if (document.compatMode == 'CSS1Compat' &&
                document.documentElement &&
                document.documentElement.offsetWidth) {
                winW = document.documentElement.offsetWidth;
                //winH = document.documentElement.offsetHeight;
            }
            if (window.innerWidth && window.innerHeight) {
                winW = window.innerWidth;
                //winH = window.innerHeight;
            }

            scwdt = winW;
            //alert(scwdt)

            $("#mainbody").css("width", scwdt - 40);
            $("#maintable").css("width", scwdt - 40);
            $("#menutable").css("width", scwdt - 40);

            */


        });


  


        $(document).ready(function () {

            if ($("#FM_visdeakmed").is(':checked') == false) {
                $("#FM_visdeakmed12").removeAttr("checked");
                $("#FM_visdeakmed12").attr("disabled", "disabled");
            }

           

            antalM = $('#FM_medarb option').size();
            antalM = antalM - 1
            $('#antalmedarblist').html(antalM)

           // winH = window.innerHeight;
            //alert(winH)

            //$("#menu - slider").css("height", winH);


            $(".closeqmnote").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });

            $(".closeqmnote").click(function () {
                $(".qmarkhelptxt").hide();
            });
            
            
          

            //Dato func jq //
            function IsValidDate(input) {
                var dayfield = input.split(".")[0];
                var monthfield = input.split(".")[1];
                var yearfield = input.split(".")[2];
                var returnval;
                var dayobj = new Date(yearfield, monthfield - 1, dayfield)
                if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
                    alert("Forkert datoformat, skriv venligst dd.mm.yyyy eller d.m.yyyy");
                    returnval = false;
                }
                else {
                    returnval = true;
                }
                return returnval;
            }


            //Make datepickers
            $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));

            $("#FM_stdato").datepicker();
            $("#FM_stdato").change(function () { if (IsValidDate($("#FM_stdato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_start_dag]").val(datesplit[0]); $("[name=FM_start_mrd]").val(datesplit[1]); $("[name=FM_start_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
            $("#FM_stdato").css("display", "inline");

            $("#FM_sldato").datepicker();
            $("#FM_sldato").change(function () { if (IsValidDate($("#FM_sldato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_slut_dag]").val(datesplit[0]); $("[name=FM_slut_mrd]").val(datesplit[1]); $("[name=FM_slut_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
            $("#FM_sldato").css("display", "inline");

       



            
            $("#sp_med").click(function () {
                
                if ($("#tr_prog_med").css('display') == "none") {
                    $("#tr_prog_med").css("visibility", "visible")
                    $("#tr_prog_med").css("display", "")
                    $.cookie('tr_medarb', '1');
                } else {
                    $("#tr_prog_med").hide('fast')
                    $.cookie('tr_medarb', '0');
                }
            });


            $("#sp_med").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });



            $("#sp_kun").click(function () {

                if ($("#tr_kun").css('display') == "none") {
                    $("#tr_kun").css("visibility", "visible")
                    $("#tr_kun").css("display", "")
                    $.cookie('tr_kun', '1');
                } else {
                    $("#tr_kun").hide('fast')
                    $.cookie('tr_kun', '0');
                }
            });

            $("#sp_kun").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });


            $("#sp_job").click(function () {

                if ($("#tr_job").css('display') == "none") {
                    $("#tr_job").css("visibility", "visible")
                    $("#tr_job").css("display", "")
                    $.cookie('tr_job', '1');
                } else {
                    $("#tr_job").hide('fast')
                    $.cookie('tr_job', '0');
                }
            });

            $("#sp_job").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });


            $(".qmarkhelp").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });
            
            $(".qmarkhelp").click(function () {
                
                var thisid = this.id

                if ($("#qmarkhelptxt_" + thisid).css('display') == "none") {
                    $("#qmarkhelptxt_" + thisid).css("visibility", "visible")
                    $("#qmarkhelptxt_" + thisid).css("display", "")
                } else {

                    $("#qmarkhelptxt_" + thisid).hide(200)

                }

            });

           

           
           

            $("#sp_ava").click(function () {

                if ($("#tr_ava").css('display') == "none") {
                    $("#tr_ava").css("visibility", "visible")
                    $("#tr_ava").css("display", "")

                    $.cookie('tr_ava', '1');
                } else {
                    $("#tr_ava").hide('fast')
                    $.cookie('tr_ava', '0');
                }
            });

            $("#sp_ava").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });



            $("#sp_pre").click(function () {

                if ($("#tr_pre").css('display') == "none") {
                    $("#tr_pre").css("visibility", "visible")
                    $("#tr_pre").css("display", "")

                    $.cookie('tr_pre', '1');
                } else {
                    $("#tr_pre").hide('fast')
                    $.cookie('tr_pre', '0');
                }
            });

            $("#sp_pre").mouseover(function () {
                $(this).css('cursor', 'pointer');
            });



            //alert($.cookie('tr_medarb'))
            if ($.cookie('tr_medarb') == '1') {
                $("#tr_prog_med").css("visibility", "visible")
                $("#tr_prog_med").css("display", "")
            }

            if ($.cookie('tr_kun') == '1') {
                $("#tr_kun").css("visibility", "visible")
                $("#tr_kun").css("display", "")
            }

            if ($.cookie('tr_job') == '1') {
                $("#tr_job").css("visibility", "visible")
                $("#tr_job").css("display", "")
            }

            if ($.cookie('tr_ava') == '1') {
                $("#tr_ava").css("visibility", "visible")
                $("#tr_ava").css("display", "")
            }

            if ($.cookie('tr_pre') == '1') {
                $("#tr_pre").css("visibility", "visible")
                $("#tr_pre").css("display", "")
            }



            $("#FM_progrp").change(function () {


                vismedarb();

            });

            $("#FM_vis_medarbejdertyper, #FM_vis_medarbejdertyper_grp").click(function () {



                vismedarb();

            });


            $("#FM_visdeakmed").click(function () {

               

                if ($("#FM_visdeakmed").is(':checked') == false) {
                    $("#FM_visdeakmed12").removeAttr("checked");
                    $("#FM_visdeakmed12").attr("disabled", "disabled");
                  
                } else {

                    $("#FM_visdeakmed12").attr("disabled", "");
                }


                vismedarb();

            });


            $("#FM_visdeakmed12").click(function () {


                vismedarb();

            });

            $("#FM_vispasmed").click(function () {


                vismedarb();

            });


            $("#FM_medarb").click(function () {


                if ($("#FM_medarb option:selected").val() == 0) {

                    $("#FM_medarb").each(function () {
                        $("#FM_medarb option").attr('selected', 'selected');

                    });


                    $("#FM_medarb option[value='0']").removeAttr("selected");

                }

            });




            function vismedarb() {

                $('#FM_medarb').children().remove();

                jq_medarb = $("#jq_userid").val();
               

                jq_progrp = $("#FM_progrp").val();

                //alert($("#FM_visdeakmed").is(':checked'))


                if ($("#FM_visdeakmed").is(':checked') == true) {
                    jq_mansat = 1
                } else {
                    jq_mansat = 0
                }


                if ($("#FM_visdeakmed12").is(':checked') == true) {
                    jq_mansat12 = 1
                } else {
                    jq_mansat12 = 0
                }


                if ($("#FM_vispasmed").is(':checked') == true) {
                    jq_mansatpas = 1
                } else {
                    jq_mansatpas = 0
                }


                if ($("#FM_vis_medarbejdertyper").is(':checked') == true) {
                    jq_mtyper = 1
                } else {
                    jq_mtyper = 0
                }

                if ($("#FM_vis_medarbejdertyper_grp").is(':checked') == true) {
                    jq_mtypergrp = 1
                } else {
                    jq_mtypergrp = 0
                }
                
                

                $.post("?jq_medarb=" + jq_medarb + "&jq_progrp=" + jq_progrp + "&jq_mansat=" + jq_mansat + "&jq_mansatpas=" + jq_mansatpas + "&jq_mansat12=" + jq_mansat12 + "&jq_mtyper=" + jq_mtyper + "&jq_mtypergrp=" + jq_mtypergrp, { control: "FN_medipgrp", AjaxUpdateField: "true" }, function (data) {


                    $("#FM_medarb").html(data);

                    

                    antalM = $('#FM_medarb option').size();
                    antalM = antalM - 1
                    $('#antalmedarblist').html(antalM)

                    //$("#FM_medarb option[1]").attr('selected', 'selected');
                    $("#FM_medarb option[value='0']").removeAttr("selected");

                    //alert("her s")

                });



                

            }

        });

       

    </script>
  
	
	<script type="text/javascript">



        

	    //$(".tregfaneblad").corner();
	    //$(".jqcorner").corner();
	    //$("#filter").corner();
	    //$("#eksport").corner();
   




function showabout(){
document.all["about"].style.visibility = "visible";
}
function hideabout(){
document.all["about"].style.visibility = "hidden";
}
function NewWin(url)    {
window.open(url, 'Help', 'width=700,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_cal(url)    {
window.open(url, 'Calender', 'width=300,height=380,scrollbars=yes,toolbar=no');    
}
function NewWin_popupaktion(url)    {
window.open(url, 'Calender', 'width=410, height=650,scrollbars=yes,toolbar=no');    
}
function NewWin_help(url)    {
window.open(url, 'Help', 'width=650,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_large(url)    {
window.open(url, 'Help', 'width=980,height=600,scrollbars=yes,toolbar=no');    
}
function NewWin_todo(url)    {
window.open(url, 'Opgavelisten', 'width=650,height=650,scrollbars=yes,toolbar=no');    
}
//Alle ovenstående popup's ændres til nedenstående//
//Gør også left og top til variable//
function popUp(URL,width,height,left,top) {
		window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
	}

function NewWin_popupcal(url)    {
	window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
}


function bgcolthisMON(divid){
document.getElementById(""+ divid +"").style.backgroundColor = "#ffffff";
}

function bgcolthisON(divid){
document.getElementById(""+ divid +"").style.backgroundColor = "#5C75AA";
}


function bgcolthisONcol(divid, bgcol) {
    document.getElementById("" + divid + "").style.backgroundColor = bgcol;
}

function bgcolthisOFF(divid, bgcol){
document.getElementById(""+ divid +"").style.backgroundColor = bgcol;
}


</script>
	
</head>






<body topmargin="0" leftmargin="0" class="regular" id="mainbody" style="width:1116px;">



<!--=======================================================
TimeOut er tænkt og produceret af OutZourCE 2002 - 2014
www.OutZourCE.dk 
Tlf: +45 2684 2000
timeout@outzource.dk
========================================================-->

