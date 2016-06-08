<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html lang="da">
<head>
	<title>TimeOut - Tid, Overblik & Fakturering</title>
	<!--<LINK REL="SHORTCUT ICON" HREF="https://outzource.dk/favicon.ico">-->	
	

           <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
    <!--<link href="../inc/menu/css/chronograph_01.less" rel="stylesheet/less" type="text/css" />-->
     <link href="../inc/menu/css/chronograph_01.css" rel="stylesheet" type="text/css" />
     

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

   
 



	<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css" />

    <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>

	<script src="../inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.timer.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.corner.js" type="text/javascript"></script>

    

 
    

    <!--  Added by Lei 17-06-2013-->
    <!-- <link rel="stylesheet" href="../inc/jquery/jquery-ui-1.10.3.custom/css/ui-lightness/jquery-ui-1.10.3.custom.min.css" /> -->     
    <script src="../inc/jquery/jquery-1.10.1.min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery-ui-1.10.3.custom/js/jquery-ui-1.10.3.custom.min.js" type="text/javascript"></script>



    <!-- MENU og CSS Martin -->
     <script src="../inc/js/less.js" type="text/javascript"></script>
    <script src="../inc/menu/js/modernizr.js" type="text/javascript"></script>
       <script src="../inc/menu/js/classie.js" type="text/javascript"></script>
    <script src="../inc/menu/js/menu.js" type="text/javascript"></script>




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

   

    
           
    </script>
  


	
</head>






<body topmargin="0" leftmargin="0" class="regular" id="mainbody">



<!--=======================================================
TimeOut er tænkt og produceret af OutZourCE 2002 - 2014
www.OutZourCE.dk 
Tlf: +45 2684 2000
timeout@outzource.dk
========================================================-->

