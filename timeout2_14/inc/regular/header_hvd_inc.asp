<!DOCTYPE html>
<!-- PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" -->

<html>
<head>
	
	<title>TimeOut - Tid, Overblik & Fakturering</title>
	<!--<LINK REL="SHORTCUT ICON" HREF="https://outzource.dk/favicon.ico">-->
	
 


	
	    <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.timer.js" type="text/javascript"></script>

	<script src="../inc/jquery/highcharts/highcharts.js" type="text/javascript"></script>
    <script src="../inc/jquery/highcharts/modules/exporting.js" type="text/javascript"></script>
	
	
  

	<%
	
    
     media = request("media")
    


	select case thisfile
	case "erp_oprfak_fs" '"fak_godkendt.asp"
	
	    select case lto 
	    case "essens"
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_verdana12.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
        case "synergi1", "intranet - local" 
        %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_arial13.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
        case "fe" 
        %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_arial13.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
        case "epi", "epi_no", "epi_sta", "epi_ab"
         %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_colibri12.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
	    case "dencker"
	    %>
	     <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
           </head>
    <body topmargin="0" leftmargin="0" class="regular">
        <%
	    case "xxx"
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak_stor.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
	    case else
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
          </head>
    <body topmargin="0" leftmargin="0" class="regular">
	    <%
	    end select
	    
	case "job_print.asp"
	
	    select case lto 
        case "essens"
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/essens_style_verdana12_1.css">
        </head>

           
            <body topmargin="0" leftmargin="0" class="regular">

        <%
        case "userminds"
	    %>

	    <LINK rel="stylesheet" type="text/css" href="../inc/style/userminds_style_trebuchet_1.css">
        </head>
        <body topmargin="0" leftmargin="0" class="regular">

         <%
	    case "synergi1", "intranet - local"
	    %>
         <LINK rel="stylesheet" type="text/css" href="../inc/style/synergi1_style_trebuchet_10.css" media="ALL">
        </head>
          
        <body topmargin="0" leftmargin="0">
        <%
           
	    case "jttek"
	    %>
         <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_jttek.css" media="ALL">
        </head>
          
        <body topmargin="0" leftmargin="0">
        <%
        case ""
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_verdana12.css">
        </head>
        <body topmargin="0" leftmargin="0" class="regular">
	     <%
	    case "xx"
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_colibri12.css">
        </head>
        <body topmargin="0" leftmargin="0" class="regular">
	    <%
	    case else
	    %>
	    <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak_stor.css">
        </head>
        <body topmargin="0" leftmargin="0" class="regular">
	    <%
	    end select
	
	case else
	%>
	<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
    </head>
    <body topmargin="0" leftmargin="0" class="regular">
	<%
	
	end select
    
    %>
	
	

<!--=======================================================
TimeOut er tænkt og produceret af OutZourCE 2002 - 2011
www.OutZourCE.dk 
Tlf: +45 2684 2000
timeout@outzource.dk
========================================================-->

