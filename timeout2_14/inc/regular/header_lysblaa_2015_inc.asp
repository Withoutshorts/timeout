<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="da">
<head>
	<title>TimeOut - Tid, Overblik & Fakturering</title>

 
  <meta charset="windows-1252">


  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="">
  <meta name="author" content="">

    
	
    <!-- Bruges til LESS styling af TO menu -->
    <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css' />
    <link href="../inc/menu/css/chronograph.css" rel="stylesheet" type="text/css" />
   
   
     <!-- Google Font: Open Sans -->
      <link rel="stylesheet" href="//fonts.googleapis.com/css?family=Open+Sans:400,400italic,600,600italic,800,800italic" />
      <link rel="stylesheet" href="//fonts.googleapis.com/css?family=Oswald:400,300,700" />

      <!-- Font Awesome CSS -->
      <link rel="stylesheet" href="css/font-awesome.min.css" />

   
       <!-- Bootstrap CSS -->
      <link rel="stylesheet" href="css/bootstrap.min.css" />

    
      <!-- Plugin CSS -->
      <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.css"/>
      <link rel="stylesheet" href="css/bootstrap-datepicker3.css" />
      <link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/jasny-bootstrap/3.1.3/css/jasny-bootstrap.min.css" />
     
    

      <!-- App CSS -->
      <link rel="stylesheet" href="css/mvpready-admin.css" />
      <link rel="stylesheet" href="css/mvpready-flat.css" />

     

     <!-- Custom styles for TimeOut  -->
      <link href="css/mpvready-style-timeout.css" rel="stylesheet" />

    
  

         <style type="text/css">

          h3.menuh3 {
          font-size: 16px;
          font-weight:300;
          letter-spacing:0.6px;
          
         }
          
         
        </style>


        <!-- HTML5 shim for IE backwards compatibility -->
        <!--[if lt IE 9]>
        <script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
        <![endif]-->
    

      <!-- Favicon -->
      <link rel="shortcut icon" href="https://outzource.dk/favicon.ico">

  

    
        
    
    
    <!-- Bootstrap core JavaScript
    ================================================== -->
    <!-- Core JS -->
   <script src="./js/libs/jquery-1.10.2.min.js"></script>
   <script src="./js/libs/bootstrap.min.js"></script>
        

    <!-- Plugin JS -->
    <%'IE <= 9.0 tjek '*** ELLERS KAN ACCORDIONS og (i) info IKKE ÅBNES og lukkes **'
    if instr(request.servervariables("HTTP_USER_AGENT"), "MSIE 9.0") <> 0 OR instr(request.servervariables("HTTP_USER_AGENT"), "MSIE 8.0") <> 0 then
          
        if request("func") = "" then 'else = rediger / opdater %>
        <script type="text/javascript" src="//cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.js"></script>
        <%end if %>

    <%else

        if request("func") = "" then 'else = rediger / opdater %>
        <script type="text/javascript" src="//cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.js"></script>
        <%end if %>

    <%end if %>

    <!--<script type="text/javascript" src="./js/datatables.min.js"></script>-->
    <script type="text/javascript" src="//cdn.datatables.net/plug-ins/1.10.9/sorting/date-dd-MMM-yyyy.js"></script>
    <script src="./js/plugins/datepicker/bootstrap-datepicker.js"></script>
    <script src="//cdnjs.cloudflare.com/ajax/libs/jasny-bootstrap/3.1.3/js/jasny-bootstrap.min.js"></script>
     
    
   
    <!-- App JS -->
    <script src="./js/mvpready-core.js"></script>
    <script src="./js/mvpready-admin.js"></script>


   

     <!-- MENU og CSS Martin SKAL KALDES EFTER JQ -->
     <script src="../inc/menu/js/modernizr.js" type="text/javascript"></script>
     <script src="../inc/menu/js/classie.js" type="text/javascript"></script>
     <script src="../inc/menu/js/menu.js" type="text/javascript"></script>

    
        







  
	
</head>

<body class=" ">



<!--=======================================================
TimeOut er tænkt og produceret af OutZourCE 2002 - 2015
www.OutZourCE.dk 
Tlf: +45 2684 2000
timeout@outzource.dk
========================================================-->

