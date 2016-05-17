<!DOCTYPE html>
<!--[if lt IE 7]>      <html class="no-js lt-ie9 lt-ie8 lt-ie7"> <![endif]-->
<!--[if IE 7]>         <html class="no-js lt-ie9 lt-ie8"> <![endif]-->
<!--[if IE 8]>         <html class="no-js lt-ie9"> <![endif]-->
<!--[if gt IE 8]><!--> <html lang="en" class="no-js"> <!--<![endif]-->
<head>
  <title>Outzource, tabeller: medarbejdertyper </title>

  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="">
  <meta name="author" content="">

  <!-- Bruges til LESS styling af TO menu -->
  <link href="css/chronograph.css" rel="stylesheet" type="text/css" />
  
  <!-- Google Font: Open Sans -->
  <link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Open+Sans:400,400italic,600,600italic,800,800italic">
  <link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Oswald:400,300,700">

  <!-- Font Awesome CSS -->
  <link rel="stylesheet" href="css/font-awesome.min.css">

  <!-- Bootstrap CSS -->
  <link rel="stylesheet" href="css/bootstrap.min.css">

  <!-- Plugin CSS -->
  <link rel="stylesheet" href="js/plugins/dataTables/dataTables.bootstrap.css">

  <!-- App CSS -->
  <link rel="stylesheet" href="css/mvpready-admin.css">
  <link rel="stylesheet" href="css/mvpready-flat.css">
  <!-- <link href="./css/custom.css" rel="stylesheet">-->

  <!-- Custom styles for TimeOut  -->
  <link href="css/mpvready-style-timeout.css" rel="stylesheet" />

  <!-- Favicon -->
  <link rel="shortcut icon" href="favicon.ico">

  <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
  <!--[if lt IE 9]>
  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
  <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
  <![endif]-->


<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"--> 







</head>

<body class=" ">



    <%call menu_2014() %>

<div id="wrapper">
  <div class="content">

   

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Budget</u>
        </h3>
        <!-- Opdater/Annuller knapper -->
        <div style="margin-top:-15px;margin-bottom:15px;">
          <button type="button" class="btn btn-secondary btn-sm pull-right"><b>Opdatér</b></button>
          <button type="button" class="btn btn-danger btn-sm pull-right" style="margin-right:10px;" /><b>Annullér</b></button>
          <div class="clearfix"></div>
        </div>
        <div class="portlet-body">




               <!-- Base data -->
              <section>
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseThree">
                                    Basedata
                                </a>
                            </h4>
                        </div> <!-- /.panel-heading -->
                     
                           <div id="collapseThree" class="panel-collapse collapse">
                            <div class="panel-body">
                              
                                 
                                        <div class="row">
                                            <div class="col-sm-2">
                                               Budget titel
                                            </div>
                                               
                                            <div class="col-sm-4">
                                                <input type="text" class="form-control text-right input-sm input-small">
						                       
                                            </div>
                                            <div class="col-sm-2">

                                                Total Budget
                                            </div>
                                              <div class="col-sm-4">
                                                <input type="text" class="form-control text-right input-sm input-small">
						                       
                                        
                                    </div><!--/.row -->

                                              <div class="row">
                                            <div class="col-sm-2">
                                               Contact
                                            </div>
                                               
                                            <div class="col-sm-4">
                                                <input type="text" class="form-control text-right input-sm input-small">
						                       
                                            </div>
                                            <div class="col-sm-2">

                                                Extra funding
                                            </div>
                                              <div class="col-sm-4">
                                                <input type="text" class="form-control text-right input-sm input-small">
						                       
                                        
                                    </div><!--/.row -->


                                   
                                       
                              
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->


            </section>






            <!-- NAVN / SORTERING / ID -->
            <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-2">
                            <h4 class="panel-title-well">Base data</h4>
                        </div>
                        <!-- Titel -->
                        <div class="col-lg-6">
                            <div class="row">
                                <div class="col-lg-2 pad-t5" style="padding-left:10px;">Budget titel:</div>
                                <div class="col-lg-4">   
                                    <input type="text" class="form-control input-small" value="" />
                                </div>
                                <div class="col-lg-2 pad-t5" style="padding-left:10px;">Total Budget:</div>
                                <div class="col-lg-4">   
                                    <input type="text" class="form-control input-small" value="" />
                                </div>

                            </div>
                        </div>


                    
                      
                    </div>
                </div>
            
            
            
            
            </section>

            <!-- NORMERET TID -->
            <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-2">
                            <h4 class="panel-title-well">Normeret tid&nbsp;<a data-toggle="modal" href="#styledModalNormTid">
                                <span class="fa fa-info-circle"></span></a></h4>
                        </div>

                        <!-- Tooltip Normeret tid -->
                        <div id="styledModalNormTid" class="modal modal-styled fade">
                            <div class="modal-dialog">
                                <div class="modal-content" style="border:none !important;padding:0;">
                                    <div class="modal-header">
                                        <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                        <h5 class="modal-title">Hjælp til normeret tid</h5>
                                    </div>
                                    <div class="modal-body">
                                        <p>Husk at tilpasse ferie optjent for medarbejdere af denne type, hvis der ændres medarbejdertype.</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <!-- /.Tooltip Normeret tid -->
                        <div class="col-lg-10 col-padL-25">   
                            <div class="row">                                
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Mandag</div>
                                    <input type="text" class="form-control text-right input-sm input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Tirsdag</div>
                                    <input type="text" class="form-control text-right input-sm input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Onsdag</div>
                                    <input type="text" class="form-control text-right input-sm input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Torsdag</div>
                                    <input type="text" class="form-control text-right input-sm input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Fredag</div>
                                    <input type="text" class="form-control text-right input-sm input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week">
                                    <div class="text-left pad-b5">Lørdag</div>
                                    <input type="text" class="form-control text-right input-sm input-red input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week col-pad-week-son">
                                    <div class="text-left pad-b5">Søndag</div>
                                    <input type="text" class="form-control text-right input-sm input-red input-small" />
                                </div>
                                <div class="col-sm-1 text-center col-pad-week" style="margin-right:0;margin-left:25px;">
                                    <div class="text-left pad-b5">Timetotal</div>
                                    <input type="text" disabled="disabled" class="form-control text-center input-sm input-small" placeholder="84,00" />
                                </div>
                                <div class="col-sm-1"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </section>
            <!-- Timepris generel -->
            <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-2">
                            <h4 class="panel-title-well">Generel timepris&nbsp;&nbsp;<a data-toggle="modal" href="#styledModalGenTimepr">
                                <span class="fa fa-info-circle"></span></a></h4>
                                <!-- Tooltip Generel Timepris -->
                                <div id="styledModalGenTimepr" class="modal modal-styled fade">
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title">Hjælp til Generel timepris</h5>
                                            </div>
                                            <div class="modal-body">
                                                Den generelle timepris er den timepris der automatisk bliver tildelt på et job ved joboprettelse.<br/><br/>De andre timepriser(ses ved at klikke på linket "<a href=".details" data-toggle="collapse" style="cursor: pointer;">Flere timepriser</a>"), kan vælges ved redigering af job, eller de kan forvælges på stam-aktiviteter.
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- /.Tooltip Generel Timepris -->
                        </div>
                        <div class="col-sm-1" style="margin-left:-5px;">
                            <input type="text" class="form-control input-small text-right" value="300" />
                        </div>
                        <div class="dropdown col-sm-1 pad-0 pad-l10">
                            <select class="form-control slc-small" style="width:76%;">
						        <option>DKK</option>
						        <option>EUR (7,4556)</option>
						        <option>SEK (0,8934)</option>
					        </select>
                        </div>
                        <div class="col-sm-6 pad-l15">
                            <!-- Åbn/luk Flere timepriser -->
                            <div style="margin-top:2px;" class="text-left">
                                <a href=".details" data-toggle="collapse" style="cursor: pointer;color:#333;">
                                    <div class="details collapse in">Se flere timepriser<i class="fa fa-caret-down" style="margin-top:-23px;margin-left:17px;font-size:22px;color:#d9534f;"></i></div>
                                    <div class="details collapse">Luk flere timepriser<i class="fa fa-caret-up" style="margin-top:-25px;font-size:22px;color:#d9534f;"></i></div>
                                </a>
                            </div>
                        </div>
                        <div class="col-sm-4"></div>
                    </div>
                </div>
                <div class="details collapse" style="margin-top:-7px;">
                    <div class="well well-white">
                        <!-- Timepris 2+3 -->
                        <div class="row" style="margin-left:170px;width:85%">
                            <label class="col-sm-1 text-right pad-0 pad-t5">Timepris 2:</label>
                            <div class="col-sm-1 pad-0 pad-l10 pad-r5"><input type="text" class="form-control input-small" /></div>
                            <div class="dropdown col-sm-1 pad-0">
                                <select class="form-control slc-small wdth-80pct">
						            <option>DKK</option>
						            <option>EUR (7,4556)</option>
						            <option>SEK (0,8934)</option>
					            </select>                                            
                            </div>
                            <label class="col-sm-1 text-right pad-0 pad-t5">Timepris 3:</label>
                            <div class="col-sm-1 pad-0 pad-l10 pad-r5"><input type="text" class="form-control input-small" /></div>
                            <div class="dropdown col-sm-1 pad-0">
                                <select class="form-control slc-small wdth-80pct">
						            <option>DKK</option>
						            <option>EUR (7,4556)</option>
						            <option>SEK (0,8934)</option>
					            </select>                                            
                            </div>
                            <label class="col-sm-1 text-right pad-0 pad-t5">Timepris 4:</label>
                            <div class="col-sm-1 pad-0 pad-l10 pad-r5"><input type="text" class="form-control input-small" /></div>
                            <div class="dropdown col-sm-1 pad-0">
                                <select class="form-control slc-small wdth-80pct">
						            <option>DKK</option>
						            <option>EUR (7,4556)</option>
						            <option>SEK (0,8934)</option>
					            </select>                                            
                            </div>

                            <label class="col-sm-1 text-right pad-0 pad-t5">Timepris 5:</label>
                            <div class="col-sm-1 pad-0 pad-l10 pad-r5"><input type="text" class="form-control input-small" /></div>
                            <div class="dropdown col-sm-1 pad-0">
                                <select class="form-control slc-small wdth-80pct">
						            <option>DKK</option>
						            <option>EUR (7,4556)</option>
						            <option>SEK (0,8934)</option>
					            </select>                                            
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- Intern kostpris -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">
                            Intern kostpris
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                          <div class="panel-body">
                            <div class="row">
                                <div class="col-sm-2"></div>
                                <div class="col-sm-1 marg-lft-2">
                                    <input type="text" class="form-control text-right input-small" />
                                </div>
                                <div class="col-sm-1" style="padding:8px 0 0 0">DKK</div>
                                <div class="col-sm-7" style="margin-top:-4px;margin-left:-50px;">
                                    <div class="checkbox">
                                        <label>
                                            <input type="checkbox" value="" />
                                            Opdater interne kostpriser, for denne type, på alle eksisterende <u>åbne</u> job.
                                        </label>
                                    </div>
                                </div>
                                <div class="col-sm-1"></div>
                            </div>
                          </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->
                    <!-- Grupper -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseThree">
                                    Grupper
                                </a>
                            </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseThree" class="panel-collapse collapse">
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-lg-2"></div>
                                    <div class="col-lg-5 marg-lft-2">
                                        <div class="row">
                                            <div class="col-sm-4">
                                                Søster-gruppe&nbsp;&nbsp;<a data-toggle="modal" href="#styledModalSstGrp"><span class="fa fa-info-circle"></span></a> 
                                            </div>
                                                <!-- Tooltip Søstergruppe -->
                                                <div id="styledModalSstGrp" class="modal modal-styled fade">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title">Hjælp til Søstergruppe</h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                Skal denne medarbejdertype medregnes under en anden medarbejder-type i budget og omsætning? Hovedgruppe følger evt. valgte søstergruppe.
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                                <!-- /.Tooltip Søstergruppe -->
                                            <div class="col-sm-6">
                                                <select class="form-control slc-small">
						                            <option>Nej</option>
						                            <option>Designer</option>
						                            <option>Marketing</option>
					                            </select>
                                            </div>
                                            <div class="col-sm-2"></div>
                                        </div>
                                    </div>
                                    <div class="col-lg-5">
                                        <div class="row">
                                            <div class="col-sm-4">
                                                Hovedgruppe&nbsp;&nbsp;<a data-toggle="modal" href="#styledModalHvdGrp"><span class="fa fa-info-circle"></span></a>
                                            </div>
                                                <!-- Tooltip Hovedgruppe -->
                                                <div id="styledModalHvdGrp" class="modal modal-styled fade">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title">Hjælp til Hovedgruppe</h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                Skal denne medarbejdertype medregnes under en anden medarbejder-type i budget og omsætning? Hovedgruppe følger evt. valgte søstergruppe.
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                                <!-- /.Tooltip Hovedgruppe -->
                                            <div class="col-sm-7">
                                                <select class="form-control slc-small">
                                                    <option>Medarbejdertype Hovedgruppe 1</option>
                                                    <option>//Følger søstergruppe//</option>
                                                </select>
                                            </div>
                                            <div class="col-sm-1"></div>
                                        </div>
                                    </div>
                                </div>
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                    
                    <!-- Afslut ugekriterie -->
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFour">
                                    Afslut ugekriterie
                                </a>
                            </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFour" class="panel-collapse collapse">
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-sm-2"></div>
                                    <div class="col-sm-5 marg-lft-2">Minimumskrav til timeantal - medarbejdertypen kan afslutte sin uge når:</div>
                                    <div class="col-sm-2" style="padding:0 10px 0 0;">
                                        <select class="form-control slc-small">
						                    <option>Ingen (kan altid afslutte)</option>
						                    <option>Realiserede timer</option>
                                            <option>Realiserede fakturérbare timer</option>
					                    </select>
                                    </div>
                                    <div class="col-sm-1" style="padding:5px 0 0 15px;">udgør min.:</div>
                                    <div class="col-sm-1" style="padding:0 5px 0 0;"><input type="text" class="form-control text-center input-small" style="width:50%;" /></div>
                                    <div class="col-sm-1" style="padding:5px 0 0 0;margin-left:-35px;">% af normtid.</div>
                                </div>
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->


                    <!-- Opdatér stamaktivitetsgruppe -->
                    <div class="panel panel-default" style="margin-bottom:-25px;">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFive">
                                    Opdatér stamaktivitetsgruppe
                                </a>
                            </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFive" class="panel-collapse collapse">
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-sm-2"></div>
                                    <div class="col-sm-4 marg-lft-2">
                                        <select multiple class="form-control slc-small">
                                            <option>Standard 2014</option>
                                            <option>Timeout</option>
                                            <option>TO implementering</option>
                                            <option>Udviklingsprojekt</option>
                                        </select>
                                    </div>
                                    <div class="col-sm-5" style="margin-top:-10px;">
                                        <div class="checkbox">
                                            <label style="width:90%;">
                                                <input type="checkbox" value="" />
                                                Ja, opdatér timepriserne for alle medarbejdere af denne type på alle aktiviteterne, i stamaktivitets-grupperne til venstre:
                                            </label>
                                        </div>
                                    </div>
                                    <div class="col-sm-1"></div>
                                </div>
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                </div> <!-- /.accordion -->
            </section> <!-- /.demo-section -->

        </div> <!-- /.portlet-body -->
      </div> <!-- /.portlet -->
    </div> <!-- /.container -->
  </div> <!-- .content -->
</div> <!-- /#wrapper -->






<!-- Bootstrap core JavaScript
================================================== -->
<!-- Core JS -->
<script src="./js/libs/jquery-1.10.2.min.js"></script>
<script src="./js/libs/bootstrap.min.js"></script>

<!--[if lt IE 9]>
<script src="./js/libs/excanvas.compiled.js"></script>
<![endif]-->

<!-- Plugin JS -->
<script src="./js/plugins/dataTables/jquery.dataTables.js"></script>
<script src="./js/plugins/dataTables/dataTables.bootstrap.js"></script>

<!-- App JS -->
<script src="./js/mvpready-core.js"></script>
<script src="./js/mvpready-admin.js"></script>

<!-- Plugin JS -->
<script src="./js/demos/table_demo.js"></script>




</body>
</html>