<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>TimeOut - Tid, Overblik & Fakturering</title>
	<!--<LINK REL="SHORTCUT ICON" HREF="https://outzource.dk/favicon.ico">-->
	

	
	<%
	
	
	
	
	select case lto
		case "kringit"%>
		<LINK rel="stylesheet" type="text/css" href="inc/style/timeout_style_print_fak.css">
		<%
		header = "<img src='ill/kringit/kring_value.gif' alt='' border='0'>"
		headerBg = "#C5161C"
		%>
		<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=headerBg%>" width="100%">
        <tr>
        <td><%=header%></td>
        </tr>
        </table>
		<%
		case else%>
		<LINK rel="stylesheet" type="text/css" href="inc/style/timeout_style_print.css">
		<%
		header = "<img src='ill/ur.gif' width='152' height='53' alt='' border='0'>"
		headerBg = "#003399"
		end select
		
	%>
		



         <link href="inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>
	<script src="inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
    <script src="inc/jquery/jquery.timer.js" type="text/javascript"></script>
    <script src="inc/jquery/jquery.corner.js" type="text/javascript"></script>




		<script language="javascript" type="text/javascript">
		

        	function onfocus(){

       

        $(document).ready(function() {

                //if ($("#lto").val() == 'intranet - local' || $("#lto").val() == 'demo') {
                  
                //  $("#FM_ref").fucus()
                 // } else {
                

                //}

                   

                   $("#FM_email").keyup(function () {
              

                   $("#logindiv").css("display", "");
                    $("#logindiv").css("visibility", "visible");
                  $("#logindiv").show(fast);

                 });

        });



        

      

             if (document.getElementById("lto").value == 'intranet - local' || document.getElementById("lto").value == 'demo') {

                 document.getElementById("FM_ref").focus()

             } else {
           <%
           if len(Request.Cookies("login")("usrval")) <> 0 then%>
           document.getElementById("loginsubmit").focus()
           <%else%>
            document.getElementById("login").focus()
           <%end if%>
           }
        }


	
        </script>
	
</head>


<%
if Request("attempt") <> 1 then
%>
<body topmargin="0" leftmargin="0" onload="onfocus();" class="login">
<%
else
%>
<body topmargin="0" leftmargin="0" class="login">
<%
end if
%>

