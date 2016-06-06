<%
thisfile = "printversion.asp"
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<script type="text/javascript">
    function changeZoom2(oSel) {
    //alert("her")
  	newZoom= oSel.options[oSel.selectedIndex].innerText
	printarea.style.zoom=newZoom+'%';
	//oCode.innerText='zoom: '+newZoom+'';	
}

/***************************************Lei Rotate function 30-04-2013**************************/
function LeiRotate() {
    var div = document.getElementById("printarea");
<<<<<<< .mine
    div.className = "rotate";
    var widthHalf = div.clientWidth / 2;
    var heightHalf = div.clientHeight / 2;
    div.setAttribute("style", "position:absolute; zoom: 100%; left:40px; top:180px; visibility:visible;margin-top:" + (widthHalf - heightHalf) + "px");
=======

    var cls = div.className;
    if (cls.indexOf("rotate2") !== -1) {
        div.className = "rotate3";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "position:absolute; zoom: 100%; left:20px; top:100px; visibility:visible;margin-top:" + (widthHalf - heightHalf) + "px");
    }
    else if (cls.indexOf("rotate3") !== -1) {
        div.className = "";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "position:absolute; zoom: 100%; left:20px; top:100px; visibility:visible;");
    }
    else if (cls.indexOf("rotate") !== -1) {
        div.className = "rotate2";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "position:absolute; zoom: 100%; left:20px; top:100px; visibility:visible;");
    }
    else {
        div.className = "rotate";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "position:absolute; zoom: 100%; left:20px; top:100px; visibility:visible;margin-top:" + (widthHalf - heightHalf) + "px");
    }    
>>>>>>> .r53
}
</SCRIPT>

<!--#include file="../inc/regular/header_hvd_css3_html5_inc_Lei.asp"-->


			
    <style type="text/css">
    .rotate
    {
        -webkit-transform:rotate(-90deg);
        -moz-transform:rotate(-90deg);
        -o-transform:rotate(-90deg);
        -ms-transform:rotate(-90deg);
    }
    </style> 

<%	
    
    media = request("media")
	txt1 = request.form("txt1")
	'txt2 = request.form("txt2")
	
	Dim BigTextArea

	For I = 1 To Request.Form("BigTextArea").Count
  		BigTextArea = BigTextArea & Request.Form("BigTextArea")(I)
	Next
	
	txt20 = request.form("txt20")
	txt = txt1 & BigTextArea & txt20
	datointerval = request("datointerval")
%> 

<!--
<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
	<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
	<td bgcolor="#FFFFFF" align=right><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
</tr>
</table>
-->

<%oimg = "ikon_grandtotal_48.png"
	oleft = 20
	otop = 10
	owdt = 400
	oskrift = "Grandtotal, timer realiseret"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift) %>

      

<<<<<<< .mine
    <div style="position:relative; zoom: 100%; left:20px; top:0px; visibility:visible;" bgcolor="#ffffff">
<h4>Periode: <%=datointerval%></h4>
  <p>
=======
			
    <style type="text/css">
    .rotate
    {
        -webkit-transform:rotate(90deg);
        -moz-transform:rotate(90deg);
        -o-transform:rotate(90deg);
        -ms-transform:rotate(90deg);
    }
    .rotate2
    {
        -webkit-transform:rotate(180deg);
        -moz-transform:rotate(180deg);
        -o-transform:rotate(180deg);
        -ms-transform:rotate(180deg);
    }
    .rotate3
    {
        -webkit-transform:rotate(270deg);
        -moz-transform:rotate(270deg);
        -o-transform:rotate(270deg);
        -ms-transform:rotate(270deg);
    }
    </style>       

    
	<div id="printarea" name="printarea" style="position:absolute; zoom: 100%; left:20px; top:100px; visibility:visible;" bgcolor="#ffffff">
    
	<h5>Periode: <%=datointerval%> <span style="font-size:10px; line-height:12px; font-style:normal; font-weight:lighter;"><br />Zoom: Brug [CTRL] + / - for at zoome.</span></h5>
     <p>
>>>>>>> .r53
            <input id="btnRotate" type="button" value="Rotate" onclick="LeiRotate();" />
     </p>
    </div>
	<div id="printarea" name="printarea" style="position:absolute; zoom: 100%; left:40px; top:180px; visibility:visible;" bgcolor="#ffffff">
    
	
   
	<%=txt%><br><br>
	<!--
	<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a></div>
	-->
	
	
	<br><br><br>&nbsp;
	</div>
	
	<br><br>&nbsp;

     <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
		
	
<!--#include file="../inc/regular/footer_inc.asp"-->