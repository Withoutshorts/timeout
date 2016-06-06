<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="overblik" style="position:absolute; left:190; top:90; visibility:visible;">
	<img src="../ill/header_resoblik.gif" width="755" height="48" alt="" border="0">
	<br>
	<br>
	<a href="jbpla_w.asp?menu=res">Ressource Kalender</a><br>
	Opbygget som en <b>uge-oversigts kalender</b>.<br>
	Denne kalender buruges når du skal booke en medarbejder nogle timer på daglig basis,<br>
	f.eks i en <b>support situation</b>, eller til et møde.
	Denne kalender kan synkroniseres med jeres Exchange server,<br> så de bookinger der oprettes her også oprettes i den enkelte medarbejders Outlook.<br>
	<br><br>
	
	<a href="ressource_belaeg_jbpla.asp">Ressource Timer</a><br>
	Overblik og <b>tildeling af ressource-timer</b> på job for hver enkelt medarbejder.<br> 
	Bruges også som f.eks <b>ferie kalender</b>.
	<br>
	<br><br>
	
	<a href="ressource_belaeg.asp">Ressource Belægning</a><br>
	Det <b>samlede overblik for hver enkelt medarbejder</b>. <br>
	Balance imellem tildelte og forbruge timer iforhold til medarbejderens normerede arbejds-uge.
	<br>
	<br><br>
	
	<a href="jbpla_k.asp">Job Belastning</a><br>
	Overblik over hvor mange ressouce timer der er <b>tildelt de enkelte job</b>,<br>
	samt hvornår der er <b>plads i ordrebøgerne</b> til nye opgaver, eller om
	der evt. ansættes flere folk.<br>
	<br><br>
	
</div>
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
