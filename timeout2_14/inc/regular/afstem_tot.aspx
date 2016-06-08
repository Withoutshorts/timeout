<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Sub page_load()
        load.Text = "Afstemning"
         
    End Sub
</script>



<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	%>


    <asp:Label runat="server" Text="Label" id="load"></asp:Label>
   
<%end if 'validering %>
<!--#include file="../inc/regular/footer_inc.asp"-->
