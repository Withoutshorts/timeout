

<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>

<!-- frames -->
	<%
	
	
	if len(request("FM_kunde")) <> 0 then
	kid = request("FM_kunde")
	else
	    if len(request.cookies("erp")("kid")) <> 0 then
	    kid = request.cookies("erp")("kid")
	    else
	    kid = 0
	    end if
	end if

	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	    if len(request.cookies("erp")("jid")) <> 0 then
	    jobid = request.cookies("erp")("jid")
	    else
	    jobid = 0
	    end if
	end if
	
	
	if len(request("FM_aftale")) <> 0 then
	aftaleid = request("FM_aftale")
	else
	    if len(request.cookies("erp")("aid")) <> 0 then
	    aftaleid = request.cookies("erp")("aid")
	    else
	    aftaleid = 0
	    end if
	end if
	
	if request("reset") = 1 then
	reset = 1 'request("reset")
	else
	reset = 0
	end if
	
	thisfile = "erp_oprfak_fs"
	
	
	
	
	
	    dontshowDD = 1
	    %>

        <!--#include file="inc/weekselector_s.asp"-->
	   
	  
	
	
	
	
	<frameset rows="126,*" framespacing="0" border="0" id="erp1">
    
    <frame src="erp_top.asp" name="top" id="top" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	        
	         <frameset cols="290,*" framespacing="0" id="erp6">
		            
		            
		        <frameset rows="160,197,*" framespacing="0" id="erp2">
    		            
    		            <frame src="erp_opr_faktura_kontakter.asp?FM_kunde=<%=kid%>" name="erp2_0" id="erp2_0" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    
	                    <%if cint(reset) = 1 then %>
	                    <frame src="erp_opr_faktura_kontojob.asp?FM_usedatokri=1&FM_kunde=<%=kid%>&FM_job=<%=jobid%>&FM_aftale=<%=aftaleid%>&&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" name="erp2_1" id="erp2_1" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%else %>
	                    <frame src="erp_opr_faktura_blank.asp" name="erp2_1" id="erp2_1" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%end if %>
	                    
	                    
	                    <%if cint(reset) = 1 then %>
	                    <frame src="erp_opr_faktura.asp?FM_usedatokri=1&FM_kunde=<%=kid%>&FM_job=<%=jobid%>&FM_aftale=<%=aftaleid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" name="erp2_2" id="erp2_2" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%else %>
	                    <frame src="erp_opr_faktura_blank.asp" name="erp2_2" id="erp2_2" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%end if %>
	                    <!--<frame src="erp_opr_faktura_dato.asp" name="erp2_3" id="erp2_3" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">-->
    	               
    	                
		        </frameset>
		        
		        
		        <%if cint(reset) = 1 then 
		        
		        if len(request("fid")) <> 0 then
		        fid = request("fid")
		        else
		        fid = 0
		        end if
		        
		        if len(request("func")) <> 0 then
		        func = request("func")
		        else
		        func = ""
		        end if
		        
		        %>
		        <frame src="erp_fak.asp?func=<%=func%>&id=<%=fid%>&FM_usedatokri=1&FM_kunde=<%=kid%>&FM_job=<%=jobid%>&FM_aftale=<%=aftaleid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" name="erp3" id="erp3" frameborder="0" scrolling="auto" marginwidth="0" marginheight="0">
	            <%else %>
	            <frame src="erp_opr_faktura_blank.asp" name="erp3" id="erp3" frameborder="0" scrolling="auto" marginwidth="0" marginheight="0">
	            <%end if %>
		        
		        
		    </frameset>
    
    
                   
                   
		    
		  
		    
    </frameset>
	
	
	<%end if %>
	






