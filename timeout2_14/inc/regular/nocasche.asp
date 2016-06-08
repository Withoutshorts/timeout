<%
Response.Expires = 15
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","public"
Response.CacheControl="Public"

%>
