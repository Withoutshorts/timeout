
<% 

Dim objXMLHTTP_global, objXMLDOM_global, i_global, strHTML_global
Dim Address_global, Latitude_global, Longitude_global
Dim oNode_global, oNodes_global
Dim sXPathQuery_global

Set objXMLDom_global = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_global = Server.CreateObject("Msxml2.ServerXMLHTTP")



select case lto
case "epi", "epi_no", "epi_sta", "epi_ab"
objXmlHttp_global.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/global_sprog_epi.xml", False
'case "fk"
'objXmlHttp_global.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/global_sprog_fk.xml", False
case "tec", "esn"
objXmlHttp_global.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/global_sprog_tec.xml", False
'case "intranet - local" 
'objXmlHttp_global.open "GET", "http://localhost/timeout_xp/inc/xml/global_sprog_fk.xml", False
case else
'objXmlHttp_global.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/global_sprog.xml", False
objXmlHttp_global.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/global_sprog.xml", False
'objXmlHttp_global.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/global_sprog.xml", False
'objXmlHttp_global.open "GET", "http://localhost/inc/xml/global_sprog.xml", False
'objXmlHttp_global.open "GET", "http://localhost/timeout_xp/inc/xml/global_sprog_tec.xml", False
end select

'objXmlHttp_global.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/global_sprog.xml", False
'objXmlHttp_global.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/global_sprog.xml", False

objXmlHttp_global.send


Set objXmlDom_global = objXmlHttp_global.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_global = Nothing





sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case sprog
case 1
sXPathQuery_global = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_global = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_global = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery_global = "//sprog/dk"
Session.LCID = 1030
end select

Set oNode_global = objXmlDom_global.documentElement.selectSingleNode(sXPathQuery_global)
Address_global = oNode_global.Text

Set oNodes_global = objXmlDom_global.documentElement.selectNodes(sXPathQuery_global)

    For Each oNode_global in oNodes_global

          global_txt_119 = oNode_global.selectSingleNode("txt_119").Text
          
         
          global_txt_120 = oNode_global.selectSingleNode("txt_120").Text
          global_txt_121 = oNode_global.selectSingleNode("txt_121").Text
          global_txt_122 = oNode_global.selectSingleNode("txt_122").Text
          global_txt_123 = oNode_global.selectSingleNode("txt_123").Text
          global_txt_124 = oNode_global.selectSingleNode("txt_124").Text
          global_txt_125 = oNode_global.selectSingleNode("txt_125").Text
          global_txt_126 = oNode_global.selectSingleNode("txt_126").Text
          global_txt_127 = oNode_global.selectSingleNode("txt_127").Text
          global_txt_128 = oNode_global.selectSingleNode("txt_128").Text
          
          global_txt_129 = oNode_global.selectSingleNode("txt_129").Text
          global_txt_130 = oNode_global.selectSingleNode("txt_130").Text
          global_txt_131 = oNode_global.selectSingleNode("txt_131").Text
          global_txt_132 = oNode_global.selectSingleNode("txt_132").Text
          global_txt_133 = oNode_global.selectSingleNode("txt_133").Text
          global_txt_134 = oNode_global.selectSingleNode("txt_134").Text
          global_txt_135 = oNode_global.selectSingleNode("txt_135").Text
          global_txt_136 = oNode_global.selectSingleNode("txt_136").Text
          global_txt_137 = oNode_global.selectSingleNode("txt_137").Text
          global_txt_138 = oNode_global.selectSingleNode("txt_138").Text
          global_txt_139 = oNode_global.selectSingleNode("txt_139").Text
          global_txt_140 = oNode_global.selectSingleNode("txt_140").Text
          global_txt_141 = oNode_global.selectSingleNode("txt_141").Text
          global_txt_142 = oNode_global.selectSingleNode("txt_142").Text
          global_txt_143 = oNode_global.selectSingleNode("txt_143").Text
          global_txt_144 = oNode_global.selectSingleNode("txt_144").Text
          global_txt_145 = oNode_global.selectSingleNode("txt_145").Text
          global_txt_146 = oNode_global.selectSingleNode("txt_146").Text
          global_txt_147 = oNode_global.selectSingleNode("txt_147").Text
          global_txt_148 = oNode_global.selectSingleNode("txt_148").Text
          global_txt_149 = oNode_global.selectSingleNode("txt_149").Text
          global_txt_150 = oNode_global.selectSingleNode("txt_150").Text
          global_txt_151 = oNode_global.selectSingleNode("txt_151").Text
          global_txt_152 = oNode_global.selectSingleNode("txt_152").Text
          global_txt_153 = oNode_global.selectSingleNode("txt_153").Text
          global_txt_154 = oNode_global.selectSingleNode("txt_154").Text
          global_txt_155 = oNode_global.selectSingleNode("txt_155").Text
          global_txt_156 = oNode_global.selectSingleNode("txt_156").Text
          global_txt_157 = oNode_global.selectSingleNode("txt_157").Text
          global_txt_158 = oNode_global.selectSingleNode("txt_158").Text
          global_txt_159 = oNode_global.selectSingleNode("txt_159").Text
          global_txt_160 = oNode_global.selectSingleNode("txt_160").Text
          global_txt_161 = oNode_global.selectSingleNode("txt_161").Text
          global_txt_162 = oNode_global.selectSingleNode("txt_162").Text
          global_txt_163 = oNode_global.selectSingleNode("txt_163").Text
          global_txt_164 = oNode_global.selectSingleNode("txt_164").Text
          global_txt_165 = oNode_global.selectSingleNode("txt_165").Text
          global_txt_166 = oNode_global.selectSingleNode("txt_166").Text
          global_txt_167 = oNode_global.selectSingleNode("txt_167").Text
          global_txt_168 = oNode_global.selectSingleNode("txt_168").Text
          global_txt_169 = oNode_global.selectSingleNode("txt_169").Text
          global_txt_170 = oNode_global.selectSingleNode("txt_170").Text
          
          global_txt_171 = oNode_global.selectSingleNode("txt_171").Text
          global_txt_172 = oNode_global.selectSingleNode("txt_172").Text
          global_txt_173 = oNode_global.selectSingleNode("txt_173").Text
          global_txt_174 = oNode_global.selectSingleNode("txt_174").Text  
          
          global_txt_175 = oNode_global.selectSingleNode("txt_175").Text 
          global_txt_176 = oNode_global.selectSingleNode("txt_176").Text 

          global_txt_177 = oNode_global.selectSingleNode("txt_177").Text
          global_txt_178 = oNode_global.selectSingleNode("txt_178").Text
    
         global_txt_179 = oNode_global.selectSingleNode("txt_179").Text

        global_txt_180 = oNode_global.selectSingleNode("txt_180").Text
        global_txt_181 = oNode_global.selectSingleNode("txt_181").Text
        global_txt_182 = oNode_global.selectSingleNode("txt_182").Text
     
        global_txt_183 = oNode_global.selectSingleNode("txt_183").Text
        global_txt_184 = oNode_global.selectSingleNode("txt_184").Text
        
        global_txt_185 = oNode_global.selectSingleNode("txt_185").Text
        global_txt_186 = oNode_global.selectSingleNode("txt_186").Text
        global_txt_187 = oNode_global.selectSingleNode("txt_187").Text

        global_txt_188 = oNode_global.selectSingleNode("txt_188").Text
        global_txt_189 = oNode_global.selectSingleNode("txt_189").Text
        global_txt_190 = oNode_global.selectSingleNode("txt_190").Text
        global_txt_191 = oNode_global.selectSingleNode("txt_191").Text

        global_txt_192 = oNode_global.selectSingleNode("txt_192").Text
        global_txt_193 = oNode_global.selectSingleNode("txt_193").Text

        global_txt_194 = oNode_global.selectSingleNode("txt_194").Text
        global_txt_195 = oNode_global.selectSingleNode("txt_195").Text

    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>