
<% 

Dim objXMLHTTP, objXMLDOM, i, strHTML

Set objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp.open "GET", "http://localhost/inc/xml/tsa_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/tsa_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/tsa_sprog.xml", False
objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/tsa_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/tsa_sprog.xml", False

objXmlHttp.send


Set objXmlDom = objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp = Nothing



Dim Address, Latitude, Longitude
Dim oNode, oNodes
Dim sXPathQuery


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
sXPathQuery = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set oNode = objXmlDom.documentElement.selectSingleNode(sXPathQuery)
Address = oNode.Text

Set oNodes = objXmlDom.documentElement.selectNodes(sXPathQuery)

For Each oNode in oNodes

     tsa_txt_001 = oNode.selectSingleNode("txt_1").Text
     tsa_txt_002 = oNode.selectSingleNode("txt_2").Text
     tsa_txt_003 = oNode.selectSingleNode("txt_3").Text
     tsa_txt_003 = oNode.selectSingleNode("txt_3").Text
    tsa_txt_004 = oNode.selectSingleNode("txt_4").Text
    tsa_txt_005 = oNode.selectSingleNode("txt_5").Text

    tsa_txt_006 = oNode.selectSingleNode("txt_6").Text
    tsa_txt_007 = oNode.selectSingleNode("txt_7").Text
    tsa_txt_008 = oNode.selectSingleNode("txt_8").Text
    tsa_txt_009 = oNode.selectSingleNode("txt_9").Text
    tsa_txt_010 = oNode.selectSingleNode("txt_10").Text
    tsa_txt_011 = oNode.selectSingleNode("txt_11").Text
    
    tsa_txt_012 = oNode.selectSingleNode("txt_12").Text
    tsa_txt_013 = oNode.selectSingleNode("txt_13").Text
    tsa_txt_014 = oNode.selectSingleNode("txt_14").Text
    tsa_txt_015 = oNode.selectSingleNode("txt_15").Text
    tsa_txt_016 = oNode.selectSingleNode("txt_16").Text
    tsa_txt_017 = oNode.selectSingleNode("txt_17").Text
    
    tsa_txt_018 = oNode.selectSingleNode("txt_18").Text
    tsa_txt_019 = oNode.selectSingleNode("txt_19").Text
    tsa_txt_020 = oNode.selectSingleNode("txt_20").Text
    tsa_txt_021 = oNode.selectSingleNode("txt_21").Text
    tsa_txt_022 = oNode.selectSingleNode("txt_22").Text
    tsa_txt_023 = oNode.selectSingleNode("txt_23").Text
    tsa_txt_024 = oNode.selectSingleNode("txt_24").Text
    
    tsa_txt_025 = oNode.selectSingleNode("txt_25").Text
    tsa_txt_026 = oNode.selectSingleNode("txt_26").Text
    tsa_txt_027 = oNode.selectSingleNode("txt_27").Text
    tsa_txt_028 = oNode.selectSingleNode("txt_28").Text
    tsa_txt_029 = oNode.selectSingleNode("txt_29").Text
    tsa_txt_030 = oNode.selectSingleNode("txt_30").Text
    tsa_txt_031 = oNode.selectSingleNode("txt_31").Text
    tsa_txt_032 = oNode.selectSingleNode("txt_32").Text
    tsa_txt_033 = oNode.selectSingleNode("txt_33").Text
    tsa_txt_034 = oNode.selectSingleNode("txt_34").Text
    
     tsa_txt_035 = oNode.selectSingleNode("txt_35").Text
      tsa_txt_036 = oNode.selectSingleNode("txt_36").Text
       tsa_txt_037 = oNode.selectSingleNode("txt_37").Text
       
        tsa_txt_038 = oNode.selectSingleNode("txt_38").Text
         tsa_txt_039 = oNode.selectSingleNode("txt_39").Text
         tsa_txt_040 = oNode.selectSingleNode("txt_40").Text
         tsa_txt_041 = oNode.selectSingleNode("txt_41").Text
         tsa_txt_042 = oNode.selectSingleNode("txt_42").Text
          tsa_txt_043 = oNode.selectSingleNode("txt_43").Text
          tsa_txt_044 = oNode.selectSingleNode("txt_44").Text
          
          tsa_txt_045 = oNode.selectSingleNode("txt_45").Text
          tsa_txt_046 = oNode.selectSingleNode("txt_46").Text
          tsa_txt_047 = oNode.selectSingleNode("txt_47").Text
          tsa_txt_048 = oNode.selectSingleNode("txt_48").Text
          tsa_txt_049 = oNode.selectSingleNode("txt_49").Text
          tsa_txt_050 = oNode.selectSingleNode("txt_50").Text
          
          tsa_txt_051 = oNode.selectSingleNode("txt_51").Text
          tsa_txt_052 = oNode.selectSingleNode("txt_52").Text
          tsa_txt_053 = oNode.selectSingleNode("txt_53").Text
          tsa_txt_054 = oNode.selectSingleNode("txt_54").Text
          tsa_txt_055 = oNode.selectSingleNode("txt_55").Text
          tsa_txt_056 = oNode.selectSingleNode("txt_56").Text
          tsa_txt_057 = oNode.selectSingleNode("txt_57").Text
          tsa_txt_058 = oNode.selectSingleNode("txt_58").Text
          tsa_txt_059 = oNode.selectSingleNode("txt_59").Text
          tsa_txt_060 = oNode.selectSingleNode("txt_60").Text
          tsa_txt_061 = oNode.selectSingleNode("txt_61").Text
          tsa_txt_062 = oNode.selectSingleNode("txt_62").Text
          tsa_txt_063 = oNode.selectSingleNode("txt_63").Text
          tsa_txt_064 = oNode.selectSingleNode("txt_64").Text
          tsa_txt_065 = oNode.selectSingleNode("txt_65").Text
          tsa_txt_066 = oNode.selectSingleNode("txt_66").Text
          
          tsa_txt_067 = oNode.selectSingleNode("txt_67").Text
          tsa_txt_068 = oNode.selectSingleNode("txt_68").Text
          tsa_txt_069 = oNode.selectSingleNode("txt_69").Text
          tsa_txt_070 = oNode.selectSingleNode("txt_70").Text
          tsa_txt_071 = oNode.selectSingleNode("txt_71").Text
          tsa_txt_072 = oNode.selectSingleNode("txt_72").Text
          tsa_txt_073 = oNode.selectSingleNode("txt_73").Text
          tsa_txt_074 = oNode.selectSingleNode("txt_74").Text
          tsa_txt_075 = oNode.selectSingleNode("txt_75").Text
          tsa_txt_076 = oNode.selectSingleNode("txt_76").Text
          tsa_txt_077 = oNode.selectSingleNode("txt_77").Text
          tsa_txt_078 = oNode.selectSingleNode("txt_78").Text
          tsa_txt_079 = oNode.selectSingleNode("txt_79").Text
        
          
          tsa_txt_080 = oNode.selectSingleNode("txt_80").Text
          tsa_txt_081 = oNode.selectSingleNode("txt_81").Text
          tsa_txt_082 = oNode.selectSingleNode("txt_82").Text
          tsa_txt_083 = oNode.selectSingleNode("txt_83").Text
          
          tsa_txt_084 = oNode.selectSingleNode("txt_84").Text
          tsa_txt_085 = oNode.selectSingleNode("txt_85").Text
          tsa_txt_086 = oNode.selectSingleNode("txt_86").Text
          
          tsa_txt_087 = oNode.selectSingleNode("txt_87").Text
          
          tsa_txt_088 = oNode.selectSingleNode("txt_88").Text
          tsa_txt_089 = oNode.selectSingleNode("txt_89").Text
          tsa_txt_090 = oNode.selectSingleNode("txt_90").Text
          
          tsa_txt_091 = oNode.selectSingleNode("txt_91").Text
          tsa_txt_092 = oNode.selectSingleNode("txt_92").Text
          tsa_txt_093 = oNode.selectSingleNode("txt_93").Text
          tsa_txt_094 = oNode.selectSingleNode("txt_94").Text
          tsa_txt_095 = oNode.selectSingleNode("txt_95").Text
          tsa_txt_096 = oNode.selectSingleNode("txt_96").Text
          tsa_txt_097 = oNode.selectSingleNode("txt_97").Text
          tsa_txt_098 = oNode.selectSingleNode("txt_98").Text
          tsa_txt_099 = oNode.selectSingleNode("txt_99").Text
          
          tsa_txt_100 = oNode.selectSingleNode("txt_100").Text
          tsa_txt_101 = oNode.selectSingleNode("txt_101").Text
          tsa_txt_102 = oNode.selectSingleNode("txt_102").Text
          tsa_txt_103 = oNode.selectSingleNode("txt_103").Text
          
          tsa_txt_104 = oNode.selectSingleNode("txt_104").Text
          tsa_txt_105 = oNode.selectSingleNode("txt_105").Text
          tsa_txt_106 = oNode.selectSingleNode("txt_106").Text
         
          
          tsa_txt_116 = oNode.selectSingleNode("txt_116").Text
          tsa_txt_117 = oNode.selectSingleNode("txt_117").Text
          tsa_txt_118 = oNode.selectSingleNode("txt_118").Text
          tsa_txt_119 = oNode.selectSingleNode("txt_119").Text
         
          tsa_txt_120 = oNode.selectSingleNode("txt_120").Text
          tsa_txt_121 = oNode.selectSingleNode("txt_121").Text
          tsa_txt_122 = oNode.selectSingleNode("txt_122").Text
          tsa_txt_123 = oNode.selectSingleNode("txt_123").Text
          
          tsa_txt_124 = oNode.selectSingleNode("txt_124").Text
           tsa_txt_125 = oNode.selectSingleNode("txt_125").Text
            tsa_txt_126 = oNode.selectSingleNode("txt_126").Text
            tsa_txt_127 = oNode.selectSingleNode("txt_127").Text
            tsa_txt_128 = oNode.selectSingleNode("txt_128").Text
            
            tsa_txt_129 = oNode.selectSingleNode("txt_129").Text
            tsa_txt_130 = oNode.selectSingleNode("txt_130").Text
            tsa_txt_131 = oNode.selectSingleNode("txt_131").Text
            tsa_txt_132 = oNode.selectSingleNode("txt_132").Text
            tsa_txt_133 = oNode.selectSingleNode("txt_133").Text
            
            tsa_txt_134 = oNode.selectSingleNode("txt_134").Text
            tsa_txt_135 = oNode.selectSingleNode("txt_135").Text
            tsa_txt_136 = oNode.selectSingleNode("txt_136").Text
            tsa_txt_137 = oNode.selectSingleNode("txt_137").Text
            tsa_txt_138 = oNode.selectSingleNode("txt_138").Text
            tsa_txt_139 = oNode.selectSingleNode("txt_139").Text
            
            tsa_txt_140 = oNode.selectSingleNode("txt_140").Text
            tsa_txt_141 = oNode.selectSingleNode("txt_141").Text
            
            tsa_txt_142 = oNode.selectSingleNode("txt_142").Text
            tsa_txt_143 = oNode.selectSingleNode("txt_143").Text
            tsa_txt_144 = oNode.selectSingleNode("txt_144").Text
            tsa_txt_145 = oNode.selectSingleNode("txt_145").Text
            tsa_txt_146 = oNode.selectSingleNode("txt_146").Text
            
            tsa_txt_147 = oNode.selectSingleNode("txt_147").Text
            tsa_txt_148 = oNode.selectSingleNode("txt_148").Text
            tsa_txt_149 = oNode.selectSingleNode("txt_149").Text
            
            tsa_txt_150 = oNode.selectSingleNode("txt_150").Text
            tsa_txt_151 = oNode.selectSingleNode("txt_151").Text
            tsa_txt_152 = oNode.selectSingleNode("txt_152").Text
            tsa_txt_153 = oNode.selectSingleNode("txt_153").Text
            tsa_txt_154 = oNode.selectSingleNode("txt_154").Text
            tsa_txt_155 = oNode.selectSingleNode("txt_155").Text
            tsa_txt_156 = oNode.selectSingleNode("txt_156").Text
            tsa_txt_157 = oNode.selectSingleNode("txt_157").Text
            tsa_txt_158 = oNode.selectSingleNode("txt_158").Text
            tsa_txt_159 = oNode.selectSingleNode("txt_159").Text
            
            tsa_txt_160 = oNode.selectSingleNode("txt_160").Text
            tsa_txt_161 = oNode.selectSingleNode("txt_161").Text
            tsa_txt_162 = oNode.selectSingleNode("txt_162").Text
            tsa_txt_163 = oNode.selectSingleNode("txt_163").Text
            tsa_txt_164 = oNode.selectSingleNode("txt_164").Text
            tsa_txt_165 = oNode.selectSingleNode("txt_165").Text
            tsa_txt_166 = oNode.selectSingleNode("txt_166").Text
            tsa_txt_167 = oNode.selectSingleNode("txt_167").Text
            tsa_txt_168 = oNode.selectSingleNode("txt_168").Text
            tsa_txt_169 = oNode.selectSingleNode("txt_169").Text
            
            tsa_txt_170 = oNode.selectSingleNode("txt_170").Text
            tsa_txt_171 = oNode.selectSingleNode("txt_171").Text
            tsa_txt_172 = oNode.selectSingleNode("txt_172").Text
            tsa_txt_173 = oNode.selectSingleNode("txt_173").Text
            tsa_txt_174 = oNode.selectSingleNode("txt_174").Text
            tsa_txt_175 = oNode.selectSingleNode("txt_175").Text
            tsa_txt_176 = oNode.selectSingleNode("txt_176").Text
            tsa_txt_177 = oNode.selectSingleNode("txt_177").Text
            tsa_txt_178 = oNode.selectSingleNode("txt_178").Text
            tsa_txt_179 = oNode.selectSingleNode("txt_179").Text
            
            tsa_txt_180 = oNode.selectSingleNode("txt_180").Text
            tsa_txt_181 = oNode.selectSingleNode("txt_181").Text
            tsa_txt_182 = oNode.selectSingleNode("txt_182").Text
            tsa_txt_183 = oNode.selectSingleNode("txt_183").Text
            tsa_txt_184 = oNode.selectSingleNode("txt_184").Text
            tsa_txt_185 = oNode.selectSingleNode("txt_185").Text
            tsa_txt_186 = oNode.selectSingleNode("txt_186").Text
            tsa_txt_187 = oNode.selectSingleNode("txt_187").Text
            
            tsa_txt_188 = oNode.selectSingleNode("txt_188").Text
            tsa_txt_189 = oNode.selectSingleNode("txt_189").Text
            
            tsa_txt_190 = oNode.selectSingleNode("txt_190").Text
            tsa_txt_191 = oNode.selectSingleNode("txt_191").Text
            tsa_txt_192 = oNode.selectSingleNode("txt_192").Text
            tsa_txt_193 = oNode.selectSingleNode("txt_193").Text
            tsa_txt_194 = oNode.selectSingleNode("txt_194").Text
            tsa_txt_195 = oNode.selectSingleNode("txt_195").Text
            tsa_txt_196 = oNode.selectSingleNode("txt_196").Text
            tsa_txt_197 = oNode.selectSingleNode("txt_197").Text
            tsa_txt_198 = oNode.selectSingleNode("txt_198").Text
            tsa_txt_199 = oNode.selectSingleNode("txt_199").Text
            
            tsa_txt_200 = oNode.selectSingleNode("txt_200").Text
            tsa_txt_201 = oNode.selectSingleNode("txt_201").Text
            tsa_txt_202 = oNode.selectSingleNode("txt_202").Text
            tsa_txt_203 = oNode.selectSingleNode("txt_203").Text
            tsa_txt_204 = oNode.selectSingleNode("txt_204").Text
            tsa_txt_205 = oNode.selectSingleNode("txt_205").Text
            tsa_txt_206 = oNode.selectSingleNode("txt_206").Text
            tsa_txt_207 = oNode.selectSingleNode("txt_207").Text
            tsa_txt_208 = oNode.selectSingleNode("txt_208").Text
            tsa_txt_209 = oNode.selectSingleNode("txt_209").Text
            
            tsa_txt_210 = oNode.selectSingleNode("txt_210").Text
            tsa_txt_211 = oNode.selectSingleNode("txt_211").Text
            tsa_txt_212 = oNode.selectSingleNode("txt_212").Text
            tsa_txt_213 = oNode.selectSingleNode("txt_213").Text
            tsa_txt_214 = oNode.selectSingleNode("txt_214").Text
            tsa_txt_215 = oNode.selectSingleNode("txt_215").Text
            tsa_txt_216 = oNode.selectSingleNode("txt_216").Text
            tsa_txt_217 = oNode.selectSingleNode("txt_217").Text
            tsa_txt_218 = oNode.selectSingleNode("txt_218").Text
            tsa_txt_219 = oNode.selectSingleNode("txt_219").Text
            
            tsa_txt_220 = oNode.selectSingleNode("txt_220").Text
            tsa_txt_221 = oNode.selectSingleNode("txt_221").Text
            tsa_txt_222 = oNode.selectSingleNode("txt_222").Text
            tsa_txt_223 = oNode.selectSingleNode("txt_223").Text
            tsa_txt_224 = oNode.selectSingleNode("txt_224").Text
            tsa_txt_225 = oNode.selectSingleNode("txt_225").Text
            tsa_txt_226 = oNode.selectSingleNode("txt_226").Text
            tsa_txt_227 = oNode.selectSingleNode("txt_227").Text
            tsa_txt_228 = oNode.selectSingleNode("txt_228").Text
            tsa_txt_229 = oNode.selectSingleNode("txt_229").Text
            tsa_txt_230 = oNode.selectSingleNode("txt_230").Text
            tsa_txt_231 = oNode.selectSingleNode("txt_231").Text
            tsa_txt_232 = oNode.selectSingleNode("txt_232").Text
            tsa_txt_233 = oNode.selectSingleNode("txt_233").Text
            tsa_txt_234 = oNode.selectSingleNode("txt_234").Text
            tsa_txt_235 = oNode.selectSingleNode("txt_235").Text
            tsa_txt_236 = oNode.selectSingleNode("txt_236").Text
            tsa_txt_237 = oNode.selectSingleNode("txt_237").Text
            tsa_txt_238 = oNode.selectSingleNode("txt_238").Text
            tsa_txt_239 = oNode.selectSingleNode("txt_239").Text
            tsa_txt_240 = oNode.selectSingleNode("txt_240").Text
            tsa_txt_241 = oNode.selectSingleNode("txt_241").Text
            
            tsa_txt_242 = oNode.selectSingleNode("txt_242").Text
            tsa_txt_243 = oNode.selectSingleNode("txt_243").Text
            tsa_txt_244 = oNode.selectSingleNode("txt_244").Text
            tsa_txt_245 = oNode.selectSingleNode("txt_245").Text
            tsa_txt_246 = oNode.selectSingleNode("txt_246").Text
            tsa_txt_247 = oNode.selectSingleNode("txt_247").Text
            tsa_txt_248 = oNode.selectSingleNode("txt_248").Text
            tsa_txt_249 = oNode.selectSingleNode("txt_249").Text
            tsa_txt_250 = oNode.selectSingleNode("txt_250").Text
            tsa_txt_251 = oNode.selectSingleNode("txt_251").Text
            tsa_txt_252 = oNode.selectSingleNode("txt_252").Text
            tsa_txt_253 = oNode.selectSingleNode("txt_253").Text
            tsa_txt_254 = oNode.selectSingleNode("txt_254").Text
            tsa_txt_255 = oNode.selectSingleNode("txt_255").Text
            tsa_txt_256 = oNode.selectSingleNode("txt_256").Text
            tsa_txt_257 = oNode.selectSingleNode("txt_257").Text
            tsa_txt_258 = oNode.selectSingleNode("txt_258").Text
            tsa_txt_259 = oNode.selectSingleNode("txt_259").Text
            tsa_txt_260 = oNode.selectSingleNode("txt_260").Text
            tsa_txt_261 = oNode.selectSingleNode("txt_261").Text
            tsa_txt_262 = oNode.selectSingleNode("txt_262").Text
            tsa_txt_263 = oNode.selectSingleNode("txt_263").Text
            tsa_txt_264 = oNode.selectSingleNode("txt_264").Text
            tsa_txt_265 = oNode.selectSingleNode("txt_265").Text
            tsa_txt_266 = oNode.selectSingleNode("txt_266").Text
            tsa_txt_267 = oNode.selectSingleNode("txt_267").Text
            tsa_txt_268 = oNode.selectSingleNode("txt_268").Text
            tsa_txt_269 = oNode.selectSingleNode("txt_269").Text
            tsa_txt_270 = oNode.selectSingleNode("txt_270").Text
            tsa_txt_271 = oNode.selectSingleNode("txt_271").Text
            tsa_txt_272 = oNode.selectSingleNode("txt_272").Text
            tsa_txt_273 = oNode.selectSingleNode("txt_273").Text
            tsa_txt_274 = oNode.selectSingleNode("txt_274").Text
            tsa_txt_275 = oNode.selectSingleNode("txt_275").Text
            tsa_txt_276 = oNode.selectSingleNode("txt_276").Text
            tsa_txt_277 = oNode.selectSingleNode("txt_277").Text
            tsa_txt_278 = oNode.selectSingleNode("txt_278").Text
            tsa_txt_279 = oNode.selectSingleNode("txt_279").Text
            tsa_txt_280 = oNode.selectSingleNode("txt_280").Text
             tsa_txt_281 = oNode.selectSingleNode("txt_281").Text
              tsa_txt_282 = oNode.selectSingleNode("txt_282").Text
               tsa_txt_283 = oNode.selectSingleNode("txt_283").Text
             tsa_txt_284 = oNode.selectSingleNode("txt_284").Text
             tsa_txt_285 = oNode.selectSingleNode("txt_285").Text
             tsa_txt_286 = oNode.selectSingleNode("txt_286").Text
             tsa_txt_287 = oNode.selectSingleNode("txt_287").Text
            tsa_txt_288 = oNode.selectSingleNode("txt_288").Text
            tsa_txt_289 = oNode.selectSingleNode("txt_289").Text
            
            tsa_txt_290 = oNode.selectSingleNode("txt_290").Text
            tsa_txt_291 = oNode.selectSingleNode("txt_291").Text
            tsa_txt_292 = oNode.selectSingleNode("txt_292").Text
            tsa_txt_293 = oNode.selectSingleNode("txt_293").Text
            tsa_txt_294 = oNode.selectSingleNode("txt_294").Text
            tsa_txt_295 = oNode.selectSingleNode("txt_295").Text
            tsa_txt_296 = oNode.selectSingleNode("txt_296").Text
            tsa_txt_297 = oNode.selectSingleNode("txt_297").Text
            tsa_txt_298 = oNode.selectSingleNode("txt_298").Text
            tsa_txt_299 = oNode.selectSingleNode("txt_299").Text
            tsa_txt_300 = oNode.selectSingleNode("txt_300").Text
            tsa_txt_301 = oNode.selectSingleNode("txt_301").Text
            tsa_txt_302 = oNode.selectSingleNode("txt_302").Text
            tsa_txt_303 = oNode.selectSingleNode("txt_303").Text
            tsa_txt_304 = oNode.selectSingleNode("txt_304").Text
            tsa_txt_305 = oNode.selectSingleNode("txt_305").Text
             tsa_txt_306 = oNode.selectSingleNode("txt_306").Text
             tsa_txt_307 = oNode.selectSingleNode("txt_307").Text
             tsa_txt_308 = oNode.selectSingleNode("txt_308").Text
             tsa_txt_309 = oNode.selectSingleNode("txt_309").Text
             tsa_txt_310 = oNode.selectSingleNode("txt_310").Text
             tsa_txt_311 = oNode.selectSingleNode("txt_311").Text
             tsa_txt_312 = oNode.selectSingleNode("txt_312").Text
             tsa_txt_313 = oNode.selectSingleNode("txt_313").Text
             tsa_txt_314 = oNode.selectSingleNode("txt_314").Text
             tsa_txt_315 = oNode.selectSingleNode("txt_315").Text
             tsa_txt_316 = oNode.selectSingleNode("txt_316").Text
             
             tsa_txt_317 = oNode.selectSingleNode("txt_317").Text
             tsa_txt_318 = oNode.selectSingleNode("txt_318").Text
             
             tsa_txt_319 = oNode.selectSingleNode("txt_319").Text
             tsa_txt_320 = oNode.selectSingleNode("txt_320").Text
             tsa_txt_321 = oNode.selectSingleNode("txt_321").Text
             tsa_txt_322 = oNode.selectSingleNode("txt_322").Text
              tsa_txt_323 = oNode.selectSingleNode("txt_323").Text
              tsa_txt_324 = oNode.selectSingleNode("txt_324").Text
              tsa_txt_325 = oNode.selectSingleNode("txt_325").Text
              tsa_txt_326 = oNode.selectSingleNode("txt_326").Text
              tsa_txt_327 = oNode.selectSingleNode("txt_327").Text
              tsa_txt_328 = oNode.selectSingleNode("txt_328").Text
              tsa_txt_329 = oNode.selectSingleNode("txt_329").Text
              
              tsa_txt_330 = oNode.selectSingleNode("txt_330").Text
              tsa_txt_331 = oNode.selectSingleNode("txt_331").Text
              tsa_txt_332 = oNode.selectSingleNode("txt_332").Text
              tsa_txt_333 = oNode.selectSingleNode("txt_333").Text
              tsa_txt_334 = oNode.selectSingleNode("txt_334").Text
              tsa_txt_335 = oNode.selectSingleNode("txt_335").Text
              tsa_txt_336 = oNode.selectSingleNode("txt_336").Text
              tsa_txt_337 = oNode.selectSingleNode("txt_337").Text
              tsa_txt_338 = oNode.selectSingleNode("txt_338").Text
              tsa_txt_339 = oNode.selectSingleNode("txt_339").Text
              tsa_txt_340 = oNode.selectSingleNode("txt_340").Text
              
              tsa_txt_341 = oNode.selectSingleNode("txt_341").Text
              tsa_txt_342 = oNode.selectSingleNode("txt_342").Text
              tsa_txt_343 = oNode.selectSingleNode("txt_343").Text
              tsa_txt_344 = oNode.selectSingleNode("txt_344").Text
              tsa_txt_345 = oNode.selectSingleNode("txt_345").Text
              tsa_txt_346 = oNode.selectSingleNode("txt_346").Text
              
              tsa_txt_347 = oNode.selectSingleNode("txt_347").Text
              tsa_txt_348 = oNode.selectSingleNode("txt_348").Text
              tsa_txt_349 = oNode.selectSingleNode("txt_349").Text
              tsa_txt_350 = oNode.selectSingleNode("txt_350").Text
              tsa_txt_351 = oNode.selectSingleNode("txt_351").Text
              tsa_txt_352 = oNode.selectSingleNode("txt_352").Text
              tsa_txt_353 = oNode.selectSingleNode("txt_353").Text
              
              tsa_txt_354 = oNode.selectSingleNode("txt_354").Text
              tsa_txt_355 = oNode.selectSingleNode("txt_355").Text
              tsa_txt_356 = oNode.selectSingleNode("txt_356").Text
              tsa_txt_357 = oNode.selectSingleNode("txt_357").Text
              tsa_txt_358 = oNode.selectSingleNode("txt_358").Text
              tsa_txt_359 = oNode.selectSingleNode("txt_359").Text
              tsa_txt_360 = oNode.selectSingleNode("txt_360").Text
              tsa_txt_361 = oNode.selectSingleNode("txt_361").Text
              tsa_txt_362 = oNode.selectSingleNode("txt_362").Text
              tsa_txt_363 = oNode.selectSingleNode("txt_363").Text
              tsa_txt_364 = oNode.selectSingleNode("txt_364").Text
              
              tsa_txt_365 = oNode.selectSingleNode("txt_365").Text
              tsa_txt_366 = oNode.selectSingleNode("txt_366").Text
              tsa_txt_367 = oNode.selectSingleNode("txt_367").Text
              tsa_txt_368 = oNode.selectSingleNode("txt_368").Text
              tsa_txt_369 = oNode.selectSingleNode("txt_369").Text
              tsa_txt_370 = oNode.selectSingleNode("txt_370").Text
              tsa_txt_371 = oNode.selectSingleNode("txt_371").Text
              tsa_txt_372 = oNode.selectSingleNode("txt_372").Text
              tsa_txt_373 = oNode.selectSingleNode("txt_373").Text
              tsa_txt_374 = oNode.selectSingleNode("txt_374").Text
              tsa_txt_375 = oNode.selectSingleNode("txt_375").Text
              tsa_txt_376 = oNode.selectSingleNode("txt_376").Text

              tsa_txt_377 = oNode.selectSingleNode("txt_377").Text
              tsa_txt_378 = oNode.selectSingleNode("txt_378").Text
              tsa_txt_379 = oNode.selectSingleNode("txt_379").Text
              tsa_txt_380 = oNode.selectSingleNode("txt_380").Text
              tsa_txt_381 = oNode.selectSingleNode("txt_381").Text
              tsa_txt_382 = oNode.selectSingleNode("txt_382").Text
              tsa_txt_383 = oNode.selectSingleNode("txt_383").Text
                
              tsa_txt_384 = oNode.selectSingleNode("txt_384").Text
              tsa_txt_385 = oNode.selectSingleNode("txt_385").Text
              tsa_txt_386 = oNode.selectSingleNode("txt_386").Text
              tsa_txt_387 = oNode.selectSingleNode("txt_387").Text
              tsa_txt_388 = oNode.selectSingleNode("txt_388").Text

              tsa_txt_389 = oNode.selectSingleNode("txt_389").Text
              tsa_txt_390 = oNode.selectSingleNode("txt_390").Text

              tsa_txt_391 = oNode.selectSingleNode("txt_391").Text
              tsa_txt_392 = oNode.selectSingleNode("txt_392").Text
              tsa_txt_393 = oNode.selectSingleNode("txt_393").Text

              tsa_txt_394 = oNode.selectSingleNode("txt_394").Text
              tsa_txt_395 = oNode.selectSingleNode("txt_395").Text
              tsa_txt_396 = oNode.selectSingleNode("txt_396").Text
        
              tsa_txt_397 = oNode.selectSingleNode("txt_397").Text
                tsa_txt_398 = oNode.selectSingleNode("txt_398").Text
                tsa_txt_399 = oNode.selectSingleNode("txt_399").Text
                tsa_txt_400 = oNode.selectSingleNode("txt_400").Text
                tsa_txt_401 = oNode.selectSingleNode("txt_401").Text

              tsa_txt_402 = oNode.selectSingleNode("txt_402").Text

             tsa_txt_403 = oNode.selectSingleNode("txt_403").Text
             tsa_txt_404 = oNode.selectSingleNode("txt_404").Text

             tsa_txt_405 = oNode.selectSingleNode("txt_405").Text
             tsa_txt_406 = oNode.selectSingleNode("txt_406").Text
             tsa_txt_407 = oNode.selectSingleNode("txt_407").Text
             tsa_txt_408 = oNode.selectSingleNode("txt_408").Text
             tsa_txt_409 = oNode.selectSingleNode("txt_409").Text
             tsa_txt_410 = oNode.selectSingleNode("txt_410").Text
            tsa_txt_411 = oNode.selectSingleNode("txt_411").Text
    
            tsa_txt_412 = oNode.selectSingleNode("txt_412").Text
    tsa_txt_413 = oNode.selectSingleNode("txt_413").Text
    tsa_txt_414 = oNode.selectSingleNode("txt_414").Text
    tsa_txt_415 = oNode.selectSingleNode("txt_415").Text
    tsa_txt_416 = oNode.selectSingleNode("txt_416").Text
    tsa_txt_417 = oNode.selectSingleNode("txt_417").Text
    tsa_txt_418 = oNode.selectSingleNode("txt_418").Text
    tsa_txt_419 = oNode.selectSingleNode("txt_419").Text
    tsa_txt_420 = oNode.selectSingleNode("txt_420").Text

     tsa_txt_421 = oNode.selectSingleNode("txt_421").Text
    tsa_txt_422 = oNode.selectSingleNode("txt_422").Text
    tsa_txt_423 = oNode.selectSingleNode("txt_423").Text
    tsa_txt_424 = oNode.selectSingleNode("txt_424").Text
    tsa_txt_425 = oNode.selectSingleNode("txt_425").Text

    tsa_txt_426 = oNode.selectSingleNode("txt_426").Text

    tsa_txt_427 = oNode.selectSingleNode("txt_427").Text
    tsa_txt_428 = oNode.selectSingleNode("txt_428").Text
    tsa_txt_429 = oNode.selectSingleNode("txt_429").Text
    tsa_txt_430 = oNode.selectSingleNode("txt_430").Text

    tsa_txt_431 = oNode.selectSingleNode("txt_431").Text
    tsa_txt_432 = oNode.selectSingleNode("txt_432").Text

    tsa_txt_433 = oNode.selectSingleNode("txt_433").Text
    tsa_txt_434 = oNode.selectSingleNode("txt_434").Text
    tsa_txt_435 = oNode.selectSingleNode("txt_435").Text
    tsa_txt_436 = oNode.selectSingleNode("txt_436").Text
    tsa_txt_437 = oNode.selectSingleNode("txt_437").Text
    tsa_txt_438 = oNode.selectSingleNode("txt_438").Text
    tsa_txt_439 = oNode.selectSingleNode("txt_439").Text

    tsa_txt_440 = oNode.selectSingleNode("txt_440").Text
    tsa_txt_441 = oNode.selectSingleNode("txt_441").Text
    tsa_txt_442 = oNode.selectSingleNode("txt_442").Text
    tsa_txt_443 = oNode.selectSingleNode("txt_443").Text
    tsa_txt_444 = oNode.selectSingleNode("txt_444").Text
    tsa_txt_445 = oNode.selectSingleNode("txt_445").Text
    tsa_txt_446 = oNode.selectSingleNode("txt_446").Text
    tsa_txt_447 = oNode.selectSingleNode("txt_447").Text
    tsa_txt_448 = oNode.selectSingleNode("txt_448").Text
    tsa_txt_449 = oNode.selectSingleNode("txt_449").Text

    tsa_txt_450 = oNode.selectSingleNode("txt_450").Text
    tsa_txt_451 = oNode.selectSingleNode("txt_451").Text
    tsa_txt_452 = oNode.selectSingleNode("txt_452").Text
    tsa_txt_453 = oNode.selectSingleNode("txt_453").Text
    tsa_txt_454 = oNode.selectSingleNode("txt_454").Text
    tsa_txt_455 = oNode.selectSingleNode("txt_455").Text
    tsa_txt_456 = oNode.selectSingleNode("txt_456").Text
    tsa_txt_457 = oNode.selectSingleNode("txt_457").Text
    tsa_txt_458 = oNode.selectSingleNode("txt_458").Text
    tsa_txt_459 = oNode.selectSingleNode("txt_459").Text

    tsa_txt_460 = oNode.selectSingleNode("txt_460").Text
    tsa_txt_461 = oNode.selectSingleNode("txt_461").Text
    tsa_txt_462 = oNode.selectSingleNode("txt_462").Text
    tsa_txt_463 = oNode.selectSingleNode("txt_463").Text
    tsa_txt_464 = oNode.selectSingleNode("txt_464").Text
    tsa_txt_465 = oNode.selectSingleNode("txt_465").Text
    tsa_txt_466 = oNode.selectSingleNode("txt_466").Text
    tsa_txt_467 = oNode.selectSingleNode("txt_467").Text
    tsa_txt_468 = oNode.selectSingleNode("txt_468").Text
    tsa_txt_469 = oNode.selectSingleNode("txt_469").Text

    tsa_txt_470 = oNode.selectSingleNode("txt_470").Text
    tsa_txt_471 = oNode.selectSingleNode("txt_471").Text
    tsa_txt_472 = oNode.selectSingleNode("txt_472").Text
    tsa_txt_473 = oNode.selectSingleNode("txt_473").Text
    tsa_txt_474 = oNode.selectSingleNode("txt_474").Text
    tsa_txt_475 = oNode.selectSingleNode("txt_475").Text
    tsa_txt_476 = oNode.selectSingleNode("txt_476").Text
    tsa_txt_477 = oNode.selectSingleNode("txt_477").Text
    tsa_txt_478 = oNode.selectSingleNode("txt_478").Text
    tsa_txt_479 = oNode.selectSingleNode("txt_479").Text

    tsa_txt_480 = oNode.selectSingleNode("txt_480").Text
    tsa_txt_481 = oNode.selectSingleNode("txt_481").Text
    tsa_txt_482 = oNode.selectSingleNode("txt_482").Text
    tsa_txt_483 = oNode.selectSingleNode("txt_483").Text
    tsa_txt_484 = oNode.selectSingleNode("txt_484").Text
    tsa_txt_485 = oNode.selectSingleNode("txt_485").Text
    tsa_txt_486 = oNode.selectSingleNode("txt_486").Text
    tsa_txt_487 = oNode.selectSingleNode("txt_487").Text
    tsa_txt_488 = oNode.selectSingleNode("txt_488").Text
    tsa_txt_489 = oNode.selectSingleNode("txt_489").Text

    tsa_txt_490 = oNode.selectSingleNode("txt_490").Text
    tsa_txt_491 = oNode.selectSingleNode("txt_491").Text
    tsa_txt_492 = oNode.selectSingleNode("txt_492").Text
    tsa_txt_493 = oNode.selectSingleNode("txt_493").Text
    tsa_txt_494 = oNode.selectSingleNode("txt_494").Text
    tsa_txt_495 = oNode.selectSingleNode("txt_495").Text
    tsa_txt_496 = oNode.selectSingleNode("txt_496").Text
    tsa_txt_497 = oNode.selectSingleNode("txt_497").Text
    tsa_txt_498 = oNode.selectSingleNode("txt_498").Text
    tsa_txt_499 = oNode.selectSingleNode("txt_499").Text

    tsa_txt_500 = oNode.selectSingleNode("txt_500").Text
    tsa_txt_501 = oNode.selectSingleNode("txt_501").Text
    tsa_txt_502 = oNode.selectSingleNode("txt_502").Text
    tsa_txt_503 = oNode.selectSingleNode("txt_503").Text
    tsa_txt_504 = oNode.selectSingleNode("txt_504").Text
    tsa_txt_505 = oNode.selectSingleNode("txt_505").Text
    tsa_txt_506 = oNode.selectSingleNode("txt_506").Text
    tsa_txt_507 = oNode.selectSingleNode("txt_507").Text
    tsa_txt_508 = oNode.selectSingleNode("txt_508").Text
    tsa_txt_509 = oNode.selectSingleNode("txt_509").Text
    
    tsa_txt_510 = oNode.selectSingleNode("txt_510").Text
    tsa_txt_511 = oNode.selectSingleNode("txt_511").Text
    tsa_txt_512 = oNode.selectSingleNode("txt_512").Text
    tsa_txt_513 = oNode.selectSingleNode("txt_513").Text
    tsa_txt_514 = oNode.selectSingleNode("txt_514").Text
    tsa_txt_515 = oNode.selectSingleNode("txt_515").Text
    tsa_txt_516 = oNode.selectSingleNode("txt_516").Text
    tsa_txt_517 = oNode.selectSingleNode("txt_517").Text
    tsa_txt_518 = oNode.selectSingleNode("txt_518").Text
    tsa_txt_519 = oNode.selectSingleNode("txt_519").Text

    tsa_txt_520 = oNode.selectSingleNode("txt_520").Text
    tsa_txt_521 = oNode.selectSingleNode("txt_521").Text
    tsa_txt_522 = oNode.selectSingleNode("txt_522").Text
    tsa_txt_523 = oNode.selectSingleNode("txt_523").Text
    tsa_txt_524 = oNode.selectSingleNode("txt_524").Text
    tsa_txt_525 = oNode.selectSingleNode("txt_525").Text
    tsa_txt_526 = oNode.selectSingleNode("txt_526").Text
    tsa_txt_527 = oNode.selectSingleNode("txt_527").Text
    tsa_txt_528 = oNode.selectSingleNode("txt_528").Text

    tsa_txt_529 = oNode.selectSingleNode("txt_529").Text

    tsa_txt_530 = oNode.selectSingleNode("txt_530").Text
    tsa_txt_531 = oNode.selectSingleNode("txt_531").Text
    tsa_txt_532 = oNode.selectSingleNode("txt_532").Text
    
next




%>