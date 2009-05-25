NB. first load wdooo.ijs
NB.
NB. =========================================================
NB. xslt using pcall
NB. error handling not yet implemented

xslt_win=: 4 : 0
p=. '' conew 'wdooo'
try.
  try.
    'xbase xtemp'=. olecreate__p 'MSXML2.DOMDocument.6.0'
    'ybase ytemp'=. olecreate__p 'MSXML2.DOMDocument.6.0'
  catch.
    try.
      'xbase xtemp'=. olecreate__p 'MSXML2.DOMDocument.4.0'
      'ybase ytemp'=. olecreate__p 'MSXML2.DOMDocument.4.0'
    catch.
      try.
        'xbase xtemp'=. olecreate__p 'MSXML2.DOMDocument.3.0'
        'ybase ytemp'=. olecreate__p 'MSXML2.DOMDocument.3.0'
      catch. smoutput 'MSXML v3 or 4 is required' throw. end.
    end.
  end.
  oleset__p xbase ; 'async' ; 0
  olemethod__p xbase ; 'loadXML' ; x
  oleset__p ybase ; 'async' ; 0
  olemethod__p ybase ; 'loadXML' ; y
  r=. olevalue__p VT_DISPATCH olemethod__p ybase ; 'transformNode' ; xbase
catch.
  smoutput 'error ',oleqer__p ''
end.
destroy__p ''
r
)


xslt_linux=: 4 : 0                                                                                              
  host=. 2!:0                                                                                                  
  tmpsty=. '/tmp/xlststy'                                                                                      
  tmpf=. '/tmp/xlstfile'                                                                                       
  (<tmpsty) 1!:2~ x                                                                                            
  (<tmpf) 1!:2~ y                                                                                              
  host 'xsltproc ', tmpsty, ' ', tmpf                                                                          
)

3 : 0 ''
  if. UNAME -: 'Win' do.
    xslt_z_ =: xslt_win_ptaraxml_
  elseif. UNAME -: 'Linux' do.
    xslt_z_ =: xslt_linux_ptaraxml_
  end.
''
)
