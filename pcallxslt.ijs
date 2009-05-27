NB. xslt using pcall
require 'general/pcall/disp'

parseError=: 3 : 0
  pe=. disp 'parseError' get__y''
  if. 0~: 0".c=. 'errorCode' get__pe'' do. 
    line=. 'line' get__pe''
    pos=. 'linePos' get__pe''
    src=. 'srcText' get__pe''
    t=. 'Error ',c,' at ',line,',',pos
    t=. t,LF, 'reason' get__pe
    destroy__pe''
    if. #src do.
      t=. t,LF,src
      t=. t,LF,(}.(0".pos)#' '),'^'
    end.
    1[smoutput t
  else.
   destroy__pe''  
   0 end.
)

xslt_win2=: 4 : 0
  try.
    try.
      qx=. disp 'MSXML2.DOMDocument.6.0'
      qy=. disp 'MSXML2.DOMDocument.6.0'
    catch.
      try.
        qx=. disp 'MSXML2.DOMDocument.4.0'
        qy=. disp 'MSXML2.DOMDocument.4.0'
      catch.
        try.
          qx=. disp 'MSXML2.DOMDocument.3.0'
          qy=. disp 'MSXML2.DOMDocument.3.0'
        catch. smoutput 'MSXML v3, 4 or 6 is required' throw. end.
      end.
    end.
    'async' put__qx 0
    'loadXML' do__qx x
    NB. if. parseError qx do. throw. end.
    'async' put__qy 0
    'loadXML' do__qy y
    NB. if. parseError qy do. throw. end.
    try.
      r=. 'transformNode' do__qy <<P__qx
NB.       destroy__qx''
NB.       destroy__qy''
    catch. smoutput 'error qer'  NB. what should go here?
    throw. end.
  catcht. r=. '' end.
  r
)
