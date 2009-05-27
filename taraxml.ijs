NB. Reading Excel 2007 OpenXML format (.xlsx) workbooks
NB.  retrieve contents of specified sheets
NB. built from project: ~Addons/tables/taraxml/taraxml

require 'arc/zip/zfiles'
NB. require 'xml/xslt'

NB. =========================================================
NB. Workbook object 
NB.  - methods/properties for Workbook

coclass 'oxmlwkbook'
coinsert 'ptaraxml'

caps=. a. {~ 65+i.26  NB. uppercase letters
nums=. a. {~ 48+i.10  NB. numerals
errnum=: >.--:2^32     NB. error is lowest negative integer
RMSTRG=: 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

NB. ---------------------------------------------------------
NB. Methods for oxmlwkbook

create=: 3 : 0
  FLN=: y   NB. Store filename as global
  NB. read sheet names and store as global
  SHEETNAMES=: getSheetNames zread 'xl/workbook.xml';FLN

  NB. read shared strings and store as global
  SHSTRINGS=: getStrings zread 'xl/sharedStrings.xml';FLN
)

destroy=: codestroy

NB. getColIdx v Calculates column index from A1 format
NB. 26 = getColIdx 'AA5'
getColIdx=: ([: <: 26 #. (' ',caps) i. _5 {. (' ',nums) -.~ ])"1

NB. getRowIdx v Calculates row index from A1 format
getRowIdx=: ([: <: (' ',caps) 0&".@-.~ ])"1

NB. getSheetNames v Reads sheet names in OpenXML workbook
NB. result: list of boxed sheet names in workbook
NB. y is: literal list of XML from workbook.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default WKBOOKSTY
getSheetNames=: 3 : 0
  WKBOOKSTY getSheetNames y
:
  res=. x xslt (RMSTRG;'') stringreplace y
  {:"1 ] _2]\ <;._2 res 
)


NB. getStrings v Reads shared strings in OpenXML workbook
NB. result: list of boxed shared strings in workbook
NB. y is: literal list of XML from sharedStrings.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHSTRGSTY
getStrings=: 3 : 0
 SHSTRGSTY getStrings y
 :
  res=. <;._2 x xslt (RMSTRG;'') stringreplace y
)

NB. getSheet v Reads sheet contents from a sheet in an OpenXML workbook
NB. result: table of boxed contents from worksheet
NB. y is: literal list of XML from sheet?.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHEETSTY
getSheet=: 3 : 0 
  SHEETSTY getSheet y
 :
  res=. _3]\<;._2 x xslt (RMSTRG;'') stringreplace y
  cellidx=. > 0 {"1  res
  cellidx=. (getRowIdx ,. getColIdx) cellidx
  strgmsk=. (<,'s') = 1 {"1  res
  cellval=. 2 {"1  res
  br=. >: >./cellidx

  strgs=. (SHSTRINGS,a:) {~  (#SHSTRINGS)&".> strgmsk#cellval
  validx=. I.-.strgmsk
  if. -. GETSTRG do.
    cellval=. (errnum&". &.> validx{cellval) validx} cellval
  end.
  cellval=. strgs (I. strgmsk)} cellval
  cellval=. cellval (<"1 cellidx)} br$a:
)


NB. ---------------------------------------------------------
NB. XSLT for transforming XML files in OpenXML workbook

Note 'XML hierachy of interest'
workbook                 NB. workbook
  sheets                 NB. worksheets
    sheet name= sheetID= NB. worksheet, name and ids attributes
)

WKBOOKSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="sheet">
        <x:value-of select="@sheetId" /><x:text>&#127;</x:text>
        <x:value-of select="@name" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest'
sst                 NB. sharedstrings
  si  xml:space     NB. string instance, if empty rather than not set then xml:space="preserve"
    t               NB. contains text for string instance
)

SHSTRGSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="t">
        <x:value-of select="." /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest'
worksheet             NB. contains worksheet info
  dimension ref=      NB. ref gives size of matrix
  sheetData           NB. contains sheet data
    row  r= spans=    NB. contains data for row 'r' (eg. ("1") over cols 'spans' (eg. "1:9")
      c  r= t=        NB. contains cell info for ref 'r' (eg. "B2") and type 't' (eg. "s" - string)
        v             NB. contains value for cell (if string then is index into si array in sharedStrings.xml)
)

SHEETSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="c">
        <x:value-of select="@r" /><x:text>&#127;</x:text>
        <x:value-of select="@t" /><x:text>&#127;</x:text>
        <x:value-of select="v" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)


Note 'Testing'
WKBOOKXML=: jpath '~Addons/tables/taraxml/test/workbook.xml'
_2]\ <;._2 WKBOOKSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread WKBOOKXML
_2]\ <;._2 WKBOOKSTY xslt (RMSTRG;'') stringreplace fread WKBOOKXML

SHSTRGXML=: jpath '~Addons/tables/taraxml/test/sharedStrings.xml'
<;._2 SHSTRGSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread SHSTRGXML
<;._2 SHSTRGSTY xslt (RMSTRG;'') stringreplace fread SHSTRGXML

Note 'Testing'
SHEET1XML=: jpath '~Addons/tables/taraxml/test/sheet1.xml'
SHEET2XML=: jpath '~Addons/tables/taraxml/test/sheet2.xml'
_3]\ <;._2 SHEETSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread SHEET1XML
_3]\ <;._2 SHEETSTY xslt (RMSTRG;'') stringreplace fread SHEET1XML
_3]\ <;._2 SHEETSTY xslt (RMSTRG;'') stringreplace fread SHEET2XML
)


NB. =========================================================
NB. Define User Interface verbs
coclass 'ptaraxml'

NB.*readxlxsheets v Reads one or more sheets from an Excel file
NB. returns: 2-column matrix with a row for each sheet
NB.       0{"1 boxed list of sheet names
NB.       1{"1 boxed list of boxed matrices of sheet contents
NB. y is: 1 or 2-item boxed list:
NB.       0{ filename of Excel workbook
NB.       1{ [default 0] optional switch to return all cells contents as strings
NB. x is: one of [default is 0]:
NB.       * numeric list of indicies of sheets to return
NB.       * boxed list of sheet names to return
NB.       * '' - return all sheets
NB. EG:   0 readxlsheets 'test.xls'
NB. reads Excel Version 2007, OpenXML (.xlsx)
readxlxsheets=: 3 : 0
  0 readxlxsheets y
:
try.
  'fln strng'=. 2{.!.(<0) boxopen y
  shts=. boxopen x
  (msg=. 'file not found') assert fexist fln

  nb=. fln conew 'oxmlwkbook'
  GETSTRG__nb=: strng
  shtnames=. SHEETNAMES__nb
  if.     a: -: shts               do. shtidx=. i. #shtnames NB. x is ''
  elseif. *./ 1 4 e.~ 3!:0 &> shts do. shtidx=. > shts       NB. x is int list
  elseif. do. shtidx=. shts i.&(tolower&.>"_)~ shtnames  NB. case insensitive
  end.
  (msg=. 'worksheet not found') assert shtidx < #shtnames
  shts=. shtidx { shtnames
  sheets=. ,&'.xml' &.> 'xl/worksheets/sheet'&, &.> 8!:0 >: shtidx

  msg=. 'error reading worksheet'
  content=. getSheet__nb@([: zread ;&fln) each sheets
  destroy__nb''
  shts,.content
catch.
  coerase <'nb'
  smoutput 'readxlxsheets: ',msg
end.
)

NB.*readxlxsheetnames v Reads sheet names from Excel workbook
NB. returns: boxed list of sheet names
NB. y is: Excel file name
NB. eg: readxlsheetnames 'test.xls'
NB. read Excel Version 2007
readxlxsheetnames=: getSheetNames_oxmlwkbook_@zread@('xl/workbook.xml'&;)


NB. =========================================================
NB. Export to z locale
readxlxsheets_z_=: readxlxsheets_ptaraxml_
readxlxsheetnames_z_=: readxlxsheetnames_ptaraxml_


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


NB. first load wdooo.ijs
NB.
NB. =========================================================


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
  r=. olevalue__p VT_DISPATCH__p olemethod__p ybase ; 'transformNode' ; xbase
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
    xslt_z_ =: xslt_win2_ptaraxml_
  elseif. UNAME -: 'Linux' do.
    xslt_z_ =: xslt_linux_ptaraxml_
  end.
''
)

