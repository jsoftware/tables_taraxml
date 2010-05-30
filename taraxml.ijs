NB. 
NB. built from project: ~Addons/tables/taraxml/taraxml
NB. =========================================================
NB. tables/taraxml
NB. Reading Excel 2007 OpenXML format (.xlsx) workbooks
NB.  retrieve contents of specified sheets


require 'arc/zip/zfiles xml/xslt'

3 : 0 ''  
  NB. always use pcall version on Windows
  NB. wd version crashes on big sheets
  if. 'Win'-:UNAME do.
    load 'xml/xslt/win_pcall'
  end.
)

NB. =========================================================
NB. Workbook object 
NB.  - methods/properties for a Workbook

coclass 'oxmlwkbook'
coinsert 'ptaraxml'


NB. ---------------------------------------------------------
NB. XSLT for transforming XML files in OpenXML workbook

Note 'XML hierachy of interest for workbook.xml'
workbook                 NB. workbook
  sheets                 NB. worksheets
    sheet name= sheetID= NB. worksheet, name and ids attributes
)

NB. Retrieve worksheet names from xl/workbook.xml
WKBOOKSTY=: 0 : 0
<x:stylesheet  version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
    <x:output method="text" encoding="UTF-8" />
    <x:template match="t:sheet">
        <x:value-of select="@name" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest for sharedStrings.xml'
sst                 NB. sharedstrings
  si  xml:space     NB. string instance, if empty rather than not set then xml:space="preserve"
    t               NB. contains text for string instance
)

NB. Retrieve shared strings from xl/sharedStrings.xml
SHSTRGSTY=: 0 : 0
<x:stylesheet version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
    <x:output method="text" encoding="UTF-8"/>
    <x:template match="t:t">
        <x:value-of select="." /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest for sheet?.xml'
worksheet             NB. contains worksheet info
  dimension ref=      NB. ref gives size of matrix
  sheetData           NB. contains sheet data
    row  r= spans=    NB. contains data for row 'r' (eg. ("1") over cols 'spans' (eg. "1:9")
      c  r= t=        NB. contains cell info for ref 'r' (eg. "B2") and type 't' (eg. "s" - string)
        v             NB. contains value for cell (if string then is index into si array in sharedStrings.xml)
)

NB. Retrieve worksheet contents from xl/worksheets/sheet?.xml
SHEETSTY=: 0 : 0
<x:stylesheet version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
        <x:output method="text"/>
    <x:template match="t:c">
        <x:value-of select="@r" /><x:text>&#127;</x:text>
        <x:value-of select="@t" /><x:text>&#127;</x:text>
        <x:value-of select="t:v" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)


caps=. a. {~ 65+i.26  NB. uppercase letters
nums=. a. {~ 48+i.10  NB. numerals
errnum=: >.--:2^32     NB. error is lowest negative integer

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
getSheetNames=: [: <;._2 WKBOOKSTY&xslt

NB. getStrings v Reads shared strings in OpenXML workbook
NB. result: list of boxed shared strings in workbook
NB. y is: literal list of XML from sharedStrings.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHSTRGSTY
getStrings=: [: <;._2 SHSTRGSTY&xslt

NB. getSheet v Reads sheet contents from a sheet in an OpenXML workbook
NB. result: table of boxed contents from worksheet
NB. y is: literal list of XML from sheet?.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHEETSTY
getSheet=: 3 : 0 
  SHEETSTY getSheet y
 :
  res=. _3]\<;._2 x xslt y
  cellidx=. > 0 {"1  res
  strgmsk=. (<,'s') = 1 {"1  res
  cellval=. 2 {"1  res
  erase 'res'
  cellidx=. (getRowIdx ,. getColIdx) cellidx
  br=. >: >./cellidx

  strgs=. (SHSTRINGS,a:) {~  (#SHSTRINGS)&".> strgmsk#cellval
  validx=. I.-.strgmsk
  if. -. GETSTRG do.
    cellval=. (errnum&". &.> validx{cellval) validx} cellval
  end.
  cellval=. strgs (I. strgmsk)} cellval
  cellval=. cellval (<"1 cellidx)} br$a:
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
  elseif. do. shtidx=. shts i.&(tolower&.>"_)~ shtnames      NB. case insensitive
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

