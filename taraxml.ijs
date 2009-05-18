NB. Reading Excel 2007 OpenXML format (.xlsx) workbooks
NB.  retrieve contents of specified sheets
NB. built from project: ~Addons/tables/taraxml/taraxml

require 'xml/sax arc/zip/zfiles'

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
  content=. nb&getSheet@([: zread ;&fln) each sheets
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
readxlxsheetnames=: getSheetNames@zread@('xl/workbook.xml'&;)


NB. =========================================================
NB. Workbook object 
NB.  - stores properties of Workbook

coclass 'oxmlwkbook'
coinsert 'ptaraxml'

create=: 3 : 0
  FLN=: y   NB. Store filename as global
  NB. read sheet names and store as global
  SHEETNAMES=: getSheetNames zread 'xl/workbook.xml';FLN

  NB. read shared strings and store as global
  SHSTRINGS=: getStrings zread 'xl/sharedStrings.xml';FLN
)

destroy=: codestroy


NB. =========================================================
NB. Reading workbook (xl/workbook.xml)
NB. reads list of worksheet names as boxed strings from workbook.xml

saxclass 'oxmlbook'

Note 'Testing'
process_oxmlbook_ fread jpath '~addons/tables/taraxml/test/workbook.xml'
)

Note 'XML hierachy of interest'
workbook                 NB. workbook
  sheets                 NB. worksheets
    sheet name= sheetID= NB. worksheet, name and ids attributes
)

startDocument=: 3 : 0
  SHTNAMES=: ''
  SHTIDS=: ''
  S=: ''
)

startElement=: 4 : 0
  e=. y
  S=: S,<e
  if. e -: 'sheet' do.  NB. append name of sheet
    SHTNAMES=: SHTNAMES , DEL ,~ x getAttribute 'name'
    SHTIDS=: SHTIDS , DEL ,~ x getAttribute 'sheetId'
  end.
  empty''
)

endElement=: 3 : 0
  S=: }:S
)

endDocument=: 3 : 0
  <;._2 SHTNAMES
)

process=: 3 : 0
  p=. '' conew >coname''
  startDocument__p ''
  parse__p y
  res=. endDocument__p''
  destroy__p''
  codestroy__p''
  res
)

getSheetNames_ptaraxml_=: process_oxmlbook_


NB. =========================================================
NB. Reading Shared Strings (xl/sharedStrings.xml)
NB. reads list of boxed strings from sharedStrings

saxclass 'oxmlstrings'

Note 'Testing'
process_oxmlstrings_ fread jpath '~addons/tables/taraxml/test/sharedStrings.xml'
)

Note 'XML hierachy of interest'
sst                 NB. sharedstrings
  si  xml:space     NB. string instance, if empty rather than not set then xml:space="preserve"
    t               NB. contains text for string instance
)

startDocument=: 3 : 0
  SHSTRNG=: ''
  S=: ''
)

startElement=: 4 : 0
  S=: S,<y
  empty''
)

characters=: 3 : 0
  s2=. _2{.S
  if. s2 -: ;:'si t' do.
    SHSTRNG=: SHSTRNG, y , DEL
  end.
)

endElement=: 3 : 0
  S=: }:S
)

endDocument=: 3 : 0
  <;._2 SHSTRNG
)

process=: 3 : 0
  p=. '' conew >coname''
  startDocument__p ''
  parse__p y
  res=. endDocument__p''
  destroy__p''
  codestroy__p''
  res
)

getStrings_ptaraxml_=: process_oxmlstrings_


NB. =========================================================
NB. Reading Worksheets (xl/worksheets/sheet?.xml)
NB. reads data from a worksheet to a boxed matrix

saxclass 'oxmlsheet'

Note 'Testing'
process_oxmlsheet_ fread jpath '~addons/tables/taraxml/test/sheet1.xml'
process_oxmlsheet_ fread jpath '~addons/tables/taraxml/test/sheet2.xml'
)

Note 'XML hierachy of interest'
worksheet             NB. contains worksheet info
  dimension ref=      NB. ref gives size of matrix
  sheetData           NB. contains sheet data
    row  r= spans=    NB. contains data for row 'r' (eg. ("1") over cols 'spans' (eg. "1:9")
      c  r= t=        NB. contains cell info for ref 'r' (eg. "B2") and type 't' (eg. "s" - string)
        v             NB. contains value for cell (if string then is index into si array in sharedStrings.xml)
)

caps=. a. {~ 65+i.26  NB. uppercase letters
nums=. a. {~ 48+i.10  NB. numerals

NB.*getColIdx v Calculates column index from A1 format
NB. 26 = getColIdx 'AA5'
getColIdx=: ([: <: 26 #. (' ',caps) i. _5 {. (' ',nums) -.~ ])"1
getRowIdx=: ([: <: (' ',caps) 0&".@-.~ ])"1

startDocument=: 3 : 0
  S=: ''
)

startElement=: 4 : 0
  S=: S,<y
  if. y -: 'dimension' do.  NB. initialize size of matrix for sheet
    'TL BR'=: (getRowIdx ,. getColIdx);._2 ':',~ x getAttribute 'ref'
    'nrows ncols'=. >: -/ BR,:TL
    SHEET=: (nrows,ncols)$ a:
  elseif. y -: 'row' do. NB.
    ROWIDX=: <: 0&". x getAttribute 'r'
    COLIDX=: ''
    VALS=: ''
    ISSTRG=: ''
  elseif. y -: (,'c') do.
    cellref=. x getAttribute 'r'
    ISSTRG=: ISSTRG, (,'s') -: x getAttribute 't'
    COLIDX=: COLIDX, getColIdx cellref
  end.
  empty''
)

characters=: 3 : 0
  s2=. _2{.S
  if. s2 -: ;:'c v' do.
    VALS=: VALS, < _9999&". y
  end.
)

endElement=: 3 : 0
  if. y -: 'row' do. NB. update SHEET matrix
    VALS=: (,SHSTRINGS {~ ISSTRG#VALS) (I.ISSTRG)} VALS
    rcidx=. (ROWIDX;COLIDX) - &.> TL
    SHEET=: VALS (<rcidx)}SHEET
  end.
  S=: }:S
)

endDocument=: 3 : 0
  SHEET=: ({.~ [: - TL + $) SHEET
  8!:0^:(GETSTRG"_) SHEET
)

process=: 4 : 0
  p=. '' conew >coname''
  coinsert__p x
  startDocument__p ''
  parse__p y
  res=. endDocument__p''
  destroy__p''
  codestroy__p''
  res
)

getSheet_ptaraxml_=: process_oxmlsheet_


NB. =========================================================
NB. Export to z locale
readxlxsheets_z_=: readxlxsheets_ptaraxml_
readxlxsheetnames_z_=: readxlxsheetnames_ptaraxml_

