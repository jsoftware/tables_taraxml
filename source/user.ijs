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
