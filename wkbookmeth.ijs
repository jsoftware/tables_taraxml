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
