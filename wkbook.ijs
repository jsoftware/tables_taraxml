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
  cellval=. errnum&".> 2 {"1  res
  br=. >: >./cellidx

  strgs=. SHSTRINGS {~ strgmsk#cellval
  cellval=. strgs (I. strgmsk)} <"0 cellval
  cellval=. cellval (<"1 cellidx)} br$a:

  8!:0^:(GETSTRG"_) cellval  NB. convert all to Strings if specified
)
