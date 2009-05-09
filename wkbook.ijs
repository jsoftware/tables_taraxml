NB. =========================================================
NB. Workbook object 
NB.  - stores properties of Workbook

coclass 'oxmlwkbook'
coinsert 'taraxml'

create=: 3 : 0
  FLN=: y   NB. Store filename as global
  NB. read sheet names and store as global
  SHEETNAMES=: getSheetNames zread 'xl/workbook.xml';FLN

  NB. read shared strings and store as global
  SHSTRINGS=: getStrings zread 'xl/sharedStrings.xml';FLN
)

destroy=: codestroy
