caps=. a. {~ 65+i.26  NB. uppercase letters
nums=. a. {~ 48+i.10  NB. numerals
errnum=: >.--:2^32     NB. error is lowest negative integer

NB. ---------------------------------------------------------
NB. Methods for oxmlwkbook

create=: 3 : 0
  FLN=: y   NB. Store filename as global
  NB. read sheet names and store as global
  SHEETNAMES=: getSheetNames (zreadproc`zread_zfiles_@.(-.TARAXMLCMDLINE)) 'xl/workbook.xml';FLN

  NB. read shared strings and store as global
  SHSTRINGS=: getStrings (zreadproc`zread_zfiles_@.(-.TARAXMLCMDLINE)) 'xl/sharedStrings.xml';FLN
)

destroy=: codestroy

hostcmd=: ([: 2!:0 '('"_ , ] , ' || true)'"_)`spawn_jtask_@.IFWIN

NB. used if TARAXMLCMDLINE
NB. require command line utility unzip
NB. use ~temp
zreadproc=: 3 : 0
'FN ZN'=. y
if. 1~:ftype ZN do. 1 return. end.
mkdir_j_ tmp
hostcmd unzipcmd, ' -o -qq "',(winpathsep^:IFWIN ZN),'" -d "',(winpathsep^:IFWIN tmp),'" "',(winpathsep^:IFWIN FN),'"', IFUNIX#' 2>/dev/null'
r=. fread f=. tmp,'/',FN
ferase ::0: f
r
)

NB. used if TARAXMLCMDLINE
NB. require command line utility xsltproc
NB. use ~temp
xsltproc=: 4 : 0
x fwrite <style=. tmp,'/xmlstyle'
y fwrite <file=. tmp,'/xmlfile'
if. IFWIN do.
  if. 1~:ftype f=. jpath '~addons/tables/taraxml/lib/xsltproc.exe' do.
    if. 1~:ftype f=. jpath '~bin/xsltproc.exe' do.
      f=. 'xsltproc.exe'
    end.
  end.
  a=. hostcmd '"',f,'" "',(winpathsep^:IFWIN style),'" "',(winpathsep^:IFWIN file),'"'
else.
  a=. hostcmd 'xsltproc "',(winpathsep^:IFWIN style),'" "',(winpathsep^:IFWIN file),'"', IFUNIX#' 2>/dev/null'
end.
NB. check BOM
if. (255 254{a.)-:2{.a do.
  a=. 8 u: 6 u: 2}.a
end.
ferase ::0: style
ferase ::0: file
assert. 0-.@-: a [ 'xsltproc error'
a
)

NB. getColIdx v Calculates column index from A1 format
NB. 26 = getColIdx 'AA5'
getColIdx=: ([: <: 26 #. (' ',caps) i. _5 {. (' ',nums) -.~ ])"1

NB. getRowIdx v Calculates row index from A1 format
getRowIdx=: ([: <: (' ',caps) 0&".@-.~ ])"1

NB. getSheetNames v Reads sheet names in OpenXML workbook
NB. result: list of boxed sheet names in workbook
NB. y is: literal list of XML from workbook.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default WKBOOKSTY
getSheetNames=: [: <;._2 WKBOOKSTY&(xsltproc`xslt_pxslt_@.(-.TARAXMLCMDLINE))

NB. getStrings v Reads shared strings in OpenXML workbook
NB. result: list of boxed shared strings in workbook
NB. y is: literal list of XML from sharedStrings.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHSTRGSTY
getStrings=: [: <;._2 SHSTRGSTY&(xsltproc`xslt_pxslt_@.(-.TARAXMLCMDLINE))

NB. getSheet v Reads sheet contents from a sheet in an OpenXML workbook
NB. result: table of boxed contents from worksheet
NB. y is: literal list of XML from sheet?.xml in OpenXML workbook
NB. x is: Optional XSLT to use to transform XML. Default SHEETSTY
getSheet=: 3 : 0
  SHEETSTY getSheet y
 :
  res=. _3]\<;._2 x xsltproc`xslt_pxslt_@.(-.TARAXMLCMDLINE) y
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
