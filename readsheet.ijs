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

getSheet_taraxml_=: process_oxmlsheet_
