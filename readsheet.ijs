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
spelm=: >.--:2^32     NB. sparse element is lowest negative integer

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
    'nrows ncols'=. >: BR
    SHEET=: 1$. (nrows,ncols);0 1;spelm
    STRGMSK=: 1$. nrows,ncols
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
    VALS=: VALS, spelm&". y
  end.
)

endElement=: 3 : 0
  if. y -: 'row' do. NB. update SHEET matrix
    NB.STRGMSK=: 1 (<ROWIDX;I.ISSTRG)}STRGMSK
    NB.STRGMSK=: ISSTRG (<ROWIDX;COLIDX)}STRGMSK 
    STRGMSK=: 1 (<ROWIDX;ISSTRG#COLIDX)}STRGMSK
    NB.VALS=: (,SHSTRINGS {~ ISSTRG#VALS) (I.ISSTRG)} VALS
    NB. rcidx=. (ROWIDX;COLIDX) - &.> TL
    SHEET=: VALS (<ROWIDX;COLIDX)}SHEET
  end.
  S=: }:S
)

endDocument=: 3 : 0
  NB. convert to dense boxed matrix and amend strings.
  strgs=. $.inv SHEET
  strgs=. (-.STRGMSK)} strgs ,: #SHSTRINGS
  STRGMSK=: STRGMSK +. SHEET=spelm
  strgs=. strgs { SHSTRINGS,a:
  SHEET=: <"0 $.inv SHEET
  SHEET=: STRGMSK} SHEET,: strgs
  NB. SHEET=: ({.~ [: - TL + $) SHEET  NB. handle offset TL
  8!:0^:(GETSTRG"_) SHEET  NB. convert all to Strings if specified
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
