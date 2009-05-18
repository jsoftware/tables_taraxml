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
