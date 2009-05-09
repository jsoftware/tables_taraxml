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
    SHSTRNG=: SHSTRNG, < y
  end.
)

endElement=: 3 : 0
  S=: }:S
)

endDocument=: 3 : 0
  SHSTRNG
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

getStrings_taraxml_=: process_oxmlstrings_
