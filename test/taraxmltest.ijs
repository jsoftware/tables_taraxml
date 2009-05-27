NB. ---------------------------------------------------------
NB. Tests for readxlxsheets
Note 'To run all tests:'
  load 'tables/taraxml'
  load 'tables/taraxml/test/taraxmltest'
)

loc=. 3 : '> (4!:4 <''y'') { 4!:3 $0'
PATH=. getpath_j_ loc''

NB. Test data
x3=: <"0 i.12 20            NB. boxed int array
x4=: 4 2$'abcd';'kdisd';'eiij asj' NB. boxed char array
x4a=: 4 2$'abcd';'';'eiij asj' NB. boxed char array with empty
x5=: 4 2$'abcd';54;(ucp^:IFWIN32 '鳴箏金粟柱');2;4.4 NB. boxed mixed array
x6=: (,:-) 5 50000 0.005 0.0000005 500000000 500000000.5  NB. numeric format
x7=: 'data 1';'data 2'      NB. boxed char vector
x9=: <"0 ]15.6 12.9 54.33   NB. boxed flt vector
x12=: 'Box Int array';<x3   NB. With sheetnames
x13=: 'Box Chr array';<x4
x14=: 'Box Mix array';<x5
x20=: 'Offset';<(,.5#<''),(4#<''),"1 x3 NB. offset topleft corner by 5 rows and 4 cols
char=: ('per Angusta';'ad Augusta'),~;:'Lorem ipsum dolor sit consectetur'
idx=: 2 2;2 3;2 4;2 5;3 2;3 3;3 5;4 2;4 4;4 5;5 2;5 3;5 4;5 5
x21=: 'Diff Mix array';<(14$char) idx} <"0 i.8 9
x22=: 'Number Fmt';< <"0 x6
x23=: (x12,x13,x14,x20,x21,:x22)

fnme=: PATH,'test.xlsx'

NB.Tests
test2=: 3 : 0
  locs=. conames 1
  assert x12 -: readxlxsheets fnme
  assert x13 -: 'Box Chr array' readxlxsheets fnme
  assert x14 -: 2 readxlxsheets fnme
  assert (x13,:x14) -: 1 2 readxlxsheets fnme
  assert x20 -: 'Offset' readxlxsheets fnme
  assert (x21,:x20) -: ('DiFF mix array';'OFFSET') readxlxsheets fnme
  assert x22 -: 'Number Fmt' readxlxsheets fnme
  assert x23 -: '' readxlxsheets fnme
  assert ('Box Mix array';<8!:0 (0j16": &.> (<2;0){ x5) (<2;0)} x5) -: 2 readxlxsheets fnme;<1
  assert locs -: conames 1  NB. check all objects cleaned up
  'taraxmltest passed'
)

smoutput test2 ''
