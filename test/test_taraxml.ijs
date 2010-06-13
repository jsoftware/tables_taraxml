NB. Tests for taraxml

Note 'To run all test scripts:'
  load 'tables/taraxml'
  load 'tables/taraxml/test/test_taraxml'
)

loc=. 3 : '> (4!:4 <''y'') { 4!:3 $0'
PATH=. getpath_j_ jpathsep loc''

NB. -------------------------------------------------------
NB. scripts for testing

load PATH,'taraxmlread.ijs'
