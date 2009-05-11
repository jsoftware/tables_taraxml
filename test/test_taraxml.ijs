NB. Tests for taraxml

Note 'To run all tests:'
  load 'tables/taraxml'
  load 'tables/taraxml/test/test_taraxml'
)

loc=. 3 : '> (4!:4 <''y'') { 4!:3 $0'
PATH=. getpath_j_ loc''

NB. -------------------------------------------------------
NB. scripts for testing

load PATH,'taraxmltest.ijs'
