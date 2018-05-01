NB. build

writesource_jp_ '~Addons/tables/taraxml/source';'~Addons/tables/taraxml/taraxml.ijs'

(jpath '~addons/tables/taraxml/taraxml.ijs') (fcopynew ::0:) jpath '~Addons/tables/taraxml/taraxml.ijs'

f=. 3 : 0
(jpath '~Addons/tables/taraxml/',y) fcopynew jpath '~Addons/tables/taraxml/source/',y
(jpath '~addons/tables/taraxml/',y) (fcopynew ::0:) jpath '~Addons/tables/taraxml/source/',y
)

mkdir_j_ jpath '~addons/tables/taraxml/test'
f 'manifest.ijs'
f 'history.txt'
f 'test/taraxmlread.ijs'
f 'test/test.ijs'
f 'test/test.xlsx'
