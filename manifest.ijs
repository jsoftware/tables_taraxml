NB. TaraXML manifest

CAPTION=: 'Platform independent system for reading OpenXML (Excel 2007 *.xlsx) files'

DESCRIPTION=: 0 : 0
The TaraXML addon reads files in Microsoft Excel's OpenXML format. For reading and writing older non-XML Excel formats see the Tara addon.

TaraXML depends on the arc/zip, xml/xslt and general/pcall addons.
Alternatively TaraXML can use unzip and xsltproc instead of the above 3 addons by setting TARAXMLCMDLINE_z_=:1
TaraXML was developed by Ric Sherlock and Bill Lam. 
)

VERSION=: '1.0.5'

RELEASE=: 'j602 j801'

DEPENDS=: 0 : 0
arc/zip/zfiles
general/pcall
xml/xslt
)

FILES=: 0 : 0
history.txt
manifest.ijs
taraxml.ijs
test/taraxmlread.ijs
test/test_taraxml.ijs
test/test.xlsx
)
