NB. TaraXML manifest

CAPTION=: 'Platform independent system for reading OpenXML (Excel 2007 *.xlsx) files'

DESCRIPTION=: 0 : 0
The TaraXML addon reads files in Microsoft Excel's OpenXML format. For reading and writing older non-XML Excel formats see the Tara addon.

TaraXML depends on a command line transformation utility.
Linux: xsltproc which should be available in various linux distro.
Windows: copy lib/* to ~bin folder
TaraXML was developed by Ric Sherlock and Bill Lam. 
)

VERSION=: '1.0.15'

RELEASE=: ''

FOLDER=: 'tables/taraxml'

FILES=: 0 : 0
history.txt
manifest.ijs
taraxml.ijs
test/taraxmlread.ijs
test/test.ijs
test/test.xlsx
)

FILESWIN64=: 0 : 0
lib/libexslt-0.dll
lib/libiconv-2.dll
lib/libwinpthread-1.dll
lib/libxml2-2.dll
lib/libxslt-1.dll
lib/xsltproc.exe
lib/zlib1.dll
)
