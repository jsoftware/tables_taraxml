NB. TaraXML manifest

CAPTION=: 'Platform independent system for reading OpenXML (Excel 2007 *.xlsx) files'

DESCRIPTION=: 0 : 0
The TaraXML addon reads files in Microsoft Excel's OpenXML format. For reading and writing older non-XML Excel formats see the Tara addon.

TaraXML depends on a command line transformation utility.
Linux: xsltproc which should be available in various linux distro.
Windows: copy lib/* to ~bin folder
TaraXML was developed by Ric Sherlock and Bill Lam. 
)

VERSION=: '1.0.11'

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

FILEWIN64=: 0 : 0
libexslt-0.dll
libiconv-2.dll
libwinpthread-1.dll
libxml2-2.dll
libxslt-1.dll
xsltproc.exe
zlib1.dll
)
