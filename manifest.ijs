NB. TaraXML manifest

CAPTION=: 'Platform independent system for reading OpenXML (Excel 2007 *.xlsx) files'

DESCRIPTION=: 0 : 0
The TaraXML addon reads files in Microsoft Excel's OpenXML format. For reading and writing older non-XML Excel formats see the Tara addon.

TaraXML depends on a command line transformation utility.
Linux: xsltproc which should be available in various linux distro.
Windows: msxsl.exe available from http://www.microsoft.com/en-us/download/details.aspx?id=21714
  msxsl.exe depends on msxml4 (not msxml6!) http://www.microsoft.com/en-us/download/details.aspx?id=19662
TaraXML was developed by Ric Sherlock and Bill Lam. 
)

VERSION=: '1.0.10'

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
