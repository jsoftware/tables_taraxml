NB. ---------------------------------------------------------
NB. XSLT for transforming XML files in OpenXML workbook

Note 'XML hierachy of interest'
workbook                 NB. workbook
  sheets                 NB. worksheets
    sheet name= sheetID= NB. worksheet, name and ids attributes
)

WKBOOKSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="sheet">
        <x:value-of select="@sheetId" /><x:text>&#127;</x:text>
        <x:value-of select="@name" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest'
sst                 NB. sharedstrings
  si  xml:space     NB. string instance, if empty rather than not set then xml:space="preserve"
    t               NB. contains text for string instance
)

SHSTRGSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="t">
        <x:value-of select="." /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest'
worksheet             NB. contains worksheet info
  dimension ref=      NB. ref gives size of matrix
  sheetData           NB. contains sheet data
    row  r= spans=    NB. contains data for row 'r' (eg. ("1") over cols 'spans' (eg. "1:9")
      c  r= t=        NB. contains cell info for ref 'r' (eg. "B2") and type 't' (eg. "s" - string)
        v             NB. contains value for cell (if string then is index into si array in sharedStrings.xml)
)

SHEETSTY=: 0 : 0
<x:stylesheet xmlns:x="http://www.w3.org/1999/XSL/Transform" version="1.0">
        <x:output method="text"/>
    <x:template match="c">
        <x:value-of select="@r" /><x:text>&#127;</x:text>
        <x:value-of select="@t" /><x:text>&#127;</x:text>
        <x:value-of select="v" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)


Note 'Testing'
WKBOOKXML=: jpath '~Addons/tables/taraxml/test/workbook.xml'
_2]\ <;._2 WKBOOKSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread WKBOOKXML
_2]\ <;._2 WKBOOKSTY xslt (RMSTRG;'') stringreplace fread WKBOOKXML

SHSTRGXML=: jpath '~Addons/tables/taraxml/test/sharedStrings.xml'
<;._2 SHSTRGSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread SHSTRGXML
<;._2 SHSTRGSTY xslt (RMSTRG;'') stringreplace fread SHSTRGXML

Note 'Testing'
SHEET1XML=: jpath '~Addons/tables/taraxml/test/sheet1.xml'
SHEET2XML=: jpath '~Addons/tables/taraxml/test/sheet2.xml'
_3]\ <;._2 SHEETSTY xslt (' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';'') stringreplace fread SHEET1XML
_3]\ <;._2 SHEETSTY xslt (RMSTRG;'') stringreplace fread SHEET1XML
_3]\ <;._2 SHEETSTY xslt (RMSTRG;'') stringreplace fread SHEET2XML
)
