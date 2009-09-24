NB. ---------------------------------------------------------
NB. XSLT for transforming XML files in OpenXML workbook

Note 'XML hierachy of interest for workbook.xml'
workbook                 NB. workbook
  sheets                 NB. worksheets
    sheet name= sheetID= NB. worksheet, name and ids attributes
)

NB. Retrieve worksheet names from xl/workbook.xml
WKBOOKSTY=: 0 : 0
<x:stylesheet  version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
    <x:output method="text" encoding="UTF-8" />
    <x:template match="t:sheet">
        <x:value-of select="@name" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest for sharedStrings.xml'
sst                 NB. sharedstrings
  si  xml:space     NB. string instance, if empty rather than not set then xml:space="preserve"
    t               NB. contains text for string instance
)

NB. Retrieve shared strings from xl/sharedStrings.xml
SHSTRGSTY=: 0 : 0
<x:stylesheet version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
    <x:output method="text" encoding="UTF-8"/>
    <x:template match="t:t">
        <x:value-of select="." /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)

Note 'XML hierachy of interest for sheet?.xml'
worksheet             NB. contains worksheet info
  dimension ref=      NB. ref gives size of matrix
  sheetData           NB. contains sheet data
    row  r= spans=    NB. contains data for row 'r' (eg. ("1") over cols 'spans' (eg. "1:9")
      c  r= t=        NB. contains cell info for ref 'r' (eg. "B2") and type 't' (eg. "s" - string)
        v             NB. contains value for cell (if string then is index into si array in sharedStrings.xml)
)

NB. Retrieve worksheet contents from xl/worksheets/sheet?.xml
SHEETSTY=: 0 : 0
<x:stylesheet version="1.0"
   xmlns:x="http://www.w3.org/1999/XSL/Transform"
   xmlns:t="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   exclude-result-prefixes="x t"
>
        <x:output method="text"/>
    <x:template match="t:c">
        <x:value-of select="@r" /><x:text>&#127;</x:text>
        <x:value-of select="@t" /><x:text>&#127;</x:text>
        <x:value-of select="t:v" /><x:text>&#127;</x:text>
    </x:template>
    <x:template match="text()" />
</x:stylesheet>
)
