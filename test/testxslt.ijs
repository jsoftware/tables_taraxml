NB. =========================================================
NB. Test the XSLT transformations on the individual XML documents in a workbook.

load 'tables/taraxml xml/xslt'

WKBOOKXML=: jpath '~Addons/tables/taraxml/test/workbook.xml'
SHSTRGXML=: jpath '~Addons/tables/taraxml/test/sharedStrings.xml'
SHEET1XML=: jpath '~Addons/tables/taraxml/test/sheet1.xml'
SHEET2XML=: jpath '~Addons/tables/taraxml/test/sheet2.xml'

Note 'Testing'
 <;._2 WKBOOKSTY_oxmlwkbook_ xslt fread WKBOOKXML
 <;._2 SHSTRGSTY_oxmlwkbook_ xslt fread SHSTRGXML
 _3]\ <;._2 SHEETSTY_oxmlwkbook_ xslt fread SHEET1XML
 _3]\ <;._2 SHEETSTY_oxmlwkbook_ xslt fread SHEET2XML
)
