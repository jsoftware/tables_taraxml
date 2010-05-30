NB. =========================================================
NB. tables/taraxml
NB. Reading Excel 2007 OpenXML format (.xlsx) workbooks
NB.  retrieve contents of specified sheets

require 'arc/zip/zfiles xml/xslt'

3 : 0 ''  
  NB. always use pcall version on Windows
  NB. wd version crashes on big sheets
  if. 'Win'-:UNAME do.
    load 'xml/xslt/win_pcall'
  end.
)
