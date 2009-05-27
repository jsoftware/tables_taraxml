require 'arc/zip/zfiles'
require 'xml/xslt'
3 : 0 ''  
  NB. always use pcall version 
  NB. wd version crashes on big sheets
  if. 'Win'-:UNAME do.
    load 'xml/xslt/win_pcall'
  end.
)
