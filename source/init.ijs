NB. =========================================================
NB. tables/taraxml
NB. Reading Excel 2007 OpenXML format (.xlsx) workbooks
NB.  retrieve contents of specified sheets

TARAXMLCMDLINE_z_=: 1 NB. (TARAXMLCMDLINE_z_"_)^:(0=4!:0<'TARAXMLCMDLINE_z_') 0~:FHS

require^:(-.TARAXMLCMDLINE) 'arc/zip/zfiles xml/xslt'

NB. always use pcall version on Windows
NB. wd version crashes on big sheets
NB. require^:(IFWIN>TARAXMLCMDLINE) 'xml/xslt/win_pcall'
