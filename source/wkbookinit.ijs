NB. =========================================================
NB. Workbook object 
NB.  - methods/properties for a Workbook

coclass 'oxmlwkbook'
coinsert 'ptaraxml'

3 : 0''
if. IFWIN do.
  if. 1=ftype '~tools/zip/unzip.exe' do.
    unzipcmd=: winpathsep jpath '~tools/zip/unzip.exe'
  else.
    unzipcmd=: 'unzip.exe'
  end.
else.
  unzipcmd=: 'unzip'
end.
EMPTY
)
