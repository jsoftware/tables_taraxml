NB. =========================================================
require 'general/pcall general/pcall/ole32'

NB. =========================================================
NB. error constant

(18!:55 :: 0:) <'olecomerrorh'
coclass 'olecomerrorh'
coinsert 'pole32'

DFH=: 3 : 0
if. '0x'-:2{.y=. }:^:('L'={:y) y do.
  d=. 0
  for_nib. ('0123456789abcdef'&i.) tolower 2}.y do.
    d=. nib (23 b.) 4 (33 b.) d
  end.
else.
  0&". y
end.
)

cheaderconst=: ''&$: : (4 : 0)
if. #x do.
  ({.x)=: {.("1) y
  ({:x)=: DFH&> {:("1) y
end.
,(>{.("1) y),("1) ' =: ',("1) (":@DFH&> {:("1) y) ,("1) LF
)

olecomerrmsg=: 3 : 0
if. y e. OLECOMERRVAL do. ; (,&' ')&.> OLECOMERRCODE #~ OLECOMERRVAL = y else. 'Other error: ', ":y end.
)

(0!:100) ('OLECOMERRCODE' ; 'OLECOMERRVAL') cheaderconst (<;._2)@(,&' ') ;._2 (0 : 0)
S_OK 0
CO_E_ALREADYINITIALIZED 0x800401F1
CO_E_APPDIDNTREG 0x800401FE
CO_E_APPNOTFOUND 0x800401F5
CO_E_APPSINGLEUSE 0x800401F6
CO_E_BAD_PATH 0x80080004
CO_E_CANTDETERMINECLASS 0x800401F2
CO_E_CLASSSTRING 0x800401F3
CO_E_CLASS_CREATE_FAILED 0x80080001
CO_E_DLLNOTFOUND 0x800401F8
CO_E_ERRORINAPP 0x800401F7
CO_E_ERRORINDLL 0x800401F9
CO_E_IIDSTRING 0x800401F4
CO_E_NOTINITIALIZED 0x800401F0
CO_E_OBJISREG 0x800401FC
CO_E_OBJNOTCONNECTED 0x800401FD
CO_E_OBJNOTREG 0x800401FB
CO_E_OBJSRV_RPC_FAILURE 0x80080006
CO_E_RELEASED 0x800401FF
CO_E_SERVER_EXEC_FAILURE 0x80080005
CO_E_SERVER_STOPPING 0x80080008
CO_E_WRONGOSFORAPP 0x800401FA
DISP_E_ARRAYISLOCKED 0x8002000D
DISP_E_BADCALLEE 0x80020010
DISP_E_BADINDEX 0x8002000B
DISP_E_BADPARAMCOUNT 0x8002000E
DISP_E_BADVARTYPE 0x80020008
DISP_E_DIVBYZERO 0x80020012
DISP_E_EXCEPTION 0x80020009
DISP_E_MEMBERNOTFOUND 0x80020003
DISP_E_NONAMEDARGS 0x80020007
DISP_E_NOTACOLLECTION 0x80020011
DISP_E_OVERFLOW 0x8002000A
DISP_E_PARAMNOTFOUND 0x80020004
DISP_E_PARAMNOTOPTIONAL 0x8002000F
DISP_E_TYPEMISMATCH 0x80020005
DISP_E_UNKNOWNINTERFACE 0x80020001
DISP_E_UNKNOWNLCID 0x8002000C
DISP_E_UNKNOWNNAME 0x80020006
E_ABORT 0x80004004
E_ACCESSDENIED 0x80070005
E_FAIL 0x80004005
E_HANDLE 0x80070006
E_INVALIDARG 0x80070057
E_NOINTERFACE 0x80004002
E_NOTIMPL 0x80004001
E_OUTOFMEMORY 0x8007000E
E_PENDING 0x8000000A
E_POINTER 0x80004003
E_UNEXPECTED 0x8000FFFF
TYPE_E_AMBIGUOUSNAME 0x8002802C
TYPE_E_BADMODULEKIND 0x800288BD
TYPE_E_BUFFERTOOSMALL 0x80028016
TYPE_E_CANTCREATETMPFILE 0x80028CA3
TYPE_E_CANTLOADLIBRARY 0x80029C4A
TYPE_E_CIRCULARTYPE 0x80029C84
TYPE_E_DLLFUNCTIONNOTFOUND 0x8002802F
TYPE_E_DUPLICATEID 0x800288C6
TYPE_E_ELEMENTNOTFOUND 0x8002802B
TYPE_E_INCONSISTENTPROPFUNCS 0x80029C83
TYPE_E_INVALIDID 0x800288CF
TYPE_E_INVALIDSTATE 0x80028029
TYPE_E_INVDATAREAD 0x80028018
TYPE_E_IOERROR 0x80028CA2
TYPE_E_LIBNOTREGISTERED 0x8002801D
TYPE_E_NAMECONFLICT 0x8002802D
TYPE_E_OUTOFBOUNDS 0x80028CA1
TYPE_E_QUALIFIEDNAMEDISALLOWED 0x80028028
TYPE_E_REGISTRYACCESS 0x8002801C
TYPE_E_SIZETOOBIG 0x800288C5
TYPE_E_TYPEMISMATCH 0x80028CA0
TYPE_E_UNDEFINEDTYPE 0x80028027
TYPE_E_UNKNOWNLCID 0x8002802E
TYPE_E_UNSUPFORMAT 0x80028019
TYPE_E_WRONGTYPEKIND 0x8002802A
)

NB. =========================================================
NB. wd syntax interface to openoffice.org

(18!:55 :: 0:) <'wdooo'
coclass 'wdooo'
coinsert 'olecomerrorh'
coinsert 'pole32'

NB. 3 : 0''
NB. a=. ;:'VT_EMPTY VT_NULL VT_I2 VT_I4  VT_R4 VT_R8 VT_CY VT_DATE'
NB. a=. a, ;:'VT_BSTR VT_DISPATCH VT_ERROR VT_BOOL'
NB. a=. a, ;:'VT_VARIANT VT_UNKNOWN VT_DECIMAL'
NB. a=. a, ;:'VT_PTR VT_SAFEARRAY VT_CARRAY VT_USERDEFINED'
NB. a=. a, ;:'VT_VECTOR VT_ARRAY VT_BYREF VT_TYPEMASK'
NB. for_ai. a do. ((>ai),'_z_')=: ".>ai end.
NB. i. 0 0
NB. )

NB. prototype
VariantInit=: 'oleaut32 VariantInit > n *'&cd
SafeArrayCreateVector=: 'oleaut32 SafeArrayCreateVector > i s i i'&cd
SafeArrayPutElement=: 'oleaut32 SafeArrayPutElement > i i *i *'&cd

NB. useful constants
S_OK=: 0

DISPID_PROPERTYPUT=: _3
dispidNamed=: 2&ic DISPID_PROPERTYPUT
pdispidNamed=: symdat@symget < 'dispidNamed'
iid_idisp=: 0 4 2 0 0 0 0 0 192 0 0 0 0 0 0 70{a.  NB. {00020400-0000-0000-c000-000000000046}

NB. Flags for IDispatch::Invoke
DISPATCH_METHOD=: 1
DISPATCH_PROPERTYGET=: 2
DISPATCH_PROPERTYPUT=: 4
DISPATCH_PROPERTYPUTREF=: 8

oleerrno=: S_OK
init=: 0

create=: 3 : 0
'ole32 CoInitialize > i x'&cd^:IFCONSOLE 0
oleerrno=: S_OK
init=: 0
)

destroy=: 3 : 0
if. init do.
  VariantClear <<temp
  memf temp
  base iuRelease ''
end.
'ole32 CoUninitialize > n'&cd^:IFCONSOLE ''
codestroy''
)

NB. ---------------------------------------------------------
NB. private members

dispid=: 4 : 0
assert. x~:0
y=. uucp y
nm=. ,symdat symget <,'y'
hr=. x idGetIDsOfNames GUID_NULL;nm;1;0;r=. ,_1
hr, r
)

makevariant=: 4 : 0
assert. x =&# y
if. 0=#y do. 0 return. end.
vargs=. mema 16 * #y
for_i. i.#y do.
  VariantInit <<arr=. vargs + 16 * i
  s=. >i{y
  (>i{x) memw arr, 0, 1, 4
  select. 16bfff (17 b.) i{x
  case. VT_BOOL do.
    ((s=0){_1 0) memw arr, 8, 1, 4
  case. VT_BSTR do.
    bstr=. SysAllocStringLen (];#) uucp ,s
    bstr memw arr, 8, 1, 4
  case. VT_I4 do.
    s memw arr, 8, 1, 4
  case. VT_R8 do.
    s memw arr, 8, 1, 8
  case. VT_UNKNOWN;VT_DISPATCH do.
    if. 0=#s do.  NB. shorthand for NULL
      0 memw arr, 8, 1, 4
    else.
      s memw arr, 8, 1, 4
    end.
  end.
end.
vargs
)

makedispparms=: 4 : 0
dispparams=. mema 16
(4#0) memw dispparams, 0, 4, 4
(x makevariant&|. y) memw dispparams, 0, 1, 4
(#y) memw dispparams, 8, 1, 4
dispparams
)

freedispparms=: 3 : 0
'a b c d'=. memr y, 0, 4, 4
if. a do.
  VariantClear@<@<"0 a+16*i.#c
  memf a
end.
memf y
)

oleinvoke=: 1 : 0
'' (m oleinvoke) y
:
'disp name'=. 2{. y
args=. 2}.y
oleerrno=: S_OK
if. 0=#x do. x=. (VT_BSTR, VT_BSTR, VT_I4, VT_I4, VT_R8, VT_UNKNOWN) {~ 2 131072 1 4 8 i. (3!:0&> args) end.
newdisp=. 0
if. disp=temp do.  NB. pass prev temp for further invoke
  if. (VT_UNKNOWN, VT_DISPATCH) -.@e.~ {.oletype temp do. 13!:8[3 [ oleerrno=: DISP_E_TYPEMISMATCH end.
  newdisp=. 1
  '' iuAddRef~ disp=. {. memr temp, 8, 1, 4
end.
if. S_OK~: 0{:: 'hr id'=. disp dispid name do. 13!:8[3 [ oleerrno=: hr end.
VariantClear <<temp
dispparams=. x makedispparms args
if. m=DISPATCH_PROPERTYPUT do.
  pdispidNamed memw dispparams, 4, 1, 4
  1 memw dispparams, 12, 1, 4  NB. Number of named arguments
end.
if. S_OK~: hr=. disp idInvoke id ; GUID_NULL ; 0 ; m ; (<dispparams) ; (<temp) ; 0 ; 0 do. 13!:8[3 [ oleerrno=: hr end.
freedispparms dispparams
if. newdisp do. disp iuRelease '' end.
temp
)

NB. ---------------------------------------------------------
NB. public members

NB. 'base temp'=. olecreate progid
olecreate=: 3 : 0
NB. create object and get idispatch, temp
oleerrno=: S_OK
if. S_OK= hr=. CLSIDFromProgID`CLSIDFromString@.('{'={.@>@{.) y ; guid=. 1#GUID do.
  if. S_OK= hr=. CoCreateInstance guid ; 0 ; CTX ; iid_idisp ; p=. ,_2 do.
    base=: {.p
    init=: 1
NB. temp result holder
    VariantInit <<temp=: mema 16
    rz=. base, temp
  end.
end.
if. S_OK~: hr do. 13!:8[3 [ oleerrno=: hr end.
rz
)

NB. y: name ; args
NB. x: args type   (optional)
olemethod=: DISPATCH_METHOD oleinvoke
oleget=: DISPATCH_PROPERTYGET oleinvoke
oleset=: DISPATCH_PROPERTYPUT oleinvoke
olesetref=: DISPATCH_PROPERTYPUTREF oleinvoke

NB. interface=. oleid temp
oleid=: 3 : 0
oleerrno=: S_OK
if. (VT_UNKNOWN, VT_DISPATCH) -.@e.~ {.oletype y do. 13!:8[3 [ oleerrno=: DISP_E_TYPEMISMATCH end.
'' iuAddRef~ d=. {. memr y, 8, 1, 4
d
)

NB. release interface created by oleid
olerelease=: 3 : 0
y iuRelease ''
)

NB. equivalent of wd'qer'
oleqer=: 3 : 0
olecomerrmsg oleerrno
)

NB. retrieve type of variant
NB. return 4-element vector: basictype isvector isarray isbyref
oletype=: 3 : 0
vt=. {. _1&ic memr y, 0, 2, 2
vt0=. vt ((17 b.) (26 b.)) VT_VECTOR (23 b.) VT_ARRAY (23 b.) VT_BYREF
vt0, 0~: vt (17 b.) VT_VECTOR, VT_ARRAY, VT_BYREF
)

NB. retrieve value of variant
olevalue=: 3 : 0
'vt vector array byref'=. oletype y
if. byref do. y=. {. memr y, 8, 1, 4 end.
select. vt
case. VT_R4 do. {. _1&fc memr y, 8, 4, 2
case. VT_R8 do. {. memr y, 8, 1, 8
case. VT_BSTR do. 6 u: memr b, 0, ({.memr b, _4 1 4), 2 [ b=. {.memr y, 8 1 4
case. do. {. memr y, 8, 1, 4
end.
)

NB. make safearray
NB. x VT_...
NB. y elements (may be empty)
NB. return 0 if failed
olevector=: 4 : 0
elms=. y
vt=. x
propVals=. SafeArrayCreateVector vt ; 0 ; #elms
failure=. 0
for_i. i.#elms do.
  if. S_OK&~: hr=. SafeArrayPutElement propVals ; (,i) (;<) <i{elms do.
    failure=. 1 break.
  end.
end.
if. 0=failure do.
  propVals
else.
  for_elm. elms do. elm iuRelease '' end.
  VariantClear <<propVals
  0
end.
)

