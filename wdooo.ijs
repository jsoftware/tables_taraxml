NB. =========================================================
NB. ------------ include files from pcall start -------------

coclass 'pcall'

NB.*acall c address call using function pointer
NB.   func=: FuncAddr acall 'i x *w ...'
NB.   res=. func arg1;arg2;...
acall=: 2 : '(''0 '',(":m),'' > '',n)&(15!:0)'

NB.*icall c interface call using v-table
NB.   iuQueryInterface=: IU_QueryInterface icall 'i x  *c *i' @ ;
NB.   res=. iObj iuQueryInterface arg1;arg2;...
icall=: 2 : '(''1 '',(":m),'' > '',n)&(15!:0)'

NB.*idef d interface index table definition
NB.   'IU_'idef IUnknown=: ;:'QueryInterface AddRef Release'
idef_z_=: 4 : '((x&,)&.>y)=: i.#y'

acall_z_=: acall_pcall_
icall_z_=: icall_pcall_

NB.*symdat v pointer to J name data, used in structs
NB.   symdat symget <'name'
symdat_z_=: 3 : 0   NB.
had=. {.memr y,(IF64{4 8),1,JPTR
had+{.memr had,0,1,JINT
)

coclass 'pole32'

CLSIDFromProgID=: 'ole32 CLSIDFromProgID  > i *w *c'&cd
CLSIDFromString=: 'ole32 CLSIDFromString  > i *w *c'&cd
CoCreateInstance=: 'ole32 CoCreateInstance > i *c i i *c *i'&cd
CoGetObject=: 'ole32 CoGetObject      > i *w i *c *i'&cd

VariantClear=: 'oleaut32 VariantClear > i *i'&cd
VariantChangeType=: 'oleaut32 VariantChangeType > i *i *i i i'&cd
SysFreeString=: 'oleaut32 SysFreeString > i i'&cd
SysAllocStringLen=: 'oleaut32 SysAllocStringLen > i *w i'&cd
CoGetActiveObject=: 'oleaut32 GetActiveObject     > i *c i *i'&cd

GUID=: 'WWWWXXYYZZZZZZZZ'
GUID_NULL=: (#GUID) # 0{a.
VAR1=: 'VtR1R2R3Valu'
VAR2=: 'VtR1R2R3Val1Val2'
DISPPARAMS=: 'ArgsNamdCArgCNmd'
TYPEATTR=: GUID,'LcidReseCtorDtorScheSinsTpknCfCvCtSvBaTfMjMnAliaIdld'
TYPEDESC=: 'LptdVt00'
PARAMDESC=: 'LppdPf00'
ELEMDESC=: TYPEDESC,PARAMDESC
FUNCDESC=: 'MbidScodParmFunkInvkCalcCpCoOvCs',ELEMDESC,'Ff00'
IID_IUnknown=: '{00000000-0000-0000-C000-000000000046}'
IID_IDispatch=: '{00020400-0000-0000-C000-000000000046}'

'CLSCTX_INPROC_SERVER CLSCTX_LOCAL_SERVER'=: 16b0001 16b0004
CTX=: CLSCTX_INPROC_SERVER+CLSCTX_LOCAL_SERVER

'VT_EMPTY VT_NULL VT_I2 VT_I4  VT_R4 VT_R8 VT_CY VT_DATE'=: i.8
'VT_BSTR VT_DISPATCH VT_ERROR VT_BOOL'=: 8+i.4
'VT_VARIANT VT_UNKNOWN VT_DECIMAL'=: 12+i.3
'VT_PTR VT_SAFEARRAY VT_CARRAY VT_USERDEFINED'=: 26+i.4
'VT_VECTOR VT_ARRAY VT_BYREF VT_TYPEMASK'=: 16b1000 16b2000 16b4000 16b0fff

'DISP_METH DISP_GET DISP_PUT DISP_SET'=: 1 2 4 8

'IU_'idef IUnknown=: ;:'QueryInterface AddRef Release'
'ID_'idef IDispatch=: IUnknown,;:;<;._2(0 : 0)
  GetTypeInfoCount GetTypeInfo GetIDsOfNames Invoke
)
'IT_'idef ITypeInfo=: IUnknown,;:;<;._2(0 : 0)
  GetTypeAttr GetTypeComp GetFuncDesc GetVarDesc GetNames
  GetRefTypeOfImplType GetImplTypeFlags GetIDsOfNames Invoke
  GetDocumentation GetDllEntry GetRefTypeInfo AddressOfMember
  CreateInstance GetMops GetContainingTypeLib ReleaseTypeAttr
  ReleaseFuncDesc ReleaseVarDesc
)

iuQueryInterface=: IU_QueryInterface icall 'i x  *c *i' @ ;
iuAddRef=: IU_AddRef icall 'i x' @ [
iuRelease=: IU_Release icall 'i x' @ [
idGetIDsOfNames=: ID_GetIDsOfNames icall 'i x  *c *i i i *i' @ ;
idGetTypeInfo=: ID_GetTypeInfo icall 'i x  i i *i' @ ;
idInvoke=: ID_Invoke icall 'i x  i *c i s *i *i i i' @ ;
itGetTypeAttr=: IT_GetTypeAttr icall 'i x  *i' @ ;
itReleaseTypeAttr=: IT_ReleaseTypeAttr icall 'i x  i' @ ;
itGetFuncDesc=: IT_GetFuncDesc icall 'i x  i *i' @ ;
itReleaseFuncDesc=: IT_ReleaseFuncDesc icall 'i x  i' @ ;
itGetNames=: IT_GetNames icall 'i x  i *i i *i' @ ;
itGetDocumentation=: IT_GetDocumentation icall 'i x  i *i *i i i' @ ;
itGetRefTypeInfo=: IT_GetRefTypeInfo icall 'i x  i *i' @ ;

hex8=: ,~ '00000000' }.~ #
hfd8=: '0x' , hex8@hfd
herr=: hfd8 assert 0 <: ]

CreateObject=: 3 : 0
IID_IDispatch CreateObject y
:
herr CoCreateInstance (GetGuid y) ; 0 ; CTX ; (GetGuid x) ; p=. ,_2
{.p
)

GetObject=: 3 : 0
IID_IDispatch GetObject y
:
herr CoGetObject y ; 0 ; (GetGuid x) ; p=. ,_2
{.p
)

GetActiveObject=: 3 : 0
IID_IDispatch GetActiveObject y
:
herr CoGetActiveObject (GetGuid y) ; 0 ; p=. ,_2
{.p
)

GetGuid=: 3 : 0
f=. CLSIDFromProgID`CLSIDFromString@.('{'={.y)
herr f y ; guid=. 1#GUID
guid
)

h=: ([: ;:^:_1"1 [: <"1 hfd)@(([: , _4 (_2&(3!:4))@|.\ ])^:(2=3!:0))
si=: I.@E.~   NB. TYPEATTR si 'Cfun'
us=: 0&(3!:4)
mi=: [: {.@memr ,&(0 1,JINT)
mc=: ,&0@] memr@, ,&JCHAR@[
mI=: 4 : '{.memr y,x,1,JINT'
mS=: 4 : '{.us memr y,x,2,JCHAR'
and=: 17 b.

GetStr=: 3 : 0
if. 0=y do. ''return. end.
len=. mi _4+y    NB. BSTR length
val=. len mc y   NB. BSTR char pairs value
8 u: 6 u: val
)

GetStrSafeFree=: 3 : 0
if. 0=y=. {.y do. '' return. end.
r=. GetStr y
SysFreeString y
r
)

AllocStr=: 3 : 0
SysAllocStringLen y;#y
)

VariantAlloc=: 3 : '(2-2)#~4%~#VAR2'
VariantStr=: GetStr@(2&{)

BoolVar=: 3 : 'VT_BOOL,0,(_1 0{~0-:{.y),0'
IntVar=: 3 : 'VT_I4,0,({.y),0'
PtrVar=: 3 : 'VT_UNKNOWN,0,({.y),0'
FloatVar=: 3 : 'VT_R8,0,_2(3!:4)2(3!:5){.y'
StrVar=: 3 : 'VT_BSTR,0,(AllocStr y),0'
EmptyVar=: 3 : '({.y,VT_EMPTY),0,0,_1'

JVar=: 3 : 0
if. y-:a: do. EmptyVar'' return. end.
if. y-:<0 do. EmptyVar VT_NULL return. end.
if. 0<L.y do. PtrVar >y return. end.
select. 3!:0 y
case. 1;4 do. IntVar y
case. 8 do. FloatVar y
case. do. StrVar ,":,y
end.
)

VarJ=: 3 : 0
select. VT_TYPEMASK and {.y
case. VT_EMPTY do. a:
case. VT_NULL do. <0
case. VT_I2 do. _1(3!:4)1(3!:4) 2{y
case. VT_I4 do. 2{y
case. VT_DISPATCH;VT_UNKNOWN do. <2{y
case. VT_R4 do. _1(3!:5)2(3!:4) 2{y
case. VT_R8 do. _2(3!:4)2(3!:4) 2 3{y
case. VT_BSTR do. VariantStr y
case. VT_BOOL do. 0~:2{y
case. do. VariantStr y [ VariantChangeType y;y;0;VT_BSTR
end.
)

VTSTR=: ; <@cut;._2 (0 : 0)
void null short long float double CURRENCY DATE
BSTR IDispatch* SCODE boolean VARIANT IUnknown* WCHAR .
char BYTE WORD DWORD int64 uint64 int UINT
void HRESULT PTR SAFEARRAY CARRAY USERDEFINED LPSTR LPWSTR
. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
FILETIME BLOB STREAM STORAGE STREAMED_OBJECT STORED_OBJECT BLOB_OBJECT
CF CLSID BAD_TYPE
)

VtStr=: 3 : 'VTSTR >@{~ (<:#VTSTR) <. y and 16bfff'

TypeDesc=: 0&$: : (4 : 0)"0
if. 0=y do. 'void' return. end.
select. vt=. 16bfff and 4 mS y
case. VT_PTR do. '*',~x TypeDesc 0 mI y return.
case. VT_USERDEFINED do. x RefDesc 0 mI y return.
case. do. VtStr vt
end.
)

RefDesc=: 4 : 0
herr x itGetRefTypeInfo y ; rt=. ,_2
rt=. {.rt
r=. >@{. rt GetDoc _1
rt iuRelease ''
r
)

FuncDesc=: 4 : 0"0
herr x itGetFuncDesc y ; pfd=. ,_2
mid=. 0 mI pfd
if. c=. (FUNCDESC si 'Cp') mS pfd do.    NB. cParams, params count
  p=. (FUNCDESC si 'Parm') mI pfd        NB. Params
  r=. x <@TypeDesc p+(#ELEMDESC)*i.c
else. r=. '' end.
r=. r,~<x TypeDesc {.(FUNCDESC si 'Lptd') + pfd  NB. return type
herr x itReleaseFuncDesc {.pfd

res=. (c+1)#_1
herr x itGetNames mid ; res ; (#res) ; ,_1
res=. 0 (I.res=_1)}res
r (, ' '&,)&.> <@GetStrSafeFree"0 res    NB. names of arguments
)

GetDoc=: 4 : 0"0
herr x itGetDocumentation y ; (name=. ,_2) ; (doc=. ,_2) ; 0 ; 0
<@GetStrSafeFree"0 name,doc           NB. name;doc
)

FuncDoc=: 4 : 0"0
herr x itGetFuncDesc y ; pfd=. ,_2
mid=. 0 mI pfd
herr x itReleaseFuncDesc {.pfd
x GetDoc mid
)


NB. Global Interface Table
NB. Marchalling interface pointers between appartments and threads

CLSID_StdGlobalInterfaceTable=: '{00000323-0000-0000-C000-000000000046}'
IID_IGlobalInterfaceTable=: '{00000146-0000-0000-C000-000000000046}'

'GIT_'idef IGlobalInterfaceTable=: IUnknown,;:;<;._2(0 : 0)
  RegisterInterfaceInGlobal RevokeInterfaceFromGlobal GetInterfaceFromGlobal
)

gitRegisterInterfaceInGlobal=: GIT_RegisterInterfaceInGlobal icall 'i x  i *c *i' @ ;
gitRevokeInterfaceFromGlobal=: GIT_RevokeInterfaceFromGlobal icall 'i x  i' @ ;
gitGetInterfaceFromGlobal=: GIT_GetInterfaceFromGlobal icall 'i x  i *c *i' @ ;

gitGet=: 3 : 0
IID_IDispatch gitGet y
:
git=. IID_IGlobalInterfaceTable CreateObject CLSID_StdGlobalInterfaceTable
herr CLSIDFromString x ; iid=. 1#GUID
herr git gitGetInterfaceFromGlobal y;iid;p=. ,2-2
git iuRelease ''
{.p
)

NB. ------------ include files from pcall end ---------------
NB. =========================================================

NB. =========================================================
NB. error constant

(18!:55 :: 0:) <'olecomerrorh'
coclass 'olecomerrorh'

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

3 : 0''
a=. ;:'VT_EMPTY VT_NULL VT_I2 VT_I4  VT_R4 VT_R8 VT_CY VT_DATE'
a=. a, ;:'VT_BSTR VT_DISPATCH VT_ERROR VT_BOOL'
a=. a, ;:'VT_VARIANT VT_UNKNOWN VT_DECIMAL'
a=. a, ;:'VT_PTR VT_SAFEARRAY VT_CARRAY VT_USERDEFINED'
a=. a, ;:'VT_VECTOR VT_ARRAY VT_BYREF VT_TYPEMASK'
for_ai. a do. ((>ai),'_z_')=: ".>ai end.
i. 0 0
)

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

