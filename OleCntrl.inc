
;--- this is the support file for OLE controls in ASM
;--- all the support code is hold in macros
;--- no external files will be linked to your source

;--------------------------------------------------------------------
;--- brief overview of macros defined by olecntrl.inc:
;-------- 1. macro on coclass level (one instance/coclass)
;--- BEGIN_CONNECTION_POINT_MAP		start table of outgoing interfaces
;--- CONNECTION_POINT_ENTRY			entry in table of outgoing interfaces
;--- END_CONNECTION_POINT_MAP		end table of outgoing interfaces
;--- BEGIN_COM_MAP					start table of interfaces
;--- COM_INTERFACE_ENTRY			entry in table of interfaces
;--- END_COM_MAP					end table of interfaces
;--- DEFINE_STD_COM_METHODS			includes code for standard interfaces
;--- DEFINE_DISPATCH				includes code for IDispatch
;--- STD_COM_CONSTRUCTOR			initialization code for standard interfaces
;--- STD_COM_DESTRUCTOR				clean up code for standard interfaces
;--- DEFINE_KNOWN_INTERFACES		define table of interfaces known by QI
;--- FIREEVENT						fire event (moved to DISPHLP.INC)
;-------- 2. macro on type library level (one instance/module)
;--- BEGIN_OBJECT_MAP				start table of object entries
;--- END_OBJECT_MAP					end table of object entries
;--- DEFINE_CLASSFACTORY			includes code for IClassFactory class
;--- DEFINE_COMHELPER				includes code for some utility functions
;--- DEFINE_GETCLASSOBJECT			includes code for DllGetClassObject
;--- DEFINE_REGISTERSERVER			includes code for DllRegisterServer
;--- DEFINE_UNREGISTERSERVER		includes code for DllUnregisterServer
;--- DEFINE_CANUNLOADNOW			includes code for DllCanUnloadNow
;--- DEFINE_DLLMAIN					includes code for DllMain
;-------- there exist other macros inside, but they should regarded
;-------- as internal only
;--------------------------------------------------------------------

;--------------------------------------------------------------------
;--- there exist some "exits" inside the macros (simple ifdefs). 
;--- To activate such an exit you only need to define it in your code.
;--- These are the exits currently defined
;--- (NOTE: every exit has the classname as suffix appended,
;--- so read the "@" in the next lines as "@"+classname):

;--- IUnknown:
;---	OnQueryInterface@ this_, riid, ppReturn
;---	OnAddRef@ this_
;---	OnRelease@ this_
;--- IViewObject:
;---	OnDraw@ this_, ptd, hicTargetDev, hdc, pRectBounds, prcWBounds
;--- IOleObject:
;---	OnInPlaceActivate@, this_, hwndParent, lprcPosRect
;---	OnVerbPrimary@, this_, lpmsg, pOleClientSite, hwndParent, lprcPosRect
;---	OnDoVerb@ this_, iVerb, lpmsg, pOleClientSite, hwndParent, lprcPosRect
;--- IDispatch:
;---	OnInvoke@, this_, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr
;--- IPersistStreamInit:
;---	OnInitNew@, this_
;--- IPersistStreamInit:
;---	OnGetSizeMax@, this_, pSize
;--- ISpecifyPropertyPages:
;---	GetPages@, this_, pPages	;MUST be defined if interface is to be supported
;--------------------------------------------------------------------
;--- to cause the macros to log memory allocs + frees, define
;--- equate _ALLOCDEBUG in the source containing DEFINE_DLLMAIN.
;---
;--- If your OnInvoke exit proc hasnt handled the event, set eax to -1
;--- and return
;--------------------------------------------------------------------

;History
;--- 03/28/04: FIREEVENT macro deactivated (is defined in DISPHLP.INC)
;--- 03/27/04: bugfix in macro STD_COM_DESTRUCTOR for multiple CP case (thanks Wayne)
;--- 03/21/04: added code in macro DEFINE_PERSISTTREAMINIT to load/save props of VT_DISPATCH type
;--- 09/15/03: DEFINE_FIREEVENTHELPER macro deleted (now in DISPHLP.INC)
;--- 07/31/03: FIREEVENT, DEFINE_FIREEVENTHELPER macros added
;--- 03/02/03: OLEIVERB_SHOW now activates control
;--- 01/02/03: first version

;--------------------------------------------------------------------
;--- define equates, typedefs + structures
;--------------------------------------------------------------------

ifndef ?AGGREGATION
?AGGREGATION	=	1	;default: support Aggregation
endif
ifndef ?EVENTSUPPORT
?EVENTSUPPORT	=	1
endif

;--- flags

DISPATCH_NO_TYPELIB_MEMBERS equ 1
DISPATCH_NO_TYPEINFO		equ 2

;*** prototypes of functions from COMHELPER macro

IsInterfaceSupported proto :ptr IID,:ptr ptr IID, :dword, :ptr, :ptr LPUNKNOWN
ComPtrAssign		proto :ptr LPUNKNOWN, :LPUNKNOWN
WideCharToLocal		proto :LPSTR, :LPWSTR, :DWORD
LocalToWideChar		proto :LPWSTR,:LPSTR,  :DWORD
FireEvent			proto c :ptr, :LPCONNECTIONPOINT, :VARARG

;--------------------------------------------------------------------
;--- BEGIN_OBJECT_MAP: defines object map containing all coclasses
;--- defined in current module
;--------------------------------------------------------------------

BEGIN_OBJECT_MAP macro name
name label ObjectEntry
??ObjectMap equ name	;;ensures only 1 map is defined
	endm

;--------------------------------------------------------------------
;--- END_OBJECT_MAP: marks end of object table
;--------------------------------------------------------------------
END_OBJECT_MAP macro
??ObjectMapItems equ ($ - ??ObjectMap) / sizeof ObjectEntry
	endm

;--------------------------------------------------------------------
;--- BEGIN_CONNECTION_POINT_MAP: start of connection point map
;--------------------------------------------------------------------
BEGIN_CONNECTION_POINT_MAP macro name
EventIID_&name label REFIID
??Tablename textequ <name>
	endm

;--------------------------------------------------------------------
;--- END_CONNECTION_POINT_MAP: end of connection point map
;--------------------------------------------------------------------
END_CONNECTION_POINT_MAP macro
%NUMCP_&??Tablename equ ($ - EventIID_&??Tablename) / sizeof REFIID
	endm

;--------------------------------------------------------------------
;--- CONNECTION_POINT_ENTRY: entry in connection point map
;--------------------------------------------------------------------
CONNECTION_POINT_ENTRY macro refiid
	REFIID refiid
	endm

;--------------------------------------------------------------------
;--- BEGIN_COM_MAP: global equate ??classname is defined (currently
;--- not used)
;--------------------------------------------------------------------

BEGIN_COM_MAP macro name
??classname equ name
%QueryInterface@&??classname proto :ptr ??classname, :REFIID, :ptr ptr
%AddRef@&??classname		proto :ptr ??classname
%Release@&??classname		proto :ptr ??classname
	endm

;--------------------------------------------------------------------
;--- COM_INTERFACE_ENTRY: define the vtable entry for this
;--- interface + define equate INCLUDE_xxx for later code insertion
;--------------------------------------------------------------------

COM_INTERFACE_ENTRY macro name, interface, Flags
ifb <interface>
_&name	name <>
else
union

%ifndef INCLUDE_&??classname&_IUnknown
ife ?AGGREGATION
_IUnknown	IUnknown <>
%INCLUDE_&??classname&_IUnknown equ 1
MEMBER _IUnknown
endif
endif

_&name	name <>
_&interface	interface <>
ends
endif

%ifndef INCLUDE_&??classname&_IUnknown
if ?AGGREGATION
_IUnknown	IUnknown <>
%INCLUDE_&??classname&_IUnknown equ 1
MEMBER _IUnknown
endif
endif


%INCLUDE_&??classname&_&name equ 1
MEMBER _&name
ifnb <interface>
MEMBER _&interface
endif

ifnb <Flags>
%&??classname&_&name&_Flags equ Flags
else
%&??classname&_&name&_Flags equ 0
endif

	endm

;--------------------------------------------------------------------
;--- COM_INTERFACE_ENTRY_IID: 
;--- currently oldname is NOT used
;--------------------------------------------------------------------

COM_INTERFACE_ENTRY_IID macro name, oldname
	org $-4
_&name	name <>
	endm	

;--------------------------------------------------------------------
;--- COM_INTERFACE_ENTRY_EX: 
;--------------------------------------------------------------------

COM_INTERFACE_ENTRY_EX macro nameslist:VARARG
union
    for name,<nameslist>
_&name	name <>
	endm
ends
	endm	

;--------------------------------------------------------------------
;--- END_COM_MAP: define all internal members now
;--- so these variables will come immediately after the vtables
;--------------------------------------------------------------------

END_COM_MAP macro

ObjRefCount			DWORD	?
	MEMBER ObjRefCount

%ifdef INCLUDE_&??classname&_IConnectionPointContainer
%ifndef ??classname&_cntCP
%??classname&_cntCP equ 1
endif
%pConnectionPoint	LPCONNECTIONPOINT	??classname&_cntCP dup (?)
	MEMBER pConnectionPoint
endif

%ifdef INCLUDE_&??classname&_IUnknown
if ?AGGREGATION
pUnkOuter			LPUNKNOWN ?
	MEMBER pUnkOuter
endif
endif

%ifdef INCLUDE_&??classname&_IDispatch
%ife ??classname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
%ife ??classname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
refIID_TypeLib		REFIID	?
lcid				LCID	?
dwMajorVer			DWORD	?
dwMinorVer			DWORD	?
	MEMBER refIID_TypeLib, lcid, dwMajorVer, dwMinorVer
endif
pTypeInfo			LPTYPEINFO	?
	MEMBER pTypeInfo
endif
else
%ifdef INCLUDE_&??classname&_IProvideClassInfo
pTypeInfo			LPTYPEINFO	?
	MEMBER pTypeInfo
endif
endif

%ifdef INCLUDE_&??classname&_IOleObject
pOleClientSite		LPOLECLIENTSITE		?
hWnd				HWND	?
dwFlags				DWORD	?
	MEMBER hWnd, pOleClientSite, dwFlags
endif

;--- values of dwFlags

OLEOBJF_INITIALIZED	equ 1
OLEOBJF_UIACTIVATED	equ 2

%ifdef INCLUDE_&??classname&_IOleInPlaceObject
%ifndef INCLUDE_&??classname&_IOleObject
.err <COM_INTERFACE_ENTRY IOleInPlaceObject requires entry IOleObject as well>
endif
endif

%ifdef INCLUDE_&??classname&_IOleInPlaceActiveObject
%ifndef INCLUDE_&??classname&_IOleObject
.err <COM_INTERFACE_ENTRY IOleInPlaceActiveObject requires entry IOleObject as well>
endif
%ifndef INCLUDE_&??classname&_IOleInPlaceObject
.err <COM_INTERFACE_ENTRY IOleInPlaceActiveObject requires entry IOleInPlaceObject as well>
endif
endif


%ifdef INCLUDE_&??classname&_IOleObject

isDirty				DWORD	?
pOleAdviseHolder	LPOLEADVISEHOLDER	?

	MEMBER isDirty
	MEMBER pOleAdviseHolder

endif

%ifdef INCLUDE_&??classname&_IDataObject

pDataAdviseHolder	LPDATAADVISEHOLDER	?

	MEMBER pDataAdviseHolder

endif

%ifdef INCLUDE_&??classname&_IOleObject

himetricExtent		SIZEL	{}
pixelExtent			SIZEL	{}
advf				DWORD	?
pAdviseSink			LPADVISESINK		?

	MEMBER  himetricExtent, pixelExtent, advf, pAdviseSink

%elseifdef INCLUDE_&??classname&_IViewObject2

himetricExtent		SIZEL	{}
pixelExtent			SIZEL	{}
advf				DWORD	?
pAdviseSink			LPADVISESINK		?

	MEMBER  himetricExtent, pixelExtent, advf, pAdviseSink

endif

	endm

;--------------------------------------------------------------------
;--- DEFINE_STD_COM_METHODS: with this macro all code necessary 
;--- for the standard interfaces of the object will be included
;--------------------------------------------------------------------

DEFINE_STD_COM_METHODS macro name

;--- parms supInterfaces + ?IFTABSIZE are defined by DEFINE_KNOWN_INTERFACES
;--- so currently define each coclass in its own source module

	DEFINE_UNKNOWN name, offset supInterfaces, ?IFTABSIZE

ifdef INCLUDE_&name&_IOleObject
	DEFINE_OLEOBJECT name
endif
ifdef INCLUDE_&name&_IOleInPlaceObject
	DEFINE_OLEINPLACEOBJECT name
endif
ifdef INCLUDE_&name&_IOleInPlaceActiveObject
	DEFINE_OLEINPLACEACTIVEOBJECT name
endif
ifdef INCLUDE_&name&_IOleControl
	DEFINE_OLECONTROL name
endif
ifdef INCLUDE_&name&_IDataObject
	DEFINE_DATAOBJECT name
endif
ifdef INCLUDE_&name&_IPersistStorage
	DEFINE_PERSISTSTORAGE name
endif
ifdef INCLUDE_&name&_IPersistStreamInit
	DEFINE_PERSISTSTREAMINIT name, offset BagTab, NUMBAGS
endif
ifdef INCLUDE_&name&_IPersistPropertyBag
	DEFINE_PERSISTPROPERTYBAG name, offset BagTab, NUMBAGS
endif
ifdef INCLUDE_&name&_IViewObject2
	DEFINE_VIEWOBJECT2 name
endif
ifdef INCLUDE_&name&_IProvideClassInfo
	DEFINE_PROVIDECLASSINFO name
endif
ifdef INCLUDE_&name&_ISpecifyPropertyPages
	DEFINE_SPECIFYPROPERTYPAGES name
endif
ifdef INCLUDE_&name&_ICategorizeProperties
	DEFINE_CATEGORIZE_PROPERTIES name, offset Categories, NUMCATEGORIES
endif
ifdef INCLUDE_&name&_IConnectionPointContainer
	DEFINE_CONNECTIONPOINTCONTAINER name
;--------------------------------------------------------------------
;--- the following 3 interfaces could possibly
;--- be moved to macro COMHELPER (1 instance/module)
;--------------------------------------------------------------------
	DEFINE_ENUMCONNECTIONPOINTS
	DEFINE_CONNECTIONPOINT
	DEFINE_ENUMCONNECTIONS
endif
ifdef INCLUDE_&name&_IRunnableObject
	DEFINE_RUNNABLEOBJECT name
endif

ifdef INCLUDE_&name&_IOleObject
	DEFINE_SENDVIEWCHANGE name
elseifdef INCLUDE_&name&_IViewObject2
	DEFINE_SENDVIEWCHANGE name
endif

	endm

;--------------------------------------------------------------------
;--- STD_COM_CONSTRUCTOR: this macro includes all code necessary 
;--- to initialize the members of included standard COM interfaces
;--- in the class constructor
;--------------------------------------------------------------------

STD_COM_CONSTRUCTOR macro objectname, pguidTypeLib, verMajor, verMinor
ifdef INCLUDE_&objectname&_IOleObject
	mov	m__IOleObject,				OFFSET COleObjectVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IOleInPlaceObject
	mov	m__IOleInPlaceObject,		OFFSET COleInPlaceObjectVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IOleInPlaceActiveObject
	mov	m__IOleInPlaceActiveObject,	OFFSET COleInPlaceActiveObjectVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IOleControl
	mov	m__IOleControl,				OFFSET COleControlVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IDataObject
	mov	m__IDataObject,				OFFSET CDataObjectVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IProvideClassInfo
	mov	m__IProvideClassInfo,		OFFSET CProvideClassInfoVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IPersistStorage
	mov	m__IPersistStorage,			OFFSET CPersistStorageVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IPersistStreamInit
	mov	m__IPersistStreamInit,		OFFSET CPersistStreamInitVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IPersistPropertyBag
	mov	m__IPersistPropertyBag,		OFFSET CPersistPropertyBagVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IViewObject2
	mov	m__IViewObject2,			OFFSET CViewObject2Vtbl_&objectname
endif
ifdef INCLUDE_&objectname&_ISpecifyPropertyPages
	mov	m__ISpecifyPropertyPages,	OFFSET CSpecifyPropertyPagesVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_ICategorizeProperties
	mov	m__ICategorizeProperties,	OFFSET CCategorizePropertiesVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IConnectionPointContainer
	mov	m__IConnectionPointContainer,OFFSET CConnectionPointContainerVtbl_&objectname
endif
ifdef INCLUDE_&objectname&_IRunnableObject
	mov	m__IRunnableObject,			OFFSET CRunnableObjectVtbl_&objectname
endif

	mov m_ObjRefCount,1

ifdef INCLUDE_&objectname&_IDispatch
ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
	mov eax, pguidTypeLib
	mov m_refIID_TypeLib, eax
	xor eax, eax
	mov ax, verMajor
	mov m_dwMajorVer, eax
	mov ax, verMinor
	mov m_dwMinorVer,eax
else
	push 0
    invoke LoadRegTypeLib, pguidTypeLib, verMajor,\
   				verMinor, LANG_NEUTRAL, esp
	.if (eax == S_OK)
		mov ecx, [esp]
		invoke vf(ecx, ITypeLib, GetTypeInfoOfGuid), objectname&_IDispatch_TypeInfo_IID, addr m_pTypeInfo
		mov ecx, [esp]
		invoke vf(ecx, IUnknown, Release)
	.endif
	pop eax
endif
endif
else
ifdef INCLUDE_&objectname&_IProvideClassInfo
	push 0
    invoke LoadRegTypeLib, pguidTypeLib, verMajor,\
   				verMinor, LANG_NEUTRAL, esp
	.if (eax == S_OK)
		mov ecx, [esp]
		invoke vf(ecx, ITypeLib, GetTypeInfoOfGuid), addr CLSID_&objectname&, addr m_pTypeInfo
		mov ecx, [esp]
		invoke vf(ecx, IUnknown, Release)
	.endif
	pop eax
endif
endif

if ?AGGREGATION
	mov	m__IUnknown, OFFSET CUnknownVtbl_&objectname
	mov eax, pUnkOuter
	.if (eax)
		mov m_pUnkOuter,eax
	.else
		lea eax,m__IUnknown
		mov m_pUnkOuter,eax
	.endif
endif

ifdef INCLUDE_&objectname&_IOleObject
	mov eax, 3969
	mov m_himetricExtent.cx_, eax
	mov m_himetricExtent.cy, eax
	mov eax, 150
	mov m_pixelExtent.cx_, eax
	mov m_pixelExtent.cy, eax
endif

	inc g_DllRefCount

	endm


;--------------------------------------------------------------------
;--- STD_COM_DESTRUCTOR: this macro includes all code necessary 
;--- to do clean up the included standard COM members
;--------------------------------------------------------------------

STD_COM_DESTRUCTOR macro objectname

	dec g_DllRefCount

ifdef INCLUDE_&objectname&_IConnectionPointContainer
if NUMCP_&objectname eq 1
	.if (m_pConnectionPoint)
		invoke vf(m_pConnectionPoint, IConnectionPoint, Release)
		mov m_pConnectionPoint,NULL
	.endif
else
	mov ecx, NUMCP_&objectname
	lea eax, m_pConnectionPoint
	.while (ecx)
		.if (dword ptr [eax])
			push eax
			push ecx
			invoke vf([eax], IConnectionPoint, Release)
			pop ecx
			pop eax
			mov dword ptr [eax],NULL
		.endif
		add eax, sizeof DWORD
		dec ecx
	.endw
endif
endif

ifdef INCLUDE_&objectname&_IDispatch
;;%ife ??classname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
	.if (m_pTypeInfo)
		invoke vf(m_pTypeInfo,ITypeInfo,Release)
		mov m_pTypeInfo,NULL
	.endif
;;endif
endif

ifdef INCLUDE_&objectname&_IOleObject
	.if (m_pOleClientSite)
		invoke vf(m_pOleClientSite,IUnknown,Release)
		mov m_pOleClientSite,NULL
	.endif
endif

ifdef INCLUDE_&objectname&_IDataObject
	.if (m_pDataAdviseHolder)
		invoke vf(m_pDataAdviseHolder, IUnknown, Release)
	.endif
endif

if defined(INCLUDE_&objectname&_IOleObject) or defined(INCLUDE_&objectname&_IViewObject2)
	.if (m_pAdviseSink)
		invoke vf(m_pAdviseSink,IUnknown,Release)
		mov m_pAdviseSink,NULL
	.endif
endif

ifdef INCLUDE_&objectname&_IOleObject
	.if (m_pOleAdviseHolder)
		invoke vf(m_pOleAdviseHolder,IUnknown,Release)
		mov m_pOleAdviseHolder,NULL
	.endif
endif
	endm


;--------------------------------------------------------------------------
;--- DEFINE_INTERFACE: this macro will define an entry in the 
;--- table of known interfaces (known by QueryInterface of the object)
;--------------------------------------------------------------------------

DEFINE_INTERFACE	macro name,	ofsvtbl
	dd offset IID_&name&	, ofsvtbl
	endm


;--------------------------------------------------------------------------
;--- DEFINE_KNOWN_INTERFACES: this macro includes all 
;--- standard COM interfaces known by the object (returned by QueryInterface)
;--------------------------------------------------------------------------

DEFINE_KNOWN_INTERFACES macro objectname, dispinterfaces:VARARG

supInterfaces label dword
ifnb <dispinterfaces>
	for parm,<dispinterfaces>
		DEFINE_INTERFACE parm,				objectname&._&parm
	endm
endif
;--- aggregation uses an extra entry for IUnknown vtable
if ?AGGREGATION
	DEFINE_INTERFACE IUnknown,				objectname&._IUnknown
else
	DEFINE_INTERFACE IUnknown,				0
endif
ifdef INCLUDE_&objectname&_IOleObject
	DEFINE_INTERFACE IOleObject,			objectname&._IOleObject
endif
ifdef INCLUDE_&objectname&_IOleInPlaceObject
	DEFINE_INTERFACE IOleWindow,			objectname&._IOleInPlaceObject
	DEFINE_INTERFACE IOleInPlaceObject,		objectname&._IOleInPlaceObject
endif
ifdef INCLUDE_&objectname&_IOleInPlaceActiveObject
;;---- this interface shouldnd be returned by QueryInterface
;;	DEFINE_INTERFACE IOleInPlaceActiveObject,objectname&._IOleInPlaceActiveObject
endif
ifdef INCLUDE_&objectname&_IOleControl
	DEFINE_INTERFACE IOleControl,			objectname&._IOleControl
endif
ifdef INCLUDE_&objectname&_IDispatch
	DEFINE_INTERFACE IDispatch,				objectname&._IDispatch
endif
ifdef INCLUDE_&objectname&_IDataObject
	DEFINE_INTERFACE IDataObject,			objectname&._IDataObject
endif
ifdef INCLUDE_&objectname&_IProvideClassInfo
	DEFINE_INTERFACE IProvideClassInfo,		objectname&._IProvideClassInfo
endif
ifdef INCLUDE_&objectname&_IPersistStorage
	DEFINE_INTERFACE IPersist,				objectname&._IPersistStorage
	DEFINE_INTERFACE IPersistStorage,		objectname&._IPersistStorage
endif
ifdef INCLUDE_&objectname&_IPersistStreamInit
	DEFINE_INTERFACE IPersistStreamInit,	objectname&._IPersistStreamInit
endif
ifdef INCLUDE_&objectname&_IPersistPropertyBag
	DEFINE_INTERFACE IPersistPropertyBag,	objectname&._IPersistPropertyBag
endif
ifdef INCLUDE_&objectname&_IViewObject2
	DEFINE_INTERFACE IViewObject,			objectname&._IViewObject2
	DEFINE_INTERFACE IViewObject2,			objectname&._IViewObject2
endif
ifdef INCLUDE_&objectname&_ISpecifyPropertyPages
	DEFINE_INTERFACE ISpecifyPropertyPages,	objectname&._ISpecifyPropertyPages
endif
ifdef INCLUDE_&objectname&_ICategorizeProperties
	DEFINE_INTERFACE ICategorizeProperties,	objectname&._ICategorizeProperties
endif
ifdef INCLUDE_&objectname&_IConnectionPointContainer
	DEFINE_INTERFACE IConnectionPointContainer,objectname&._IConnectionPointContainer
endif
ifdef INCLUDE_&objectname&_IRunnableObject
	DEFINE_INTERFACE IRunnableObject,		objectname&._IRunnableObject
endif
?IFTABSIZE equ ($ - supInterfaces)/ (2 * sizeof dword)
	endm


;--------------------------------------------------------------------------
;--- DEFINE_CLASSFACTORY: this macro will define a IClassFactory object.
;--- Only one for the module is needed (so constructor is public to
;--- avoid multiple definitions
;--------------------------------------------------------------------------

DEFINE_CLASSFACTORY macro 
local Destroy_, QueryInterface_, AddRef_, Release_, CreateInstance_, LockServer_, supInterfaces_

ifndef CLASSFACTORY_INCLUDED

CLASSFACTORY_INCLUDED equ 1

LPOBJECTENTRY typedef ptr ObjectEntry

CClassFactory struct
vtbl		dd ?
ObjRefCount dd ?
pClass		LPOBJECTENTRY ?
CClassFactory ends

	.const

CClassFactoryVtbl IClassFactoryVtbl {QueryInterface_,\
	AddRef_, Release_, CreateInstance_, LockServer_}

	.code

;------ constructor ClassFactory, return addr of object in eax (NULL == error)

;;Create@CClassFactory PROC public constructor:LPCONSTRUCTOR
Create@CClassFactory PROC public pClass:ptr ObjectEntry

	DebugOut "Create@ClassFactory"

	invoke	LocalAlloc,LMEM_FIXED or LMEM_ZEROINIT,sizeof CClassFactory
	.if (eax == NULL)
		DebugOut "Create@ClassFactory failed"
		ret
	.endif
				
	mov	[eax].CClassFactory.vtbl,OFFSET CClassFactoryVtbl

	mov ecx, pClass
	mov [eax].CClassFactory.pClass, ecx

	inc g_DllRefCount

	mov	[eax].CClassFactory.ObjRefCount, 1

	ret
Create@CClassFactory ENDP

;------ destructor ClassFactory, return void

Destroy_ PROC this_:ptr CClassFactory

	DebugOut "Destroy@ClassFactory"

	invoke LocalFree, this_

	dec g_DllRefCount

	ret
Destroy_ ENDP		


;--- IClassFactory interface


	.const

supInterfaces_ label dword
	dd offset IID_IUnknown,0
	dd offset IID_IClassFactory,0
IFTABSIZE_ equ ($ - supInterfaces_)/ (2 * sizeof dword)

	.code

QueryInterface_ PROC this_:ptr CClassFactory,riid:ptr IID,ppReturn:ptr

	invoke IsInterfaceSupported, riid, offset supInterfaces_, IFTABSIZE_,  this_, ppReturn
	ret

QueryInterface_ ENDP


AddRef_ PROC this_:ptr CClassFactory

	DebugOut "IClassFactory::AddRef"

	mov	eax, this_
	inc	[eax].CClassFactory.ObjRefCount
	mov	eax, [eax].CClassFactory.ObjRefCount
	ret

AddRef_ ENDP


Release_ PROC this_:ptr CClassFactory

	DebugOut "IClassFactory::Release"

	mov	eax, this_
	dec	[eax].CClassFactory.ObjRefCount

	mov eax,[eax].CClassFactory.ObjRefCount
	.if (eax == 0)
		invoke Destroy_, this_
		xor eax,eax
	.endif
	ret

Release_ ENDP


CreateInstance_ PROC pThis:ptr CClassFactory, pUnkOuter:LPUNKNOWN,
					riid:ptr IID,ppObject:ptr LPUNKNOWN

local	pObject:ptr objectname

	DebugOut "IClassFactory::CreateInstance"

	mov	eax, ppObject
	mov	DWORD PTR [eax], NULL

if ?AGGREGATION
;------------- if pUnkOuter != NULL riid MUST be IID_IUnknown!
	.if (pUnkOuter != NULL)
		invoke IsEqualGUID, riid, addr IID_IUnknown
		.if (eax == FALSE)
			DebugOut "IClassFactory::CreateInstance failed (riid != IID_IUnknown)"
			return CLASS_E_NOAGGREGATION
		.endif
	.endif
else
	.if (pUnkOuter != NULL)
		DebugOut "IClassFactory::CreateInstance failed (pUnkOuter != Null)"
		return CLASS_E_NOAGGREGATION
	.endif
endif

;------------- call constructor
	mov ecx, pThis
	mov ecx,[ecx].CClassFactory.pClass
	invoke [ecx].ObjectEntry.constructor, ecx, pUnkOuter

	.if (eax == NULL)
		DebugOut "IClassFactory::CreateInstance failed (constructor returned NULL)"
		return E_OUTOFMEMORY
	.endif

;--- constructor has returned an LPUNKNOWN

	mov pObject,eax

	invoke vf(pObject,IUnknown,QueryInterface), riid, ppObject
    push eax
	invoke vf(pObject,IUnknown,Release)
    pop eax    
	ret

CreateInstance_ ENDP


LockServer_ PROC pThis:ptr CClassFactory, bLockServer:DWORD

	DebugOut "IClassFactory::LockServer(%X)", bLockServer

    .if (bLockServer)
        inc g_DllRefCount
    .else
        dec g_DllRefCount
    .endif
	return S_OK

LockServer_ ENDP

endif

	endm


;--------------------------------------------------------------------------
;--- DEFINE_UNKNOWN: this macro defines an aggregatable IUnknown
;--------------------------------------------------------------------------

ifndef DEFINE_UNKNOWN_DEFINED

DEFINE_UNKNOWN_DEFINED equ 1

DEFINE_UNKNOWN macro objectname, iftab, iftabsize
local ?QueryInterface, ?AddRef, ?Release

	.const

	align 4

CUnknownVtbl_&objectname label IUnknownVtbl 
if ?AGGREGATION
	IUnknownVtbl {\
		_NDQueryInterface@&objectname,\
		_NDAddRef@&objectname,\
		_NDRelease@&objectname}
else
	IUnknownVtbl {\
		QueryInterface@&objectname,\
		AddRef@&objectname,\
		Release@&objectname}
endif

	.code

if ?AGGREGATION

;--- QueryInterface@&objectname, AddRef@....
;--- is jumped to from all internal interfaces

QueryInterface@&objectname PROC public this_:ptr objectname,riid:ptr IID,ppReturn:ptr ptr
	mov eax,this_
	invoke vf([eax].&objectname&.pUnkOuter, IUnknown, QueryInterface), riid, ppReturn
	ret
QueryInterface@&objectname endp

AddRef@&objectname PROC public this_:ptr objectname
	mov eax,this_
	invoke vf([eax].&objectname&.pUnkOuter, IUnknown, AddRef)
	ret
AddRef@&objectname endp

Release@&objectname PROC public this_:ptr objectname
	mov eax,this_
	invoke vf([eax].&objectname&.pUnkOuter, IUnknown, Release)
	ret
Release@&objectname endp

;--- _NDQueryInterface@&objectname, _NDAddRef@....
;--- are entries in out IUnknown vtable

_NDQueryInterface@&objectname::
	sub dword ptr [esp+4], objectname&._IUnknown
	jmp NDQueryInterface@&objectname
_NDAddRef@&objectname::
	sub dword ptr [esp+4], objectname&._IUnknown
	jmp NDAddRef@&objectname
_NDRelease@&objectname::
	sub dword ptr [esp+4], objectname&._IUnknown
	jmp NDRelease@&objectname


?QueryInterface	catstr <NDQueryInterface>,<@>,<objectname>
?AddRef			catstr <NDAddRef>,<@>,<objectname>
?Release		catstr <NDRelease>,<@>,<objectname>

else

?QueryInterface	catstr <QueryInterface>,<@>,<objectname>
?AddRef			catstr <AddRef>,<@>,<objectname>
?Release		catstr <Release>,<@>,<objectname>

endif

;*** IUnknown::QueryInterface - the real, nondelegated QueryInterface

?QueryInterface PROC public this_:ptr objectname,riid:ptr IID,ppReturn:ptr ptr

ifdef _DEBUG
local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE
endif

	invoke IsInterfaceSupported, riid, iftab, iftabsize,  this_, ppReturn
ifdef OnQueryInterface@&objectname
	invoke OnQueryInterface@&objectname, this_, riid, ppReturn
endif
ifndef ?NOQILOG
ifdef _DEBUG
    push eax
	invoke StringFromGUID2,riid, addr wszIID, LENGTHOF wszIID
	invoke wsprintf, addr szKey, CStr("Interface\%S"), addr wszIID
	invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey 
	.if (eax == ERROR_SUCCESS)
		mov dwSize, sizeof szKey
		invoke RegQueryValueEx, hKey, CStr(""), NULL, NULL, addr szKey, addr dwSize
		invoke RegCloseKey, hKey
	.else
		mov szKey,0
	.endif
    pop eax
	DebugOut "IUnknown::QueryInterface@&objectname&(%S[%s])=%X", addr wszIID, addr szKey, eax
endif
endif
	ret

?QueryInterface ENDP

;*** IUnknown::AddRef - the real, undelegated AddRef method

?AddRef PROC public this_:ptr objectname

ifdef OnAddRef@&objectname
	invoke OnAddRef@&objectname, this_
endif
	mov eax,this_
	inc [eax].&objectname&.ObjRefCount
	mov	eax, [eax].&objectname&.ObjRefCount
	ret
?AddRef ENDP		

;*** IUnknown::Release - the real, undelegated Release method

?Release PROC public this_:ptr objectname

ifdef OnRelease@&objectname
	invoke OnRelease@&objectname, this_
endif
	mov eax,this_
	dec [eax].&objectname&.ObjRefCount
	mov eax,[eax].&objectname&.ObjRefCount
	.if (eax == 0)
		invoke Destroy@&objectname&, this_
		xor eax,eax
	.endif
	ret

?Release ENDP	

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_PROVIDECLASSINFO: this macro defines an IProvideClassInfo
;--------------------------------------------------------------------------

ifndef PROVIDECLASSINFO_INCLUDED

PROVIDECLASSINFO_INCLUDED equ 1

;;externdef CProvideClassInfoVtbl:		IProvideClassInfoVtbl

DEFINE_PROVIDECLASSINFO macro objectname
local QueryInterface_, AddRef_, Release_ , GetClassInfo_

	.const

;--- IProvideClassInfo vtable

CProvideClassInfoVtbl_&objectname label IProvideClassInfoVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
    dd GetClassInfo_


    .code

;--- IProvideClassInfo interface

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IProvideClassInfo
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IProvideClassInfo
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IProvideClassInfo
	jmp Release@&objectname

GetClassInfo_:
	sub dword ptr [esp+4], objectname&._IProvideClassInfo
GetClassInfo@&objectname PROC uses __this this_:ptr objectname, ppTI:ptr LPTYPEINFO

local	pTypeLib:LPTYPELIB
local	dwIndex:DWORD

	mov __this,this_

    .IF (!ppTI)
        mov eax, E_POINTER
    .else
ifdef &objectname&_IDispatch_Flags 
ife &objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
		invoke LoadRegTypeLib,m_refIID_TypeLib,\
							m_dwMajorVer,\
							m_dwMinorVer,\
							LANG_NEUTRAL, addr pTypeLib
else
		invoke vf(m_pTypeInfo, ITypeInfo, GetContainingTypeLib), addr pTypeLib, addr dwIndex
endif
		.if (eax == S_OK)
			invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid),addr CLSID_&objectname&, ppTI
			push eax
			invoke vf(pTypeLib, ITypeLib, Release)
			pop eax
		.endif
else
		mov eax, m_pTypeInfo
		mov ecx, ppTI
		mov [ecx], eax
		.if (eax)
			invoke vf(eax, IUnknown, AddRef)
			mov eax, S_OK
		.else
			mov eax, E_FAIL
		.endif
endif
	.endif
	DebugOut "IProvideClassInfo::GetClassInfo, hr=%X", eax
    ret

GetClassInfo@&objectname ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_SPECIFYPROPERTYPAGES: defines an ISpecifyPropertyPages
;--------------------------------------------------------------------------

ifndef SPECIFYPROPERTYPAGES_INCLUDED

SPECIFYPROPERTYPAGES_INCLUDED equ 1


DEFINE_SPECIFYPROPERTYPAGES macro objectname
local QueryInterface_, AddRef_, Release_ 

;;externdef CSpecifyPropertyPagesVtbl_&objectname:	ISpecifyPropertyPagesVtbl

externdef CLSID_StockFontPage:GUID
externdef CLSID_StockColorPage:GUID
externdef CLSID_StockPicturePage:GUID

	.const

CSpecifyPropertyPagesVtbl_&objectname label ISpecifyPropertyPagesVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
    dd GetPages

;--- ISpecifyPropertyPages methods

    .code

QueryInterface_:
	sub dword ptr [esp+4], objectname&._ISpecifyPropertyPages
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._ISpecifyPropertyPages
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._ISpecifyPropertyPages
	jmp Release@&objectname


GetPages PROC  this_:ptr objectname, pPages:ptr CAUUID

	sub this_,objectname&._ISpecifyPropertyPages

	DebugOut "ISpecifyPropertyPages::GetPages"
    
    .IF (!pPages)
        return E_POINTER
    .ENDIF
;---------------------- this exist MUST exist since without it this
;---------------------- interface is pointless
	invoke GetPages@&objectname, this_, pPages
	ret

GetPages ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_VIEWOBJECT2: defines IViewObject + IViewObject2
;--------------------------------------------------------------------------


ifndef VIEWOBJECT2_INCLUDED

VIEWOBJECT2_INCLUDED equ 1


DEFINE_VIEWOBJECT2 macro objectname
local QueryInterface_, AddRef_, Release_
local Draw_, SetAdvise_, GetAdvise_, GetExtent_

;;externdef CViewObject2Vtbl_&objectname:				IViewObject2Vtbl

	.const

CViewObject2Vtbl_&objectname label IViewObject2Vtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd Draw_, GetColorSet, Freeze, Unfreeze, SetAdvise_, GetAdvise_
	dd GetExtent_

SendViewChange@&objectname	PROTO :ptr objectname

	.code

;--------------------------------------------------------------------------
;IViewObject interface
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IViewObject2
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IViewObject2
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IViewObject2
	jmp Release@&objectname

Draw_:
	sub dword ptr [esp+4], objectname&._IViewObject2
Draw$ PROC uses __this this_:ptr objectname, dwAspect:DWORD, lindex:SDWORD, 
				pvAspect:ptr, ptd:ptr, hicTargetDev:HDC, 
				hdc:HDC, pRectBounds:ptr RECT, prcWBounds:ptr RECT,
				pfnContinue:ptr ptr, dwContinue:DWORD 
LOCAL	pRect:ptr RECT

    mov __this,this_

	mov eax,pRectBounds
	.if (!eax)
        return E_INVALIDARG
	.endif
	mov pRect,eax
	DebugOut "IViewObject::Draw (%d,%d,%d,%d)",\
		[eax].RECT.left,[eax].RECT.top,[eax].RECT.right,[eax].RECT.bottom
    .IF dwAspect != DVASPECT_CONTENT
        return DV_E_DVASPECT
    .ENDIF
ifdef OnDraw@&objectname
	invoke OnDraw@&objectname, this_, ptd, hicTargetDev, hdc, pRectBounds, prcWBounds
endif
    ret

Draw$ ENDP

;--------------------------------------------------------------------------

GetColorSet PROC this_:ptr objectname, dwAspect:DWORD, lindex:SDWORD, 
                           pvAspect:ptr, ptd:ptr, hicTargetDev:HDC, 
                           ppColorSet:ptr ptr LOGPALETTE

	DebugOut "IViewObject::GetColorSet"
    return E_NOTIMPL

GetColorSet	ENDP

;--------------------------------------------------------------------------

Freeze PROC this_:ptr objectname, dwAspect:DWORD, lindex:SDWORD, 
                           pvAspect:ptr, pdwFreeze:ptr DWORD 
	DebugOut "IViewObject::Freeze"
    return E_NOTIMPL

Freeze ENDP

;--------------------------------------------------------------------------

Unfreeze PROC this_:ptr objectname, dwFreeze:DWORD

	DebugOut "IViewObject::Unfreeze"
    return E_NOTIMPL

Unfreeze ENDP

;--------------------------------------------------------------------------

SetAdvise_:
	sub dword ptr [esp+4], objectname&._IViewObject2
SetAdvise$ PROC uses __this this_:ptr objectname, aspect:DWORD, advf:DWORD, pViewAdviseSink:LPADVISESINK

    mov __this,this_

	DebugOut "IViewObject::SetAdvise(%u, %u, %X)", aspect, advf, pViewAdviseSink
    .IF aspect != DVASPECT_CONTENT
        mov eax, DV_E_DVASPECT
        ret
    .ENDIF
	invoke ComPtrAssign,addr m_pAdviseSink, pViewAdviseSink
    mov eax, advf
    mov m_advf, eax
    .IF (eax & ADVF_PRIMEFIRST)
        invoke SendViewChange@&objectname, __this
    .ENDIF
    return S_OK

SetAdvise$ ENDP

;--------------------------------------------------------------------------

GetAdvise_:
	sub dword ptr [esp+4], objectname&._IViewObject2
GetAdvise$ PROC uses __this this_:ptr objectname, pdwAspects:ptr DWORD, pAdvf:ptr DWORD, ppViewAdviseSink:ptr LPADVISESINK

    mov __this,this_

	DebugOut "IViewObject::GetAdvise"

    mov ecx,ppViewAdviseSink
    .IF (!ecx)
        return E_POINTER
    .ENDIF

	mov dword ptr [ecx],0
	invoke ComPtrAssign, ecx, m_pAdviseSink

    mov ecx, pdwAspects
    mov eax, [ecx]
    .IF eax == DVASPECT_CONTENT
        mov ecx, pAdvf
        mov eax, [ecx]
        .IF (eax)
            mov eax, m_advf
            mov [ecx], eax
        .ENDIF
    .ENDIF
    return S_OK

GetAdvise$ ENDP	

;--------------------------------------------------------------------------
;IViewObject2 (only 1 extra method)
;--------------------------------------------------------------------------

GetExtent_:
	sub dword ptr [esp+4], objectname&._IViewObject2
GetExtent$ PROC this_:ptr objectname, aspect:DWORD, lindex:SDWORD, 
                           ptd:ptr, pSize:ptr SIZEL 

	DebugOut "IViewObject2::GetExtent"

    mov ecx, pSize 
    .IF (!ecx)
        mov eax, E_POINTER
    .ELSEIF (aspect != DVASPECT_CONTENT) 
        mov eax, DV_E_DVASPECT
	.ELSE
	    mov edx,this_
		mov eax, [edx].&objectname&.himetricExtent.cx_
	    mov [ecx].SIZEL.cx_, eax    
		mov eax, [edx].&objectname&.himetricExtent.cy
	    mov [ecx].SIZEL.cy, eax    
		DebugOut "IViewObject2::GetExtent returns x=%u, y=%u",\
			[ecx].SIZEL.cx_, [ecx].SIZEL.cy
		mov eax, S_OK
    .ENDIF
	ret

GetExtent$ ENDP

	endm

endif

;--------------------------------------------------------------------------
;--- DEFINE_COMHELPER: only 1 instance of this code is needed for a 
;--- module. So all procs are declared public
;--------------------------------------------------------------------------

DEFINE_COMHELPER macro

ifndef COMHELPER_INCLUDED

COMHELPER_INCLUDED equ 1

ifndef lstrlenW
lstrlenW	proto pStr:ptr word
endif

	.code

;--------------------------------------------------------------------------

;*** scan interface tab and see if requested iface is in there

IsInterfaceSupported proc public uses ebx esi edi pReqIF:ptr IID, pIFTab:ptr ptr IID, dwEntries:dword, pThis:ptr, ppReturn:ptr LPUNKNOWN
	
	mov ecx,dwEntries
	mov esi,pIFTab
	mov ebx,0
	.while (ecx)
		lodsd
		mov edi,eax
		lodsd
		mov edx,eax
		mov eax,esi
		mov esi,pReqIF
		push ecx
		mov ecx,4
		repz cmpsd
		pop ecx
		.if (ZERO?)
			mov ebx,edx
			add ebx,pThis
			.break
		.endif
		mov esi,eax
		dec ecx
	.endw
	mov ecx,ppReturn
	mov [ecx],ebx

	.if (ebx)
		invoke vf(ebx,IUnknown,AddRef)
		mov eax,S_OK
	.else
		mov eax,E_NOINTERFACE
	.endif
	ret

IsInterfaceSupported endp

;--------------------------------------------------------------------------

ComPtrAssign proc public uses ebx pp:ptr LPUNKNOWN, lp:LPUNKNOWN
	.if (lp != NULL)
		invoke vf(lp,IUnknown,AddRef)
	.endif
	mov ebx,pp
	.if (dword ptr [ebx])
		invoke vf([ebx],IUnknown,Release)
	.endif
	mov eax,lp
	mov [ebx],eax
	ret
ComPtrAssign endp

;--------------------------------------------------------------------------

WideCharToLocal PROC public pLocal:LPSTR, pWide:LPWSTR, dwChars:DWORD

	mov	eax, pLocal
	mov	BYTE PTR [eax], 0

ifdef UNICODE
	invoke lstrcpyn, pLocal, pWide, dwChars
else
	invoke WideCharToMultiByte, CP_ACP, 0, pWide, -1, pLocal, dwChars, NULL, NULL
endif

	invoke lstrlen, pLocal

	ret
WideCharToLocal ENDP

;--------------------------------------------------------------------------

LocalToWideChar PROC public pWide:LPWSTR, pLocal:LPSTR,  dwChars:DWORD

	mov	eax, pWide
	mov	WORD PTR [eax], 0

ifdef UNICODE
	invoke lstrcpyn, pWide, pLocal, dwChars
else
	invoke MultiByteToWideChar, CP_ACP, 0, pLocal, -1, pWide, dwChars
endif

	invoke lstrlenW, pWide

	ret
LocalToWideChar ENDP

if ?EVENTSUPPORT
;;	DEFINE_FIREEVENTHELPER
endif

endif

	endm


;--------------------------------------------------------------------------
;--- DEFINE_GETCLASSOBJECT: defines proc DllGetClassObject (which has to be
;--- exported)
;--------------------------------------------------------------------------

REGSTRUCT struct
lpszSubKey		LPSTR  ? ;
lpszValueName   LPSTR  ? ;
lpszData		LPSTR  ? ;
REGSTRUCT ends

LPREGSTRUCT typedef ptr REGSTRUCT

;*** definition of CEventSink class (used by CConnectionPoint)

LPEVENTSINK typedef ptr CEventSink

CEventSink struct
m_pNext		LPEVENTSINK ?
m_pEvent	LPUNKNOWN	?
CEventSink ends

protoConstructor typedef proto :ptr ObjectEntry, :LPUNKNOWN
LPCONSTRUCTOR typedef ptr protoConstructor

ObjectEntry struct
pClsId		REFGUID ?
pLibId		REFGUID ?
dwVerMajor	SWORD ?
dwVerMinor	SWORD ?	
pRegKeys	LPREGSTRUCT ?
constructor	LPCONSTRUCTOR ?
ObjectEntry ends

DEFINE_GETCLASSOBJECT macro pClsTab

Create@CClassFactory	PROTO :ptr ObjectEntry

	.code

DllGetClassObject PROC public uses esi rclsid:ptr CLSID,riid:ptr IID,ppReturn:ptr

local pClassFactory:ptr IClassFactory
local hResult:dword

	mov	eax, ppReturn
	mov	DWORD PTR [eax], 0

	mov esi, pClsTab
	mov ecx, ??ObjectMapItems
	.while (ecx)
		push ecx
		invoke	IsEqualGUID, rclsid, [esi].ObjectEntry.pClsId
		pop ecx
		.break .if (eax)
		add esi, sizeof ObjectEntry
		dec ecx
	.endw
	.if (!ecx)
		return CLASS_E_CLASSNOTAVAILABLE
	.endif

;;	invoke	Create@CClassFactory, esi].ObjectEntry.constructor
	invoke	Create@CClassFactory, esi
	.if (eax == NULL)
		return E_OUTOFMEMORY
	.endif
	mov pClassFactory,eax

	invoke vf(pClassFactory,IClassFactory,QueryInterface),riid,ppReturn
	mov	hResult, eax

	invoke vf(pClassFactory,IClassFactory,Release)

	return hResult

DllGetClassObject ENDP

	endm


;--------------------------------------------------------------------------
;--- DEFINE_CATEGORIZEPROPERTIES: defines ICategorizeProperties
;--------------------------------------------------------------------------

Category struct
dwID		dword ?
pszName		dword ?
pDispArray	dword ?
cntArray	dword ?
Category ends


ifndef ICategorizePropertiesVtbl

;;externdef stdcall IID_ICategorizeProperties:IID

BEGIN_INTERFACE ICategorizeProperties
    ; IUnknown methods
    STDMETHOD       QueryInterface, riid:ptr IID, ppv:DWORD
    STDMETHOD       AddRef
    STDMETHOD       Release
    ; IServiceProvider methods
    STDMETHOD       MapPropertyToCategory	,:DISPID, :ptr DWORD
    STDMETHOD       GetCategoryName			,:DWORD, :LCID, :ptr BSTR
END_INTERFACE

LPCATEGORIZEPROPERTIES typedef ptr ICategorizeProperties
endif


ifndef CATEGORIZE_PROPERTIES_DEFINED

CATEGORIZE_PROPERTIES_DEFINED equ 1


DEFINE_CATEGORIZE_PROPERTIES macro objectname, cattab, numcat
local QueryInterface_, AddRef_, Release_ 

;;externdef CCategorizePropertiesVtbl_&objectname:	ICategorizePropertiesVtbl

; ICategorizeProperties interface
;----------------------------------------------------------------------

    .const

CCategorizePropertiesVtbl_&objectname ICategorizePropertiesVtbl {\
	QueryInterface_, AddRef_, Release_,\
	MapPropertyToCategory, GetCategoryName}

;*** this IID isnt defined in older standard Libs, so define it here

IID_ICategorizeProperties GUID <04D07FC10H, 0F931H, 011CEH, <0B0H, 001H, 000H, 0AAH, 000H, 068H, 084H, 0E5H>>

    .code

;--------------------------------------------------------------------------
;ICategorizeProperties interface
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._ICategorizeProperties
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._ICategorizeProperties
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._ICategorizeProperties
	jmp Release@&objectname


MapPropertyToCategory PROC uses ebx edi this_:ptr objectname, dispid:DISPID, ppropcat:ptr DWORD

	DebugOut "ICategorizeProperties::MapPropertyToCategory, DispID=%X", dispid

	mov eax,dispid
    mov ebx, cattab
	mov edx, numcat
	.while (edx)
		mov edi,[ebx].Category.pDispArray
		mov ecx,[ebx].Category.cntArray
		repnz scasd
		.if (ZERO?)
			mov ecx, ppropcat
			mov eax,[ebx].Category.dwID
			mov [ecx], eax
			return S_OK
		.endif
		add ebx,sizeof Category
		dec edx	
	.endw
    return E_FAIL

MapPropertyToCategory ENDP

;--------------------------------------------------------------------------

GetCategoryName PROC this_:ptr objectname, propcat:DWORD, lcid:LCID, pbstrName:ptr BSTR

LOCAL Buf[MAX_PATH]:BYTE
LOCAL wBuf[MAX_PATH]:WORD

	DebugOut "ICategorizeProperties::GetCategoryName"

	mov eax, propcat
    mov edx, cattab
	mov ecx, numcat
	.while (ecx)
		.if (eax == [edx].Category.dwID)
			invoke lstrcpy, addr Buf, [edx].Category.pszName
			invoke LocalToWideChar, ADDR wBuf, ADDR Buf, MAX_PATH
			invoke SysAllocString, ADDR wBuf
			mov ecx, pbstrName 
			mov [ecx], eax
			return S_OK
	    .ENDIF
		add edx,sizeof Category
		dec ecx
	.endw
    return E_FAIL

GetCategoryName ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_OLEINPLACEOBJECT: defines IOleInPlaceObject + IOleWindow
;--------------------------------------------------------------------------

ifndef OLEINPLACEOBJECT_DEFINED

OLEINPLACEOBJECT_DEFINED equ 1


DEFINE_OLEINPLACEOBJECT macro objectname
local QueryInterface_, AddRef_, Release_
local GetWindow_, ContextSensitiveHelp_
local InPlaceDeactivate_, InPlaceDeactivate$, UIDeactivate_, UIDeactivate$
local SetObjectRects_, ReactivateAndUndo_

;;externdef COleInPlaceObjectVtbl_&objectname:	IOleInPlaceObjectVtbl

	.const

COleInPlaceObjectVtbl_&objectname label IOleInPlaceObjectVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetWindow_, ContextSensitiveHelp_
	dd InPlaceDeactivate_, UIDeactivate_, SetObjectRects_
	dd ReactivateAndUndo_

	.code

;--------------------------------------------------------------------------
;IOleWindow interface (also used by IOleInPlaceObject interface)
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
	jmp Release@&objectname

;--------------------------------------------------------------------------

GetWindow_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
GetWindow@&objectname proc this_:ptr objectname,phWnd:ptr HWND

    mov ecx,this_

	DebugOut "IOleWindow::GetWindow, hWnd=%X",[ecx].&objectname&.hWnd

	mov eax,[ecx].&objectname&.hWnd
	mov ecx,phWnd
	mov [ecx],eax
	.if (eax)
		mov eax,S_OK
	.else
		mov eax,E_UNEXPECTED
	.endif
	ret

GetWindow@&objectname endp

;--------------------------------------------------------------------------

ContextSensitiveHelp_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
ContextSensitiveHelp@&objectname proc this_:ptr objectname,fEnterMode:dword

	DebugOut "IOleWindow::ContextSensitiveHelp"
	return E_NOTIMPL

ContextSensitiveHelp@&objectname endp


;--------------------------------------------------------------------------
;IOleInPlaceObject interface
;--------------------------------------------------------------------------

UIDeactivate_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
UIDeactivate$ proc uses __this this_:ptr objectname

LOCAL pOleInPlaceSite:LPOLEINPLACESITE
LOCAL pOleInPlaceFrame:LPOLEINPLACEFRAME

    mov __this,this_

	DebugOut "IOleInPlaceObject::UIDeactivate"
	.if ((m_dwFlags & OLEOBJF_UIACTIVATED) && m_pOleClientSite)
   		invoke vf(m_pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleInPlaceFrame, ADDR pOleInPlaceFrame
    	.if (eax == S_OK)
   			invoke vf(pOleInPlaceFrame,IOleInPlaceFrame,SetActiveObject), NULL, NULL
    		invoke vf(pOleInPlaceFrame,IOleInPlaceFrame,Release)
	    .endif
		invoke vf(m_pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleInPlaceSite, ADDR pOleInPlaceSite
		.if (eax == S_OK)
			invoke vf(pOleInPlaceSite,IOleInPlaceSite,OnUIDeactivate), FALSE
			invoke vf(pOleInPlaceSite,IOleInPlaceSite,Release)
		.endif
		and m_dwFlags, not OLEOBJF_UIACTIVATED
	.endif

	return S_OK

UIDeactivate$ endp

;---

InPlaceDeactivate_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
InPlaceDeactivate$ proc uses __this this_:ptr objectname

LOCAL pOleInPlaceSite:LPOLEINPLACESITE

	DebugOut "IOleInPlaceObject::InPlaceDeactivate"

    mov __this,this_

	invoke UIDeactivate$, __this

;------------------------ delete the window
    .if (m_hWnd)
	    invoke DestroyWindow,m_hWnd
        mov m_hWnd,NULL
    .endif

;------------------------ tell container we are deactivated
	.if (m_pOleClientSite)
		invoke vf(m_pOleClientSite, IOleInPlaceSite, QueryInterface), addr IID_IOleInPlaceSite, ADDR pOleInPlaceSite
		.if (eax == S_OK)
			invoke vf(pOleInPlaceSite,IOleInPlaceSite,OnInPlaceDeactivate)
			invoke vf(pOleInPlaceSite,IOleInPlaceSite,Release)
		.endif
	.endif
	return S_OK

InPlaceDeactivate$ endp

;--------------------------------------------------------------------------

SetObjectRects_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
SetObjectRects$ proc uses esi __this this_:ptr objectname,lprcPosRect:ptr RECT,lprcClipRect:ptr RECT

    mov __this,this_

	mov esi,lprcPosRect
	.if (esi)
		mov ecx,[esi].RECT.right
		sub ecx,[esi].RECT.left
	    mov m_pixelExtent.cx_,ecx
		mov edx,[esi].RECT.bottom
		sub edx,[esi].RECT.top
	    mov m_pixelExtent.cy,edx
		.if (m_hWnd)
			invoke SetWindowPos, m_hWnd, NULL,\
				[esi].RECT.left, [esi].RECT.top,\
				ecx, edx, SWP_NOZORDER or SWP_NOACTIVATE
;;			invoke InvalidateRect, m_hWnd, 0, 0
		.endif
ifdef _DEBUG
		mov eax,lprcPosRect
		mov ecx,lprcClipRect
		.if (!ecx)
			mov ecx,eax
		.endif
		DebugOut "IOleInPlaceObject::SetObjectRects([%d,%d,%d,%d],[%d,%d,%d,%d])",\
			[eax].RECT.left,[eax].RECT.top,[eax].RECT.right,[eax].RECT.bottom,\
			[ecx].RECT.left,[ecx].RECT.top,[ecx].RECT.right,[ecx].RECT.bottom
endif
	.endif

	return S_OK

SetObjectRects$ endp

;--------------------------------------------------------------------------

ReactivateAndUndo_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceObject
ReactivateAndUndo$ proc this_:ptr objectname

	DebugOut "IOleInPlaceObject::ReactivateAndUndo"
    mov eax, INPLACE_E_NOTUNDOABLE
	ret

ReactivateAndUndo$ endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_OLEINPLACEACTIVEOBJECT: defines IOleInPlaceActiveObject
;--------------------------------------------------------------------------


ifndef OLEINPLACEACTIVEOBJECT_DEFINED

OLEINPLACEACTIVEOBJECT_DEFINED equ 1


DEFINE_OLEINPLACEACTIVEOBJECT macro objectname
local QueryInterface_, AddRef_, Release_, GetWindow_, ContextSensitiveHelp_

;;externdef COleInPlaceActiveObjectVtbl_&objectname:	IOleInPlaceActiveObjectVtbl

	.const

COleInPlaceActiveObjectVtbl_&objectname label IOleInPlaceActiveObjectVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetWindow_, ContextSensitiveHelp_
	dd TranslateAccelerator_, OnFrameWindowActivate
	dd OnDocWindowActivate, ResizeBorder, EnableModeless

	.code

;--------------------------------------------------------------------------
;IOleInPlaceActiveObject interface
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceActiveObject
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceActiveObject
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceActiveObject
	jmp Release@&objectname

GetWindow_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceActiveObject
	jmp GetWindow@&objectname

ContextSensitiveHelp_:
	sub dword ptr [esp+4], objectname&._IOleInPlaceActiveObject
	jmp ContextSensitiveHelp@&objectname

TranslateAccelerator_ proc this_:ptr objectname, lpMsg:ptr MSG

;	@AdjustThis

	return S_FALSE

TranslateAccelerator_ endp

OnFrameWindowActivate proc this_:ptr objectname, fActivate:dword

;	@AdjustThis

	return S_OK

OnFrameWindowActivate endp

OnDocWindowActivate proc this_:ptr objectname, fActivate:dword

;	@AdjustThis

	return S_OK

OnDocWindowActivate endp

ResizeBorder proc this_:ptr objectname, lpRect:ptr RECT, pUIWindow: ptr,fFrameWindow:dword

;	@AdjustThis
	return S_OK

ResizeBorder endp

EnableModeless proc this_:ptr objectname, fActivate:dword

;	@AdjustThis
	return S_OK

EnableModeless endp


	endm

endif



;--------------------------------------------------------------------------
;--- DEFINE_PERSISTSTORAGE: defines IPersist + IPersistStorage
;--------------------------------------------------------------------------

ifndef PERSISTSTORAGE_DEFINED

PERSISTSTORAGE_DEFINED equ 1


DEFINE_PERSISTSTORAGE macro objectname
local QueryInterface_, AddRef_, Release_
local GetClassID_, IsDirty_, InitNew_, Load_, Save_, wszname
local SaveCompleted_, HandsOffStorage_

;;externdef CPersistStorageVtbl_&objectname:			IPersistStorageVtbl

	.const

CPersistStorageVtbl_&objectname label IPersistStorageVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetClassID_, IsDirty_, InitNew_, Load_, Save_
	dd SaveCompleted_, HandsOffStorage_

wszName	dw L("Contents"),0

	.code

;--------------------------------------------------------------------------
;IPersist 
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
	jmp Release@&objectname


GetClassID_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
GetClassID PROC uses esi edi this_:ptr objectname, pClassID:ptr GUID

	DebugOut "IPersist::GetClassID"

	mov edi,pClassID
	.IF (!edi)
		mov eax, E_POINTER
	.else
		mov esi, offset CLSID_&objectname
		movsd
		movsd
		movsd
		movsd
		xor eax, eax			; return S_OK	 
	.ENDIF
	ret

GetClassID ENDP

;--------------------------------------------------------------------------
;IPersistStorage 
;--------------------------------------------------------------------------

IsDirty_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
IsDirty@&objectname proc this_:ptr objectname

	DebugOut "IPersistStorage::IsDirty"
	mov ecx,this_
	lea ecx,[ecx].&objectname&._IPersistStreamInit
	invoke vf( ecx, IPersistStreamInit, IsDirty)
	ret

IsDirty@&objectname endp

;--------------------------------------------------------------------------

Load_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
Load@&objectname proc this_:ptr objectname, pStorage:LPSTORAGE

LOCAL pStream:LPSTREAM

ifdef _DEBUG
local statstg:STATSTG
	invoke vf(pStorage, IStorage, Stat), addr statstg, STATFLAG_NONAME
	DebugOut "IPersistStorage::Load(%X), statstg.grfMode=%X", pStorage, statstg.grfMode
endif
	
	invoke vf(pStorage, IStorage, OpenStream), offset wszName,\
			NULL, STGM_READ or STGM_SHARE_EXCLUSIVE, NULL, addr pStream
	.if (EAX == S_OK)
		mov ecx,this_
		lea ecx,[ecx].&objectname&._IPersistStreamInit
		invoke vf(ecx, IPersistStreamInit, Load), pStream
		push eax
		invoke vf(pStream,IStream,Release)
		pop eax
ifdef _DEBUG
	.else
		DebugOut "IPersistStorage::Load IStorage::OpenStream failed hr=%X",eax
endif
	.endif
	DebugOut "IPersistStorage::Load exit, hr=%X",eax
	ret

Load@&objectname endp

;--------------------------------------------------------------------------

Save_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
Save@&objectname proc this_:ptr objectname,pStorage:LPSTORAGE, fSameAsLoad:dword

LOCAL pStream:LPSTREAM

ifdef _DEBUG
local statstg:STATSTG
	invoke vf(pStorage, IStorage, Stat), addr statstg, STATFLAG_NONAME
	DebugOut "IPersistStorage::Save(%X, %X), statstg.grfMode=%X", pStorage, fSameAsLoad, statstg.grfMode
endif

	invoke vf(pStorage, IStorage, CreateStream), offset wszName,\
		STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE,\
		NULL,NULL,addr pStream
	.if (EAX == S_OK)
		mov ecx,this_
		lea ecx,[ecx].&objectname&._IPersistStreamInit
		invoke vf(ecx, IPersistStreamInit, Save), pStream, TRUE
		push eax
		invoke vf(pStream, IStream, Release)
		pop eax 
	.endif
	DebugOut "IPersistStorage::Save exit, hr=%X",eax
	ret

Save@&objectname endp

;--------------------------------------------------------------------------

InitNew_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
InitNew@&objectname proc this_:ptr objectname, pStorage:LPSTORAGE

	DebugOut "IPersistStorage::InitNew(%X)", pStorage

;	invoke WriteFmtUserTypeStg, pStorage, CF_TEXT, usertype

	mov ecx,this_
	lea ecx,[ecx].&objectname&._IPersistStreamInit
	invoke vf(ecx, IPersistStreamInit, InitNew)

	ret

InitNew@&objectname endp

;--------------------------------------------------------------------------

SaveCompleted_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
SaveCompleted@&objectname proc this_:ptr objectname,pStorage:LPSTORAGE

	DebugOut "IPersistStorage::SaveCompleted(%X)", pStorage
	return S_OK

SaveCompleted@&objectname endp

;--------------------------------------------------------------------------

HandsOffStorage_:
	sub dword ptr [esp+4], objectname&._IPersistStorage
HandsOffStorage@&objectname proc this_:ptr objectname

	DebugOut "IPersistStorage::HandsOffStorage"
	return S_OK

HandsOffStorage@&objectname endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_PERSISTSTREAMINIT: defines IPersistStreamInit
;--------------------------------------------------------------------------

ifndef PERSISTSTREAMINIT_DEFINED

PERSISTSTREAMINIT_DEFINED equ 1

DEFINE_PERSISTSTREAMINIT macro objectname, bagtab, NUMBAGS
local QueryInterface_, AddRef_, Release_
local GetClassID_, IsDirty_, Load_, Save_, GetSizeMax_, InitNew_
local IsDirty$, Load$, Save$, InitNew$

;;externdef CPersistStreamInitVtbl_&objectname:		IPersistStreamInitVtbl

SendViewChange@&objectname		PROTO :ptr objectname

	.const

CPersistStreamInitVtbl_&objectname label IPersistStreamInitVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetClassID_, IsDirty_, Load_, Save_, GetSizeMax_, InitNew_

	.code

;--------------------------------------------------------------------------
;IPersistStreamInit 
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
	jmp Release@&objectname


GetClassID_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
	jmp GetClassID

IsDirty_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
IsDirty$ PROC this_:ptr objectname

	DebugOut "IPersistStreamInit::IsDirty"
	mov ecx,this_
	.IF ([ecx].&objectname&.isDirty)
		xor eax, eax	; mov eax, S_OK
	.ELSE
		mov eax, S_FALSE
	.ENDIF
	ret

IsDirty$ ENDP

;--------------------------------------------------------------------------

Load_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
Load$ PROC uses __this esi this_:ptr objectname, pStream:ptr IStream

local bCount:DWORD
local dwSize:DWORD
local clsid:CLSID

	DebugOut "IPersistStreamInit::Load"

	mov __this,this_

	mov esi,bagtab	
	mov ecx,NUMBAGS
	.while (ecx)
		push ecx
		mov ax,[esi].BagEntry.varType
		.if (ax == VT_I2)
			mov dl,2
		.else
			mov dl,4
		.endif
		movzx edx,dl
		mov dwSize,edx
		mov ecx,[esi].BagEntry.dwOffset
		add ecx, __this
;----------------------------------- 21.3.2004: VT_DISPATCH case added
		.if (ax == VT_DISPATCH)
			invoke OleLoadFromStream, pStream, addr IID_IDispatch, ecx 
			DebugOut " OleLoadFromStream returned %X", eax
		.else
			invoke vf(pStream, IStream, Read), ecx, dwSize, ADDR bCount
		.endif
ifdef _DEBUG
		mov ecx,[esi].BagEntry.dwOffset
		add ecx,__this
		DebugOut "loaded %u bytes: %X", bCount, dword ptr [ecx]
endif
		add esi,sizeof BagEntry
		pop ecx
		dec ecx
	.endw
	
	invoke SendViewChange@&objectname, __this

	return S_OK

Load$ ENDP

;--------------------------------------------------------------------------

Save_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
Save$ PROC uses ebx esi this_:ptr objectname, pStream:ptr IStream, fClearDirty:DWORD

LOCAL	bCount:DWORD
local	dwSize:dword
local	pPersistStream:ptr IPERSISTSTREAM

	DebugOut "IPersistStreamInit::Save enter"

	mov __this,this_

	.IF fClearDirty
		mov m_isDirty, FALSE
	.ENDIF

	mov esi, bagtab
	mov ecx,NUMBAGS
	.while (ecx)
		push ecx
		mov ax,[esi].BagEntry.varType
		.if (ax == VT_I2)
			mov dl,2
		.else
			mov dl,4
		.endif
		movzx edx,dl
		mov dwSize,edx
		mov ecx,[esi].BagEntry.dwOffset
		add ecx, __this
;----------------------------------- 21.3.2004: VT_DISPATCH case added
		.if (ax == VT_DISPATCH)
        	mov ecx, [ecx]
			invoke vf(ecx, IUnknown, QueryInterface), addr IID_IPersistStream, addr pPersistStream
			DebugOut " QueryInterface(IPersistStream) returned %X", eax
			.if (eax == S_OK)
				invoke OleSaveToStream, pPersistStream, pStream
				DebugOut " OleSaveToStream returned %X", eax
				invoke vf(pPersistStream, IUnknown, Release)
			.endif
		.else
			invoke vf(pStream, IStream, Write), ecx, dwSize, ADDR bCount
		.endif
ifdef _DEBUG
		mov ecx, [esi].BagEntry.dwOffset
		add ecx,__this
		DebugOut "saved %u bytes: %X", bCount, dword ptr [ecx]
endif
		add esi,sizeof BagEntry
		pop ecx
		dec ecx
	.endw

	DebugOut "IPersistStreamInit::Save exit"

	return S_OK

Save$ ENDP

;--------------------------------------------------------------------------

GetSizeMax_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
GetSizeMax@&objectname PROC this_:ptr objectname, pSize:ptr QWORD 

	DebugOut "IPersistStreamInit::GetSizeMax"

	mov eax, pSize
	.IF !eax
		return E_POINTER
	.ENDIF
ifdef OnGetSizeMax@&objectname
	invoke OnGetSizeMax@&objectname, this_, pSize
else
	mov dword ptr [eax+0], ?SAVESIZE
	mov dword ptr [eax+4], 0
endif
	return S_OK

GetSizeMax@&objectname ENDP

;--------------------------------------------------------------------------

InitNew_:
	sub dword ptr [esp+4], objectname&._IPersistStreamInit
InitNew$ PROC uses __this this_:ptr objectname

	DebugOut "IPersistStreamInit::InitNew"

	mov __this,this_

	.if (m_dwFlags & OLEOBJF_INITIALIZED)
		return E_UNEXPECTED
	.endif
ifdef OnInitNew@&objectname
	invoke OnInitNew@&objectname, __this
endif
	
	mov m_isDirty, TRUE

	.if (eax == S_OK)
		or m_dwFlags, OLEOBJF_INITIALIZED
	.endif
	ret

InitNew$ ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_PERSISTPROPERTYBAG: defines IPersistPropertyBag
;--------------------------------------------------------------------------

ifndef PERSISTPROPERTYBAG_DEFINED

BagEntry struct
varType		VARTYPE ?		;type of property (variant type!)
dwOffset	dword	?		;offset in class structure
pwszName	LPWSTR	?		;name of property (for vb, in dbc)
BagEntry ends


PERSISTPROPERTYBAG_DEFINED equ 1

DEFINE_PERSISTPROPERTYBAG macro objectname, bagtab, NUMBAGS
local QueryInterface_, AddRef_, Release_
local GetClassID_, InitNew_, Load_, Save_
local InitNew$, Load$, Save$

;;externdef CPersistPropertyBagVtbl_&objectname:		IPersistPropertyBagVtbl

	.const

CPersistPropertyBagVtbl_&objectname label IPersistPropertyBagVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetClassID_, InitNew_, Load_, Save_

	.code
;--------------------------------------------------------------------------
;IPersistPropertyBag   
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
	jmp Release@&objectname
GetClassID_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
	jmp GetClassID


InitNew_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
InitNew$ PROC this_:ptr objectname

	DebugOut "IPersistPropertyBag::InitNew"
	mov ecx,this_
if 1
	lea ecx,[ecx].&objectname&._IPersistStreamInit
	invoke vf(ecx, IPersistStreamInit, InitNew) 
else
	invoke vf(addr [ecx].&objectname&.PersistStreamInit, IPersistStreamInit, InitNew) 
endif
	ret
InitNew$ ENDP


;--------------------------------------------------------------------------

Load_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag
Load$ PROC uses __this esi this_:ptr objectname, pPropBag:LPPROPERTYBAG, pErrorLog:LPERRORLOG

LOCAL var:VARIANT

	DebugOut "IPersistPropertyBag::Load(%X, %X)", pPropBag, pErrorLog

	mov __this,this_
	
	mov esi, bagtab	
	mov ecx,NUMBAGS
	.while (ecx)
		push ecx
		invoke VariantInit, addr var
		mov ax,[esi].BagEntry.varType
		mov var.vt, ax
		mov var.lVal, 0
		DebugOut " calling IPropertyBag::Read(%S)", [esi].BagEntry.pwszName
		invoke vf(pPropBag, IPropertyBag, Read), [esi].BagEntry.pwszName, ADDR var, NULL
		.if SUCCEEDED(eax)
			mov eax, var.lVal
			mov edx,[esi].BagEntry.dwOffset
			.if (var.vt == VT_I2)
				mov word ptr [ebx+edx], ax
			.else
				mov dword ptr [ebx+edx], eax
			.endif
		.endif
		add esi,sizeof BagEntry
		pop ecx
		dec ecx
	.endw

	return S_OK

Load$ ENDP

;--------------------------------------------------------------------------

Save_:
	sub dword ptr [esp+4], objectname&._IPersistPropertyBag

Save$ PROC uses __this esi this_:ptr objectname, pPropBag:LPPROPERTYBAG, fClearDirty:DWORD, fSaveAllProperties:DWORD

LOCAL var:VARIANT

	DebugOut "IPersistPropertyBag::Save(%X, %X, %X)", pPropBag, fClearDirty, fSaveAllProperties

	mov __this,this_

	.IF fClearDirty
		mov m_isDirty, FALSE
	.ENDIF	  

	mov esi, bagtab
	mov ecx,NUMBAGS
	.while (ecx)
		push ecx
		invoke VariantInit, addr var
		mov ax,[esi].BagEntry.varType
		mov var.vt, ax
		mov edx,[esi].BagEntry.dwOffset
		mov eax,dword ptr [ebx+edx]
		mov var.lVal, eax
		DebugOut " calling IPropertyBag::Write(%S)", [esi].BagEntry.pwszName
		invoke vf(pPropBag, IPropertyBag, Write), [esi].BagEntry.pwszName, ADDR var
		add esi,sizeof BagEntry
		pop ecx
		dec ecx
	.endw

	return S_OK

Save$ ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_OLECONTROL: defines IOleControl
;--------------------------------------------------------------------------


ifndef OLECONTROL_DEFINED

OLECONTROL_DEFINED equ 1


DEFINE_OLECONTROL macro objectname
local QueryInterface, AddRef, Release
local GetControlInfo_, OnMnemonic_, OnAmbientPropertyChange_, FreezeEvents_ 

;;externdef COleControlVtbl_&objectname:	IOleControlVtbl

	.const

COleControlVtbl_&objectname label IOleControlVtbl
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd GetControlInfo_, OnMnemonic_, OnAmbientPropertyChange_, FreezeEvents_

    .code

;--------------------------------------------------------------------------
; IOleControl
;--------------------------------------------------------------------------

QueryInterface:
	sub dword ptr [esp+4], objectname&._IOleControl
	jmp QueryInterface@&objectname
AddRef:
	sub dword ptr [esp+4], objectname&._IOleControl
	jmp AddRef@&objectname
Release:
	sub dword ptr [esp+4], objectname&._IOleControl
	jmp Release@&objectname


GetControlInfo_:
	sub dword ptr [esp+4], objectname&._IOleControl

GetControlInfo PROC uses __this this_:ptr objectname, pCI:ptr CONTROLINFO

	mov __this,this_

	DebugOut "IOleControl::GetControlInfo"

    return E_NOTIMPL

GetControlInfo ENDP

OnMnemonic_:
	sub dword ptr [esp+4], objectname&._IOleControl

OnMnemonic@&objectname PROC uses __this this_:ptr objectname, pMsg:ptr MSG

	mov __this,this_

	DebugOut "IOleControl::OnMnemonic"

    return S_OK

OnMnemonic@&objectname ENDP


OnAmbientPropertyChange_:
	sub dword ptr [esp+4], objectname&._IOleControl

OnAmbientPropertyChange@&objectname PROC uses __this this_:ptr objectname, dispID:DISPID

	mov __this,this_

	DebugOut "IOleControl::OnAmbientPropertyChange(%d)", dispID

    return S_OK

OnAmbientPropertyChange@&objectname ENDP

FreezeEvents_:
	sub dword ptr [esp+4], objectname&._IOleControl

FreezeEvents@&objectname PROC uses __this this_:ptr objectname, bFreeze:BOOL

	mov __this,this_

	DebugOut "IOleControl::FreezeEvents(%u)", bFreeze

    return S_OK

FreezeEvents@&objectname ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_DATAOBJECT: defines IDataObject
;--------------------------------------------------------------------------


ifndef DATAOBJECT_DEFINED

DATAOBJECT_DEFINED equ 1


DEFINE_DATAOBJECT macro objectname
local QueryInterface, AddRef, Release
local GetData_, GetDataHere_, QueryGetData_, GetCanonicalFormatEtc_
local SetData_, EnumFormatEtc_, DAdvise_, DUnadvise_, EnumDAdvise_

	.const

if 0
CDataObjectVtbl_&objectname IDataObjectVtbl {QueryInterface, AddRef, Release,\
    GetData_, GetDataHere_, QueryGetData_, GetCanonicalFormatEtc_,\
    SetData_, EnumFormatEtc_, DAdvise_, DUnadvise_, EnumDAdvise_}
else
CDataObjectVtbl_&objectname label IDataObjectVtbl
	dd QueryInterface, AddRef, Release
    dd GetData_, GetDataHere_, QueryGetData_, GetCanonicalFormatEtc_
    dd SetData_, EnumFormatEtc_, DAdvise_, DUnadvise_, EnumDAdvise_
endif
    .code

;--------------------------------------------------------------------------
; IDataObject
;--------------------------------------------------------------------------

QueryInterface:
	sub dword ptr [esp+4], objectname&._IDataObject
	jmp QueryInterface@&objectname
AddRef:
	sub dword ptr [esp+4], objectname&._IDataObject
	jmp AddRef@&objectname
Release:
	sub dword ptr [esp+4], objectname&._IDataObject
	jmp Release@&objectname

GetData_:
	sub dword ptr [esp+4], objectname&._IDataObject
GetData@&objectname PROC uses __this this_:ptr objectname, pFormatetc:ptr FORMATETC, pmedium:ptr STGMEDIUM

	mov __this,this_
ifdef _DEBUG
	mov ecx, pFormatetc
	movzx edx, [ecx].FORMATETC.cfFormat
	DebugOut "IDataObject::GetData(%X[%X, %X, %X, %X, %X], %X)",\
		ecx, edx, [ecx].FORMATETC.ptd, [ecx].FORMATETC.dwAspect,\
		[ecx].FORMATETC.lindex, [ecx].FORMATETC.tymed, pmedium
endif
    return DV_E_FORMATETC

GetData@&objectname ENDP


GetDataHere_:
	sub dword ptr [esp+4], objectname&._IDataObject
GetDataHere@&objectname PROC uses __this this_:ptr objectname, pFormatetc:ptr FORMATETC, pmedium:ptr STGMEDIUM

	mov __this,this_
	DebugOut "IDataObject::GetDataHere(%X, %X)", pFormatetc, pmedium
    return DV_E_FORMATETC

GetDataHere@&objectname ENDP


QueryGetData_:
	sub dword ptr [esp+4], objectname&._IDataObject
QueryGetData@&objectname PROC uses __this this_:ptr objectname, pFormatetc:ptr FORMATETC

	mov __this,this_

	DebugOut "IDataObject::QueryGetData(%X)", pFormatetc

    return DV_E_FORMATETC

QueryGetData@&objectname ENDP


GetCanonicalFormatEtc_:
	sub dword ptr [esp+4], objectname&._IDataObject
GetCanonicalFormatEtc@&objectname PROC uses __this this_:ptr objectname,
		pFormatetcIn:ptr FORMATETC, pFormatetcOut:ptr FORMATETC

	mov __this,this_
	DebugOut "IDataObject::GetCanonicalFormatEtc(%X, %X)", pFormatetcIn, pFormatetcOut
    return E_UNEXPECTED

GetCanonicalFormatEtc@&objectname ENDP

SetData_:
	sub dword ptr [esp+4], objectname&._IDataObject
SetData@&objectname PROC uses __this this_:ptr objectname,
		pFormatetc:ptr FORMATETC, pmedium:ptr STGMEDIUM, fRelease:BOOL

	mov __this,this_
	DebugOut "IDataObject::SetData(%X, %X, %u)", pFormatetc, pmedium, fRelease
    return E_NOTIMPL

SetData@&objectname ENDP

EnumFormatEtc_:
	sub dword ptr [esp+4], objectname&._IDataObject
EnumFormatEtc@&objectname PROC uses __this this_:ptr objectname, dwDirection:DWORD, ppenumFormatetc: ptr ptr IEnumFORMATETC

	mov __this,this_

	DebugOut "IDataObject::EnumFormatEtc(%u, %X)", dwDirection, ppenumFormatetc
	mov ecx, ppenumFormatetc
	mov dword ptr [ecx], NULL
    return E_NOTIMPL

EnumFormatEtc@&objectname ENDP

DAdvise_:
	sub dword ptr [esp+4], objectname&._IDataObject
DAdvise@&objectname PROC uses __this this_:ptr objectname, pFormatetc:ptr FORMATETC,
				advf:DWORD, pAdvSink: LPADVISESINK, pdwConnection:ptr DWORD

	mov __this,this_
	DebugOut "IDataObject::DAdvise(%X, %u, %X, %X)", pFormatetc, advf, pAdvSink, pdwConnection
	.if (!m_pDataAdviseHolder)
		invoke CreateDataAdviseHolder, addr m_pDataAdviseHolder
	.endif
	.if (m_pDataAdviseHolder)
		lea ecx, m__IDataObject
		invoke vf(m_pDataAdviseHolder, IDataAdviseHolder, Advise), ecx, pFormatetc, advf, pAdvSink, pdwConnection
		ret
	.else
		mov ecx, pdwConnection
		mov dword ptr [ecx], NULL
	.endif
	return OLE_E_ADVISENOTSUPPORTED

DAdvise@&objectname ENDP

DUnadvise_:
	sub dword ptr [esp+4], objectname&._IDataObject
DUnadvise@&objectname PROC uses __this this_:ptr objectname, dwConnection:DWORD

	mov __this,this_
	DebugOut "IDataObject::DUnadvise(%X)", dwConnection
	.if (m_pDataAdviseHolder)
		invoke vf(m_pDataAdviseHolder, IDataAdviseHolder, Unadvise), dwConnection
		ret
	.endif
    return OLE_E_ADVISENOTSUPPORTED

DUnadvise@&objectname ENDP

EnumDAdvise_:
	sub dword ptr [esp+4], objectname&._IDataObject
EnumDAdvise@&objectname PROC uses __this this_:ptr objectname, ppenumAdvise:ptr ptr IEnumSTATDATA

	mov __this,this_
	DebugOut "IDataObject::EnumDAdvise(%X)", ppenumAdvise
	.if (m_pDataAdviseHolder)
		invoke vf(m_pDataAdviseHolder, IDataAdviseHolder, EnumAdvise), ppenumAdvise
		ret
	.endif
	mov ecx, ppenumAdvise
	mov dword ptr [ecx], NULL
    return OLE_E_ADVISENOTSUPPORTED

EnumDAdvise@&objectname ENDP

	endm

endif



;--------------------------------------------------------------------------
;--- DEFINE_OLEOBJECT: defines IOleObject
;--------------------------------------------------------------------------

ifndef OLEOBJECT_DEFINED

OLEOBJECT_DEFINED equ 1

DEFINE_OLEOBJECT macro objectname
local QueryInterface_, AddRef_, Release_
local SetClientSite_, GetClientSite_, SetHostNames_
local Close_,  SetMoniker_, GetMoniker_
local DoVerb_, OnInPlaceActivate_
local Advise_, Unadvise_, EnumAdvise_, GetExtent_, SetExtent_

	.const

COleObjectVtbl_&objectname label IOleObjectVtbl 
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd SetClientSite_, GetClientSite_, SetHostNames_
	dd Close_, SetMoniker_, GetMoniker_
	dd InitFromData, GetClipboardData_, DoVerb_
	dd EnumVerbs, Update, IsUpToDate
	dd GetUserClassID, GetUserType
	dd SetExtent_, GetExtent_
	dd Advise_, Unadvise_, EnumAdvise_
	dd GetMiscStatus, SetColorScheme

	.code

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IOleObject
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IOleObject
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IOleObject
	jmp Release@&objectname

;--------------------------------------------------------------------------
;IOleObject interface
;--------------------------------------------------------------------------

SetClientSite_:
	sub dword ptr [esp+4], objectname&._IOleObject

SetClientSite@&objectname PROC	uses __this this_:ptr objectname, pOleClientSite:LPOLECLIENTSITE

	mov __this,this_

	DebugOut "IOleObject::SetClientSite(%X)", pOleClientSite

	invoke ComPtrAssign, addr m_pOleClientSite, pOleClientSite
	return S_OK

SetClientSite@&objectname ENDP

;--------------------------------------------------------------------------

GetClientSite_:
	sub dword ptr [esp+4], objectname&._IOleObject

GetClientSite@&objectname PROC this_:ptr objectname, ppOleClientSite:ptr LPOLECLIENTSITE

	DebugOut "IOleObject::GetClientSite"

	mov ecx,ppOleClientSite
	.IF (!ecx)
		return E_POINTER
	.ENDIF
	mov dword ptr [ecx],NULL
	mov eax,this_
	invoke ComPtrAssign, ecx, [eax].&objectname&.pOleClientSite
	return S_OK

GetClientSite@&objectname ENDP
 
;--------------------------------------------------------------------------

SetHostNames_:
	sub dword ptr [esp+4], objectname&._IOleObject

SetHostNames@&objectname PROC this_:ptr objectname, pwszContainerApp:ptr WORD, pwszContainerObj:ptr WORD

ifdef _DEBUG
	push esi
	push edi
	sub esp, 256
	.if (pwszContainerApp)
		lea esi, [esp+128]
		invoke WideCharToMultiByte, CP_ACP, 0, pwszContainerApp, -1, esi, 128, 0, 0
	.else
		mov esi, CStr("NULL")
	.endif
	.if (pwszContainerObj)
		mov edi, esp
		invoke WideCharToMultiByte, CP_ACP, 0, pwszContainerObj, -1, edi, 128, 0, 0
	.else
		mov edi, CStr("NULL")
	.endif
	DebugOut "IOleObject::SetHostNames(%s,%s)", esi, edi
	add esp, 256
	pop edi
	pop esi
endif
	return S_OK

SetHostNames@&objectname ENDP

;--------------------------------------------------------------------------
;* Object changes from "running" to "loaded"

Close_:
	sub dword ptr [esp+4], objectname&._IOleObject

Close@&objectname PROC uses __this this_:ptr objectname, dwSaveOption:DWORD 

	DebugOut "enter IOleObject::Close(%X)", dwSaveOption

	mov __this,this_

;;;--------------------- deactivated 1.8.2003
;;;	.if (m_hWnd)

	.IF (dwSaveOption == OLECLOSE_SAVEIFDIRTY) || (dwSaveOption == OLECLOSE_PROMPTSAVE)
		DebugOut "Dirty flag=%u", m_isDirty
		.IF (m_pOleClientSite && m_isDirty)
			DebugOut "calling IOleClientSite::SaveObject"
			invoke vf(m_pOleClientSite, IOleClientSite, SaveObject)
			.IF (m_pOleAdviseHolder) 
				DebugOut "calling IOleAdviseHolder::SendOnSave"
				invoke vf(m_pOleAdviseHolder, IOleAdviseHolder, SendOnSave)
			.ENDIF
		.ENDIF
	.ENDIF

	.IF (m_hWnd)
		lea ecx,m__IOleInPlaceObject
		invoke vf(ecx, IOleInPlaceObject, InPlaceDeactivate)
	.ENDIF

;--- call OleAdviseHolder or IAdviseSink, but not both!
if 1								
	.IF (m_pOleAdviseHolder) 
		DebugOut "calling IOleAdviseHolder::SendOnClose"
		invoke vf(m_pOleAdviseHolder, IOleAdviseHolder, SendOnClose)
	.ENDIF
else
	.IF (m_pAdviseSink) 
		invoke vf(m_pAdviseSink, IAdviseSink, OnClose)
	.ENDIF
endif
;;;	.endif

;------------------------ dont release client site, this is done
;------------------------ by the host itself by calling SetClientSite(0)
if 0
	.IF (m_pOleClientSite)
		invoke vf(m_pOleClientSite, IOleClientSite, Release)
	.ENDIF
endif

	DebugOut "exit IOleObject::Close"

	return S_OK

Close@&objectname ENDP

;--------------------------------------------------------------------------

SetMoniker_:
	sub dword ptr [esp+4], objectname&._IOleObject

SetMoniker@&objectname PROC uses __this this_:ptr objectname, dwWhichMoniker:DWORD, pmk:ptr IMoniker

	DebugOut "IOleObject::SetMoniker"
if 0
	mov __this,this_
	.IF (m_pOleAdviseHolder) 
		invoke vf(m_pOleAdviseHolder, IOleAdviseHolder, SendOnRename), pmk
	.ENDIF
endif
	return E_NOTIMPL

SetMoniker@&objectname ENDP

;--------------------------------------------------------------------------

GetMoniker_:
	sub dword ptr [esp+4], objectname&._IOleObject

GetMoniker@&objectname PROC uses __this this_:ptr objectname, dwAssign:DWORD, dwWhichMoniker:DWORD, ppMoniker:ptr ptr

	DebugOut "IOleObject::GetMoniker"
	mov __this,this_
	mov ecx, ppMoniker
	@mov dword ptr [ecx], 0
	mov eax, E_NOTIMPL
	.if (m_pOleClientSite)
		invoke vf(m_pOleClientSite, IOleClientSite, GetMoniker),\
			OLEGETMONIKER_ONLYIFTHERE, OLEWHICHMK_OBJFULL, ppMoniker
	.endif
	ret

GetMoniker@&objectname ENDP
 
;--------------------------------------------------------------------------

InitFromData PROC this_:ptr objectname, pDataObject:LPDATAOBJECT, fCreation:BOOL, dwReserved:DWORD

	DebugOut "IOleObject::InitFromData(%X, %X, %X)", pDataObject, fCreation, dwReserved
	return E_NOTIMPL

InitFromData ENDP

;--------------------------------------------------------------------------

GetClipboardData_ PROC this_:ptr objectname, dwReserved:DWORD, ppDataObject:ptr LPDATAOBJECT

	DebugOut "IOleObject::GetClipboardData(%X, %X)", dwReserved, ppDataObject
	mov eax,ppDataObject
	.if (eax)
		xor ecx, ecx
		mov [eax],ecx
	.endif
	return E_NOTIMPL

GetClipboardData_ ENDP

;--------------------------------------------------------------------------

ShowPropertyPage proc hwndParent:HWND

LOCAL cauuid:CAUUID
LOCAL pUnknown:LPUNKNOWN

ifdef INCLUDE_&objectname&_ISpecifyPropertyPages
	lea ecx, m__ISpecifyPropertyPages
	invoke vf(ecx, ISpecifyPropertyPages, GetPages),addr cauuid
	.if (eax == S_OK)
		mov pUnknown,__this
		invoke OleCreatePropertyFrame, hwndParent, 32, 32, \
				CStrW(L("Properties")), 1, addr pUnknown,\
				cauuid.cElems,cauuid.pElems, 0, 0, 0
		mov eax,S_OK
	.else
		mov eax, OLEOBJ_S_CANNOT_DOVERB_NOW 
	.endif
else
	mov eax, OLEOBJ_S_CANNOT_DOVERB_NOW 
endif
	ret

ShowPropertyPage endp

;--------------------------------------------------------------------------

;*** private method OnInPlaceActivate

OnInPlaceActivate_ proc pOleClientSite:LPOLECLIENTSITE, hwndParent:HWND, lprcPosRect:ptr RECT

LOCAL pOleInPlaceSite:LPOLEINPLACESITE

	.if (m_hWnd)		;if hWnd exists object is already activated
        return S_OK
    .endif
	invoke vf(pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleInPlaceSite, ADDR pOleInPlaceSite
	.if (eax == S_OK)
		invoke vf(pOleInPlaceSite, IOleInPlaceSite, CanInPlaceActivate)
		.if (eax == S_OK)
			invoke vf(pOleInPlaceSite, IOleInPlaceSite, OnInPlaceActivate)
ifdef OnInPlaceActivate@&objectname
			invoke OnInPlaceActivate@&objectname, __this, hwndParent, lprcPosRect
endif
			invoke vf(pOleClientSite,IOleClientSite, ShowObject)	;new 2.2.2002
			mov eax, S_OK
		.else
			mov eax,OLEOBJ_S_CANNOT_DOVERB_NOW	
		.endif
		push eax
		invoke vf(pOleInPlaceSite, IOleInPlaceSite, Release)
		pop eax
	.else
		mov eax,OLEOBJ_S_CANNOT_DOVERB_NOW	
	.endif
    ret

OnInPlaceActivate_ endp

OnUIActivate@&objectname proc pOleClientSite:LPOLECLIENTSITE

LOCAL pOleInPlaceSite:LPOLEINPLACESITE
LOCAL pOleInPlaceFrame:LPOLEINPLACEFRAME

	.if (!(m_dwFlags & OLEOBJF_UIACTIVATED))
  		invoke vf(pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleInPlaceSite, ADDR pOleInPlaceSite
	   	.if (eax == S_OK)
    		invoke vf(pOleInPlaceSite,IOleInPlaceSite, OnUIActivate)
	    	invoke vf(pOleInPlaceSite,IOleInPlaceSite, Release)
		.endif
		invoke vf(pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleInPlaceFrame, ADDR pOleInPlaceFrame
		.if (eax == S_OK)
			invoke vf(pOleInPlaceFrame,IOleInPlaceFrame,SetActiveObject), addr m__IOleInPlaceActiveObject, NULL
			invoke vf(pOleInPlaceFrame,IOleInPlaceFrame,Release)
		.endif
		or m_dwFlags, OLEOBJF_UIACTIVATED
	.endif
	mov eax,S_OK
	ret
OnUIActivate@&objectname endp

;--------------------------------------------------------------------------

DoVerb_:
	sub dword ptr [esp+4], objectname&._IOleObject

DoVerb@&objectname PROC uses __this this_:ptr objectname, iVerb:SDWORD, lpmsg:ptr MSG,\
	 pOleClientSite:LPOLECLIENTSITE, lindex:DWORD, hwndParent:HWND, lprcPosRect:ptr RECT

LOCAL pOleControlSite:LPOLECONTROLSITE

	mov __this,this_

	DebugOut "enter IOleObject::DoVerb Verb=%d",iVerb

	mov eax,iVerb

	.IF (eax == OLEIVERB_PRIMARY)

ifdef OnVerbPrimary@&objectname
		invoke OnVerbPrimary@&objectname, __this, lpmsg,\
						pOleClientSite, hwndParent, lprcPosRect
else
		mov eax,S_OK
endif

	.ELSEIF (eax == OLEIVERB_SHOW)

        invoke OnInPlaceActivate_, pOleClientSite, hwndParent, lprcPosRect
		.if (m_hWnd)
			invoke ShowWindow,m_hWnd,SW_SHOWNOACTIVATE
		.endif
		invoke vf(pOleClientSite, IOleClientSite, ShowObject)
		mov eax,S_OK		;return S_OK, ignore returncode from ShowObject

	.ELSEIF (eax == OLEIVERB_HIDE)

		.if (m_dwFlags & OLEOBJF_UIACTIVATED)
			lea ecx,m__IOleInPlaceObject
			invoke vf(ecx, IOleInPlaceObject, UIDeactivate)
		.endif
		.if (m_hWnd)
			invoke ShowWindow,m_hWnd,SW_HIDE
		.endif
		mov eax, S_OK

	.ELSEIF (eax == OLEIVERB_UIACTIVATE)

        invoke OnInPlaceActivate_, pOleClientSite, hwndParent, lprcPosRect
        .if (eax == S_OK)
			invoke OnUIActivate@&objectname, pOleClientSite
        .endif

	.ELSEIF (eax == OLEIVERB_INPLACEACTIVATE)

        invoke OnInPlaceActivate_, pOleClientSite, hwndParent, lprcPosRect

	.ELSEIF (eax == OLEIVERB_PROPERTIES)

		invoke vf(pOleClientSite, IOleClientSite, QueryInterface), addr IID_IOleControlSite, ADDR pOleControlSite
		.IF SUCCEEDED(eax)
			invoke vf(pOleControlSite, IOleControlSite, ShowPropertyFrame)
			push eax
			invoke vf(pOleControlSite, IOleControlSite, Release)
			pop eax
			.if (eax != S_OK)
				invoke ShowPropertyPage, hwndParent	
			.endif
		.ELSE
			invoke ShowPropertyPage, hwndParent	
		.ENDIF

	.elseif (eax >= 0)

ifdef OnDoVerb@&objectname
		invoke OnDoVerb@&objectname, __this, iVerb, lpmsg,\
						pOleClientSite, hwndParent, lprcPosRect
else
		mov eax,OLEOBJ_S_INVALIDVERB
endif
	.else

		mov eax, E_NOTIMPL

	.ENDIF
	DebugOut "exit IOleObject::DoVerb Verb=%d, hr=%X",iVerb,eax
	ret

DoVerb@&objectname ENDP

;--------------------------------------------------------------------------

EnumVerbs PROC this_:ptr objectname, ppIEnumOleVerb:ptr ptr IEnumOleVerb

	DebugOut "enter IOleObject::EnumVerbs"
	invoke OleRegEnumVerbs, addr CLSID_&objectname,	ppIEnumOleVerb
	DebugOut "exit IOleObject::EnumVerbs, hr=%X",eax
	ret
EnumVerbs ENDP

;--------------------------------------------------------------------------

Update PROC this_:ptr objectname

	DebugOut "IOleObject::Update"
	return S_OK
Update ENDP

;--------------------------------------------------------------------------

IsUpToDate PROC this_:ptr objectname

	DebugOut "IOleObject::IsUpToDate"
	return S_OK
IsUpToDate ENDP

;--------------------------------------------------------------------------

GetUserClassID PROC uses esi edi this_:ptr objectname, pClsid:ptr GUID

	DebugOut "IOleObject::GetUserClassID"

	mov edi, pClsid 
	.IF (!edi)
		mov eax, E_POINTER
	.else
		mov esi, offset CLSID_&objectname
		movsd
		movsd
		movsd
		movsd
		xor eax, eax			; return S_OK
	.ENDIF
	ret

GetUserClassID ENDP

;--------------------------------------------------------------------------

GetUserType PROC this_:ptr objectname, dwFormOfType:DWORD, pszUserType:ptr ptr word

	DebugOut "enter IOleObject::GetUserType Type=%X",dwFormOfType

	invoke OleRegGetUserType, addr CLSID_&objectname, dwFormOfType, pszUserType

	DebugOut "exit IOleObject::GetUserType, hr=%X",eax
	ret
GetUserType ENDP


HIMETRIC_PER_INCH   equ 2540

if 0
MAP_PIX_TO_LOGHIM macro x,ppli
   ( (HIMETRIC_PER_INCH*(x) + ((ppli)>>1)) / (ppli) )
   endm
endif

MAP_LOGHIM_TO_PIX macro x, ppli
	mov eax, ppli
	mov ecx, x
	mul ecx
	add eax, HIMETRIC_PER_INCH/2
	cdq
	mov ecx, HIMETRIC_PER_INCH
	div ecx
	endm

HiMetricToPixel proc lpSizeInHiMetric:ptr SIZEL, lpSizeInPix: ptr SIZEL

local	nPixelsPerInchX:DWORD	;// Pixels per logical inch along width
local	nPixelsPerInchY:DWORD	;// Pixels per logical inch along height
local	hDCScreen:HDC

	invoke GetDC, NULL
	mov hDCScreen, eax
	invoke GetDeviceCaps, hDCScreen, LOGPIXELSX
	mov nPixelsPerInchX, eax
	invoke GetDeviceCaps, hDCScreen, LOGPIXELSY
	mov nPixelsPerInchY, eax
	invoke ReleaseDC, NULL, hDCScreen

	mov ecx, lpSizeInHiMetric
	MAP_LOGHIM_TO_PIX [ecx].SIZEL.cx_, nPixelsPerInchX
	mov ecx, lpSizeInPix
	mov [ecx].SIZEL.cx_, eax

	mov ecx, lpSizeInHiMetric
	MAP_LOGHIM_TO_PIX [ecx].SIZEL.cy, nPixelsPerInchY
	mov ecx, lpSizeInPix
	mov [ecx].SIZEL.cy, eax
	ret
HiMetricToPixel endp

;--------------------------------------------------------------------------

SetExtent_:
	sub dword ptr [esp+4], objectname&._IOleObject

SetExtent@&objectname PROC uses __this this_:ptr objectname, dwAspect:DWORD, pSizel:ptr SIZEL

	mov eax,pSizel
	.IF (!eax)
		return E_POINTER
	.ENDIF

	DebugOut "IOleObject::SetExtent,x=%u,y=%u",[eax].SIZEL.cx_,[eax].SIZEL.cy

	.IF (dwAspect != DVASPECT_CONTENT) 
		return E_FAIL
	.ENDIF

	mov __this,this_

	mov ecx, pSizel
	mov eax,[ecx].SIZEL.cx_
	mov edx,[ecx].SIZEL.cy
	.if ((eax != m_himetricExtent.cx_) || (edx != m_himetricExtent.cy))
		mov m_isDirty, TRUE
	.endif
	mov m_himetricExtent.cx_, eax
	mov m_himetricExtent.cy, edx

	invoke HiMetricToPixel, ecx, addr m_pixelExtent

	.if (m_hWnd)
		invoke SetWindowPos,m_hWnd,NULL,0,0,\
			m_pixelExtent.cx_,\
			m_pixelExtent.cy,\
			SWP_NOZORDER or SWP_NOMOVE or SWP_NOACTIVATE
	.endif
	DebugOut "IOleObject::SetExtent,Pixel x=%u,y=%u", m_pixelExtent.cx_, m_pixelExtent.cy
	return S_OK

SetExtent@&objectname ENDP

;--------------------------------------------------------------------------

GetExtent_:
	sub dword ptr [esp+4], objectname&._IOleObject

GetExtent@&objectname PROC this_:ptr objectname, dwAspect:DWORD, pSizel:ptr SIZEL 

	DebugOut "IOleObject::GetExtent"
	mov ecx, pSizel 
	.IF (!ecx)
		mov eax, E_POINTER
	.ELSEIF (dwAspect != DVASPECT_CONTENT) 
		mov eax, E_INVALIDARG
	.ELSE
		mov edx,this_
		mov eax, [edx].&objectname&.himetricExtent.cx_
		mov [ecx].SIZEL.cx_, eax
		mov eax, [edx].&objectname&.himetricExtent.cy
		mov [ecx].SIZEL.cy, eax
		DebugOut "IOleObject::GetExtent x=%u, y=%u",\
				[ecx].SIZEL.cx_, [ecx].SIZEL.cy
		xor eax, eax			; return S_OK
	.ENDIF
	ret
GetExtent@&objectname ENDP

;--------------------------------------------------------------------------

Advise_:
	sub dword ptr [esp+4], objectname&._IOleObject

Advise@&objectname PROC uses __this this_:ptr objectname, pAdvSink:ptr IAdviceSink, pdwConnection:ptr DWORD

	DebugOut "IOleObject::Advise(%X, %X)", pAdvSink, pdwConnection

	mov __this,this_

	mov eax,pdwConnection
	mov dword ptr [eax],NULL

	.if (!m_pOleAdviseHolder)
	    invoke CreateOleAdviseHolder, ADDR m_pOleAdviseHolder
		DebugOut "CreateOleAdviseHolder returned %X", eax
		.if (FAILED(eax))
			ret
		.endif
	.endif

	invoke vf(m_pOleAdviseHolder, IOleAdviseHolder, Advise),\
			pAdvSink, pdwConnection
	ret

Advise@&objectname ENDP

;--------------------------------------------------------------------------

Unadvise_:
	sub dword ptr [esp+4], objectname&._IOleObject

Unadvise@&objectname PROC this_:ptr objectname, dwConnection:DWORD 

	DebugOut "IOleObject::Unadvise(%X)", dwConnection
	mov ecx,this_
	.if ([ecx].&objectname&.pOleAdviseHolder)
		invoke vf([ecx].&objectname&.pOleAdviseHolder, IOleAdviseHolder, Unadvise), \
			dwConnection
	.else
		mov eax,E_FAIL
	.endif
	ret
Unadvise@&objectname ENDP

;--------------------------------------------------------------------------

EnumAdvise_:
	sub dword ptr [esp+4], objectname&._IOleObject

EnumAdvise@&objectname PROC this_:ptr objectname, ppEnumAdvise:DWORD


	DebugOut "IOleObject::EnumAdvise(%X)", ppEnumAdvise
	mov eax,ppEnumAdvise
	mov dword ptr [eax], NULL
	mov ecx,this_
	.if ([ecx].&objectname&.pOleAdviseHolder)
		invoke vf([ecx].&objectname&.pOleAdviseHolder, IOleAdviseHolder, EnumAdvise),\
			ppEnumAdvise
	.else
		mov eax,E_FAIL
	.endif
	ret
EnumAdvise@&objectname ENDP

;--------------------------------------------------------------------------

GetMiscStatus PROC this_:ptr objectname, dwAspect:DWORD, pdwStatus:ptr DWORD 

	DebugOut "enter IOleObject::GetMiscStatus"
	invoke OleRegGetMiscStatus, addr CLSID_&objectname, dwAspect, pdwStatus
	DebugOut "exit IOleObject::GetMiscStatus, hr=%X",eax
	ret
GetMiscStatus ENDP

;--------------------------------------------------------------------------

SetColorScheme PROC this_:ptr objectname, plogpalette:DWORD

	DebugOut "IOleObject::SetColorScheme"
	return E_NOTIMPL

SetColorScheme ENDP

	endm

endif


DEFINE_SENDVIEWCHANGE macro objectname

SendViewChange@&objectname PROC public uses __this this_:ptr objectname

	mov __this,this_

    .IF (m_pAdviseSink)
        invoke vf( m_pAdviseSink, IAdviseSink, OnViewChange), DVASPECT_CONTENT, -1
	    .IF ( m_advf & ADVF_ONLYONCE)
		    invoke vf( m_pAdviseSink, IUnknown, Release)
			mov m_pAdviseSink, NULL
	        mov m_advf, NULL
	    .ENDIF
    .ENDIF
ifdef INCLUDE_&objectname&_IOleObject
	.if (m_hWnd)
		invoke InvalidateRect, m_hWnd,0,1
	.endif
endif
    ret

SendViewChange@&objectname ENDP

	endm


;--------------------------------------------------------------------------
;--- DEFINE_CONNECTIONPOINTCONTAINER: defines IConnectionPointContainer
;--------------------------------------------------------------------------

ifndef CONNECTIONPOINTCONTAINER_DEFINED

CONNECTIONPOINTCONTAINER_DEFINED equ 1

DEFINE_CONNECTIONPOINTCONTAINER macro objectname	;; pEventIIDs, numCPs
local QueryInterface_, AddRef_, Release_, InitCP_, EnumConnectionPoints_, FindConnectionPoint_

;;externdef CConnectionPointContainerVtbl_&objectname:IConnectionPointContainerVtbl

Create@CEnumConnectionPoints proto :LPUNKNOWN, :ptr LPCONNECTIONPOINTS, :DWORD

	.const

CConnectionPointContainerVtbl_&objectname label IConnectionPointContainerVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd EnumConnectionPoints_, FindConnectionPoint_

	.code

;--------------------------------------------------------------------------
;IConnectionPointContainer interface
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IConnectionPointContainer
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IConnectionPointContainer
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IConnectionPointContainer
	jmp Release@&objectname


Create@CConnectionPoint proto :LPUNKNOWN, :ptr IID


InitCP_ proc uses __this esi edi this_:ptr objectname

	mov __this,this_

	.erre NUMCP_&objectname eq LENGTHOF objectname&.pConnectionPoint, <classname_cntCP must match entries in CONNECTION_POINT_MAP>

	mov ecx, NUMCP_&objectname
	mov esi, offset EventIID_&objectname
	lea edi, m_pConnectionPoint
	.while (ecx)
		push ecx
		lodsd
		.if (!(dword ptr [edi]))
			invoke Create@CConnectionPoint, __this, eax
			mov [edi], eax
		.endif
		add edi, 4
		pop ecx
		dec ecx
	.endw
	ret
InitCP_ endp


EnumConnectionPoints_:
	sub dword ptr [esp+4], objectname&._IConnectionPointContainer

EnumConnectionPoints PROC this_:ptr objectname, ppIEnumConnectionPoints:ptr ptr


	DebugOut "IConnectionPointContainer::EnumConnectionPoints"
	.if (ppIEnumConnectionPoints == NULL)
		return E_POINTER
	.endif

	invoke InitCP_, this_

	mov eax,this_
	lea eax, [eax].&objectname&.pConnectionPoint
	invoke Create@CEnumConnectionPoints, this_, eax, NUMCP_&objectname

	mov ecx,ppIEnumConnectionPoints
	mov [ecx],eax
	.if (eax)
		return S_OK
	.else
		return E_OUTOFMEMORY
	.endif

EnumConnectionPoints endp

;--------------------------------------------------------------------------

FindConnectionPoint_:
	sub dword ptr [esp+4], objectname&._IConnectionPointContainer

FindConnectionPoint PROC uses __this esi edi this_:ptr objectname, pRefIID: ptr IID, ppConnectionPoint:ptr ptr

local	dwNumCPs:DWORD
local	iid:IID
ifdef _DEBUG
local	wszIID[40]:word
endif

	mov __this,this_
	mov eax,ppConnectionPoint
	.if (eax == NULL)
		mov eax, E_POINTER
		jmp done
	.endif
	mov dword ptr [eax],NULL

	invoke InitCP_, __this

	lea esi,m_pConnectionPoint
	mov dwNumCPs, NUMCP_&objectname
	.while (dwNumCPs)
		lodsd
		mov edi, eax
		invoke vf(edi, IConnectionPoint, GetConnectionInterface), addr iid
		.if (eax == S_OK)
			invoke IsEqualGUID, addr iid, pRefIID
			.if (eax)
				invoke ComPtrAssign, ppConnectionPoint, edi
				mov eax, S_OK
				jmp done
			.endif
		.endif
		dec dwNumCPs
	.endw
	mov eax, CONNECT_E_NOCONNECTION
done:
ifdef _DEBUG
	push eax
	invoke StringFromGUID2, pRefIID, addr wszIID, LENGTHOF wszIID
	pop eax
	DebugOut "IConnectionPointContainer::FindConnectionPoint(%S)=%X", addr wszIID, eax
endif
	ret

FindConnectionPoint endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_ENUMCONNECTIONPOINTS: macro needs no parameter
;--------------------------------------------------------------------------

ifndef ENUMCONNECTIONPOINTS_DEFINED

ENUMCONNECTIONPOINTS_DEFINED equ 1


DEFINE_ENUMCONNECTIONPOINTS macro
local Destroy_, QueryInterface_, AddRef_, Release_, Next_, Skip_, Reset_, Clone_, supInterfaces_

	.const

CEnumConnectionPointsVtbl label IEnumConnectionPointsVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd	Next_, Skip_, Reset_, Clone_

	.code

;--------------------------------------------------------------------------
;IEnumConnectionPoints interface
;--------------------------------------------------------------------------

ifndef pIEnumConnectionPointsVtbl
pIEnumConnectionPointsVtbl typedef ptr IEnumConnectionPointsVtbl
endif

CEnumConnectionPoints struct
EnumConnectionPoints	pIEnumConnectionPointsVtbl ?
ObjRefCount			DWORD	?
pObject				DWORD	?
dwIndex				DWORD	?
pCPTab				LPVOID	?
dwSize				DWORD	?
CEnumConnectionPoints ends

Create@CEnumConnectionPoints proc uses ebx pObject:LPUNKNOWN, pCPs:ptr LPCONNECTIONPOINTS, dwSize:DWORD

	DebugOut "Create@CEnumConnectionPoints"

	invoke LocalAlloc,LMEM_FIXED,sizeof CEnumConnectionPoints
	.if (eax == NULL)
		ret
	.endif
	mov ebx, eax
	assume ebx:ptr CEnumConnectionPoints

	mov [ebx].EnumConnectionPoints,offset CEnumConnectionPointsVtbl
	mov [ebx].ObjRefCount,1
	mov eax,pObject
	mov [ebx].pObject,eax
	invoke vf(eax, IUnknown, AddRef)
	mov [ebx].dwIndex,0
	mov eax, dwSize
	mov [ebx].dwSize, eax
	mov eax, pCPs
	mov [ebx].pCPTab, eax

	inc g_DllRefCount

	return ebx
	assume ebx:nothing

Create@CEnumConnectionPoints endp


Destroy_ PROC uses ebx this_:ptr CEnumConnectionPoints

	DebugOut "Destroy@CEnumConnectionPoints"

	mov eax,this_
	invoke vf([eax].CEnumConnectionPoints.pObject, IUnknown, Release)
	invoke LocalFree,this_
	dec g_DllRefCount
	ret

Destroy_ ENDP

	.const

supInterfaces_ label dword
	dd offset IID_IUnknown,0
	dd offset IID_IEnumConnectionPoints,0
IFTABSIZE_ equ ($ - supInterfaces_)/ (2 * sizeof dword)

	.code

QueryInterface_ PROC this_:ptr CEnumConnectionPoints,riid:ptr IID,ppReturn:ptr
	invoke IsInterfaceSupported, riid, offset supInterfaces_, IFTABSIZE_,  this_, ppReturn
	ret
QueryInterface_ ENDP	


AddRef_ PROC this_:ptr CEnumConnectionPoints

	mov eax,this_
	assume eax:ptr CEnumConnectionPoints

	inc [eax].ObjRefCount
	mov	eax, [eax].ObjRefCount
	ret
	assume eax:nothing

AddRef_ ENDP			


Release_ PROC this_:ptr CEnumConnectionPoints
	
	mov eax,this_
	assume eax:ptr CEnumConnectionPoints

	dec [eax].ObjRefCount
	mov eax,[eax].ObjRefCount
	.if (eax == 0)
		invoke Destroy_, this_
		xor eax,eax
	.endif
	ret
	assume eax:nothing

Release_ ENDP


Next_ PROC uses ebx esi edi this_:ptr CEnumConnectionPoints, dwElements:dword, ppICP:ptr LPCONNECTIONPOINT, pdwFetched:ptr dword

local	dwReturn:dword

	DebugOut "IEnumConnectionPoints::Next"

	mov ebx,this_
	assume ebx:ptr CEnumConnectionPoints

	.if (dwElements > 1 && !pdwFetched)
		return E_INVALIDARG
	.endif

	mov edi,ppICP
	.if (!edi)
		return E_POINTER
	.endif

	mov dwReturn,0

;	mov esi,[ebx].pObject
	mov ecx,[ebx].dwIndex
	mov esi,[ebx].pCPTab
	lea esi,[esi+ecx*4]

	.while (dwElements && (ecx < [ebx].dwSize))
		mov dword ptr [edi], NULL
		push ecx
		invoke ComPtrAssign, edi, [esi]
		pop ecx
		add edi, sizeof LPCONNECTIONPOINT
		add esi, sizeof LPCCONNECTIONPOINT
		inc ecx
		dec dwElements
		inc [ebx].dwIndex
		inc dwReturn
	.endw

	mov eax,pdwFetched
	.if (eax)
		mov ecx,dwReturn
		mov [eax],ecx
	.endif

	.if (dwElements == 0)
		mov eax,S_OK
	.else
		mov eax,S_FALSE
	.endif
	ret
	assume ebx:nothing

Next_ endp


;--------------------------------------------------------------------------

Skip_ PROC this_:ptr CEnumConnectionPoints,cConn:dword
	return E_FAIL
Skip_ endp

;--------------------------------------------------------------------------

Reset_ PROC this_:ptr CEnumConnectionPoints
	
	DebugOut "IEnumConnectionPoints::Reset"
	mov ecx,this_
	mov [ecx].CEnumConnectionPoints.dwIndex,0
	return S_OK
Reset_ endp

;--------------------------------------------------------------------------

Clone_ PROC this_:ptr CEnumConnectionPoints,ppIECP:ptr LPENUMCONNECTIONPOINTS
	return E_FAIL
Clone_ endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_CONNECTIONPOINT: macro needs no parameter
;--------------------------------------------------------------------------

ifndef CONNECTIONPOINT_DEFINED

CONNECTIONPOINT_DEFINED equ 1


DEFINE_CONNECTIONPOINT macro
local Destroy_, QueryInterface_, AddRef_, Release_,GetConnectionInterface_,GetConnectionPointContainer_,Advise_, Unadvise_, EnumConnections_,supInterfaces_

;*** definition of CConnectionPoint class

CConnectionPoint struct
ConnectionPoint	IConnectionPoint <>
ObjRefCount		DWORD	?
refIID			DWORD	?
pFirstSink		LPEVENTSINK	?
pObject			LPUNKNOWN ?
CConnectionPoint ends

LPCCONNECTIONPOINT typedef ptr CConnectionPoint

;;	MEMBER ConnectionPoint, ObjRefCount, refIID, pFirstSink, pObject

CConnectionPointVtbl label IConnectionPointVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd GetConnectionInterface_, GetConnectionPointContainer_
	dd Advise_, Unadvise_, EnumConnections_

;--------------------------------------------------------------------------
;IConnectionPoint interface
;--------------------------------------------------------------------------

;;__this	textequ <ebx>
;;_this	textequ <[__this].CConnectionPoint>

;;	MEMBER ConnectionPoint, ObjRefCount, refIID, pFirstSink, pObject


Create@CConnectionPoint proc uses ebx pObject:LPUNKNOWN, riid:ptr IID

	DebugOut "Create@CConnectionPoint"

	invoke LocalAlloc, LMEM_FIXED or LMEM_ZEROINIT, sizeof CConnectionPoint
	.if (!eax)
		ret
	.endif
	mov ebx, eax
	assume ebx:ptr CConnectionPoint

	mov [ebx].ConnectionPoint,offset CConnectionPointVtbl
	mov [ebx].ObjRefCount,1
	mov eax,pObject
	mov [ebx].pObject,eax
	mov eax,riid
	mov [ebx].refIID,eax
	mov [ebx].pFirstSink,NULL

	return ebx
	assume ebx:nothing

Create@CConnectionPoint endp

Destroy_ PROC uses ebx this_:ptr CConnectionPoint

	DebugOut "Destroy@CConnectionPoint"

	mov ebx,this_
	assume ebx:ptr CConnectionPoint

;	mov [ebx].refIID,NULL

	mov eax, [ebx].pFirstSink
	assume eax:ptr CEventSink
	mov [ebx].pFirstSink, NULL

	.while (eax)
		push [eax].m_pNext
		invoke LocalFree, eax
		pop eax
	.endw
	invoke LocalFree, ebx
	ret
	assume eax:nothing
	assume ebx:nothing

Destroy_ ENDP


	.const

supInterfaces_ label dword
	dd offset IID_IUnknown,0
	dd offset IID_IConnectionPoint,0
IFTABSIZE_ equ ($ - supInterfaces_)/ (2 * sizeof dword)

	.code

QueryInterface_ PROC this_:ptr CConnectionPoint,riid:ptr IID,ppReturn:ptr
	invoke IsInterfaceSupported, riid, offset supInterfaces_, IFTABSIZE_,  this_, ppReturn
	ret
QueryInterface_ ENDP	


AddRef_ PROC this_:ptr CConnectionPoint


	mov eax,this_
	assume eax:ptr CConnectionPoint

	inc [eax].ObjRefCount

;	DebugOut "IConnectionPoint::AddRef %u", [eax].ObjRefCount

	mov	eax, [eax].ObjRefCount
	ret
	assume eax:nothing

AddRef_ ENDP			


Release_ PROC this_:ptr CConnectionPoint
	

	mov eax,this_
	assume eax:ptr CConnectionPoint

	dec [eax].ObjRefCount

;	DebugOut "IConnectionPoint::Release %u", [eax].ObjRefCount

	mov eax,[eax].ObjRefCount
	.if (eax == 0)
		invoke Destroy_,this_
		xor eax,eax
	.endif
	ret
	assume eax:nothing

Release_ ENDP


GetConnectionInterface_ PROC uses esi edi this_:ptr CConnectionPoint,ppIID:ptr ptr IID

	DebugOut "IConnectionPoint::GetConnectionInterface"

	mov edi,ppIID
	.if (edi == NULL)
		return E_POINTER
	.endif
	mov eax, this_
	mov esi, [eax].CConnectionPoint.refIID
	movsd
	movsd
	movsd
	movsd
	return S_OK

GetConnectionInterface_ endp

;--------------------------------------------------------------------------

GetConnectionPointContainer_ PROC this_:ptr CConnectionPoint,ppICPContainer:ptr IConnectionPointContainer

	DebugOut "IConnectionPoint::GetConnectionPointContainer"

	.if (ppICPContainer == NULL)
		return E_POINTER
	.endif
	mov ecx,this_
	invoke vf([ecx].CConnectionPoint.pObject, IUnknown, QueryInterface),\
		addr IID_IConnectionPointContainer, ppICPContainer
	ret

GetConnectionPointContainer_ endp


;--------------------------------------------------------------------------


Advise_ PROC uses ebx esi this_:ptr CConnectionPoint,pIUnknown:ptr IUnknown,pdwCookie:ptr dword

	DebugOut "IConnectionPoint::Advise"

	mov eax,pdwCookie
	.if (eax == 0)
		return E_POINTER
	.endif
	mov dword ptr [eax],0

	mov ebx,this_
	assume ebx:ptr CConnectionPoint

	invoke LocalAlloc, LMEM_FIXED or LMEM_ZEROINIT, sizeof CEventSink
	.if (eax == NULL)
		return E_OUTOFMEMORY
	.endif
	mov esi,eax
	assume esi:ptr CEventSink

	invoke vf(pIUnknown, IUnknown, QueryInterface), [ebx].refIID, addr [esi].m_pEvent

	.if (eax != S_OK)
		invoke LocalFree, esi
		DebugOut "IConnectionPoint::Advise failed with CONNECT_E_CANNOTCONNECT"
		return CONNECT_E_CANNOTCONNECT
	.endif
	lea eax,[ebx].pFirstSink
	.while ([eax].CEventSink.m_pNext)
		mov eax, [eax].CEventSink.m_pNext
	.endw
	mov [eax].CEventSink.m_pNext, esi
	mov eax,pdwCookie
	mov [eax],esi

	return S_OK
	assume ebx:nothing
	assume esi:nothing

Advise_ endp

;--------------------------------------------------------------------------

Unadvise_ PROC uses ebx esi this_:ptr CConnectionPoint, dwCookie:dword

	DebugOut "IConnectionPoint::Unadvise"

	mov ebx,this_
	assume ebx:ptr CConnectionPoint

	lea esi, [ebx].pFirstSink
	assume esi:ptr CEventSink

	.while ([esi].m_pNext)
		mov eax,[esi].m_pNext
		.if (eax == dwCookie)
		    invoke vf([eax].CEventSink.m_pEvent, IUnknown, Release)
			mov eax,[esi].m_pNext
			push [eax].CEventSink.m_pNext
			invoke LocalFree, eax
			pop [esi].m_pNext
			return S_OK
		.endif
		mov esi,[esi].m_pNext
	.endw

	return CONNECT_E_NOCONNECTION

	assume ebx:nothing
	assume esi:nothing

Unadvise_ endp

;--------------------------------------------------------------------------

Create@CEnumConnections proto :ptr CConnectionPoint, :LPEVENTSINK

EnumConnections_ PROC this_:ptr CConnectionPoint, ppIEnumConnections:ptr LPENUMCONNECTIONS

;	DebugOut "IConnectionPoint::EnumConnections"
	mov eax, this_
	invoke Create@CEnumConnections, eax, [eax].CConnectionPoint.pFirstSink
	mov ecx, ppIEnumConnections
	.if (eax)
		mov [ecx], eax
		return S_OK
	.else
		mov dword ptr [ecx], NULL
		return E_FAIL
	.endif

EnumConnections_ endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_ENUMCONNECTIONS: macro needs no parameter
;--------------------------------------------------------------------------

ifndef ENUMCONNECTIONS_DEFINED

DEFINE_ENUMCONNECTIONS macro
local Destroy_, QueryInterface_, AddRef_, Release_, Next_, Skip_, Reset_, Clone_, supInterfaces_

	.const

CEnumConnectionsVtbl label IEnumConnectionsVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd Next_, Skip_, Reset_, Clone_

	.code

;--------------------------------------------------------------------------
;IEnumConnections interface
;--------------------------------------------------------------------------

ifndef pIEnumConnectionsVtbl
pIEnumConnectionsVtbl typedef ptr IEnumConnectionsVtbl
endif

CEnumConnections struct
EnumConnections	pIEnumConnectionsVtbl ?
ObjRefCount			DWORD	?
pObject				LPCONNECTIONPOINT ?
pEventSink			LPEVENTSINK	?
dwIndex				DWORD	?
CEnumConnections ends

Create@CEnumConnections proc uses ebx pObject:ptr CConnectionPoint, pEventSink:LPEVENTSINK

;	DebugOut "Create@CEnumConnections"

	invoke LocalAlloc,LMEM_FIXED,sizeof CEnumConnections
	.if (eax == NULL)
		ret
	.endif
	mov ebx, eax
	assume ebx:ptr CEnumConnections

	mov [ebx].EnumConnections,offset CEnumConnectionsVtbl
	mov [ebx].ObjRefCount,1
	mov eax, pObject
	mov [ebx].pObject, eax
	invoke vf(eax, IUnknown, AddRef)
	mov eax, pEventSink
	mov [ebx].pEventSink,eax
	mov [ebx].dwIndex,0

	inc g_DllRefCount

	return ebx
	assume ebx:nothing

Create@CEnumConnections endp


Destroy_ PROC this_:ptr CEnumConnections

;	DebugOut "Destroy@CEnumConnections"

	mov eax,this_
	invoke vf([eax].CEnumConnections.pObject, IUnknown, Release)
	invoke LocalFree,this_
	dec g_DllRefCount
	ret

Destroy_ ENDP

	.const

supInterfaces_ label dword
	dd offset IID_IUnknown,0
	dd offset IID_IEnumConnections,0
IFTABSIZE_ equ ($ - supInterfaces_)/ (2 * sizeof dword)

	.code

QueryInterface_ PROC this_:ptr CEnumConnections,riid:ptr IID,ppReturn:ptr
	invoke IsInterfaceSupported, riid, offset supInterfaces_, IFTABSIZE_,  this_, ppReturn
	ret
QueryInterface_ ENDP	


AddRef_ PROC this_:ptr CEnumConnections

	mov eax,this_
	assume eax:ptr CEnumConnections

	inc [eax].ObjRefCount
	mov	eax, [eax].ObjRefCount
	ret
	assume eax:nothing

AddRef_ ENDP			


Release_ PROC this_:ptr CEnumConnections
	
	mov eax,this_
	assume eax:ptr CEnumConnections

	dec [eax].ObjRefCount
	mov eax,[eax].ObjRefCount
	.if (eax == 0)
		invoke Destroy_, this_
		xor eax,eax
	.endif
	ret
	assume eax:nothing

Release_ ENDP


Next_ PROC uses ebx esi edi this_:ptr CEnumConnections, dwElements:dword, ppICP:ptr LPCONNECTIONPOINT, pdwFetched:ptr dword

local	dwReturn:dword

;	DebugOut "IEnumConnections::Next"

	mov ebx,this_
	assume ebx:ptr CEnumConnections

	.if (dwElements > 1 && !pdwFetched)
		return E_INVALIDARG
	.endif

	mov edi,ppICP
	.if (!edi)
		return E_POINTER
	.endif

	mov dwReturn,0

	mov esi,[ebx].pEventSink
	assume esi:ptr CEventSink
	mov ecx,[ebx].dwIndex
	.while (ecx)
		mov esi, [esi].m_pNext
		.break .if (!esi)
		dec ecx
	.endw

	.while (dwElements && (esi))
		mov dword ptr [edi], NULL
		push ecx
		invoke ComPtrAssign, edi, [esi].m_pEvent
		pop ecx
		add edi, sizeof LPDISPATCH
		mov esi, [esi].m_pNext
		inc ecx
		dec dwElements
		inc [ebx].dwIndex
		inc dwReturn
	.endw

	mov eax,pdwFetched
	.if (eax)
		mov ecx,dwReturn
		mov [eax],ecx
	.endif

	.if (dwElements == 0)
		mov eax,S_OK
	.else
		mov eax,S_FALSE
	.endif
	ret
	assume ebx:nothing
	assume esi:nothing

Next_ endp


;--------------------------------------------------------------------------

Skip_ PROC this_:ptr CEnumConnections,cConn:dword
	return E_FAIL
Skip_ endp

;--------------------------------------------------------------------------

Reset_ PROC this_:ptr CEnumConnections
	
;	DebugOut "IEnumConnections::Reset"
	mov ecx,this_
	mov [ecx].CEnumConnections.dwIndex,0
	return S_OK
Reset_ endp

;--------------------------------------------------------------------------

Clone_ PROC this_:ptr CEnumConnections,ppIECP:ptr LPENUMCONNECTIONS
	return E_FAIL
Clone_ endp

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_RUNNABLEOBJECT: defines IRunnableObject. Should not be
;--- used by server dlls
;--------------------------------------------------------------------------

ifndef RUNNABLEOBJECT_DEFINED

RUNNABLEOBJECT_DEFINED equ 1

externdef CRunnableObjectVtbl:			IRunnableObjectVtbl

DEFINE_RUNNABLEOBJECT macro objectname
local QueryInterface_, AddRef_, Release_, GetRunningClass_, Run_, IsRunning_, LockRunning_, SetContainedObject_

Category struct
dwID		dword ?
pszName		dword ?
pDispArray	dword ?
cntArray	dword ?
Category ends

    .const

CRunnableObjectVtbl IRunnableObjectVtbl {\
	QueryInterface_, AddRef_, Release_,\
	GetRunningClass_, Run_, IsRunning_, LockRunning_, SetContainedObject_}

    .code

;--------------------------------------------------------------------------
;IRunnableObject interface
;--------------------------------------------------------------------------

QueryInterface_:
	sub dword ptr [esp+4], objectname&._IRunnableObject
	jmp QueryInterface@&objectname
AddRef_:
	sub dword ptr [esp+4], objectname&._IRunnableObject
	jmp AddRef@&objectname
Release_:
	sub dword ptr [esp+4], objectname&._IRunnableObject
	jmp Release@&objectname

GetRunningClass_ PROC uses ebx edi this_:ptr objectname, lpClsId:ptr GUID

	DebugOut "IRunnableObject::GetRunningClass, lpClsId=%X", lpClsId

    return E_FAIL

GetRunningClass_ ENDP

;--------------------------------------------------------------------------

Run_ PROC this_:ptr objectname, lpbc:ptr

	DebugOut "IRunnableObject::Run, lpbc=%X", lpbc

    return S_OK

Run_ ENDP

;--------------------------------------------------------------------------

IsRunning_ PROC this_:ptr objectname

	DebugOut "IRunnableObject::IsRunning"

    return S_OK

IsRunning_ ENDP

;--------------------------------------------------------------------------

LockRunning_ PROC this_:ptr objectname, fLock:BOOL, fLastUnlockCloses:BOOL

	DebugOut "IRunnableObject::LockRunning, fLock=%u", fLock

    return S_OK

LockRunning_ ENDP

;--------------------------------------------------------------------------

SetContainedObject_ PROC this_:ptr objectname, fContained:BOOL

	DebugOut "IRunnableObject::SetContainedObject, fContained=%u", fContained

    return S_OK

SetContainedObject_ ENDP

	endm

endif


;--------------------------------------------------------------------------
;--- DEFINE_DISPATCH: defines IDispatch. The IDispatch vtable is NOT
;--- defned here, since this vtable will most likely be extended
;--- by class specific methods
;--------------------------------------------------------------------------


ifndef DISPATCH_DEFINED

DISPATCH_DEFINED equ 1

DEFINE_DISPATCH macro objectname, interfaceid
local SearchTypeInfo_

SendViewChange@&objectname		PROTO :ptr objectname

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
if objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
objectname&_IDispatch_TypeInfo_IID equ offset interfaceid
endif
endif

;--------------------------------------------------------------------------
; IDispatch interface
;--------------------------------------------------------------------------

if objectname&._IDispatch 
echo Error, vtable for IDispatch must be first!!!!
.err
endif

	.code

GetTypeInfoCount_ proc this_:ptr objectname, pCntinfo:ptr SDWORD

;	DebugOut "IDispatch::GetTypeInfoCount"

;if objectname&._IDispatch
;	sub this_, objectname&._IDispatch
;endif

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
	mov ecx, pCntinfo
	mov sdword ptr [ecx], 1		;we provide 1 type info
	return S_OK
else
	return E_FAIL
endif
GetTypeInfoCount_ endp


;--- search typeinfo of control in type library

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS

SearchTypeInfo_ proc lcid:LCID

LOCAL pTypeLib:LPTYPELIB
LOCAL pTypeInfo:LPTYPEINFO

    invoke LoadRegTypeLib, m_refIID_TypeLib, m_dwMajorVer,\
   				m_dwMinorVer, lcid, ADDR pTypeLib 
   	.if FAILED(eax)
		xor eax,eax
   		ret
   	.endif
	invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), addr interfaceid, ADDR pTypeInfo
;------------------- pTypeLib no longer needed
   	push eax
   	invoke vf(pTypeLib, ITypeLib, Release)
    pop eax
   	.if FAILED(eax)
		xor eax,eax
		ret 
    .endif
;--------------------- save the actual lcid in object data
   	mov eax, lcid
    mov m_lcid, eax
;--------------------- also save the matched pti
   	mov eax, pTypeInfo
    mov m_pTypeInfo, eax
	ret

SearchTypeInfo_ endp

endif
endif

GetTypeInfo_ proc uses __this this_:ptr objectname, iTypeInfo:DWORD, lcid:LCID, ppTInfo:ptr LPTYPEINFO
             
;	DebugOut "IDispatch::GetTypeInfo(Index=%u, LCID=%X)", iTypeInfo, lcid

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO

if objectname&._IDispatch
	sub this_, objectname&._IDispatch
endif

	mov __this, this_

	mov ecx, ppTInfo
	mov dword ptr [ecx],NULL

	.if (iTypeInfo != 0)
		return DISP_E_BADINDEX
	.endif

    mov eax, m_pTypeInfo

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
	.if (eax == NULL)
    	invoke SearchTypeInfo_, lcid
    .endif
endif

	push eax
	invoke ComPtrAssign, ppTInfo, eax
	pop eax

	.if (eax)
		mov eax,S_OK
	.else
		mov eax,DISP_E_UNKNOWNLCID
	.endif
	ret
else
	return E_FAIL
endif

GetTypeInfo_ endp

;---

GetIDsOfNames_ proc uses __this this_:ptr objectname, rrid:ptr IID, rgszNames:DWORD, cNames:DWORD, lcid:LCID, rgDispID:DWORD
			 
	DebugOut "IDispatch::GetIDsOfNames"

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO

if objectname&._IDispatch
	sub this_, objectname&._IDispatch
endif

	mov __this, this_

    mov eax, m_pTypeInfo
ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
	.if (eax == NULL)
    	invoke SearchTypeInfo_, lcid
    .endif
endif
	.if (eax)
		invoke vf( m_pTypeInfo, ITypeInfo, GetIDsOfNames), rgszNames, cNames, rgDispID
	.else
		mov eax, DISP_E_UNKNOWNLCID
	.endif
	ret
else
	return E_FAIL
endif

GetIDsOfNames_ endp

;-----------------------------------------------------------------------
;--- the main dispatcher. Most containers will call this function
;--- to set/get properties or call members. The real dispatching work
;--- inhere is done by ITypeInfo:Invoke. All we need to do is searching
;--- for a ITypeInfo pointer if we are called the first time
;-----------------------------------------------------------------------

Invoke_ proc uses __this this_:ptr objectname, dispIdMember:DISPID, riid:ptr IID, lcid:LCID, wFlags:DWORD, 
            pDispParams:ptr DISPPARAMS, pVarResult:ptr VARIANT, pExcepInfo:ptr EXCEPINFO, puArgErr:ptr DWORD

local pDispatch:LPDISPATCH

if 0 ; ifdef ?INVOKELOG
	DebugOut "IDispatch::Invoke DispID=%X", dispIdMember
endif

if objectname&._IDispatch
	sub this_, objectname&._IDispatch
endif
	mov __this, this_

;---------------------- ITypeInfo::Invoke requires a IDispatch pointer
if objectname&._IDispatch
    lea eax, m__IDispatch
    mov pDispatch,eax
else
    mov pDispatch,__this
endif

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPELIB_MEMBERS
    .if (!m_pTypeInfo)     
	    invoke SearchTypeInfo_, lcid
		.if (!eax)
			return DISP_E_MEMBERNOTFOUND
		.endif
    .endif
endif
endif
	invoke SetErrorInfo, NULL, NULL

ifdef OnInvoke@&objectname
	invoke OnInvoke@&objectname, __this, dispIdMember,\
			wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr
	.if (eax != -1)
		jmp done
	.endif
endif

ife objectname&_IDispatch_Flags and DISPATCH_NO_TYPEINFO
	invoke vf( m_pTypeInfo, ITypeInfo, Invoke_), pDispatch, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr
endif

done:

ifdef ?INVOKELOGERR
	.if (eax != S_OK)
		DebugOut "IDispatch::Invoke DispID=%X returned %X", dispIdMember, eax
	.endif
else
ifdef ?INVOKELOG
	DebugOut "IDispatch::Invoke DispID=%X(%d) returned %X", dispIdMember, dispIdMember, eax
endif
endif

	ret

Invoke_ endp

	endm

endif

;--------------------------------------------------------------------------
;--- DEFINE_DLLMAIN: define a DllMain (saves hInstance in g_hInstance)
;--------------------------------------------------------------------------

DEFINE_DLLMAIN macro
	
	.data

g_hInstance		HINSTANCE 0

	.code

ifdef _ALLOCDEBUG
prLocalAlloc typedef proto :DWORD, :DWORD
prLocalFree  typedef proto :HANDLE
LPLOCALALLOC typedef ptr prLocalAlloc
LPLOCALFREE  typedef ptr prLocalFree

		.data

g_cntLocal dd 0
g_LocalAlloc LPLOCALALLOC 0
g_LocalFree  LPLOCALFREE 0

		.code

myLocalAlloc proc uFlags:DWORD, uBytes:DWORD
		invoke g_LocalAlloc, uFlags, uBytes
		pushad
		sub esp, 128
		mov edx, esp
		inc g_cntLocal
		invoke wsprintf, edx, CStr(<"%3u LocalAlloc(%u) returned %X, called from %X",13,10>), g_cntLocal, uBytes, eax, dword ptr [ebp+4]
		invoke OutputDebugString, esp
		add esp, 128
		popad
		ret
myLocalAlloc endp

myLocalFree proc handle:HANDLE
		invoke g_LocalFree, handle
		pushad
		sub esp, 128
		mov edx, esp
		dec g_cntLocal
		invoke wsprintf, edx, CStr(<"%3u LocalFree(%X) returned %X, called from %X",13,10>), g_cntLocal, handle, eax, dword ptr [ebp+4]
		invoke OutputDebugString, esp
		add esp, 128
		popad
		ret
myLocalFree endp

endif

DllMain PROC public hInstance:HINSTANCE,dwReason:dword,lpReserved:dword

	mov	eax, dwReason
	.if (eax == DLL_PROCESS_ATTACH)
		mov		ecx, hInstance
		mov		g_hInstance, ecx
ifdef _ALLOCDEBUG
		push 0
		invoke VirtualProtect, addr _imp__LocalAlloc@8, 4, PAGE_READWRITE, esp 
		pop eax
		mov	eax, myLocalAlloc
		xchg eax, _imp__LocalAlloc@8
		mov g_LocalAlloc, eax
		mov	eax, myLocalFree
		xchg eax, _imp__LocalFree@4
		mov g_LocalFree, eax
endif
	.elseif (eax == DLL_PROCESS_DETACH)
	.endif
	mov	eax, 1
	ret
DllMain ENDP

	endm

;--------------------------------------------------------------------------
;--- DEFINE_UNREGISTERSERVER: define DllUnregisterServer (must be exported)
;--- in the debug version there are some security checks. Thats to avoid
;--- deleting all entries in HKEY_CLASSES_ROOT, CLSID or Typelib because 
;--- of an error in the table definition
;--------------------------------------------------------------------------

DEFINE_UNREGISTERSERVER macro pClsTab

	.code

DeleteKeyWithSubKeys proc uses ebx hKey:HANDLE,pszKey:ptr byte

local	szKey[MAX_PATH]:byte
local	hSubKey:HANDLE
local	filetime:FILETIME
local	dwSize:dword

		invoke RegOpenKeyEx,hKey,pszKey,NULL,KEY_ALL_ACCESS,addr hSubKey
		.if (eax == ERROR_SUCCESS)
			mov ebx,0
			.while (1)
				mov dwSize,sizeof szKey
				invoke RegEnumKeyEx,hSubKey,ebx,addr szKey,addr dwSize,NULL,NULL,NULL,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke DeleteKeyWithSubKeys,hSubKey,addr szKey
			.endw							
			invoke RegCloseKey,hSubKey
		.endif
		invoke RegDeleteKey,hKey,pszKey		;and delete subkey
		ret
DeleteKeyWithSubKeys endp


DllUnregisterServer PROC public uses ebx esi edi

local	szCLSID[40]:byte
local	wszGUID[40]:word
local	szSubKey[MAX_PATH]:byte

	mov edi, pClsTab
	mov ebx, ??ObjectMapItems
	.while (ebx)
;------------------------------ get the CLSID in string form
		invoke StringFromGUID2, [edi].ObjectEntry.pClsId, addr wszGUID,40
		invoke WideCharToLocal, addr szCLSID, addr wszGUID, sizeof szCLSID

		mov esi, [edi].ObjectEntry.pRegKeys
		.while (1)
			.if ([esi].REGSTRUCT.lpszSubKey == -1)
				.break .if ([esi].REGSTRUCT.lpszData == 0)
				invoke wsprintf, addr szSubKey, [esi].REGSTRUCT.lpszData, addr szCLSID
ifdef _DEBUG
				.if (!szSubKey)
					invoke DebugBreak
					return SELFREG_E_CLASS
				.endif
				invoke lstrcmpi, addr szSubKey, CStr("CLSID")
				.if (!eax)
					invoke DebugBreak
					return SELFREG_E_CLASS
				.endif
				invoke lstrcmpi, addr szSubKey, CStr("TYPELIB")
				.if (!eax)
					invoke DebugBreak
					return SELFREG_E_CLASS
				.endif
				invoke lstrcmpi, addr szSubKey, CStr("INTERFACE")
				.if (!eax)
					invoke DebugBreak
					return SELFREG_E_CLASS
				.endif
endif
				invoke DeleteKeyWithSubKeys, HKEY_CLASSES_ROOT, addr szSubKey
			.endif
			add esi, sizeof REGSTRUCT
		.endw
	    invoke UnRegisterTypeLib, [edi].ObjectEntry.pLibId,\
			[edi].ObjectEntry.dwVerMajor, [edi].ObjectEntry.dwVerMinor, 0, SYS_WIN32
		dec ebx
		add edi, sizeof ObjectEntry
	.endw
;------------------------------ Unregister Type Library and Interfaces


	return S_OK

DllUnregisterServer endp

	endm

;--------------------------------------------------------------------------
;--- DEFINE_REGISTERSERVER: define DllRegisterServer (must be exported)
;--------------------------------------------------------------------------

externdef g_hInstance:HINSTANCE

DEFINE_REGISTERSERVER macro pClsTab

	.code

DllRegisterServer PROC public uses ebx esi edi

local	hKey:HANDLE
local	dwDisp:dword
local   pTypeLib:LPTYPELIB
local	szCLSID[40]:byte
local	szLIBID[40]:byte
local	wszGUID[40]:word
local	szKeyPrefix[64]:byte
local	szSubKey[MAX_PATH]:byte
local	szModule[MAX_PATH]:byte
local	szData[MAX_PATH]:byte
local	wszModule[MAX_PATH]:word

	mov esi, pClsTab
	mov edi, ??ObjectMapItems
	.while (edi)
;------------------------------ get the CLSID in string form
		invoke StringFromGUID2, [esi].ObjectEntry.pClsId, addr wszGUID,40
		invoke WideCharToLocal, addr szCLSID, addr wszGUID, sizeof szCLSID

;------------------------------ get the LIBID in string form
		invoke StringFromGUID2, [esi].ObjectEntry.pLibId, addr wszGUID,40
		invoke WideCharToLocal, addr szLIBID, addr wszGUID, sizeof szLIBID

;------------------------------ get this DLL's path and file name
		invoke GetModuleFileName, g_hInstance,addr szModule,sizeof szModule

;------------------------------ register the CLSID entries
		mov ebx, [esi].ObjectEntry.pRegKeys
		assume ebx:ptr REGSTRUCT

		.while (1)
			.if ([ebx].lpszSubKey == -1)
				.break .if ([ebx].lpszData == 0)
				invoke wsprintf, addr szKeyPrefix, [ebx].lpszData, addr szCLSID
				add ebx, sizeof REGSTRUCT
				.continue
			.endif
			invoke lstrcpy, addr szSubKey, addr szKeyPrefix
			.if ([ebx].lpszSubKey)
				invoke lstrlen, addr szSubKey
				lea ecx, szSubKey
				add ecx, eax
				mov byte ptr [ecx],'\'
				inc ecx
;------------------------------ Create the sub key string.
				invoke lstrcpy, ecx, [ebx].lpszSubKey
			.endif

			invoke RegCreateKeyEx, HKEY_CLASSES_ROOT,\
					addr szSubKey,0,NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, \
					NULL, addr hKey, addr dwDisp
			.if (eax == NOERROR)
;------------------------------ if necessary, create the value string
	          .if ([ebx].lpszData)
					.if ([ebx].lpszData == -1)
						invoke lstrcpy,addr szData, addr szCLSID
					.elseif ([ebx].lpszData == -2)
						invoke lstrcpy,addr szData, addr szLIBID
					.else
					    invoke wsprintf, addr szData, [ebx].lpszData, addr szModule
					.endif
					invoke lstrlen, addr szData
					mov edx,eax
					inc edx
					invoke RegSetValueEx, hKey, [ebx].lpszValueName,\
			    			0, REG_SZ, addr szData, edx
	            .endif
				invoke	RegCloseKey, hKey
			.else
				return SELFREG_E_CLASS
			.endif

			add ebx,sizeof REGSTRUCT
		.endw

		add esi, sizeof ObjectEntry
		dec edi
	.endw

;------------------------------ Register Type Library and Interfaces

    invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,addr szModule,sizeof szModule,addr wszModule, sizeof wszModule 
	invoke LoadTypeLibEx, addr wszModule, REGKIND_REGISTER, addr pTypeLib
    .if (eax == S_OK)
	    invoke vf(pTypeLib,ITypeLib,Release)
    .endif

	return S_OK
	assume ebx:nothing

DllRegisterServer ENDP

	endm


;--------------------------------------------------------------------------
;--- DEFINE_CANUNLOADNOW: define DllCanUnloadNow (should be exported)
;--------------------------------------------------------------------------

DEFINE_CANUNLOADNOW macro

	.data

g_DllRefCount	DD 0

	.code

DllCanUnloadNow PROC public

	DebugOut "DllCanUnloadNow refcount=%u", g_DllRefCount
	xor		eax, eax
	cmp		g_DllRefCount, eax
	setne	al
	ret
DllCanUnloadNow ENDP

	endm

;--------------------------------------------------------------------------
;--- FIREEVENT: fires an event
;--------------------------------------------------------------------------

if 0	;;moved to DISPHLP.INC 

FIREEVENT macro this:req, interface, func
ifndef &interface&_&func&MDData
_TEXT$03	segment dword public 'CODE'
&interface&_&func&MDData label byte
	dd &interface&_&func&DispId
	dw &interface&_&func&MD
_TEXT$03	ends
endif
	exitm <FireEvent, addr &interface&_&func&MDData, this>
	endm
endif

