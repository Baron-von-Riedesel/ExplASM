
;--- interface viewer dll for COMView 1.7.0+ and OLEView
;--- currently viewer for IShellFolder is implemented

	.386
	.model flat, stdcall
	option casemap:none
	option proc:private

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN	equ 1
INCL_OLE2			equ 1
_WIN32_IE			equ 500h
	include windows.inc
	include commctrl.inc
	include windowsx.inc

	include shellapi.inc
	include objidl.inc
	include olectl.inc
	include shlguid.inc
	include shlobj.inc
	include shlwapi.inc
	include shobjidl.inc

	include macros.inc
	.list
	.cref

?AGGREGATION	= 0	;no aggregation
?EVENTSUPPORT	= 0	;no event support

	include olecntrl.inc
	include debugout.inc
	include rsrc.inc

INSIDE_CVViewer equ 1

	include CVViewer.inc
	include CShellBrowser.inc

	includelib  kernel32.lib
	includelib  advapi32.lib
	includelib  user32.lib
	includelib  gdi32.lib
	includelib  oleaut32.lib
	includelib  ole32.lib
	includelib  uuid.lib
	includelib  shell32.lib
	includelib  comctl32.lib

CVViewer struct

BEGIN_COM_MAP CVViewer
	COM_INTERFACE_ENTRY IInterfaceViewer
END_COM_MAP

pUnknown		LPUNKNOWN ?

CVViewer ends

;--------------------------------------------------------------------------

	.data

externdef g_DllRefCount:DWORD

	.const

CLSID_CVViewer GUID sCLSID_CVViewer


;--- the object table: defines coclasses installed by this module
;--- used by DllGetClassObject, DllRegisterServer + DllUnregisterServer

BEGIN_OBJECT_MAP ObjectMap
	ObjectEntry {\
		offset CLSID_CVViewer,\
		offset 0, 0, 0,\
		offset RegKeys_CVViewer,\
		Create@CVViewer}
END_OBJECT_MAP

;--------------------------------------------------------------------------

;--- define standard COM functions
;--- one instance for all coclasses in this module

	DEFINE_COMHELPER
	DEFINE_CLASSFACTORY
	DEFINE_GETCLASSOBJECT offset ObjectMap
	DEFINE_REGISTERSERVER offset ObjectMap
	DEFINE_UNREGISTERSERVER offset ObjectMap
	DEFINE_CANUNLOADNOW
	DEFINE_DLLMAIN		;DllMain std (saves hInstance in g_hInstance)

;-------------------------------------------------------------
;--- coclass CVViewer
;-------------------------------------------------------------

	.const

IID_IInterfaceViewer IID {0fc37e5bah, 4a8eh, 11ceh, {87h,0bh,08h,00h,36h,8dh,23h,02h}}

Description		textequ <"Interface Viewers for COMView and OLEView">

;--- registry infos for registration/unregister CVViewer coclass

RegKeys_CVViewer label REGSTRUCT
	REGSTRUCT <-1, 0, CStr("CLSID\%s")>
	REGSTRUCT <0, 0, CStr(Description)>
	REGSTRUCT <CStr("InprocServer32"), 0, CStr("%s")>
	REGSTRUCT <CStr("InprocServer32"), CStr("ThreadingModel"), CStr("Apartment")>
	REGSTRUCT <-1, 0, CStr("Interface\{000214E6-0000-0000-C000-000000000046}")>
	REGSTRUCT <CStr("OLEViewerIViewerCLSID"), 0, -1>
	REGSTRUCT <-1, 0, CStr("Interface\{000214E4-0000-0000-C000-000000000046}")>
	REGSTRUCT <CStr("OLEViewerIViewerCLSID"), 0, -1>
	REGSTRUCT <-1, 0, 0>

;--- vtable for interface IInterfaceViewer

CInterfaceViewerVtbl label dword
	IUnknownVtbl {QueryInterface@CVViewer, AddRef@CVViewer, Release@CVViewer}
	dd View

;--- define interfaces known by IUnknown::QueryInterface

	DEFINE_KNOWN_INTERFACES CVViewer, IInterfaceViewer

	.code

__this	textequ <ebx>
_this	textequ <[__this].CVViewer>

	MEMBER pUnknown

;--- standard code (won't be much here besides IUnknown methods)

	DEFINE_STD_COM_METHODS CVViewer

;--- constructor coclass CVViewer

Create@CVViewer	proc public uses esi __this pClass: ptr ObjectEntry, pUnkOuter:LPUNKNOWN
	
	invoke LocalAlloc, LMEM_FIXED or LMEM_ZEROINIT,sizeof CVViewer
	.if (eax == NULL)
		ret
	.endif
	mov __this,eax

	mov	m__IInterfaceViewer, OFFSET CInterfaceViewerVtbl

	STD_COM_CONSTRUCTOR CVViewer

	return __this

Create@CVViewer	endp

;--- destructor coclass CVViewer

Destroy@CVViewer proc public uses __this this_:ptr CVViewer

    mov __this,this_

	STD_COM_DESTRUCTOR CVViewer

	.if (m_pUnknown)
		invoke vf(m_pUnknown, IUnknown, Release)
	.endif

	invoke LocalFree, __this
    ret

Destroy@CVViewer endp

IDM_FIRST equ 1
IDM_LAST	equ 100

DisplayMenu proc hWnd:HWND, pUnknown:LPUNKNOWN

local hMenu:HMENU
local pContextMenu:LPCONTEXTMENU
local pShellExtInit:LPSHELLEXTINIT
local pt:POINT

	invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IContextMenu, addr pContextMenu
	.if (eax == S_OK)
if 0
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IShellExtInit, addr pShellExtInit
		.if (eax == S_OK)
			invoke vf(pShellExtInit, IShellExtInit, Initialize), xx,yy,zz
			invoke vf(pShellExtInit, IUnknown, Release)
		.endif
endif
		invoke CreatePopupMenu
		mov hMenu, eax
		invoke vf(pContextMenu, IContextMenu, QueryContextMenu), hMenu, \
			0, IDM_FIRST, IDM_LAST, CMF_NORMAL
		invoke GetCursorPos, addr pt
		invoke TrackPopupMenu, hMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
			pt.x, pt.y, 0, hWnd, NULL
		invoke DestroyMenu, hMenu
		invoke vf(pContextMenu, IUnknown, Release)
	.endif
	ret

DisplayMenu endp

;----------------------------------------------------------------

;--- IInterfaceViewer::View method 

View	proc uses __this this_:ptr CVViewer, hwndParent:HWND, riid:REFIID, punk:LPUNKNOWN

	mov __this, this_
	mov eax, punk
	mov m_pUnknown, eax
	invoke vf(eax, IUnknown, AddRef)

	invoke IsEqualGUID, riid, addr IID_IShellFolder
	.if (eax)
		invoke Create@CShellBrowser, punk
		.if (eax)
			invoke Show@CShellBrowser, eax, NULL;hwndParent
		.endif
		mov eax, S_OK
		ret
	.endif
	invoke IsEqualGUID, riid, addr IID_IContextMenu
	.if (eax)
		invoke DisplayMenu, hwndParent, punk
		mov eax, S_OK
		ret
	.endif
	return S_OK

View	endp

end DllMain
