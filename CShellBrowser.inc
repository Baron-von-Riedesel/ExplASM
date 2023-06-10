
?SERVICEPROVIDER	equ 1		;support IServiceProvider
?OLECOMMANDTARGET	equ 1		;support IOleCommandTarget
?DROPTARGET			equ 1		;support IDropTarget
?DRAGDROPHELPER		equ 1		;use DragDropHelper
?OLEINPLACEFRAME	equ 0		;support IOleInPlaceFrame+IOleInPlaceUIWindow
?WEBBROWSER			equ 0		;browsing in explorer for win98 (not finished)

LPDROPTARGETHELPER typedef ptr IDropTargetHelper

if ?SERVICEPROVIDER
	include servprov.inc
endif
if ?OLECOMMANDTARGET
	include docobj.inc
endif
if ?WEBBROWSER
	include exdisp.inc
endif

CShellBrowser struct

ifdef INSIDE_CShellBrowser

BEGIN_COM_MAP CShellBrowser
	COM_INTERFACE_ENTRY IShellBrowser
if ?SERVICEPROVIDER
	COM_INTERFACE_ENTRY IServiceProvider
endif
if ?OLECOMMANDTARGET
	COM_INTERFACE_ENTRY IOleCommandTarget
endif
if ?DROPTARGET
	COM_INTERFACE_ENTRY IDropTarget
endif
if ?OLEINPLACEFRAME
	COM_INTERFACE_ENTRY_EX IOleInPlaceFrame, IOleInPlaceUIWindow, IOleWindow
endif
END_COM_MAP

hWnd			HWND ?
hWndTV			HWND ?
hWndView		HWND ?
hWndSB			HWND ?
hWndFocus		HWND ?
pShellFolder	LPSHELLFOLDER ?
pShellView		LPSHELLVIEW ?
pContextMenu2	LPCONTEXTMENU2 ?
if ?OLEINPLACEFRAME
pOleInPlaceActiveObject	LPOLEINPLACEACTIVEOBJECT ?
endif
if ?WEBBROWSER
pWebBrowser		LPWEBBROWSER2 ?
endif
pMalloc			LPMALLOC ?
pStream			LPSTREAM ?
hSHNotify		HANDLE ?
dwSizeTV		DWORD ?
rect			RECT <>
bCreateView		BOOLEAN ?	;create right panel views
bGetUIObjectOf	BOOLEAN ?	;call IShellFolder::GetUIObjectOf for context menu
bCompMenus		BOOLEAN ?	;support composite menus
bStatusBar		BOOLEAN ?	;status bar is visible
bDropTarget		BOOLEAN ?	;treeview is a drop target
bDontRespond	BOOLEAN ?	;do not respond to DDE messages
bLabelEdit		BOOLEAN ?	;currently in label edit mode
bRButton		BOOLEAN ?	;drag&drop with right mouse button?

endif

CShellBrowser ends


Create@CShellBrowser		PROTO :LPUNKNOWN
Show@CShellBrowser			PROTO :ptr CShellBrowser, :HWND
TranslateAccelerator@CShellBrowser PROTO :ptr CShellBrowser, :ptr MSG
GetFullPidl@CShellBrowser	PROTO :HANDLE
Deinit@IServiceProvider		PROTO
GetFolder					PROTO :HANDLE
InsertChildItems			PROTO :HANDLE
_StrToLong					PROTO :LPSTR
FindPidl					PROTO :LPITEMIDLIST, :BOOL
NavigateToPidl				PROTO :LPITEMIDLIST
Pidl_Copy					PROTO :LPITEMIDLIST
Pidl_GetLastItem			PROTO :LPITEMIDLIST
Pidl_SkipLastItem			PROTO :LPITEMIDLIST, :ptr LPITEMIDLIST

externdef CServiceProviderVtbl: DWORD
externdef COleCommandTargetVtbl: DWORD
externdef CDropTargetVtbl: DWORD
externdef COleInPlaceFrameVtbl: DWORD

externdef g_DllRefCount:DWORD
externdef g_bGetUIObjectOf:BOOLEAN
externdef g_bCreateView:BOOLEAN
externdef g_bCompMenus:BOOLEAN
externdef g_aApplication:DWORD
externdef g_aTopic:DWORD
externdef g_rect:RECT
externdef g_szPath:BYTE
if ?DRAGDROPHELPER
externdef g_pDropTargetHelper:LPDROPTARGETHELPER
endif

;--- macros + equates

@MakeStubs macro classname, interface, memberlist:VARARG
    for member,<memberlist>
member&_:
	sub DWORD ptr [esp+4], classname&._&interface
	jmp member
	endm
	endm

@MakeStubsEx macro classname, interface, memberlist:VARARG
    for member,<memberlist>
member&_:
	sub DWORD ptr [esp+4], classname&._&interface
%	jmp member&@&classname
	endm
	endm

ifdef _DEBUG
externdef DEBUGPREFIX:LPSTR
endif
