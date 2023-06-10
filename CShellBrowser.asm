
;--- CShellBrowser object
;--- that's a simple explorer

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
	include dde.inc
	include windowsx.inc
    include wincrack.inc

	include shellapi.inc
	include objidl.inc
	include olectl.inc
	include shlguid.inc
	include shlobj.inc
	include shlwapi.inc
	include shobjidl.inc

?AGGREGATION	= 0	;no aggregation
?EVENTSUPPORT	= 0	;no event support

	include olecntrl.inc
	include debugout.inc
	include rsrc.inc
	.list
	.cref

	include macros.inc


INSIDE_CShellBrowser equ 1

	include CShellBrowser.inc


SFGAOF typedef DWORD

?IMAGELIST			equ 1		;use IExtractIcon and set icons in tree view
?RETURNCMD			equ 1

WM_SHNOTIFY			equ 401h	;window message for shell change notifications
WM_GETISHELLBROWSER equ 407h	;undocumented window message

;--- pointer to SHNOTIFYSTRUCT is in wParam of WM_SHNOTIFY

SHNOTIFYSTRUCT struct
dwItem1 DWORD ?
dwItem2 DWORD ?
SHNOTIFYSTRUCT ends


IDM_FIRST			equ FCIDM_SHVIEWFIRST
IDM_LAST			equ IDM_FIRST+100


WndProc						PROTO :HWND, :DWORD, :WPARAM, :LPARAM
WndProc@CShellBrowser		PROTO :ptr CShellBrowser, :DWORD, :WPARAM, :LPARAM
Destroy@CShellBrowser		PROTO :ptr CShellBrowser
SetShellView@CShellBrowser	PROTO :ptr CShellBrowser, :LPSHELLVIEW
;SHBindToParent				PROTO :LPITEMIDLIST, :REFIID, :ptr LPUNKNOWN, :ptr LPITEMIDLIST
OnSize						PROTO
InsertChildItems			PROTO :HANDLE
GetFolder					PROTO :HANDLE
GetDDEVariables				PROTO
OnDDEExecute				PROTO :LPSTR
TranslateAcceleratorSB		PROTO :ptr CShellBrowser, lpmsg:ptr MSG, wID:WORD


	.data

g_hAccel		HACCEL NULL
g_hCursor		HCURSOR NULL
g_hHalftoneBrush HBRUSH NULL
g_hWndFocus		HWND NULL
g_dwCount		DWORD 0
g_rect			RECT {0,0,0,0}
g_himl			HANDLE NULL
g_EditWndProc	DWORD 0
if ?DRAGDROPHELPER
g_pDropTargetHelper	LPDROPTARGETHELPER NULL
endif
g_szPath db MAX_PATH dup(9)


protoSHChangeNotifyRegister typedef proto :HWND, :DWORD, :DWORD, :DWORD, :DWORD, :ptr LPITEMIDLIST
protoSHChangeNotifyUnregister typedef proto :HANDLE
protoSHParseDisplayName	typedef proto :ptr WORD, :DWORD, :ptr LPITEMIDLIST, :SFGAOF, :ptr SFGAOF

LPFNSHCHANGENOTIFYREGISTER		typedef ptr protoSHChangeNotifyRegister
LPFNSHCHANGENOTIFYUNREGISTER	typedef ptr protoSHChangeNotifyUnregister
LPFNSHPARSEDISPLAYNAME			typedef ptr protoSHParseDisplayName

g_lpfnSHChangeNotifyRegister	LPFNSHCHANGENOTIFYREGISTER NULL
g_lpfnSHChangeNotifyUnregister	LPFNSHCHANGENOTIFYUNREGISTER NULL
g_lpfnSHParseDisplayName		LPFNSHPARSEDISPLAYNAME NULL

g_bSort			BOOLEAN TRUE
g_bRegistered	BOOLEAN FALSE
g_bGetUIObjectOf BOOLEAN TRUE
g_bCreateView	BOOLEAN FALSE
g_bCompMenus	BOOLEAN TRUE
g_bStatusBar	BOOLEAN TRUE
g_bDropTarget	BOOLEAN TRUE
g_bTrackSelect	BOOLEAN TRUE
g_bBrowseFiles	BOOLEAN TRUE

	.const

	align 4

;IID_IShellBrowser	sIID_IShellBrowser
;IID_IShellFolder	sIID_IShellFolder
;IID_IShellView		sIID_IShellView
;IID_IContextMenu	sIID_IContextMenu
if ?DRAGDROPHELPER
;CLSID_DragDropHelper sCLSID_DragDropHelper
;IID_IDropTargetHelper sIID_IDropTargetHelper
endif

;--- vtable for interface IShellBrowser

CShellBrowserVtbl label dword
	IUnknownVtbl {QueryInterface@CShellBrowser, AddRef@CShellBrowser, Release@CShellBrowser}
	dd GetWindow_, ContextSensitiveHelp
	dd InsertMenusSB, SetMenuSB, RemoveMenusSB
	dd SetStatusTextSB, EnableModelessSB, TranslateAcceleratorSB
	dd BrowseObject, GetViewStateStream, GetControlWindow
	dd SendControlMsg, QueryActiveShellView, OnViewWindowActive, SetToolbarItems


if ?OLECOMMANDTARGET
@if1 textequ <, IOleCommandTarget>
else
@if1 textequ <>
endif
if ?SERVICEPROVIDER
@if2 textequ <, IServiceProvider>
else
@if2 textequ <>
endif
if ?DROPTARGET
@if3 textequ <, IDropTarget>
else
@if3 textequ <>
endif
if ?OLEINPLACEFRAME
@if4 textequ <, IOleWindow, IOleInPlaceUIWindow, IOleInPlaceFrame>
else
@if4 textequ <>
endif

%	DEFINE_KNOWN_INTERFACES CShellBrowser, IShellBrowser @if1 @if2 @if3 @if4

szAppNameEx db " - "
szAppName db "ExplorerASM",0	

ifdef _DEBUG
DEBUGPREFIX LPSTR CStr("ExplorerASM:")
endif

szShell32 db "Shell32.dll",0

	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER _IShellBrowser, ObjRefCount
	MEMBER hWnd, hWndTV, hWndView, hWndSB, hWndFocus
	MEMBER pShellFolder, pShellView, pContextMenu2, pStream, pMalloc
	MEMBER rect, dwSizeTV, hSHNotify
	MEMBER bCreateView, bGetUIObjectOf, bCompMenus, bStatusBar
	MEMBER bDropTarget, bDontRespond, bLabelEdit
if ?OLEINPLACEFRAME
	MEMBER _IOleInPlaceFrame, pOleInPlaceActiveObject
endif
if ?WEBBROWSER
	MEMBER pWebBrowser
endif

	DEFINE_STD_COM_METHODS CShellBrowser


;--- InitApp: this proc is called only once


InitApp proc

local hShell32:HINSTANCE
local pShellIcon:LPSHELLICON
local wNullPidl:WORD
local iIndex:DWORD
local osvi:OSVERSIONINFO
local hIcon:HICON
local dwSize:DWORD
local dwType:DWORD
local szPath[MAX_PATH]:byte

		mov osvi.dwOSVersionInfoSize, sizeof OSVERSIONINFO
		invoke GetVersionEx, addr osvi

		invoke LoadLibrary, addr szShell32
		mov hShell32, eax

		.if (eax > 32)
			invoke GetProcAddress, hShell32, 2
			mov g_lpfnSHChangeNotifyRegister, eax
			invoke GetProcAddress, hShell32, 4
			mov g_lpfnSHChangeNotifyUnregister, eax
			invoke GetProcAddress, hShell32, CStr("SHParseDisplayName")
			mov g_lpfnSHParseDisplayName, eax
		.endif

		invoke LoadCursor, g_hInstance, IDC_CURSOR1
		mov g_hCursor, eax
		invoke LoadAccelerators, g_hInstance, IDR_ACCELERATOR1
		mov g_hAccel, eax
if ?IMAGELIST
if 0
		invoke GetSystemMetrics, SM_CXSMICON	;returns 0014h ?????
else
		mov eax, 0010h
endif
		.if ((osvi.dwPlatformId == VER_PLATFORM_WIN32_NT) && (osvi.dwMajorVersion >= 5))
			invoke ImageList_Create, eax, eax, ILC_COLOR32 or ILC_MASK, 4, 4
		.else
			mov g_bBrowseFiles, FALSE
			invoke ImageList_Create, eax, eax, ILC_COLORDDB or ILC_MASK, 4, 4
		.endif
		mov g_himl, eax
if 0
		mov dwSize, sizeof szPath
		invoke SHGetValue, HKEY_CLASSES_ROOT, CStr("Folder\DefaultIcon"),CStr(""),
			addr dwType, addr szPath, addr dwSize
		.if (eax == S_OK)
			mov eax, ','
			invoke StrChr, addr szPath, eax
			.if (eax)
				mov byte ptr [eax], 0
				xor eax, eax
				invoke ExtractIcon, g_hInstance, addr szPath, eax
				push eax
				invoke ImageList_AddIcon, g_himl, eax
				pop eax
				invoke DestroyIcon, eax
			.endif
		.endif
else
		invoke ExtractIcon, g_hInstance, addr szShell32, 3
		.if (eax)
			push eax
			invoke ImageList_AddIcon( g_himl, eax)
			pop eax
			invoke DestroyIcon, eax
		.endif
		invoke ExtractIcon, g_hInstance, addr szShell32, 4
		.if (eax)
			push eax
			invoke ImageList_AddIcon( g_himl, eax)
			pop eax
			invoke DestroyIcon, eax
		.endif
endif
if 0
		mov iIndex, 34
		invoke vf(m_pShellFolder, IUnknown, QueryInterface), addr IID_IShellIcon, addr pShellIcon
		.if (eax == S_OK)
			mov wNullPidl,0
			invoke vf(pShellIcon, IShellIcon, GetIconOf), addr wNullPidl, GIL_FORSHELL, addr iIndex
			invoke vf(pShellIcon, IUnknown, Release)
		.endif
		invoke ExtractIcon, g_hInstance, addr szShell32, iIndex
else
		invoke ExtractIcon, g_hInstance, addr szShell32, 34
endif
		.if (eax)
			push eax
			invoke ImageList_AddIcon( g_himl, eax)
			pop eax
			invoke DestroyIcon, eax
		.endif
endif
		invoke GetDDEVariables
if ?DRAGDROPHELPER
		invoke CoCreateInstance, addr CLSID_DragDropHelper, NULL,
			CLSCTX_INPROC_SERVER, addr IID_IDropTargetHelper, addr g_pDropTargetHelper
endif

		invoke FreeLibrary, hShell32

		ret

InitApp endp


;--- create a CShellBrowser object


Create@CShellBrowser proc public uses __this pUnknown:LPUNKNOWN


	DebugOut "Create@CShellBrowser"
	invoke LocalAlloc, LMEM_FIXED or LMEM_ZEROINIT,sizeof CShellBrowser
	.if (eax == NULL)
		ret
	.endif
	mov __this,eax

	mov m__IShellBrowser.lpVtbl, offset CShellBrowserVtbl

if ?SERVICEPROVIDER
	mov m__IServiceProvider.lpVtbl, offset CServiceProviderVtbl
endif
if ?OLECOMMANDTARGET
	mov m__IOleCommandTarget.lpVtbl, offset COleCommandTargetVtbl
endif
if ?DROPTARGET
	mov m__IDropTarget.lpVtbl, offset CDropTargetVtbl
endif
if ?OLEINPLACEFRAME
	mov m__IOleInPlaceFrame.lpVtbl, offset COleInPlaceFrameVtbl
endif

	STD_COM_CONSTRUCTOR CShellBrowser

	invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IShellFolder, addr m_pShellFolder
	.if (eax != S_OK)
		invoke Destroy@CShellBrowser, __this
		return 0
	.endif

	mov m_dwSizeTV, -1

	invoke SHGetMalloc, addr m_pMalloc

	.if (!g_dwCount)
		invoke InitApp
	.endif

	inc g_dwCount

	return __this
	align 4

Create@CShellBrowser endp

;--- delete a CShellBrowser object

Destroy@CShellBrowser proc uses __this this_:ptr CShellBrowser

	DebugOut "Destroy@CShellBrowser(%X)", this_

	mov __this, this_

	STD_COM_DESTRUCTOR CShellBrowser

if ?WEBBROWSER
	.if (m_pWebBrowser)
		invoke vf(m_pWebBrowser, IUnknown, Release)
	.endif
endif
	.if (m_pShellView)
		invoke vf(m_pShellView, IUnknown, Release)
	.endif
	.if (m_pShellFolder)
		invoke vf(m_pShellFolder, IUnknown, Release)
	.endif
	.if (m_pStream)
		invoke vf(m_pStream, IUnknown, Release)
	.endif
	.if (m_pMalloc)
		invoke vf(m_pMalloc, IUnknown, Release)
	.endif

	dec g_dwCount
	.if (!g_dwCount)
if ?DRAGDROPHELPER
		.if (g_pDropTargetHelper)
			invoke vf(g_pDropTargetHelper, IUnknown, Release)
		.endif
endif
if ?IMAGELIST
		.if (g_himl)
			invoke ImageList_Destroy, g_himl
		.endif
endif
		.if (g_aApplication)
			invoke GlobalDeleteAtom, g_aApplication
		.endif
		.if (g_aTopic)
			invoke GlobalDeleteAtom, g_aTopic
		.endif
	.endif

	invoke LocalFree, this_
	ret
	align 4

Destroy@CShellBrowser endp


Show@CShellBrowser proc public uses __this this_:ptr CShellBrowser, hWnd:HWND

local wc:WNDCLASSEX

	mov __this, this_
	.if (hWnd)
		invoke GetWindowRect, hWnd, addr m_rect
		mov eax, m_rect.right
		sub eax, m_rect.left
		mov m_rect.right, eax
		mov eax, m_rect.bottom
		sub eax, m_rect.top
		mov m_rect.bottom, eax
		add m_rect.left, 16
		add m_rect.top, 16
	.endif
	.if (!g_bRegistered)
		invoke RtlZeroMemory, addr wc, sizeof WNDCLASSEX
		mov wc.cbSize, sizeof WNDCLASSEX
		mov wc.style, 0
		mov wc.lpfnWndProc, WndProc
		mov eax, g_hInstance
		mov wc.hInstance, eax
		invoke LoadIcon, g_hInstance, IDI_ICON1
		mov wc.hIcon, eax
		invoke LoadCursor, NULL, IDC_ARROW
		mov wc.hCursor, eax
		mov wc.hbrBackground, COLOR_BTNFACE + 1
		mov wc.lpszMenuName, IDR_MENU1
		mov wc.lpszClassName, offset szAppName
		mov wc.hIconSm, NULL
		invoke RegisterClassEx, addr wc
		mov g_bRegistered, TRUE
	.endif
	invoke CreateWindowEx, 0, offset szAppName, CStr(""),
		WS_OVERLAPPEDWINDOW or WS_VISIBLE, m_rect.left, m_rect.top, m_rect.right, m_rect.bottom,\
		hWnd, NULL, g_hInstance, __this
	ret
	align 4

Show@CShellBrowser endp

;--- send key to view (but do NOT if treeview is in label edit mode)

TranslateAccelerator@CShellBrowser proc public uses __this this_:ptr CShellBrowser, lpmsg:ptr MSG

	mov __this, this_
	mov eax, S_FALSE
	.if (!m_bLabelEdit)
		mov eax, m_hWndFocus
		.if (eax == m_hWndView)
			invoke vf(m_pShellView, IShellView, TranslateAccelerator), lpmsg
			.if (eax == S_FALSE)
				invoke TranslateAcceleratorSB, __this, lpmsg, NULL
			.endif
		.else
			invoke TranslateAcceleratorSB, __this, lpmsg, NULL
			.if (eax == S_FALSE && m_pShellView)
				invoke vf(m_pShellView, IShellView, TranslateAccelerator), lpmsg
			.endif
		.endif
	.endif
	ret
	align 4

TranslateAccelerator@CShellBrowser endp


;----------------------------------------------------------------

if 0
;--- this function is not available in win9x and winnt systems

SHBindToParent proc uses esi edi pidl:LPITEMIDLIST, riid:REFIID, ppUnknown:ptr LPUNKNOWN, ppidl:ptr LPITEMIDLIST

local pShellFolder:LPSHELLFOLDER
local hr:DWORD

	xor ecx, ecx

	mov eax, ppidl
	.if (eax)
		mov [eax],ecx
	.endif
	mov eax, ppUnknown
	mov [eax], ecx

	mov edx, pidl
	.while (edx)
		movzx eax,[edx].SHITEMID.cb
		.break .if (!eax)
		mov esi, ecx
		add ecx, eax
		lea edx, [eax][edx]
	.endw
	add esi, sizeof WORD
	invoke vf(m_pMalloc, IMalloc, Alloc), esi
	.if (eax)
		mov ecx, esi
		mov esi, pidl
		mov edi, eax
		rep movsb
		mov [edi-2],cx
		mov edi, eax
		invoke vf(m_pShellFolder, IShellFolder, BindToObject), edi, NULL,\
			riid, addr pShellFolder
		mov hr, eax
		.if (eax == S_OK)
			.if (ppidl)
				mov ecx, ppidl
				mov [ecx],edi
			.else
				invoke vf(m_pMalloc, IMalloc, Free), edi
			.endif
			mov ecx, ppUnknown
			mov eax, pShellFolder
			mov [ecx], eax
		.else
			invoke vf(m_pMalloc, IMalloc, Free), edi
		.endif
		mov eax, hr
	.else
		mov eax, E_OUTOFMEMORY
	.endif
	ret
	align 4

SHBindToParent endp
endif

;----------------------------------------------------------------

;--- release IShellView

ReleaseShellView proc uses esi

	.if (m_pShellView)
		invoke vf(m_pShellView, IShellView, UIActivate), SVUIA_DEACTIVATE
		.if (m_hWndView)
			mov eax, m_hWndFocus
			.if (eax == m_hWndView)
				mov m_hWndFocus, NULL
			.endif
			invoke vf(m_pShellView, IShellView, DestroyViewWindow)
			mov m_hWndView, NULL
		.endif
		invoke vf(m_pShellView, IUnknown, Release)
		mov m_pShellView, NULL
	.endif
	ret
	align 4

ReleaseShellView endp


;--- update menu items

UpdateMenu proc

local hMenu:HMENU

	invoke GetMenu, m_hWnd
	mov hMenu, eax

	.if (m_bStatusBar)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_STATUSLINE, ecx

	.if (m_bCreateView)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_CREATEVIEWOBJECT, ecx

	.if (m_bGetUIObjectOf)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_GETUIOBJECTOF, ecx

	.if (m_bCompMenus)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_COMPMENUS, ecx

	.if (m_bDropTarget)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_DROPTARGET, ecx

	.if (g_bTrackSelect)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_TRACKSELECT, ecx

	.if (!m_bDontRespond)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_DDERESPOND, ecx

	.if (g_bBrowseFiles)
		mov ecx, MF_CHECKED
	.else
		mov ecx, MF_UNCHECKED
	.endif
	invoke CheckMenuItem, hMenu, IDM_BROWSEFILES, ecx

	ret
	align 4

UpdateMenu endp

;--- get attributes of an item

GetAttributes proc hItem:HANDLE, pdwAttributes:ptr DWORD

local pShellFolder:LPSHELLFOLDER
local pItemIDList:LPITEMIDLIST
local tvi:TV_ITEM

	mov eax, hItem
	mov tvi.hItem, eax
	mov tvi.mask_, TVIF_PARAM
	invoke TreeView_GetItem( m_hWndTV, addr tvi)
	mov eax, tvi.lParam
	.if (!eax)
		ret
	.endif
	mov pItemIDList, eax

	invoke TreeView_GetParent( m_hWndTV, hItem)
	invoke GetFolder, eax
	.if (eax)
		mov pShellFolder, eax
		invoke vf(pShellFolder, IShellFolder, GetAttributesOf), 1, addr pItemIDList,
			pdwAttributes
		push eax
		invoke vf(pShellFolder, IShellFolder, Release)
		pop eax
		.if (eax == S_OK)
			mov eax, TRUE
		.else
			xor eax, eax
		.endif
	.endif
	ret
	align 4

GetAttributes endp


if 0

;--- check if a folder is on a removeable device

IsRemoveable proc hItem:HANDLE

local sfgaof:SFGAOF

	mov sfgaof, SFGAO_REMOVABLE
	invoke GetAttributes, hItem, addr sfgaof
	.if ((eax) && (sfgaof & SFGAO_REMOVABLE))
		mov eax, TRUE
	.else
		mov eax, FALSE
	.endif
	ret
	align 4

IsRemoveable endp

endif

StrRet2String proc pItemIDList:LPITEMIDLIST, pstrret:ptr STRRET, pszText:LPSTR, dwMax:DWORD

	mov ecx, pstrret
	.if ([ecx].STRRET.uType == STRRET_WSTR)
		invoke WideCharToMultiByte, CP_ACP, 0, [ecx].STRRET.pOleStr, -1, pszText, dwMax, 0, 0
		mov ecx, pstrret
		invoke vf(m_pMalloc, IMalloc, Free), [ecx].STRRET.pOleStr
		mov eax, pszText
	.elseif ([ecx].STRRET.uType == STRRET_CSTR)
		lea eax, [ecx].STRRET.cStr
	.else
		mov eax, pItemIDList
		add eax,[ecx].STRRET.uOffset
	.endif
	ret
	align 4

StrRet2String endp

HRESULT_CODE macro
	movsx eax, ax
	endm

SortChildrenCB proc lParam1:LPARAM, lParam2:LPARAM, lParamSort:LPARAM

local pShellFolder:LPSHELLFOLDER

	mov eax, lParamSort
	mov pShellFolder, eax
	invoke vf(pShellFolder, IShellFolder, CompareIDs), 0, lParam1, lParam2
	HRESULT_CODE
	ret
	align 4

SortChildrenCB endp


?TESTMODE equ 0

if 1

InsertItem proc hItem:HANDLE, pShellFolder:LPSHELLFOLDER, pidl:LPITEMIDLIST, bValidate:BOOL

local dwRC:BOOL
local sfgaof:SFGAOF
local strret:STRRET
local tvi:TV_INSERTSTRUCT
local szText[MAX_PATH]:byte

	mov eax, hItem
	mov tvi.hParent, eax
	.if (bValidate)
		mov tvi.hInsertAfter, TVI_SORT
	.else
		mov tvi.hInsertAfter, TVI_LAST
	.endif
	lea eax, szText
	mov tvi.item.pszText, eax
	mov szText,0

	mov dwRC, FALSE
	mov strret.uType, STRRET_CSTR
	invoke vf(pShellFolder, IShellFolder, GetDisplayNameOf), pidl,\
			SHGDN_INFOLDER, addr strret
	.if (eax == S_OK)
		mov tvi.item.mask_, TVIF_TEXT or TVIF_PARAM or TVIF_CHILDREN
		mov eax, pidl
		mov tvi.item.lParam, eax
		mov sfgaof, SFGAO_HASSUBFOLDER
		.if (g_bBrowseFiles)
			or sfgaof, SFGAO_BROWSABLE or SFGAO_FOLDER
		.endif
if ?TESTMODE
		or sfgaof, SFGAO_COMPRESSED or SFGAO_REMOVABLE
endif
		invoke vf(pShellFolder, IShellFolder, GetAttributesOf), 1, addr pidl,\
			addr sfgaof
		.if ((eax == S_OK) && (sfgaof & SFGAO_HASSUBFOLDER))
			mov tvi.item.cChildren, 1
		.else
			.if (g_bBrowseFiles)
				mov tvi.item.cChildren, I_CHILDRENCALLBACK
			.else
				mov tvi.item.cChildren, 0
			.endif
		.endif
if ?IMAGELIST
		or tvi.item.mask_, TVIF_IMAGE or TVIF_SELECTEDIMAGE
		mov tvi.item.iSelectedImage, I_IMAGECALLBACK
		mov tvi.item.iImage, I_IMAGECALLBACK
endif
		.if ((!g_bBrowseFiles) || (sfgaof & (SFGAO_BROWSABLE or SFGAO_FOLDER)))
			invoke StrRet2String, pidl, addr strret, tvi.item.pszText, MAX_PATH
			mov tvi.item.pszText, eax
if ?TESTMODE
			lea ecx, szText
			.if (ecx != eax)
				mov tvi.item.pszText, ecx
				invoke wsprintf, addr szText, CStr("%s,%X"), eax, sfgaof
			.else
				invoke lstrlen, ecx
				lea ecx, szText
				add ecx, eax
				invoke wsprintf, ecx, CStr(",%X"), sfgaof
			.endif
endif
			invoke TreeView_InsertItem( m_hWndTV, addr tvi)
			mov dwRC, TRUE
		.endif
	.endif
	return dwRC

InsertItem endp


;--- returns number of childs


InsertChildItems proc public hItem:HANDLE

local pShellFolder:LPSHELLFOLDER
local pEnumIDList:LPENUMIDLIST
local pidl:LPITEMIDLIST
local dwChildren:DWORD
local sortcb:TVSORTCB
local tvi:TV_ITEM

	invoke GetFolder, hItem
	.if (!eax)
		return -1
	.endif
	mov pShellFolder, eax

	mov dwChildren, 0

	.if (g_bBrowseFiles)
		mov ecx,SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN
	.else
		mov ecx,SHCONTF_FOLDERS or SHCONTF_INCLUDEHIDDEN
	.endif
	invoke vf(pShellFolder, IShellFolder, EnumObjects_), m_hWnd, ecx, addr pEnumIDList
	.if (eax != S_OK)
		invoke vf(pShellFolder, IShellFolder, Release)
		return -1
	.endif

;--------- now enumerate the objects of the folder and insert them
;--------- into the treeview

	.while (1)
		invoke vf(pEnumIDList, IEnumIDList, Next), 1, addr pidl, NULL
		.break .if (eax != S_OK)

		invoke InsertItem, hItem, pShellFolder, pidl, FALSE
		.if (eax)
			inc dwChildren
		.endif
	.endw

	invoke vf(pEnumIDList, IUnknown, Release)
	.if (g_bSort)
		mov eax, hItem
		mov sortcb.hParent, eax
		mov sortcb.lpfnCompare, SortChildrenCB
		mov eax, pShellFolder
		mov sortcb.lParam, eax
		invoke TreeView_SortChildrenCB( m_hWndTV, addr sortcb, 0)
	.endif

	invoke vf(pShellFolder, IShellFolder, Release)

exit:
	mov eax, dwChildren
	.if (eax)
		mov eax, 1
	.endif
	mov tvi.cChildren, eax
	mov eax, hItem
	mov tvi.hItem, eax
	mov tvi.mask_, TVIF_CHILDREN or TVIF_STATE
	mov tvi.state, TVIS_EXPANDEDONCE
	mov tvi.stateMask, TVIS_EXPANDEDONCE
	invoke TreeView_SetItem( m_hWndTV, addr tvi)
	return dwChildren
	align 4

InsertChildItems endp


else

InsertChildItems proc public hItem:HANDLE

local pShellFolder:LPSHELLFOLDER
local pEnumIDList:LPENUMIDLIST
local pidl:LPITEMIDLIST
local sfgaof:SFGAOF
local dwChildren:DWORD
local strret:STRRET
local sortcb:TVSORTCB
local tvi:TV_INSERTSTRUCT
local szText[MAX_PATH]:byte

	invoke GetFolder, hItem
	.if (!eax)
		return -1
	.endif
	mov pShellFolder, eax

	mov eax, hItem
	mov tvi.hParent, eax
	mov tvi.hInsertAfter, TVI_LAST

	mov dwChildren, 0

	.if (g_bBrowseFiles)
		mov ecx,SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN
	.else	
		mov ecx,SHCONTF_FOLDERS or SHCONTF_INCLUDEHIDDEN
	.endif
	invoke vf(pShellFolder, IShellFolder, EnumObjects), m_hWnd, ecx, addr pEnumIDList
	.if (eax != S_OK)
		invoke vf(pShellFolder, IShellFolder, Release)
		return -1
	.endif

;--------- now enumerate the objects of the folder and insert them
;--------- into the treeview

	.while (1)
		invoke vf(pEnumIDList, IEnumIDList, Next), 1, addr pidl, NULL
		.break .if (eax != S_OK)
		mov strret.uType, STRRET_CSTR
		invoke vf(pShellFolder, IShellFolder, GetDisplayNameOf), pidl,\
				SHGDN_INFOLDER, addr strret
		.if (eax == S_OK)
			mov tvi.item.imask, TVIF_TEXT or TVIF_PARAM or TVIF_CHILDREN
			mov eax, pidl
			mov tvi.item.lParam, eax
			mov sfgaof, SFGAO_HASSUBFOLDER
			.if (g_bBrowseFiles)
				or sfgaof, SFGAO_BROWSABLE or SFGAO_FOLDER
			.endif
if ?TESTMODE
			or sfgaof, SFGAO_COMPRESSED or SFGAO_REMOVABLE
endif
			invoke vf(pShellFolder, IShellFolder, GetAttributesOf), 1, addr pidl,\
				addr sfgaof
			.if ((eax == S_OK) && (sfgaof & SFGAO_HASSUBFOLDER))
				mov tvi.item.cChildren, 1
			.else
				.if (g_bBrowseFiles)
					mov tvi.item.cChildren, I_CHILDRENCALLBACK
				.else
					mov tvi.item.cChildren, 0
				.endif
			.endif
if ?IMAGELIST
			or tvi.item.imask, TVIF_IMAGE or TVIF_SELECTEDIMAGE
			mov tvi.item.iSelectedImage, I_IMAGECALLBACK
			mov tvi.item.iImage, I_IMAGECALLBACK
endif
			.if ((!g_bBrowseFiles) || (sfgaof & (SFGAO_BROWSABLE or SFGAO_FOLDER)))
				invoke StrRet2String, pidl, addr strret, addr szText, sizeof szText
if ?TESTMODE
				sub esp, 256
				mov edx,esp
				invoke wsprintf, edx, CStr("%s,%X"), eax, sfgaof
				mov eax,esp
endif
				mov tvi.item.pszText, eax
				invoke TreeView_InsertItem( m_hWndTV, addr tvi)
if ?TESTMODE
				add esp,256
endif
				inc dwChildren
			.endif
		.endif
	.endw

	invoke vf(pEnumIDList, IUnknown, Release)
	.if (g_bSort)
		mov eax, tvi.hParent
		mov sortcb.hParent, eax
		mov sortcb.lpfnCompare, SortChildrenCB
		mov eax, pShellFolder
		mov sortcb.lParam, eax
		invoke TreeView_SortChildrenCB( m_hWndTV, addr sortcb, 0)
	.endif

	invoke vf(pShellFolder, IShellFolder, Release)

exit:
	mov eax, dwChildren
	.if (eax)
		mov eax, 1
	.endif
	mov tvi.item.cChildren, eax
	mov eax, hItem
	mov tvi.item.hItem, eax
	mov tvi.item.imask, TVIF_CHILDREN or TVIF_STATE
	mov tvi.item.state, TVIS_EXPANDEDONCE
	mov tvi.item.stateMask, TVIS_EXPANDEDONCE
	invoke TreeView_SetItem( m_hWndTV, addr tvi.item)
	return dwChildren
	align 4

InsertChildItems endp

endif


GetTextFromCLSID proc pGUID:ptr GUID, pStr:LPSTR, dwSize:dword

local	szStr[128]:byte
local	wszStr[40]:word
local	szGUID[40]:byte
local	hKey:HANDLE
local	dwType:dword

		mov ecx, pStr
		mov byte ptr [ecx],0
		invoke StringFromGUID2,pGUID,addr wszStr,40
		invoke WideCharToMultiByte,CP_ACP,0,addr wszStr,40,addr szGUID, sizeof szGUID,0,0 
		invoke wsprintf,addr szStr,CStr("CLSID\%s"), addr szGUID
		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szStr,0,KEY_READ,addr hKey
		.if (eax == S_OK)
			invoke RegQueryValueEx,hKey,CStr(""),NULL,addr dwType,pStr,addr dwSize
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

GetTextFromCLSID endp


RefreshView proc

local pItemIDList:LPITEMIDLIST
local pidl:LPITEMIDLIST
local pPersistFolder:LPPERSISTFOLDER
local dwSize:DWORD
local tvi:TV_INSERTSTRUCT
local hItem:HANDLE
local clsid:CLSID
local szText[128]:byte

	invoke TreeView_GetSelection( m_hWndTV)
	.if (eax)
		invoke GetFullPidl@CShellBrowser, eax
		mov pidl, eax
	.else
		mov pidl, NULL
	.endif

;;	invoke ReleaseShellView

	invoke SetWindowRedraw( m_hWndTV, FALSE)

;----------------------------- remove selection before deleteall
	invoke TreeView_SelectItem( m_hWndTV, NULL)

	invoke TreeView_DeleteAllItems( m_hWndTV)

	mov tvi.hParent, 0
	mov tvi.hInsertAfter, TVI_LAST

	invoke vf(m_pShellFolder, IUnknown, QueryInterface), addr IID_IPersistFolder, addr pPersistFolder
	.if (eax == S_OK)
		mov szText, 0
		invoke vf(pPersistFolder, IPersistFolder, GetClassID), addr clsid
		invoke GetTextFromCLSID, addr clsid, addr szText, sizeof szText
		invoke vf(pPersistFolder, IUnknown, Release)
	.else
		invoke lstrcpy, addr szText, CStr("Desktop")
	.endif
	lea eax, szText
	mov tvi.item.pszText, eax
	mov tvi.item.lParam, 0
if ?IMAGELIST
	mov tvi.item.mask_, TVIF_TEXT or TVIF_PARAM or TVIF_IMAGE or TVIF_SELECTEDIMAGE
	mov tvi.item.iSelectedImage, I_IMAGECALLBACK
	mov tvi.item.iImage, I_IMAGECALLBACK
else
	mov tvi.item.imask, TVIF_TEXT or TVIF_PARAM
endif
	invoke TreeView_InsertItem( m_hWndTV, addr tvi)
	mov hItem, eax

if 0
	invoke lstrcat, addr szText, addr szAppNameEx
	invoke SetWindowText, m_hWnd, addr szText
endif

	invoke InsertChildItems, hItem

	invoke TreeView_Expand( m_hWndTV, hItem, TVE_EXPAND )

	invoke SetWindowRedraw( m_hWndTV, TRUE)

	.if (pidl)
		invoke NavigateToPidl, pidl
		invoke vf(m_pMalloc, IMalloc, Free), pidl
	.else
		invoke TreeView_SelectItem( m_hWndTV, hItem)
	.endif

	ret
	align 4

RefreshView endp

;--- get a full pidl of a treeview item
;--- in lParam of TV_ITEM is a relative pidl (from EnumObjects) only

GetFullPidl@CShellBrowser proc public uses esi edi hItem:HANDLE

local dwESP:DWORD
local tvi:TV_ITEM

	mov eax, hItem
	mov tvi.hItem, eax
	
	mov dwESP, esp

	xor esi, esi
	xor edi, edi
	mov tvi.mask_, TVIF_PARAM
	.while (eax)
		invoke TreeView_GetItem( m_hWndTV, addr tvi)
		mov eax, tvi.lParam
		.break .if (!eax)
		movzx ecx, [eax].SHITEMID.cb
		add edi, ecx
		inc esi
		push eax
		invoke TreeView_GetParent( m_hWndTV, tvi.hItem)
		mov tvi.hItem, eax
	.endw
	.if (!edi)
		jmp exit
	.endif
	add edi, sizeof WORD
	invoke vf(m_pMalloc, IMalloc, Alloc), edi
	.if (!eax)
		jmp exit
	.endif
	mov edi, eax

	mov edx, esi
	.while (edx)
		pop esi
		movzx ecx,[esi].SHITEMID.cb
		rep movsb
		dec edx
	.endw
	mov [edi].SHITEMID.cb,0
exit:
	mov esp, dwESP
	ret
	align 4

GetFullPidl@CShellBrowser endp

;--- get an IShellFolder object of a treeview item

GetFolder proc public uses esi hItem:HANDLE

local pidl:LPITEMIDLIST
local pShellFolder:LPSHELLFOLDER

	invoke GetFullPidl@CShellBrowser, hItem
	mov pidl, eax
	.if (eax)
		mov pShellFolder, NULL
		invoke vf(m_pShellFolder, IShellFolder, BindToObject),
				pidl, NULL, addr IID_IShellFolder, addr pShellFolder
		invoke vf(m_pMalloc, IMalloc, Free), pidl
		mov eax, pShellFolder
	.else
		invoke vf(m_pShellFolder, IShellFolder, AddRef)
		mov eax, m_pShellFolder
	.endif
	ret
	align 4

GetFolder endp


GetTrueClientRect proc prect:ptr RECT

local rect:RECT

	invoke GetClientRect, m_hWnd, prect
	.if (m_bStatusBar)
		invoke GetWindowRect, m_hWndSB, addr rect
		mov ecx, rect.bottom
		sub ecx, rect.top
		mov eax, prect
		sub [eax].RECT.bottom, ecx
	.endif
	ret
	align 4

GetTrueClientRect endp

GetLeftPanelWidth proc prect:ptr RECT
	mov ecx, prect
	.if (m_dwSizeTV != -1)
		mov eax, m_dwSizeTV
	.else
		mov eax, [ecx].RECT.right
		sub eax, [ecx].RECT.left
		invoke MulDiv, eax, 3, 10
		sub eax, GRIPSIZE
	.endif
	ret
	align 4

GetLeftPanelWidth endp

;--- create a view object based on pidl stored in a treeview item

CreateViewObject proc uses esi hItem:HANDLE

local pShellFolder:LPSHELLFOLDER
local pShellView:LPSHELLVIEW
local hWndView:HWND
local dwWidth:DWORD
local rect:RECT
local dwUIActivateFlags:DWORD
local tvi:TV_ITEM
local folderset:FOLDERSETTINGS

	
	.if ((!m_bCreateView) || (hItem == NULL))
		invoke ReleaseShellView
		invoke LoadMenu, g_hInstance, IDR_MENU1
		invoke SetMenu, m_hWnd, eax
		invoke UpdateMenu
		invoke OnSize
		return TRUE
	.endif

	mov pShellView, NULL

;------------------------------- get a IShellFolder object
	invoke GetFolder, hItem
	mov pShellFolder, eax

	.if (eax)
		DebugOut "will call IShellFolder::CreateViewObject"
		invoke vf(pShellFolder, IShellFolder, CreateViewObject),
			m_hWnd, addr IID_IShellView, addr pShellView
		invoke vf(pShellFolder, IShellFolder, Release)
	.endif
	.if (pShellView)

		invoke GetTrueClientRect, addr rect
		invoke GetLeftPanelWidth, addr rect
		mov dwWidth, eax
		.if (!m_pShellView)
			invoke SetWindowPos, m_hWndTV, 0, 0, 0, dwWidth, rect.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
		.endif

		mov folderset.ViewMode, FVM_DETAILS
		mov folderset.fFlags, 0
		.if (m_pShellView)
			invoke vf(m_pShellView, IShellView, GetCurrentInfo), addr folderset
		.endif

		mov eax, dwWidth
		add eax, GRIPSIZE
		mov rect.left, eax
		DebugOut "will call IShellView::CreateViewWindow"
		mov hWndView, NULL
		invoke vf(pShellView, IShellView, CreateViewWindow),
			m_pShellView, addr folderset, __this, addr rect, addr hWndView
		.if (eax == S_OK)
			mov dwUIActivateFlags, SVUIA_ACTIVATE_NOFOCUS
			mov eax, m_hWndView
			.if (eax && (eax == m_hWndFocus))
				mov dwUIActivateFlags, SVUIA_ACTIVATE_FOCUS
			.endif
			invoke ReleaseShellView
			mov eax, hWndView
			mov m_hWndView, eax
			mov eax, pShellView
			mov m_pShellView, eax
			invoke vf(m_pShellView, IShellView, UIActivate), dwUIActivateFlags
			.if (dwUIActivateFlags == SVUIA_ACTIVATE_FOCUS)
				invoke SetFocus, m_hWndView
			.endif
;--------------------- some extensions dont create window with WS_TABSTOP
;--------------------- so add it now
			invoke GetWindowLong, m_hWndView, GWL_STYLE
			or eax, WS_TABSTOP
			invoke SetWindowLong, m_hWndView, GWL_STYLE, eax

			mov eax, TRUE
		.else
			invoke vf(pShellView, IShellView, Release)
			mov eax, FALSE
		.endif
	.else
		mov eax, FALSE
	.endif
	ret
	align 4

CreateViewObject endp

InsertNewChild proc pidlNew:LPITEMIDLIST

local hItem:HANDLE
local pidl:LPITEMIDLIST
local pidl2:LPITEMIDLIST
local pidl3:LPITEMIDLIST
local pShellFolder:LPSHELLFOLDER
local strret:STRRET
local sfgaof:SFGAOF
local tvi:TV_ITEM

		invoke Pidl_SkipLastItem, pidlNew, addr pidl
		mov pidl2, eax
		invoke FindPidl, eax, FALSE
		.if (eax)
			mov hItem, eax
			mov tvi.hItem, eax
			mov tvi.mask_, TVIF_STATE
			mov tvi.stateMask, TVIS_EXPANDEDONCE 
			invoke TreeView_GetItem( m_hWndTV, addr tvi)
;------------------------------------- if this item wasnt expanded yet, do nothing
			.if (!(tvi.state & TVIS_EXPANDEDONCE))
				jmp done
			.endif
			invoke GetFolder, hItem
			.if (eax)
				mov pShellFolder, eax

;--------------------------- this is very odd, but seems unavoidable:
;--------------------------- the new name has to parsed again to get the
;--------------------------- proper pidl

				mov strret.uType, STRRET_WSTR
				invoke vf(pShellFolder, IShellFolder, GetDisplayNameOf), pidl,
					SHGDN_INFOLDER, addr strret
				.if (eax == S_OK && (strret.uType == STRRET_WSTR))
					invoke vf(pShellFolder, IShellFolder, ParseDisplayName), m_hWnd,
						NULL, strret.pOleStr, NULL, addr pidl3, NULL
					invoke vf(m_pMalloc, IMalloc, Free), strret.pOleStr
					invoke vf(m_pMalloc, IMalloc, Free), pidl
					mov eax, pidl3
					mov pidl, eax
				.endif

				invoke InsertItem, hItem, pShellFolder, pidl, TRUE
				.if (eax)
					mov pidl, NULL
				.endif
				invoke vf(pShellFolder, IShellFolder, Release)
			.endif
		.endif
done:
		invoke vf(m_pMalloc, IMalloc, Free), pidl2
		.if (pidl)
			invoke vf(m_pMalloc, IMalloc, Free), pidl
		.endif
		ret

InsertNewChild endp

;--- WM_SHNOTIFY

OnSHNotify proc uses esi wParam:WPARAM, lParam:LPARAM

local hItem:HANDLE
local pidl:LPITEMIDLIST
local pidl2:LPITEMIDLIST
local pShellFolder:LPSHELLFOLDER
local tvi:TV_ITEM
local szText[MAX_PATH]:byte

	mov esi,wParam
ifdef _DEBUG
	mov ecx, lParam
	.if (ecx & SHCNE_RENAMEITEM)
		mov ecx, CStr("SHCNE_RENAMEITEM")
	.elseif (ecx & SHCNE_CREATE)
		mov ecx, CStr("SHCNE_CREATE")
	.elseif (ecx & SHCNE_DELETE)
		mov ecx, CStr("SHCNE_DELETE")
	.elseif (ecx & SHCNE_MKDIR)
		mov ecx, CStr("SHCNE_MKDIR")
	.elseif (ecx & SHCNE_RMDIR)
		mov ecx, CStr("SHCNE_RMDIR")
	.elseif (ecx & SHCNE_RENAMEFOLDER)
		mov ecx, CStr("SHCNE_RENAMEFOLDER")
	.elseif (ecx & SHCNE_EXTENDED_EVENT)
		mov ecx, CStr("SHCNE_EXTENDED_EVENT")
	.elseif (ecx & SHCNE_UPDATEITEM)
		mov ecx, CStr("SHCNE_UPDATEITEM")
	.else
		invoke wsprintf, addr szText, CStr("%X"), ecx
		lea ecx, szText
	.endif
	DebugOut "WM_SHNOTIFY(%s,[%X,%X],%X)", ecx, [esi].SHNOTIFYSTRUCT.dwItem1,\
		[esi].SHNOTIFYSTRUCT.dwItem2, lParam
endif
	.if (lParam & SHCNE_RMDIR)
ifdef _DEBUG
		invoke SHGetPathFromIDList, [esi].SHNOTIFYSTRUCT.dwItem1, addr szText
		.if (eax)
			DebugOut "%s", addr szText
		.endif
endif
		invoke FindPidl, [esi].SHNOTIFYSTRUCT.dwItem1, FALSE
		.if (eax)
			invoke TreeView_DeleteItem( m_hWndTV, eax)
		.endif
	.endif

	.if (lParam & SHCNE_MKDIR)
ifdef _DEBUG
		invoke SHGetPathFromIDList, [esi].SHNOTIFYSTRUCT.dwItem1, addr szText
		.if (eax)
			DebugOut "%s", addr szText
		.endif
endif
		invoke InsertNewChild, [esi].SHNOTIFYSTRUCT.dwItem1
	.endif

	.if (lParam & SHCNE_RENAMEFOLDER)
ifdef _DEBUG
		invoke SHGetPathFromIDList, [esi].SHNOTIFYSTRUCT.dwItem1, addr szText
		.if (eax)
			DebugOut "%s", addr szText
		.endif
		invoke SHGetPathFromIDList, [esi].SHNOTIFYSTRUCT.dwItem2, addr szText
		.if (eax)
			DebugOut "%s", addr szText
		.endif
endif
		mov hItem, NULL
		.if ([esi].SHNOTIFYSTRUCT.dwItem1)
			invoke FindPidl, [esi].SHNOTIFYSTRUCT.dwItem1, FALSE
			.if (eax)
				mov hItem, eax
				invoke TreeView_DeleteItem( m_hWndTV, eax)
			.endif
		.endif
		.if (hItem && [esi].SHNOTIFYSTRUCT.dwItem2)
			invoke InsertNewChild, [esi].SHNOTIFYSTRUCT.dwItem2
		.endif
	.endif
;------------------------------------ local folder now shared, refresh icon
	.if (lParam & SHCNE_NETSHARE)
if ?IMAGELIST
		invoke FindPidl, [esi].SHNOTIFYSTRUCT.dwItem1, FALSE
		.if (eax)
			mov tvi.hItem, eax
			mov tvi.mask_, TVIF_IMAGE or TVIF_SELECTEDIMAGE
			mov tvi.iImage, I_IMAGECALLBACK
			mov tvi.iSelectedImage, I_IMAGECALLBACK
			invoke TreeView_SetItem( m_hWndTV, addr tvi)
		.endif
endif
	.endif
ifdef _DEBUG
	.if (lParam & SHCNE_UPDATEITEM)
		invoke SHGetPathFromIDList, [esi].SHNOTIFYSTRUCT.dwItem1, addr szText
		.if (eax)
			DebugOut "%s", addr szText
		.endif
	.endif
endif
	ret
OnSHNotify endp


;--- WM_NOTIFY, TVN_GETDISPINFO


OnGetDispInfo proc uses esi ptvdi:ptr TV_DISPINFO

local pShellFolder:LPSHELLFOLDER
local tvi:TV_ITEM
if ?IMAGELIST
local pExtractIcon:LPEXTRACTICON
local dwFlags:DWORD
local iIndex:DWORD
local szIconFile[MAX_PATH]:BYTE
local hIconSmall:HICON
local hIconLarge:HICON
endif

	mov esi, ptvdi
	assume esi:ptr TV_DISPINFO

if 1
	.if ([esi].item.mask_ & TVIF_CHILDREN)
		invoke InsertChildItems, [esi].item.hItem
		.if (eax == -1)
			mov eax, 0
		.elseif (eax)
			mov eax, 1
		.endif
		mov [esi].item.cChildren, eax
		mov tvi.cChildren, eax
		mov eax, [esi].item.hItem
		mov tvi.hItem, eax
		mov tvi.mask_, TVIF_CHILDREN
		invoke TreeView_SetItem( m_hWndTV, addr tvi)
	.endif	
endif

if ?IMAGELIST
	.if ([esi].item.mask_ & (TVIF_IMAGE or TVIF_SELECTEDIMAGE))
		mov [esi].item.iImage, 0
		mov [esi].item.iSelectedImage, 1
		invoke TreeView_GetParent( m_hWndTV, [esi].item.hItem)
		.if (eax)
			invoke GetFolder, eax
			.if (eax)
				mov pShellFolder, eax
				invoke vf(pShellFolder, IShellFolder, GetUIObjectOf), m_hWnd,
						1, addr [esi].item.lParam,
						addr IID_IExtractIcon, NULL, addr pExtractIcon
				.if (eax == S_OK)
					.if ([esi].item.mask_ & TVIF_IMAGE)
						invoke vf(pExtractIcon, IExtractIcon, GetIconLocation), GIL_FORSHELL,
							addr szIconFile, sizeof szIconFile, addr iIndex, addr dwFlags
						.if (eax == S_OK)
							invoke vf(pExtractIcon, IExtractIcon, Extract), addr szIconFile,
								iIndex, addr hIconLarge, addr hIconSmall, 00100020h
							.if (eax == S_OK)
								invoke ImageList_AddIcon( g_himl, hIconSmall)
								mov [esi].item.iImage, eax
								invoke DestroyIcon, hIconLarge
								invoke DestroyIcon, hIconSmall
							.endif
						.endif
					.endif
					.if ([esi].item.mask_ & TVIF_SELECTEDIMAGE)
						invoke vf(pExtractIcon, IExtractIcon, GetIconLocation),
							GIL_FORSHELL or GIL_OPENICON,
							addr szIconFile, sizeof szIconFile, addr iIndex, addr dwFlags
						.if (eax == S_OK)
							invoke vf(pExtractIcon, IExtractIcon, Extract), addr szIconFile,\
								iIndex, addr hIconLarge, addr hIconSmall, 00100020h
							.if (eax == S_OK)
								invoke ImageList_AddIcon( g_himl, hIconSmall)
								mov [esi].item.iSelectedImage, eax
								invoke DestroyIcon, hIconLarge
								invoke DestroyIcon, hIconSmall
							.endif
						.endif
					.endif
					invoke vf(pExtractIcon, IExtractIcon, Release)
				.endif
				invoke vf(pShellFolder, IShellFolder, Release)
			.endif
		.else
			mov [esi].item.iImage, 2
			mov [esi].item.iSelectedImage, 2
		.endif

		mov eax, [esi].item.hItem
		mov tvi.hItem, eax
		mov ecx, [esi].item.mask_
		and ecx, TVIF_IMAGE or TVIF_SELECTEDIMAGE
		mov tvi.mask_, ecx
		mov eax,[esi].item.iImage
		mov tvi.iImage, eax
		mov eax,[esi].item.iSelectedImage
		mov tvi.iSelectedImage, eax
		invoke TreeView_SetItem( m_hWndTV, addr tvi)
	.endif

endif
	ret
	assume esi:nothing

OnGetDispInfo endp


editwndproc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

	.if (message == WM_GETDLGCODE)
		mov eax,DLGC_WANTALLKEYS
	.else
		invoke CallWindowProc, g_EditWndProc, hWnd, message, wParam, lParam
	.endif
	ret
	
editwndproc endp


;--- WM_NOTIFY, TVN_ENDLABELEDIT: rename a folder


OnEndLabelEdit proc uses esi pnmtvdi:ptr NMTVDISPINFO

local pShellFolder:LPSHELLFOLDER
local pidl:LPITEMIDLIST
local dwSize:DWORD
local bRC:BOOL
local tvi:TV_ITEM

	mov bRC, FALSE

	mov esi, pnmtvdi

;--- the parent folder is needed for SetNameOf function

	invoke TreeView_GetParent( m_hWndTV, [esi].NMTVDISPINFO.item.hItem)
	.if (eax)
		invoke GetFolder, eax
		.if (eax)
			mov pShellFolder, eax
			invoke lstrlen, [esi].NMTVDISPINFO.item.pszText
			add eax, 4
			and al,0FCh
			shl eax, 1
			mov dwSize, eax
			sub esp, eax
			shr eax, 1
			mov ecx, esp
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				[esi].NMTVDISPINFO.item.pszText, -1, ecx, eax
			mov ecx, esp
			invoke vf(pShellFolder, IShellFolder, SetNameOf), m_hWnd,
				[esi].NMTVDISPINFO.item.lParam, ecx, SHGDN_INFOLDER, addr pidl
			.if (eax == S_OK)
				invoke vf(m_pMalloc, IMalloc, Free), [esi].NMTVDISPINFO.item.lParam
				mov eax, pidl
				mov tvi.lParam, eax
				mov tvi.mask_, TVIF_PARAM
				mov eax, [esi].NMTVDISPINFO.item.hItem
				mov tvi.hItem, eax
				invoke TreeView_SetItem( m_hWndTV, addr tvi)
				mov bRC, TRUE
			.endif
			add esp, dwSize
			invoke vf(pShellFolder, IShellFolder, Release)
		.endif
	.endif
	return bRC

OnEndLabelEdit endp



ExecuteVerb proc hItem:HANDLE, lpVerb:LPSTR

local pShellFolder:LPSHELLFOLDER
local pContextMenu:LPCONTEXTMENU
local tvht:TV_HITTESTINFO
local hPopupMenu:HMENU
local sfgaof:SFGAOF
local tvi:TV_ITEM
local pt:POINT
local cmic:CMINVOKECOMMANDINFO
local szText[64]:byte

	DebugOut "EvecuteVerb enter"
	invoke TreeView_GetParent( m_hWndTV, hItem)
	.if (!eax)
		jmp exit
	.endif
	invoke GetFolder, eax
	.if (eax)
		DebugOut "EvecuteVerb: pShellFolder=%X", eax
		mov pShellFolder, eax
		mov tvi.mask_, TVIF_PARAM or TVIF_STATE or TVIF_CHILDREN 
		mov tvi.stateMask, TVIS_EXPANDED
		mov eax, hItem
		mov tvi.hItem, eax
		invoke TreeView_GetItem( m_hWndTV, addr tvi)
		DebugOut "EvecuteVerb: TreeView_GetItem=%X", eax
		.if (eax)
			mov eax, lpVerb
			mov cmic.lpVerb, eax
			invoke vf(pShellFolder, IShellFolder, GetUIObjectOf),
				m_hWnd, 1, addr tvi.lParam, addr IID_IContextMenu,
				NULL, addr pContextMenu
			DebugOut "EvecuteVerb: IShellFolder.GetUIObjectOf=%X, lpVerb=%X", eax, lpVerb
			.if (eax == S_OK)
				.if (lpVerb == NULL)
					invoke vf(pContextMenu, IContextMenu, QueryInterface),
						addr IID_IContextMenu2, addr m_pContextMenu2
					DebugOut "EvecuteVerb: IContextMeny.QueryInterface=%X", eax
					invoke CreatePopupMenu
					mov hPopupMenu, eax
					.if (tvi.state & TVIS_EXPANDED)
						mov ecx, CStr("Collapse")
					.else
						mov ecx, CStr("Expand")
					.endif
					.if (tvi.cChildren)
						mov edx, MF_ENABLED or MF_BYPOSITION or MF_STRING
					.else
						mov edx, MF_GRAYED or MF_BYPOSITION or MF_STRING
					.endif
					invoke InsertMenu, hPopupMenu, 0, edx, IDM_EXPAND, ecx
					invoke InsertMenu, hPopupMenu, 1, MF_BYPOSITION or MF_SEPARATOR, -1, NULL
					mov sfgaof, SFGAO_CANRENAME
					invoke vf(pShellFolder, IShellFolder, GetAttributesOf), 1,
							addr tvi.lParam, addr sfgaof
					mov ecx, CMF_NORMAL	or CMF_EXPLORE
					.if ((eax == S_OK) && (sfgaof & SFGAO_CANRENAME))
						or ecx, CMF_CANRENAME
					.endif
					invoke vf(pContextMenu, IContextMenu, QueryContextMenu),
						hPopupMenu, 2, IDM_FIRST, IDM_LAST, ecx
					DebugOut "EvecuteVerb: IContextMeny.QueryContextMenu=%X", eax
					.if (SUCCEEDED(eax))
						invoke SetMenuDefaultItem, hPopupMenu, 0, TRUE
						invoke GetMenuItemCount, hPopupMenu
						.if (eax == 2)
							invoke DeleteMenu, hPopupMenu, 1, MF_BYPOSITION
						.endif
						DebugOut "EvecuteVerb: calling TrackPopupMenu"
						invoke GetCursorPos, addr pt
						invoke TrackPopupMenu, hPopupMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
							pt.x, pt.y, 0, m_hWnd, NULL
						.if ((eax >= IDM_FIRST) && (eax <= IDM_LAST))
							sub eax, IDM_FIRST
							mov cmic.lpVerb, eax
							mov ecx, eax
							invoke vf(pContextMenu, IContextMenu, GetCommandString), ecx,
								GCS_VERBA, NULL, addr szText, sizeof szText
							.if (eax == S_OK)
								lea eax, szText
								mov lpVerb, eax
							.endif
						.elseif (eax == IDM_EXPAND)
							invoke TreeView_Expand( m_hWndTV, hItem, TVE_TOGGLE )
							mov lpVerb, NULL
							mov cmic.lpVerb, NULL
						.endif
					.endif
					invoke DestroyMenu, hPopupMenu
				.endif
				.if (cmic.lpVerb || lpVerb)
					mov cmic.cbSize, sizeof CMINVOKECOMMANDINFO
					mov cmic.fMask, 0
					mov eax, m_hWnd
					mov cmic.hwnd, eax
					mov cmic.lpParameters, NULL
					mov cmic.lpDirectory, NULL
					mov cmic.nShow, SW_SHOWNORMAL
					.if (lpVerb)
						invoke lstrcmpi, lpVerb, CStr("RENAME")
						.if (!eax)
							invoke TreeView_EditLabel( m_hWndTV, tvi.hItem)
							jmp done
						.endif
					.endif
					invoke vf(pContextMenu, IContextMenu, InvokeCommand), addr cmic
done:
				.endif
				invoke vf(pContextMenu, IUnknown, Release)
				.if (m_pContextMenu2)
					invoke vf(m_pContextMenu2, IUnknown, Release)
					mov m_pContextMenu2, NULL
				.endif
			.endif
		.endif
		invoke vf(pShellFolder, IUnknown, Release)
	.endif
exit:
	ret
ExecuteVerb endp


;--- WM_NOTIFY, NM_RCLICK: show context menu


OnRightClick proc pnmhdr:ptr NMHDR

local tvht:TV_HITTESTINFO

	DebugOut "OnRightClick enter"
	invoke GetCursorPos,addr tvht.pt
										; get the item below hit point
	invoke ScreenToClient, m_hWndTV, addr tvht.pt
	invoke TreeView_HitTest( m_hWndTV, addr tvht )
	.if (tvht.hItem)
		invoke ExecuteVerb, tvht.hItem, NULL
	.endif
	ret
	align 4

OnRightClick endp

;--- WM_NOTIFY, TVN_KEYDOWN

OnKeyDown proc vk:DWORD, flags:DWORD

local tvi:TV_ITEM

	invoke TreeView_GetSelection( m_hWndTV)
	mov tvi.hItem, eax
	mov tvi.mask_, TVIF_PARAM

	.if (vk == VK_DELETE)
		.if (tvi.hItem)
			invoke ExecuteVerb, tvi.hItem, CStr("delete")
		.else
			invoke MessageBeep, MB_OK
		.endif
	.endif
exit:
	ret

OnKeyDown endp

;--- WM_NOTIFY


OnNotify proc uses esi pNMHDR:ptr NMHDR

local pShellFolder:LPSHELLFOLDER
local pExtractIcon:LPEXTRACTICON
local sfgaof:SFGAOF
;local pt:POINT
local tvi:TV_ITEM

	xor eax, eax
	mov esi, pNMHDR
	.if ([esi].NMHDR.code == NM_SETFOCUS)

		mov eax, [esi].NMHDR.hwndFrom
		mov m_hWndFocus, eax
		.if (m_pShellView)
			invoke vf(m_pShellView, IShellView, UIActivate), SVUIA_ACTIVATE_NOFOCUS
		.endif

	.elseif (([esi].NMHDR.code == NM_RCLICK) && (m_bGetUIObjectOf))

		invoke OnRightClick, esi

	.elseif ([esi].NMHDR.code == TVN_KEYDOWN)

		movzx ecx, [esi].TV_KEYDOWN.wVKey
		invoke OnKeyDown, ecx, [esi].TV_KEYDOWN.flags

	.elseif ([esi].NMHDR.code == TVN_ITEMEXPANDING)

		xor eax, eax
		mov ecx, [esi].NM_TREEVIEW.itemNew.lParam
		.if (ecx && (!([esi].NM_TREEVIEW.itemNew.state & TVIS_EXPANDEDONCE)))
			invoke InsertChildItems, [esi].NM_TREEVIEW.itemNew.hItem
			.if (eax == -1)
				mov eax, 1
			.else
				xor eax, eax
			.endif
		.endif

	.elseif ([esi].NMHDR.code == TVN_DELETEITEM)

		mov ecx, [esi].NM_TREEVIEW.itemOld.lParam
		invoke vf(m_pMalloc, IMalloc, Free), ecx

	.elseif ([esi].NMHDR.code == TVN_GETDISPINFO)

		.if ([esi].TV_DISPINFO.item.mask_ & (TVIF_CHILDREN or TVIF_IMAGE or TVIF_SELECTEDIMAGE))
			invoke OnGetDispInfo, esi
		.endif

	.elseif ([esi].NMHDR.code == TVN_BEGINLABELEDIT)

		mov m_bLabelEdit, TRUE
		mov sfgaof, SFGAO_CANRENAME
		invoke GetAttributes, [esi].NMTVDISPINFO.item.hItem, addr sfgaof
		.if (sfgaof)
;------------------------------ set edit controls window proc
			invoke TreeView_GetEditControl( m_hWndTV)
			invoke SetWindowLong, eax, GWL_WNDPROC, editwndproc
			mov g_EditWndProc, eax
			xor eax, eax
		.else
			mov eax, TRUE
		.endif

	.elseif ([esi].NMHDR.code == TVN_ENDLABELEDIT)

		mov m_bLabelEdit, FALSE
		.if ([esi].NMTVDISPINFO.item.pszText)
			invoke OnEndLabelEdit, esi
			.if (!eax)
				invoke TreeView_EditLabel( m_hWndTV, [esi].NMTVDISPINFO.item.hItem)
				invoke TreeView_GetEditControl( m_hWndTV)
				mov ecx, eax
				invoke SetWindowText, ecx, [esi].NMTVDISPINFO.item.pszText
				xor eax, eax
			.endif
		.endif

	.elseif ([esi].NMHDR.code == TVN_SELCHANGING)

if 0		
;----------------------------------- if removeable, refresh folder content

		invoke IsRemoveable, [esi].NM_TREEVIEW.itemNew.hItem
		.if (eax)
			.while (1)
				invoke TreeView_GetChild( m_hWndTV, [esi].NM_TREEVIEW.itemNew.hItem)
				.break .if (!eax)
				invoke TreeView_DeleteItem( m_hWndTV, eax)
			.endw

			invoke InsertChildItems, [esi].NM_TREEVIEW.itemNew.hItem
			.if (eax == -1)
				mov eax, 1
				ret
			.elseif (eax)
				mov eax, 1
			.endif
			mov tvi.cChildren, eax
			mov tvi.imask, TVIF_CHILDREN
			mov eax, [esi].NM_TREEVIEW.itemNew.hItem
			mov tvi.hItem, eax
			invoke TreeView_SetItem( m_hWndTV, addr tvi)
		.endif
endif

		.if ([esi].NM_TREEVIEW.itemNew.hItem == NULL)
			xor eax, eax
			jmp exit
		.endif
		invoke CreateViewObject, [esi].NM_TREEVIEW.itemNew.hItem
;------------------------------ if successful, set main window title
		.if (eax)
			mov eax, [esi].NM_TREEVIEW.itemNew.hItem
			mov tvi.hItem, eax
			mov tvi.mask_, TVIF_TEXT
			sub esp, 128
			mov tvi.pszText, esp
			mov tvi.cchTextMax, 128
			invoke TreeView_GetItem( m_hWndTV, addr tvi)
			mov ecx, esp
			invoke lstrcat, ecx, addr szAppNameEx
			invoke SetWindowText, m_hWnd, esp
			add esp, 128
			xor eax, eax
		.else
			mov eax, 1
		.endif

	.endif
exit:
	ret
	align 4

OnNotify endp


aboutdlgproc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

	.if (message == WM_INITDIALOG)
		invoke GetDlgItem, hWnd, IDC_EDIT1
		mov ecx, eax
		invoke SetWindowText, ecx,
			CStr(<13,10,"ExplorerASM Version 1.2.0",13,10,"Public Domain (Japheth 2003-2007)",13,10,"http://github.com/Baron-von-Riedesel/ExplASM",13,10>)
		mov eax, 1
	.elseif (message == WM_CLOSE)
		invoke EndDialog, hWnd, 0
	.elseif (message == WM_COMMAND)
		movzx eax, word ptr wParam
		.if (eax == IDCANCEL)
			invoke PostMessage, hWnd, WM_CLOSE, 0, 0
		.endif
	.else
		xor eax, eax
	.endif
	ret
	align 4

aboutdlgproc endp

;--- WM_COMMAND


OnCommand proc wParam:WPARAM, lParam:LPARAM
		 
local hMenu:HMENU
local pidl:LPITEMIDLIST
local pOleCommandTarget:LPOLECOMMANDTARGET
local rect:RECT

		invoke GetMenu, m_hWnd
		mov hMenu, eax

		xor eax, eax
		movzx ecx, word ptr wParam+0
		.if (ecx == IDM_EXIT)

			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

		.elseif (ecx == IDOK)		;comes from labeledit

			DebugOut "IDOK, wParam=%X, lParam=%X", wParam, lParam

		.elseif (ecx == IDM_REFRESH)

			invoke RefreshView
if 0
			.if (m_pShellView)
				invoke vf(m_pShellView, IShellView, Refresh)
			.endif
endif
		.elseif (ecx == IDM_ABOUT)

			invoke DialogBoxParam, g_hInstance, IDD_DIALOG2, m_hWnd, aboutdlgproc, 0

		.elseif (ecx == IDM_STATUSLINE)

			xor m_bStatusBar, 1
			mov al, m_bStatusBar
			mov g_bStatusBar, al
			.if (m_bStatusBar)
				mov ecx, SW_SHOWNORMAL
			.else
				mov ecx, SW_HIDE
			.endif
			invoke ShowWindow, m_hWndSB, ecx
			invoke UpdateMenu
			invoke OnSize

		.elseif (ecx == IDM_CREATEVIEWOBJECT)

			xor m_bCreateView, 1
			mov al, m_bCreateView
			mov g_bCreateView, al
			invoke UpdateMenu
			invoke TreeView_GetSelection( m_hWndTV)
			.if (eax)
				invoke CreateViewObject, eax
			.endif

		.elseif (ecx == IDM_GETUIOBJECTOF)

			xor m_bGetUIObjectOf, 1
			mov al, m_bGetUIObjectOf
			mov g_bGetUIObjectOf, al
			invoke UpdateMenu

		.elseif (ecx == IDM_COMPMENUS)

			xor m_bCompMenus, 1
			mov al, m_bCompMenus
			mov g_bCompMenus, al
			.if (!m_bCompMenus)
				invoke LoadMenu, g_hInstance, IDR_MENU1
				invoke SetMenu, m_hWnd, eax
			.endif
			invoke UpdateMenu

		.elseif (ecx == IDM_DROPTARGET)

			xor m_bDropTarget, 1
			mov al, m_bDropTarget
			mov g_bDropTarget, al
			.if (al)
				lea ecx, m__IDropTarget
				invoke RegisterDragDrop, m_hWndTV, ecx
			.else
				invoke RevokeDragDrop, m_hWndTV
			.endif
			invoke UpdateMenu

		.elseif (ecx == IDM_TRACKSELECT)

			xor g_bTrackSelect, 1
			invoke GetWindowLong, m_hWndTV, GWL_STYLE
			.if (g_bTrackSelect)
				or eax, TVS_TRACKSELECT
			.else
				and eax, NOT (TVS_TRACKSELECT)
			.endif
			invoke SetWindowLong, m_hWndTV, GWL_STYLE, eax
			invoke UpdateMenu

		.elseif (ecx == IDM_DDERESPOND)

			xor m_bDontRespond, 1
			invoke UpdateMenu

		.elseif (ecx == IDM_BROWSEFILES)

			xor g_bBrowseFiles, 1
			invoke UpdateMenu
			invoke RefreshView

		.elseif (ecx == IDM_DELETE)

			invoke OnKeyDown, VK_DELETE, 0

		.else
			.if (m_pShellView)
				.if ((eax >= FCIDM_SHVIEWFIRST) && (eax <= FCIDM_SHVIEWLAST))
					invoke SendMessage, m_hWndView, WM_COMMAND, wParam, lParam
				.endif
			.endif
		.endif

		ret
		align 4

OnCommand endp

GRIPSIZE equ 4

;--- user resizes TreeView

DrawResizeLine proc uses esi hWnd:HWND , xPos:DWORD
        
local	hBrushOld:HBRUSH
local	hdc:HDC
local	rect:RECT
local	hBitmap:HBITMAP
local	pattern[8]:WORD
local	iToolbarHeight:DWORD
	
	.data
g_iOldResizeLine DWORD 0
	.code

	mov iToolbarHeight,0
	.if (!g_hHalftoneBrush)
		lea ecx, pattern
		xor esi, esi
		mov eax, 5555h
		.while (esi < 8)
			mov [ecx+esi*2], eax
			xor eax, 0FFFFh
			inc esi
		.endw
		invoke CreateBitmap, 8,8,1,1,addr pattern
		mov hBitmap,eax
		invoke CreatePatternBrush, hBitmap
		mov g_hHalftoneBrush, eax
		invoke DeleteObject, hBitmap
	.endif

	invoke GetClientRect, m_hWndTV, addr rect

	invoke GetDC, hWnd
	mov hdc, eax
	invoke SelectObject, hdc, g_hHalftoneBrush
	mov hBrushOld, eax
	mov ecx, iToolbarHeight
	add ecx, 2
	invoke PatBlt, hdc, xPos, ecx, GRIPSIZE, rect.bottom, PATINVERT

	invoke SelectObject, hdc, hBrushOld
	invoke ReleaseDC, hWnd, hdc
	mov eax, xPos
	mov g_iOldResizeLine, eax
	ret
	align 4

DrawResizeLine endp


;--- WM_SIZE


OnSize proc

local rect:RECT
local rectSB:RECT

	invoke GetClientRect, m_hWnd, addr rect
	.if (m_bStatusBar)
		invoke GetWindowRect, m_hWndSB, addr rectSB
		mov ecx, rectSB.bottom
		sub ecx, rectSB.top
		sub rect.bottom, ecx
		invoke SetWindowPos, m_hWndSB, NULL, 0, rect.bottom, rect.right, ecx, SWP_NOZORDER or SWP_NOACTIVATE
	.endif
		
	.if (!m_bCreateView)
		invoke SetWindowPos, m_hWndTV, NULL, 0, 0, rect.right, rect.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
	.else
		invoke GetLeftPanelWidth, addr rect
		mov m_dwSizeTV, eax
		mov ecx, eax
		push ecx
		invoke SetWindowPos, m_hWndTV, NULL, 0, 0, ecx, rect.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
		pop edx
		add edx, GRIPSIZE
		mov ecx, rect.right
		sub ecx, edx
		invoke SetWindowPos, m_hWndView, NULL, edx, 0, ecx, rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
	.endif
	ret
	align 4

OnSize endp

;----------------------------------------------------------
;--- WM_CREATE
;----------------------------------------------------------

OnCreate proc uses esi edi

local dummy:DWORD
;local sfs:SHELLFLAGSTATE
local dwWidth:DWORD
local dwXMax:DWORD
local dwYMax:DWORD
local rect:RECT

;-------------------------- got a window rect?
	.if (m_rect.left)
		invoke SetWindowPos, m_hWnd, NULL, m_rect.left, m_rect.top,
			m_rect.right, m_rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
	.else
;-------------------------- no. check sizes are suitable
		invoke GetSystemMetrics, SM_CXSCREEN
		mov dwXMax, eax
		invoke GetSystemMetrics, SM_CYSCREEN
		mov dwYMax, eax
		mov ecx, g_rect.right
		mov edx, g_rect.bottom
		.if ((!ecx) || (!edx) || (ecx > dwXMax) || (edx > dwYMax))
			invoke MulDiv, dwXMax, 2, 3
			mov g_rect.right, eax
			invoke MulDiv, dwYMax, 2, 3
			mov g_rect.bottom, eax
		.endif
		mov ecx, g_rect.left
		mov edx, g_rect.top
		.if (((!ecx) && (!edx)) || (ecx >= dwXMax) || (edx >= dwYMax))
			mov eax, dwXMax
			sub eax, g_rect.right
			shr eax, 1
			mov g_rect.left, eax

			mov eax, dwYMax
			sub eax, g_rect.bottom
			shr eax, 1
			mov g_rect.top, eax
		.endif
		invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top,
			g_rect.right, g_rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
	.endif

	mov ecx, WS_CHILD or WS_VISIBLE or WS_TABSTOP or TVS_HASBUTTONS or TVS_HASLINES or TVS_SHOWSELALWAYS or TVS_EDITLABELS or 4000h
	.if (g_bTrackSelect)
		or ecx, TVS_TRACKSELECT
	.endif

	invoke CreateWindowEx, WS_EX_CLIENTEDGE, CStr("SysTreeView32"), NULL, ecx,
			0,0,0,0, m_hWnd, IDC_TREE1, g_hInstance, NULL
	mov m_hWndTV, eax

if ?IMAGELIST
	.if (g_himl)
		invoke TreeView_SetImageList( m_hWndTV, g_himl, TVSIL_NORMAL)
	.endif
endif

	invoke CreateWindowEx, 0, CStr("msctls_statusbar32"), NULL,
			WS_CHILD or WS_VISIBLE or SBARS_SIZEGRIP or CCS_BOTTOM,
			0,0,0,0, m_hWnd, IDC_STATUSBAR, g_hInstance, NULL
	mov m_hWndSB, eax

	mov dwWidth,-1
	StatusBar_SetParts m_hWndSB, 1, addr dwWidth

	mov al, g_bCreateView
	mov m_bCreateView, al
	mov al,g_bGetUIObjectOf
	mov m_bGetUIObjectOf, al
	mov al, g_bCompMenus
	mov m_bCompMenus, al
	mov al, g_bStatusBar
	mov m_bStatusBar, al
	mov al, g_bDropTarget
	mov m_bDropTarget, al

	invoke UpdateMenu

	.if (!m_bStatusBar)
		invoke ShowWindow, m_hWndSB, SW_HIDE
	.endif

	.if (m_bDropTarget)
		lea ecx, m__IDropTarget
		invoke RegisterDragDrop, m_hWndTV, ecx
	.endif

	ret
	align 4

OnCreate endp


ParseDisplayName proc

local hinstShell:HINSTANCE
local pidl:LPITEMIDLIST
local sfgaof:SFGAOF
local wszPath[MAX_PATH]:word

	.if (!g_szPath)
		return 0
	.endif

	mov pidl, NULL
	invoke MultiByteToWideChar, CP_ACP, MB_PRECOMPOSED,
			addr g_szPath, -1, addr wszPath, MAX_PATH
if 0
	.if (g_lpfnSHParseDisplayName)
		invoke g_lpfnSHParseDisplayName, addr wszPath, NULL,
			addr pidl, NULL, NULL
	.endif
else
	invoke vf(m_pShellFolder, IShellFolder, ParseDisplayName), m_hWnd,
			NULL, addr wszPath, NULL, addr pidl, NULL
endif
	mov eax, pidl
	ret
ParseDisplayName endp

;----------------------------------------------------------
;--- dialog proc CShellBrowser
;----------------------------------------------------------

WndProc@CShellBrowser proc uses __this this_:ptr CShellBrowser, message:DWORD, wParam:WPARAM, lParam:LPARAM

local lo:DWORD
local hi:DWORD
local pidl:LPITEMIDLIST
local wp:WINDOWPLACEMENT
local rect:RECT
local rect2:RECT

	mov __this, this_

if 0
	.if (message == WM_SETCURSOR)
	.elseif (message == WM_NOTIFY)
	.elseif (message == WM_NCHITTEST)
	.elseif (message == WM_NCMOUSEMOVE)
	.elseif (message == WM_ENTERIDLE)
	.else
		DebugOut "DlgProc, msg=%X", message
	.endif
endif

	mov eax, message
	.if (eax == WM_CREATE)
		invoke OnCreate
		invoke SetFocus, m_hWndTV
		invoke SendMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0
		invoke ParseDisplayName
		.if (eax)
			mov pidl, eax
			invoke NavigateToPidl, pidl
			invoke vf(m_pMalloc, IMalloc, Free), pidl
		.endif
		.if (g_lpfnSHChangeNotifyRegister)
			invoke SHGetSpecialFolderLocation, m_hWnd, CSIDL_DESKTOP, addr pidl
			invoke g_lpfnSHChangeNotifyRegister, m_hWnd, 2,
				SHCNE_ALLEVENTS, WM_SHNOTIFY, 1, addr pidl
			mov m_hSHNotify, eax
			invoke vf(m_pMalloc, IMalloc, Free), pidl
		.endif

		xor eax, eax

	.elseif (eax == WM_CLOSE)

		.if (m_hSHNotify && g_lpfnSHChangeNotifyUnregister)
			invoke g_lpfnSHChangeNotifyUnregister, m_hSHNotify
		.endif
		invoke Deinit@IServiceProvider
		invoke TreeView_GetSelection( m_hWndTV)
		.if (eax)
			invoke GetFullPidl@CShellBrowser, eax
			mov pidl, eax
			mov g_szPath, 0
			invoke SHGetPathFromIDList, pidl, addr g_szPath
			invoke vf(m_pMalloc, IMalloc, Free), pidl
		.endif
		invoke ReleaseShellView
		invoke DestroyWindow, m_hWnd
		invoke Destroy@CShellBrowser, __this

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_SIZE)

		invoke OnSize

		mov wp.length_, sizeof WINDOWPLACEMENT
		invoke GetWindowPlacement, m_hWnd, addr wp
		invoke CopyRect, addr g_rect, addr wp.rcNormalPosition
		mov eax, g_rect.right
		sub eax, g_rect.left
		mov g_rect.right, eax
		mov eax, g_rect.bottom
		sub eax, g_rect.top
		mov g_rect.bottom, eax
		xor eax, eax

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_SETCURSOR)

		movzx eax, word ptr lParam+0
		mov ecx, wParam
		.if ((eax == HTCLIENT) && (ecx == m_hWnd))
			invoke SetCursor, g_hCursor
			mov eax, 1
		.else
			jmp default
		.endif

	.elseif (eax == WM_GETISHELLBROWSER)

		DebugOut "undocumented message WM_GETISHELLBROWSER received"
		mov eax, __this

	.elseif (eax == WM_SHNOTIFY)

		invoke OnSHNotify, wParam, lParam

	.elseif ((eax == WM_INITMENUPOPUP) || (eax == WM_ENTERMENULOOP) || (eax == WM_EXITMENULOOP))

		xor eax, eax
		.if (m_pContextMenu2)
			.if (message == WM_INITMENUPOPUP)
				invoke vf(m_pContextMenu2, IContextMenu2, HandleMenuMsg), message, wParam, lParam
				mov eax, 1
			.endif
		.elseif (m_hWndView)
			DebugOut "WM_INITMENUPOPUP/ENTERMENULOOP/EXITMENULOOP enter(%X)", message
			invoke SendMessage, m_hWndView, message, wParam, lParam
			DebugOut "WM_INITMENUPOPUP/ENTERMENULOOP/EXITMENULOOP exit"
		.endif

	.elseif (eax == WM_MENUSELECT)

		xor eax, eax
		movzx ecx, word ptr wParam
		.if ((ecx >= FCIDM_BROWSERFIRST) && (ecx <= FCIDM_BROWSERLAST))
			;
		.elseif (m_hWndView)
			DebugOut "WM_MENUSELECT enter(%X)", ecx
			invoke SendMessage, m_hWndView, message, wParam, lParam
			DebugOut "WM_MENUSELECT exit"
		.endif

	.elseif ((eax == WM_MEASUREITEM) || (eax == WM_DRAWITEM))

		DebugOut "WM_MEASUREITEM/WM_DRAWITEM"
		xor eax, eax
		.if (m_pContextMenu2)
			invoke vf(m_pContextMenu2, IContextMenu2, HandleMenuMsg), message, wParam, lParam
			mov eax, 1
		.elseif (m_hWndView)
			invoke SendMessage, m_hWndView, message, wParam, lParam
		.endif

	.elseif (eax == WM_LBUTTONDOWN)

		xor eax, eax
		.if (m_hWndView)
			invoke SetCapture, m_hWnd

			invoke GetWindowRect, m_hWndTV, addr rect
			invoke GetWindowRect, m_hWndView, addr rect2
			invoke UnionRect, addr rect, addr rect, addr rect2
			invoke ClipCursor, addr rect
			mov g_iOldResizeLine, -1
			movzx eax, word ptr lParam+0
			invoke DrawResizeLine, m_hWnd, eax
		.endif

	.elseif (eax == WM_LBUTTONUP)

		invoke GetCapture
		.if (eax == m_hWnd)
			invoke ClipCursor, NULL
			invoke ReleaseCapture
			invoke DrawResizeLine, m_hWnd, g_iOldResizeLine
			movzx eax, word ptr lParam
			mov m_dwSizeTV, eax
			invoke OnSize
		.endif

	.elseif (eax == WM_MOUSEMOVE)

		invoke GetCapture
		.if (eax == m_hWnd)
			invoke DrawResizeLine, m_hWnd, g_iOldResizeLine
			movzx eax, word ptr lParam+0
			invoke DrawResizeLine, m_hWnd, eax
		.endif

	.elseif (eax == WM_SETFOCUS)

		.if (m_hWndFocus)
			invoke SetFocus, m_hWndFocus
		.endif

	.elseif (eax == WM_DDE_INITIATE)

		DebugOut "WM_DDE_INITIATE, wParam=%X, lParam=%X, m_hWndView=%X", wParam, lParam, m_hWndView
		.if (!m_bDontRespond)
;------------------------------- respond to windows from this thread only
			movzx ecx, word ptr lParam+0
			movzx edx, word ptr lParam+2
			.if ((ecx == g_aApplication) && (edx == g_aTopic))
				invoke GetWindowThreadProcessId, wParam, NULL
				push eax
				invoke GetCurrentThreadId
				pop ecx
				.if (ecx == eax)
					invoke SendMessage, wParam, WM_DDE_ACK, m_hWnd, lParam
				.endif
			.endif
		.endif

	.elseif (eax == WM_DDE_EXECUTE)

		DebugOut "WM_DDE_EXECUTE, wParam=%X, lParam=%X", wParam, lParam
		invoke UnpackDDElParam, WM_DDE_EXECUTE, lParam, addr lo, addr hi
		invoke GlobalLock, hi
;------------------------------- now a fantastic amount of work is to be done
		invoke OnDDEExecute, eax
		.if (eax)
			mov lo, 8000h
		.else
			mov lo, 0
		.endif
		invoke GlobalUnlock, hi
		invoke ReuseDDElParam, lParam, WM_DDE_EXECUTE, WM_DDE_ACK, lo, hi
		invoke PostMessage, wParam, WM_DDE_ACK, m_hWnd, eax

	.elseif (eax == WM_DDE_TERMINATE)

		DebugOut "WM_DDE_TERMINATE, wParam=%X", wParam
		invoke PostMessage, wParam, WM_DDE_TERMINATE, m_hWnd, lParam

	.else
default:
		invoke DefWindowProc, m_hWnd, message, wParam, lParam
	.endif
exit:
	ret
	align 4

WndProc@CShellBrowser endp


;--- true dialog proc


WndProc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

	.if (message == WM_NCCREATE)
		mov ecx, lParam
		mov ecx,[ecx].CREATESTRUCT.lpCreateParams
		mov eax, hWnd
		mov [ecx].CShellBrowser.hWnd, eax
		push ecx
		invoke SetWindowLong, hWnd, GWL_USERDATA, ecx
		pop eax
	.else
		invoke GetWindowLong, hWnd, GWL_USERDATA
	.endif

	.if (eax)
		invoke WndProc@CShellBrowser, eax, message, wParam, lParam
	.endif
	ret
	align 4

WndProc endp

;----------------------------------------------------------------
;--- IShellBrowser methods
;----------------------------------------------------------------

GetWindow_ proc uses __this this_:ptr CShellBrowser, phwnd:ptr HWND

	mov __this, this_
	DebugOut "IShellBrowser::GetWindow(%X)",phwnd
	mov eax, m_hWnd
	mov ecx, phwnd
	mov [ecx],eax
	return S_OK
	align 4

GetWindow_ endp


ContextSensitiveHelp proc uses __this this_:ptr CShellBrowser, fEnterMode:BOOL

	mov __this, this_
	DebugOut "IShellBrowser::ContextSensitiveHelp"
	return E_UNEXPECTED
	align 4

ContextSensitiveHelp endp


InsertMenusSB proc uses __this esi this_:ptr CShellBrowser, hmenuShared:HMENU, lpMenuWidths:ptr OLEMENUGROUPWIDTHS

local hMenu:HMENU
local hSubMenu:HMENU
local mii:MENUITEMINFO

	mov __this, this_

	DebugOut "IShellBrowser::InsertMenusSB(%X)", hmenuShared

	.if (!m_bCompMenus)
		return S_OK
	.endif

	invoke LoadMenu, g_hInstance, IDR_MENU1
	mov hMenu, eax

	invoke RtlZeroMemory, addr mii, sizeof MENUITEMINFO
	mov mii.cbSize, sizeof MENUITEMINFO
	mov mii.fMask, MIIM_SUBMENU or MIIM_STATE or MIIM_TYPE or MIIM_ID
	mov mii.fType, MFT_STRING
	mov mii.fState, MFS_ENABLED

	mov esi, lpMenuWidths
	invoke GetSubMenu, hMenu, 0
	.if (eax)
		mov mii.hSubMenu, eax
		mov mii.wID, FCIDM_MENU_FILE
		mov mii.dwTypeData, CStr("&File")
		invoke InsertMenuItem, hmenuShared, -1, TRUE, addr mii
		inc dword ptr [esi+0*sizeof DWORD]
	.endif
	invoke CreatePopupMenu
	.if (eax)
		mov mii.hSubMenu, eax
		mov mii.wID, FCIDM_MENU_EDIT
		mov mii.dwTypeData, CStr("&Edit")
		invoke InsertMenuItem, hmenuShared, -1, TRUE, addr mii
		inc dword ptr [esi+0*sizeof DWORD]
	.endif

	invoke GetSubMenu, hMenu, 1
	.if (eax)
		mov hSubMenu, eax
		mov mii.fMask, MIIM_TYPE or MIIM_ID
		mov mii.wID, FCIDM_MENU_VIEW_SEP_OPTIONS
		mov mii.fType, MFT_SEPARATOR
		invoke InsertMenuItem, hSubMenu, 1, TRUE, addr mii

		mov eax, hSubMenu
		mov mii.hSubMenu, eax
		mov mii.fMask, MIIM_SUBMENU or MIIM_STATE or MIIM_TYPE or MIIM_ID
		mov mii.fType, MFT_STRING
		mov mii.wID, FCIDM_MENU_VIEW
		mov mii.dwTypeData, CStr("&View")
		invoke InsertMenuItem, hmenuShared, -1, TRUE, addr mii
		inc dword ptr [esi+2*sizeof DWORD]
	.endif

	invoke GetSubMenu, hMenu, 2
	.if (eax)
		mov mii.hSubMenu, eax
		mov mii.wID, FCIDM_MENU_TOOLS
		mov mii.dwTypeData, CStr("&Tools")
		invoke InsertMenuItem, hmenuShared, -1, TRUE, addr mii
		inc dword ptr [esi+2*sizeof DWORD]
	.endif

	invoke GetSubMenu, hMenu, 3
	.if (eax)
		mov mii.hSubMenu, eax
		mov mii.wID, FCIDM_MENU_HELP
		mov mii.dwTypeData, CStr("&Help")
		invoke InsertMenuItem, hmenuShared, -1, TRUE, addr mii
		inc dword ptr [esi+4*sizeof DWORD]
	.endif
	return S_OK
	align 4

InsertMenusSB endp


SetMenuSB proc uses __this this_:ptr CShellBrowser, hmenuShared:HMENU, dwReserved:DWORD, hwndActiveObject:HWND

	mov __this, this_
	DebugOut "IShellBrowser::SetMenuSB(%X,%X,%X)", hmenuShared, dwReserved, hwndActiveObject
	.if (!m_bCompMenus)
		return S_OK
	.endif
if 1
	.if (hmenuShared)
		invoke SetMenu, m_hWnd, hmenuShared
	.else
		invoke LoadMenu, g_hInstance, IDR_MENU1
		invoke SetMenu, m_hWnd, eax
	.endif
	invoke UpdateMenu
else
	invoke SetMenu, m_hWnd, hmenuShared
	.if (hmenuShared)
		invoke UpdateMenu
	.endif
endif
	return S_OK
	align 4

SetMenuSB endp


RemoveMenusSB proc uses __this this_:ptr CShellBrowser, hmenuShared:HMENU

	mov __this, this_
	DebugOut "IShellBrowser::RemoveMenusSB(%X)", hmenuShared
if 0
	invoke DeleteMenu, hmenuShared, FCIDM_MENU_FILE, MF_BYCOMMAND
	invoke DeleteMenu, hmenuShared, FCIDM_MENU_EDIT, MF_BYCOMMAND
	invoke DeleteMenu, hmenuShared, FCIDM_MENU_TOOLS, MF_BYCOMMAND
endif
	return S_OK
	align 4

RemoveMenusSB endp


SetStatusTextSB proc uses __this this_:ptr CShellBrowser, lpszStatusText:LPCOLESTR 

local szText[256]:byte

	mov __this, this_
	DebugOut "IShellBrowser::SetStatusTextSB(%X)", lpszStatusText
	.if ( lpszStatusText )
		invoke WideCharToMultiByte, CP_ACP, 0, lpszStatusText, -1, addr szText, sizeof szText, NULL, NULL
		lea eax, szText
	.else
		mov eax, CStr("")
	.endif
	StatusBar_SetText m_hWndSB, 0, eax
	return S_OK
	align 4

SetStatusTextSB endp


EnableModelessSB proc uses __this this_:ptr CShellBrowser, fEnable:BOOL

	mov __this, this_
	DebugOut "IShellBrowser::EnableModelessSB(%X)", fEnable
	return S_OK
	align 4

EnableModelessSB endp


TranslateAcceleratorSB proc uses __this this_:ptr CShellBrowser, lpmsg:ptr MSG, wID:WORD

	mov __this, this_
;;	DebugOut "IShellBrowser::TranslateAcceleratorSB"
	invoke TranslateAccelerator, m_hWnd, g_hAccel, lpmsg
	.if (eax)
		mov eax, S_OK
	.else
		mov eax, S_FALSE
	.endif
	ret
	align 4

TranslateAcceleratorSB endp

;--- function should allow a view to navigate to an item
;--- (inside or outside of the current folder). But is
;--- not used currently

BrowseObject proc uses __this this_:ptr CShellBrowser, pidl:LPITEMIDLIST, wFlags:DWORD

	mov __this, this_
	DebugOut "IShellBrowser::BrowseObject(%X, %X)", pidl, wFlags
	.if (wFlags & SBSP_ABSOLUTE)
		invoke NavigateToPidl, pidl
		.if (eax)
			return S_OK
		.endif
	.endif
	return E_FAIL
	align 4

BrowseObject endp


GetViewStateStream proc uses __this this_:ptr CShellBrowser, grfMode:DWORD, ppStrm:ptr LPSTREAM

	mov __this, this_
	DebugOut "IShellBrowser::GetViewStateStream(%X, %X)", grfMode, ppStrm
	.if (!m_pStream)
		invoke CreateStreamOnHGlobal, NULL, TRUE, addr m_pStream
	.endif
	mov ecx, ppStrm
	mov eax, m_pStream
	mov dword ptr [ecx], eax
	.if (eax)
		invoke vf(m_pStream, IUnknown, AddRef)
		return S_OK
	.else
		return E_OUTOFMEMORY
	.endif
	align 4

GetViewStateStream endp

;--- get control hwnds
;--- FCW_STATUS(1), FCW_TOOLBAR(2), FCW_TREE (3)

GetControlWindow proc uses __this this_:ptr CShellBrowser, id:DWORD, lphwnd:ptr HWND

	mov __this, this_
	DebugOut "IShellBrowser::GetControlWindow(%X, %X)", id, lphwnd
	.if (id == FCW_TREE)
		mov edx, m_hWndTV
		mov eax, S_OK
	.elseif (id == FCW_STATUS)
		mov edx, m_hWndSB
		mov eax, S_OK
	.else
		xor edx, edx
		mov eax, E_FAIL
	.endif
	mov ecx, lphwnd
	mov [ecx], edx
	ret
	align 4

GetControlWindow endp

;SB_SETTEXTA equ <SB_SETTEXT>
;SB_SETICON equ <WM_USER+15>

SendControlMsg proc uses __this this_:ptr CShellBrowser, id:DWORD, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM, pres:ptr DWORD

	mov __this, this_
ifdef _DEBUG
	.if (id == FCW_STATUS)
		.if ((uMsg == SB_SETTEXTW) || (uMsg == SB_SETTEXTA))
			sub esp,64*2
			mov edx, esp
			.if (uMsg == SB_SETTEXTW)
				invoke WideCharToMultiByte, CP_ACP, 0, lParam, -1, edx, 64, 0, 0
			.else
				invoke lstrcpyn, edx, lParam, 64
			.endif
			mov edx, esp
			DebugOut "IShellBrowser::SendControlMsg(FCW_STATUS,SB_SETTEXT,%u,'%s')", wParam, edx
			add esp,64*2
		.elseif (uMsg == SB_SETICON)
			DebugOut "IShellBrowser::SendControlMsg(FCW_STATUS,SB_SETICON,%u,%X)", wParam, lParam
		.elseif (uMsg == SB_SETPARTS)
			DebugOut "IShellBrowser::SendControlMsg(FCW_STATUS,SB_SETPARTS,%u)", wParam
		.else
			DebugOut "IShellBrowser::SendControlMsg(%X,%X,%X,%X)", id, uMsg, wParam, lParam
		.endif
	.else
		DebugOut "IShellBrowser::SendControlMsg(%X,%X,%X,%X)", id, uMsg, wParam, lParam
	.endif
endif
	.if (id == FCW_TREE)
		mov eax, m_hWndTV
	.elseif (id == FCW_STATUS)
		mov eax, m_hWndSB
	.else
		xor eax, eax
	.endif

	.if (eax)
		invoke SendMessage, eax, uMsg, wParam, lParam
		mov ecx, pres
		.if (ecx)
			mov [ecx], eax
		.endif
		mov eax, S_OK
	.else
		mov ecx, pres
		.if (ecx)
			mov dword ptr [ecx],0
		.endif
		mov eax, E_FAIL
	.endif
	ret
	align 4

SendControlMsg endp


QueryActiveShellView proc uses __this this_:ptr CShellBrowser, ppshv:ptr LPSHELLVIEW

	mov __this, this_
	DebugOut "IShellBrowser::QueryActiveShellView"
	mov ecx, ppshv
	mov eax, m_pShellView
	mov dword ptr [ecx], eax
	.if (eax)
		invoke vf(eax, IShellView, AddRef)
		mov eax, S_OK
	.else
		mov eax, E_FAIL
	.endif
	ret
	align 4

QueryActiveShellView endp

;--- the view window has got the focus

OnViewWindowActive proc uses __this this_:ptr CShellBrowser, pshv:LPSHELLVIEW

	mov __this, this_
	DebugOut "IShellBrowser::OnViewWindowActive(%X)", pshv
	mov eax, m_hWndView
	mov m_hWndFocus, eax
	return S_OK
	align 4

OnViewWindowActive endp


SetToolbarItems proc uses __this this_:ptr CShellBrowser, lpButtons:ptr TBBUTTON, nButtons:DWORD, uFlags:DWORD

	mov __this, this_
	DebugOut "IShellBrowser::SetToolbarItems"
	return S_OK
	align 4

SetToolbarItems endp


	end
