
;--- IDropTarget installs left explorer panel as a drop target

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

?AGGREGATION	= 0	;no aggregation
?EVENTSUPPORT	= 0	;no event support

	include olecntrl.inc
	include debugout.inc
	include rsrc.inc
	.list
	.cref

INSIDE_CShellBrowser equ 1

LPDROPTARGETHELPER typedef ptr IDropTargetHelper
LPSHELLLINK typedef ptr IShellLink

	include CShellBrowser.inc

if ?DROPTARGET

	includelib shlwapi.lib

	.const

CDropTargetVtbl label dword
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
    dd DragEnter_, DragOver_, DragLeave_, Drop_

;CLSID_ShellLink GUID sCLSID_ShellLink
;IID_IShellLink sIID_IShellLinkA

	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER hWnd, hWndTV, bRButton
	MEMBER pMalloc

	@MakeStubsEx CShellBrowser, IDropTarget, QueryInterface, AddRef, Release
	@MakeStubs CShellBrowser, IDropTarget, DragEnter, DragOver, DragLeave, Drop

DragEnter proc uses __this this_:ptr CShellBrowser, pDataObject:LPDATAOBJECT, grfKeyState:DWORD, pt:POINTL, pdwEffect:ptr DWORD

local fe:FORMATETC

	mov __this,this_

;----------------------------- check if data is acceptable

	mov fe.cfFormat, CF_HDROP
	mov fe.ptd, NULL
	mov fe.dwAspect, DVASPECT_CONTENT
	mov fe.lindex, -1
	mov fe.tymed, TYMED_HGLOBAL
	invoke	vf(pDataObject, IDataObject, QueryGetData), addr fe
	mov ecx, pdwEffect
	DebugOut "IShellBrowser::DragEnter(%X), IDataObject::QueryGetData=%X %X", grfKeyState, eax, dword ptr [ecx]
	.IF (eax != S_OK)
		mov dword ptr [ecx], DROPEFFECT_NONE
if 1
	.ELSE
		mov edx, grfKeyState
		and edx, MK_CONTROL or MK_SHIFT
		mov eax, [ecx]
		.if (edx == (MK_CONTROL or MK_SHIFT))
			and eax, DROPEFFECT_LINK
		.elseif (edx == MK_CONTROL)
			and eax,DROPEFFECT_COPY
		.else
			and eax,DROPEFFECT_MOVE
		.endif
		.if (!eax)
			mov eax, DROPEFFECT_COPY
		.endif
		and DWORD PTR [ecx], eax
endif
	.endif
	.if (grfKeyState & MK_RBUTTON)
		mov m_bRButton, TRUE
	.else
		mov m_bRButton, FALSE
	.endif
if ?DRAGDROPHELPER
	.if (g_pDropTargetHelper)
		mov ecx, pdwEffect
		invoke vf(g_pDropTargetHelper, IDropTargetHelper, DragEnter), m_hWnd, pDataObject, addr pt, [ecx]
		DebugOut "IDropTargetHelper::DragEnter=%X", eax
	.endif
endif

	return S_OK
	align 4

DragEnter endp

DragOver proc uses __this this_:ptr CShellBrowser, grfKeyState:DWORD, pt:POINTL, pdwEffect:ptr DWORD

local tvht:TV_HITTESTINFO

	mov __this,this_
	DebugOut "IShellBrowser::DragOver(%X)", grfKeyState

	mov eax, pt.x
	mov tvht.pt.x, eax
	mov eax, pt.y
	mov tvht.pt.y, eax
	invoke ScreenToClient, m_hWndTV, addr tvht.pt
	invoke TreeView_HitTest( m_hWndTV, addr tvht)
	.if (eax)
		invoke TreeView_SelectDropTarget( m_hWndTV, eax)
		mov ecx, pdwEffect
		mov eax, [ecx]
		mov edx, grfKeyState
		and edx, MK_CONTROL or MK_SHIFT
		.if (edx == (MK_CONTROL or MK_SHIFT))
			and eax, DROPEFFECT_LINK
		.elseif (edx == MK_CONTROL)
			and eax,DROPEFFECT_COPY
		.else
			and eax,DROPEFFECT_MOVE
		.endif
		.if (!eax)
			mov eax, DROPEFFECT_COPY
		.endif
	.else
		mov eax, DROPEFFECT_NONE
	.endif
	mov ecx, pdwEffect
	and DWORD PTR [ecx], eax

if ?DRAGDROPHELPER
	.if (g_pDropTargetHelper)
		invoke vf(g_pDropTargetHelper, IDropTargetHelper, DragOver), addr pt, [ecx]
ifdef _DEBUG
		mov ecx, pdwEffect
endif
		DebugOut "IDropTargetHelper::DragOver=%X, %X", eax, dword ptr [ecx]
	.endif
endif

	return S_OK
	align 4

DragOver endp

DragLeave proc uses __this this_:ptr CShellBrowser

	mov __this,this_
	DebugOut "IShellBrowser::DragLeave"

	invoke TreeView_SelectDropTarget( m_hWndTV, NULL)
if ?DRAGDROPHELPER
	.if (g_pDropTargetHelper)
		invoke vf(g_pDropTargetHelper, IDropTargetHelper, DragLeave)
		DebugOut "IDropTargetHelper::DragLeave=%X", eax
	.endif
endif
	return S_OK
	align 4

DragLeave endp

IDM_MOVE	equ FCIDM_BROWSERFIRST+0
IDM_COPY	equ FCIDM_BROWSERFIRST+1
IDM_CREATESHORTCUT	equ FCIDM_BROWSERFIRST+2
IDM_CANCEL	equ FCIDM_BROWSERFIRST+3

DisplayContextMenu proc uses esi grfKeyState:DWORD, dwEffect:DWORD

local dwDefault:DWORD
local pt:POINT

	mov dwDefault, 0
	invoke CreatePopupMenu
	mov esi, eax
	.if (dwEffect & DROPEFFECT_MOVE)
		invoke AppendMenu, esi, MF_STRING, IDM_MOVE, CStr("Move Here")
		.if (!(grfKeyState & MK_CONTROL))
			mov dwDefault, IDM_MOVE
		.endif
	.endif
	.if (dwEffect & DROPEFFECT_COPY)
		invoke AppendMenu, esi, MF_STRING, IDM_COPY, CStr("Copy Here")
		.if (grfKeyState & MK_CONTROL)
			mov dwDefault, IDM_COPY
		.endif
	.endif
	.if (dwEffect & DROPEFFECT_LINK)
		invoke AppendMenu, esi, MF_STRING, IDM_CREATESHORTCUT, CStr("Create Shortcut(s) Here")
		mov ecx, grfKeyState
		and ecx, MK_CONTROL or MK_SHIFT
		.if (ecx == MK_CONTROL or MK_SHIFT)
			mov dwDefault, IDM_CREATESHORTCUT
		.endif
	.endif
	invoke AppendMenu, esi, MF_SEPARATOR, -1, 0
	invoke AppendMenu, esi, MF_STRING, IDM_CANCEL, CStr("Cancel")
	.if (dwDefault)
		invoke SetMenuDefaultItem, esi, dwDefault, FALSE
	.endif
	invoke GetCursorPos, addr pt
	invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
		pt.x, pt.y, 0, m_hWnd, NULL
	.if (eax == IDM_COPY)
		mov eax, MK_CONTROL
	.elseif (eax == IDM_CREATESHORTCUT)
		mov eax, MK_CONTROL or MK_SHIFT
	.elseif (eax == IDM_MOVE)
		mov eax, MK_RBUTTON
	.else
		xor eax, eax
	.endif
	push eax
	invoke DestroyMenu, esi
	pop eax
	ret

DisplayContextMenu endp


CreateLink proc lpszPathObj:LPSTR, lpszPathLink:LPSTR , lpszDesc:LPSTR

local	hres:DWORD
local	pShellLink:LPSHELLLINK
local	pPersistFile:LPPERSISTFILE
local	wsz[MAX_PATH]:WORD
 
	invoke CoCreateInstance, addr CLSID_ShellLink, NULL,
		CLSCTX_INPROC_SERVER, addr IID_IShellLinkA, addr pShellLink
	.if (eax == S_OK)
 
;		 // Set the path to the shortcut target and add the 
;		 // description. 
		invoke vf(pShellLink, IShellLinkA, SetPath), lpszPathObj
		invoke vf(pShellLink, IShellLinkA, SetDescription), lpszDesc
 
;		// Query IShellLink for the IPersistFile interface for saving the 
;		// shortcut in persistent storage. 
		invoke vf(pShellLink, IShellLinkA, QueryInterface), addr IID_IPersistFile,
			addr pPersistFile
 
		.if (eax == S_OK)
			invoke MultiByteToWideChar, CP_ACP, 0, lpszPathLink, -1, addr wsz, MAX_PATH
;			// Save the link by calling IPersistFile::Save. 
			invoke vf(pPersistFile, IPersistFile, Save), addr wsz, TRUE
			invoke vf(pPersistFile, IUnknown, Release)
		.endif
		invoke vf(pShellLink, IUnknown, Release)
	.endif 
	ret
CreateLink endp


?COPYDROPFILES equ 1

;--- grfKeyState doesnt show mouse button state (because the drop
;--- gets into effect if the mouse button is released!)

Drop proc uses __this esi this_:ptr CShellBrowser, pDataObject:LPDATAOBJECT, grfKeyState:DWORD, pt:POINTL, pdwEffect:ptr DWORD

local	pidl:LPITEMIDLIST
local	fe:FORMATETC
local	dwFiles:DWORD
local	dwSize:DWORD
local	stgmedium:STGMEDIUM
local	shfos:SHFILEOPSTRUCT
local	szPath[MAX_PATH]:BYTE
local	szPathTmp[MAX_PATH]:BYTE
local	szDesc[MAX_PATH]:BYTE
local	szFile[MAX_PATH]:BYTE

	mov __this,this_
	DebugOut "IShellBrowser::Drop(%X)", grfKeyState

	mov fe.cfFormat, CF_HDROP
	mov fe.ptd, NULL
	mov fe.dwAspect,DVASPECT_CONTENT
	mov fe.lindex,-1
	mov fe.tymed, TYMED_HGLOBAL
	invoke vf(pDataObject, IDataObject, GetData), ADDR fe, ADDR stgmedium
	.if (eax != S_OK)
		jmp exit
	.endif

if ?COPYDROPFILES
;--------------------------------- get number of files dropped
	invoke DragQueryFile, stgmedium.hGlobal, -1, 0, 0
	mov dwFiles, eax
	mov dwSize, 1
	xor esi, esi
	.while (esi < dwFiles)
		invoke DragQueryFile, stgmedium.hGlobal, esi, 0, 0
		inc eax
		add dwSize, eax
		inc esi
	.endw

	invoke LocalAlloc, LMEM_FIXED, dwSize
	.if (!eax)
		jmp exit2
	.endif
	mov shfos.pFrom, eax
	mov edx, eax

	xor esi, esi
	.while (esi < dwFiles)
		push edx
		invoke DragQueryFile, stgmedium.hGlobal, esi, edx, MAX_PATH
		pop edx
		add edx, eax
		mov byte ptr [edx],0
		inc edx
		inc esi
	.endw
	mov byte ptr [edx],0
else
	mov eax, stgmedium.hGlobal			;that doesnt work with wide chars
	add eax,[eax].DROPFILES.pFiles
	mov shfos.pFrom, eax
endif
	mov eax, m_hWnd
	mov shfos.hwnd, eax

	invoke TreeView_GetDropHilight( m_hWndTV)
	invoke GetFullPidl@CShellBrowser, eax
	mov pidl, eax
	invoke SHGetPathFromIDList, pidl, addr szPath
	invoke vf(m_pMalloc, IMalloc, Free), pidl
	lea eax, szPath
	mov shfos.pTo, eax

	mov edx, grfKeyState
	.if (m_bRButton)
		mov ecx, pdwEffect
		invoke DisplayContextMenu, edx, [ecx]
		.if (!eax)
			jmp exit2
		.endif
		mov edx, eax
	.endif

	mov ecx, pdwEffect
	mov ecx, [ecx]
	and edx, MK_CONTROL or MK_SHIFT
	.if ((ecx & DROPEFFECT_LINK) && (edx == MK_CONTROL or MK_SHIFT))
		mov esi, shfos.pFrom
		.while (byte ptr [esi])
			invoke lstrlen, esi
			mov dwSize, eax
			invoke lstrcpy, addr szFile, esi
			invoke PathStripPath, addr szFile
			invoke wsprintf, addr szDesc, CStr("Shortcut of %s"), addr szFile
			invoke wsprintf, addr szPathTmp, CStr("%s\%s.lnk"), shfos.pTo, addr szFile
			invoke CreateLink, esi, addr szPathTmp, addr szDesc
			add esi, dwSize
			inc esi
		.endw
	.else
		.if ((ecx & DROPEFFECT_MOVE) && (!edx))
			mov shfos.wFunc, FO_MOVE
		.elseif (ecx & DROPEFFECT_COPY)
			mov shfos.wFunc, FO_COPY
		.else
			jmp exit3
		.endif
		mov shfos.fFlags, FOF_ALLOWUNDO
		mov shfos.fAnyOperationsAborted, FALSE
;;		mov shfos.hNameMappings, xxx
;;		mov shfos.lpszProgressTitle, xxx
		invoke SHFileOperation, addr shfos
	.endif
exit3:
if ?COPYDROPFILES
	invoke LocalFree, shfos.pFrom
endif
exit2:
	invoke ReleaseStgMedium, addr stgmedium
exit:
	invoke TreeView_SelectDropTarget( m_hWndTV, NULL)

if ?DRAGDROPHELPER
	.if (g_pDropTargetHelper)
		mov ecx, pdwEffect
		invoke vf(g_pDropTargetHelper, IDropTargetHelper, Drop), pDataObject, addr pt, [ecx]
		DebugOut "IDropTargetHelper::Drop=%X", eax
	.endif
endif

	return S_OK
	align 4

Drop endp

endif

	end
