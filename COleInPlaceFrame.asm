

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

LPDROPTARGETHELPER typedef ptr IDropTargetHelper

?AGGREGATION	= 0	;no aggregation
?EVENTSUPPORT	= 0	;no event support

	include olecntrl.inc
	include debugout.inc
	include rsrc.inc
	.list
	.cref

;protostrlenW typedef proto :ptr WORD
;externdef _imp__lstrlenW@4:PTR protostrlenW
;lstrlenW equ <_imp__lstrlenW@4>

INSIDE_CShellBrowser equ 1

	include CShellBrowser.inc

if ?OLEINPLACEFRAME

E_NOTOOLSPACE equ 800401A1h

	.const

COleInPlaceFrameVtbl label DWORD
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd offset GetWindow__
	dd offset ContextSensitiveHelp_
	dd offset GetBorder_
	dd offset RequestBorderSpace_
	dd offset SetBorderSpace_
	dd offset SetActiveObject_
	dd offset InsertMenus_
	dd offset SetMenu__
	dd offset RemoveMenus_
	dd offset SetStatusText_
	dd offset EnableModeless_
	dd offset TranslateAccelerator__

	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER hWnd, pOleInPlaceActiveObject

	@MakeStubsEx CShellBrowser, IOleInPlaceFrame, QueryInterface, AddRef, Release
	@MakeStubs CShellBrowser, IOleInPlaceFrame, GetWindow_, ContextSensitiveHelp
	@MakeStubs CShellBrowser, IOleInPlaceFrame, GetBorder, RequestBorderSpace
	@MakeStubs CShellBrowser, IOleInPlaceFrame, SetBorderSpace, SetActiveObject
	@MakeStubs CShellBrowser, IOleInPlaceFrame, InsertMenus, SetMenu_, RemoveMenus
	@MakeStubs CShellBrowser, IOleInPlaceFrame, SetStatusText, EnableModeless, TranslateAccelerator_

GetWindow_ proc uses __this this_:ptr CShellBrowser, phwnd:ptr HWND

	DebugOut "IOleInPlaceSite::GetWindow"
	mov __this, this_
	mov eax, phwnd
	mov ecx, m_hWnd
	mov [eax], ecx
	return S_OK

GetWindow_ endp


ContextSensitiveHelp proc this_:ptr CShellBrowser, fEnterMode:BYTE
	DebugOut "IOleInPlaceSite::ContextSensitiveHelp"
	return E_NOTIMPL
ContextSensitiveHelp endp


GetBorder proc uses __this this_:ptr CShellBrowser, lprectBorder:ptr RECT

	DebugOut "IOleInPlaceUIWindow::GetBorder"
	return E_NOTOOLSPACE

GetBorder endp


RequestBorderSpace proc this_:ptr CShellBrowser, pborderwidths:ptr BORDERWIDTHS

	DebugOut "IOleInPlaceUIWindow::RequestBorderSpace"
	return E_NOTOOLSPACE 

RequestBorderSpace endp


SetBorderSpace proc uses __this esi this_:ptr CShellBrowser, pborderwidths:ptr BORDERWIDTHS

	DebugOut "IOleInPlaceUIWindow::SetBorderSpace(%X)", pborderwidths
	mov eax,pborderwidths
	.if (!eax)
		return S_OK
	.else
		mov ecx,[eax].BORDERWIDTHS.left
		add ecx,[eax].BORDERWIDTHS.top
		add ecx,[eax].BORDERWIDTHS.right
		add ecx,[eax].BORDERWIDTHS.bottom
		.if (!ecx)
			return S_OK
		.endif
	.endif
	return E_UNEXPECTED

SetBorderSpace endp


SetActiveObject proc uses __this this_:ptr CShellBrowser, pActiveObject:LPOLEINPLACEACTIVEOBJECT, pszObjName:ptr WORD

	DebugOut "IOleInPlaceUIWindow::SetActiveObject(%X)", pActiveObject
	mov __this,this_
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, Release)
	.endif
	mov eax, pActiveObject
	mov m_pOleInPlaceActiveObject, eax
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, AddRef)
	.endif
	return S_OK

SetActiveObject endp


InsertMenus proc uses esi this_:ptr CShellBrowser, hmenuShared:HMENU, lpMenuWidths:ptr OLEMENUGROUPWIDTHS

	DebugOut "IOleInPlaceFrame::InsertMenus(%X)", hmenuShared
	return E_UNEXPECTED

InsertMenus endp

SetMenu_ proc uses __this this_:ptr CShellBrowser, hmenuShared:HMENU, holemenu:HANDLE, hwndActiveObject:HWND

	DebugOut "IOleInPlaceFrame::SetMenu(%X, %X, %X)", hmenuShared, holemenu, hwndActiveObject
	return E_UNEXPECTED

SetMenu_ endp

RemoveMenus proc uses __this this_:ptr CShellBrowser, hmenuShared:HMENU

	DebugOut "IOleInPlaceFrame::RemoveMenus(%X)", hmenuShared
	return E_UNEXPECTED

RemoveMenus endp

SetStatusText proc uses __this this_:ptr CShellBrowser, pszStatusText:ptr WORD

local dwSize:DWORD

	mov __this,this_
	.if (pszStatusText)
		invoke lstrlenW, pszStatusText
		add eax, 4
		and al, 0FCh
		sub esp, eax
		mov dwSize, eax
		mov edx, esp
		invoke WideCharToMultiByte, CP_ACP, 0, pszStatusText, -1, edx, dwSize, NULL, NULL
;;		invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, esp
		add esp, dwSize
	.endif

	return S_OK

SetStatusText endp

EnableModeless proc this_:ptr CShellBrowser, fEnable:DWORD

	DebugOut "IOleInPlaceFrame::EnableModeless(%u)", fEnable
	return S_OK

EnableModeless endp

TranslateAccelerator_ proc this_:ptr CShellBrowser, lpmsg:ptr MSG, wID:WORD

	DebugOut "IOleInPlaceFrame::TranslateAccelerator"
	return S_FALSE

TranslateAccelerator_ endp

endif

	end
