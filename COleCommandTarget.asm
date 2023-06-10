
;--- the IShellBrowser is requested for IOleCommandTarget
;--- currently we dont allow any commands

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

INSIDE_CShellBrowser equ 1

	include CShellBrowser.inc

if ?OLECOMMANDTARGET

	.const

COleCommandTargetVtbl label dword
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd QueryStatus_, Exec_


	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER hWnd, hWndTV

	@MakeStubsEx CShellBrowser, IOleCommandTarget, QueryInterface, AddRef, Release
	@MakeStubs CShellBrowser, IOleCommandTarget, QueryStatus, Exec

;;CGID_Explorer sCGID_Explorer

QueryStatus proc uses __this this_:ptr CShellBrowser, pguidCmdGroup:REFGUID, cCmds:DWORD, prgCmds:ptr OLECMD, pCmdText:ptr OLECMDTEXT
ifdef _DEBUG
local szGUID[40]:byte
local wszGUID[40]:word
	.if (pguidCmdGroup)
		invoke StringFromGUID2, pguidCmdGroup, addr wszGUID, 40
		invoke WideCharToMultiByte, CP_ACP, 0, addr wszGUID, -1, addr szGUID, 40, NULL, NULL
	.else
		invoke lstrcpy, addr szGUID, CStr("NULL")
	.endif
	DebugOut "CShellBrowser::QueryStatus(%s, %X)", addr szGUID, cCmds
endif
	.if (pguidCmdGroup)
		invoke IsEqualGUID, pguidCmdGroup, addr CGID_Explorer
		.if (eax)
			mov ecx, cCmds
			mov edx, prgCmds
			mov eax, pCmdText
			.while (ecx)
				pushad
				mov [edx].OLECMD.cmdf, 0
				.if (eax)
					mov [eax].OLECMDTEXT.cmdtextf,OLECMDTEXTF_NONE
					mov [eax].OLECMDTEXT.cwActual,0
				.endif
				popad
				dec ecx
				add edx, sizeof OLECMD
				.if (eax)
					add eax, sizeof OLECMDTEXT
				.endif
			.endw
			return S_OK
		.endif
	.endif
	return OLECMDERR_E_UNKNOWNGROUP

QueryStatus endp

Exec proc uses __this this_:ptr CShellBrowser, pguidCmdGroup:REFGUID, nCmdID:DWORD, nCmdExecOpt:DWORD, pvaIn:ptr VARIANT, pvaOut:ptr VARIANT
ifdef _DEBUG
local szGUID[40]:byte
local wszGUID[40]:word
	.if (pguidCmdGroup)
		invoke StringFromGUID2, pguidCmdGroup, addr wszGUID, 40
		invoke WideCharToMultiByte, CP_ACP, 0, addr wszGUID, -1, addr szGUID, 40, NULL, NULL
	.else
		invoke lstrcpy, addr szGUID, CStr("NULL")
	.endif
	DebugOut "CShellBrowser::Exec(%s, %X, %X)", addr szGUID, nCmdID, nCmdExecOpt
endif
	return OLECMDERR_E_DISABLED
Exec endp


endif

	end
