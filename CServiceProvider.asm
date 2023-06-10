
;--- the IShellBrowser object may be asked for IServiceProvider
;--- main purpose seems to get an IShellBrowser instance of 
;--- the top level explorer (SID_STopLevelBrowser)
;--- In this case we just return the current IShellBrowser object

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

	include macros.inc

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

INSIDE_CShellBrowser equ 1

	include CShellBrowser.inc

?TOPLEVELBROWSER	equ 0

if ?SERVICEPROVIDER

	.data
if ?TOPLEVELBROWSER
g_pShellBrowser LPSHELLBROWSER NULL
g_pServiceProvider LPSERVICEPROVIDER NULL
endif
g_bInit BOOLEAN FALSE

	.const

CServiceProviderVtbl label dword
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd QueryService_


if ?TOPLEVELBROWSER
CLSID_Browser GUID {0A5E46E3Ah, 8849h, 11D1h, {9Dh,8Ch,00h,0C0h,4Fh,0C9h,9Dh,61h}}
endif
if ?WEBBROWSER
CLSID_NotKnown GUID {0D7D1D00h, 6FC0h, 11D0h, {0A9h,74h,00h,0C0h,4Fh,0D7h,05h,0A2h}}
CLSID_WebBrowser GUID {8856F961h,340Ah,11D0h, {0A9h,6Bh,00h,0C0h,4Fh,0D7h,05h,0A2h}}
endif
;SID_STopWindow sSID_STopWindow

	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER hWnd, hWndTV
if ?WEBBROWSER
	MEMBER pWebBrowser
endif

	@MakeStubsEx CShellBrowser, IServiceProvider, QueryInterface, AddRef, Release
	@MakeStubs CShellBrowser, IServiceProvider, QueryService

;SID_STopLevelBrowser sSID_STopLevelBrowser


Init proc

	mov g_bInit, TRUE
if ?TOPLEVELBROWSER
	invoke CoCreateInstance, addr CLSID_Browser, NULL, CLSCTX_INPROC_SERVER,\
			addr IID_IServiceProvider, addr g_pServiceProvider
	.if (eax == S_OK)
		invoke vf(g_pServiceProvider, IServiceProvider, QueryService), addr SID_STopLevelBrowser,\
				addr IID_IShellBrowser, addr g_pShellBrowser
	.endif
	.if (!g_pShellBrowser)
		mov eax, __this
		mov g_pShellBrowser, eax
		invoke vf(g_pShellBrowser, IUnknown, AddRef)
	.endif
endif
	ret
Init endp

Deinit@IServiceProvider proc public
if ?TOPLEVELBROWSER
	.if (g_pShellBrowser)
		invoke vf(g_pShellBrowser, IUnknown, Release)
	.endif
	.if (g_pServiceProvider)
		invoke vf(g_pServiceProvider, IUnknown, Release)
	.endif
endif
	ret
Deinit@IServiceProvider endp

QueryService proc uses __this this_:ptr CShellBrowser, guidService:REFGUID, riid:REFIID, ppv:ptr LPVOID
ifdef _DEBUG
local szGUID[40]:byte
local wszGUID[40]:word
local szIID[40]:byte
local wszIID[40]:word
	invoke StringFromGUID2, guidService, addr wszGUID, 40
	invoke WideCharToMultiByte, CP_ACP, 0, addr wszGUID, -1, addr szGUID, 40, NULL, NULL
	invoke StringFromGUID2, riid, addr wszIID, 40
	invoke WideCharToMultiByte, CP_ACP, 0, addr wszIID, -1, addr szIID, 40, NULL, NULL
endif
	mov __this, this_
	.if (!g_bInit)
		invoke Init
	.endif
	invoke IsEqualGUID, guidService, addr SID_SShellBrowser
	.if (eax)
		DebugOut "CShellBrowser::QueryService(SID_SShellBrowser[%s], %s)", addr szGUID, addr szIID
		invoke vf(__this, IUnknown, QueryInterface), riid, ppv
		ret
	.endif
	invoke IsEqualGUID, guidService, addr SID_STopWindow
	.if (eax)
		DebugOut "CShellBrowser::QueryService(SID_STopWindow[%s], %s)", addr szGUID, addr szIID
		invoke vf(__this, IUnknown, QueryInterface), riid, ppv
		ret
	.endif
;------------------- without that query the statusline remains blank in XP
	invoke IsEqualGUID, guidService, addr SID_STopLevelBrowser
	.if (eax)
		DebugOut "CShellBrowser::QueryService(SID_STopLevelBrowser[%s], %s)", addr szGUID, addr szIID
if ?TOPLEVELBROWSER
		invoke vf(g_pShellBrowser, IShellBrowser, QueryInterface), riid, ppv
else
		invoke vf(__this, IUnknown, QueryInterface), riid, ppv
endif
		ret
	.endif
	DebugOut "CShellBrowser::QueryService(%s, %s)", addr szGUID, addr szIID
if ?WEBBROWSER
	invoke IsEqualGUID, guidService, addr CLSID_NotKnown
	.if (eax)
		.if (!m_pWebBrowser)
			invoke CoCreateInstance, addr CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER,\
				addr IID_IWebBrowser, addr m_pWebBrowser
		.endif
		.if (m_pWebBrowser)
			invoke vf(m_pWebBrowser, IUnknown, AddRef)
		.endif
		mov ecx,ppv
		mov eax, m_pWebBrowser
		mov [ecx], eax
		return S_OK
	.endif
endif
	mov ecx, ppv
	mov dword ptr [ecx], NULL
	return E_NOINTERFACE

QueryService endp

endif

	end
