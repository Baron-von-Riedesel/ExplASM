

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

	include debugout.inc
	include rsrc.inc
	.list
	.cref

	include CShellBrowser.inc

	includelib  kernel32.lib
	includelib  advapi32.lib
	includelib  user32.lib
	includelib  gdi32.lib
	includelib  oleaut32.lib
	includelib  ole32.lib
	includelib  shell32.lib
	includelib  comctl32.lib
	includelib  uuid.lib

;--------------------------------------------------------------------------

?MULTIWINDOW	equ 0		;explorer will create just 1 window

	.data

externdef g_hInstance:HINSTANCE
externdef g_DllRefCount:DWORD

g_DllRefCount	DWORD 0
g_hInstance		HINSTANCE 0
g_hWndDlg		HWND 0
g_oldwndproc	DWORD 0
g_saverect		RECT <>
g_argc			DWORD 0
g_argv			LPSTR NULL

	.const

szOptions db "Options",0
szWindowPos db "WndDim",0
szLastFolder db "Folder",0

	.code

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


;--- setup arguments: will set g_argc and g_argv global vars


SetArguments proc public uses esi edi ebx

local	argc:dword

		invoke GetCommandLine
		and eax,eax
		jz exit
		mov esi,eax
		xor edi,edi			;EDI will count the number of arguments
		xor edx,edx			;EDX will count the number of bytes
							;needed for the arguments
							;(not including the null terminators)
nextarg:					;<---- get next argument
		.while (1)
			lodsb
			.break .if ((al != ' ') && (al != 9))	;skip spaces and tabs
		.endw
		or al,al
		je donescanX		;done commandline scan
		inc edi 			;Another argument
		xor ebx,ebx 		;EBX will count characters in argument
		dec esi 			;back up to reload character
		push esi 			;save start of argument
		mov cl,00
		.while (1)
			lodsb
			.break .if (!al)
			.if (!cl)
				.if ((al == ' ') || (al == 9))	;white space term. argument
					push ebx 			;save argument length
					jmp nextarg
				.endif
				.if ((!ebx) && al == '"')	;starts argument with "?
					or cl,1
										;handle argument beginning with doublequote
					pop eax				;throw away old start
					push esi 			;and set new start
					.continue
				.endif
			.elseif (al == '"')
				and cl,0FEh
				.continue
			.endif

			.if ((al == '\')  && (byte ptr [esi] == '"'))
				inc esi
			.endif
			inc ebx
			inc edx 			;one more space
		.endw
		push ebx 			; save length of last argument
donescanX:
		mov argc,edi		; Store number of arguments
		add edx,edi 		; add terminator bytes
		inc edi 			; add one for NULL pointer
		shl edi,2			; every pointer takes 4 bytes
		add edx,edi 		; add that space to space for strings

		invoke LocalAlloc, LMEM_FIXED, edx
		and eax,eax
		jz exit

		mov g_argv,eax
		add edi,eax 		; edi -> behind vector table (strings)
		mov ecx,argc
		mov g_argc,ecx
		lea ebx,[edi-4]
		mov dword ptr [ebx],0 ;mark end of argv
		sub ebx,4
		mov edx,ecx
		.while (edx)
			pop ecx 		;get length
			pop esi 		;get address
			mov [ebx],edi
			sub ebx,4
			.while (ecx)
				lodsb
				.if (al == '\')
					.continue .if (byte ptr [esi] == '"')
				.endif
				stosb
				dec ecx
			.endw
			xor al,al
			stosb
			dec edx
		.endw
exit:
		ret
		align 4

SetArguments endp

LoadParms proc uses esi ebx edi

local dwNums:DWORD
local pNum:ptr DWORD
local szPath[MAX_PATH]:byte
local szText[128]:byte

	invoke GetModuleFileName, NULL, addr szPath,  sizeof szPath
	lea ecx, szPath
	mov dword ptr [ecx+eax-3],"ini"

	invoke GetPrivateProfileString, addr szOptions, addr szWindowPos,
		CStr(""), addr szText, sizeof szText, addr szPath

	.if (eax)
		lea ebx, szText
		lea edx, g_rect.right
		mov pNum, edx
		mov esi, ebx
		mov dwNums, 2
		.while (dwNums)
			mov al, [ebx]
			.if ((al == ',') || (al == 0))
				mov byte ptr [ebx],0
				push eax
				xor eax, eax
				.if (esi != ebx)
					invoke StrToLong, esi
				.endif
				mov edx, pNum
				mov [edx],eax
				add edx, 4
				mov pNum, edx
				dec dwNums
				lea esi, [ebx+1]
				pop eax
				.break .if (al == 0)
			.endif
			inc ebx
		.endw
	.endif
	lea edi, g_saverect.right
	lea esi, g_rect.right
	movsd
	movsd

	invoke GetPrivateProfileString, addr szOptions, addr szLastFolder,
		CStr(""), addr g_szPath, MAX_PATH, addr szPath

	ret

LoadParms endp

SaveParms proc uses esi edi

local szPath[MAX_PATH]:byte
local szText[128]:byte

	invoke	GetModuleFileName, NULL, addr szPath,  sizeof szPath
	lea ecx, szPath
	mov dword ptr [ecx+eax-3],"ini"

	lea esi, g_rect.right
	lea edi, g_saverect.right
	mov ecx, 2
	repe cmpsd
	.if (ZERO?)
		jmp next
	.endif
	invoke wsprintf, addr szText, CStr("%u,%u"), g_rect.right, g_rect.bottom
	invoke WritePrivateProfileString, addr szOptions, addr szWindowPos,
		addr szText, addr szPath
next:
	invoke WritePrivateProfileString, addr szOptions, addr szLastFolder,
		addr g_szPath, addr szPath
exit:
	ret
SaveParms endp

if ?MULTIWINDOW
EnumThreadWindowsCB proc hWnd:HWND, lParam:LPARAM

	invoke GetWindowLong, hWnd, GWL_WNDPROC
	.if (eax == mywndproc)
		mov ecx, lParam
		inc dword ptr [ecx]
	.endif
	return TRUE

EnumThreadWindowsCB endp

IsLastWindow proc

local dwNumWindows:DWORD

	mov dwNumWindows, 0
	invoke GetCurrentThreadId
	lea ecx, dwNumWindows
	invoke EnumThreadWindows, eax, offset EnumThreadWindowsCB, ecx
	.if (dwNumWindows <= 1)
		invoke PostQuitMessage, 0
	.endif
	ret
IsLastWindow endp
endif

mywndproc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

if 0
	.if (message == WM_INITDIALOG)
		invoke SetWindowText, hWnd, CStr("ExplorerASM")
if ?MULTIWINDOW
	.elseif (message == WM_ACTIVATE)
		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax, hWnd
			mov g_hWndDlg, eax
		.endif
endif
	.endif
endif

	.if (message == WM_DESTROY)
if ?MULTIWINDOW
		invoke IsLastWindow
else
		invoke PostQuitMessage, 0
endif
	.endif
	invoke CallWindowProc, g_oldwndproc, hWnd, message, wParam, lParam
	ret
mywndproc endp


WinMain proc hInstance:HINSTANCE,hPrevInstance:HINSTANCE,lpszCmdLine:LPSTR,iCmdShow:dword

local hWndMain:HWND
local msg:MSG
local pShellFolder:LPSHELLFOLDER
local pShellBrowser:LPSHELLBROWSER
local iccx:INITCOMMONCONTROLSEX

if 1
	mov iccx.dwSize,sizeof INITCOMMONCONTROLSEX
	mov iccx.dwICC, ICC_WIN95_CLASSES
	invoke InitCommonControlsEx,addr iccx
else
	invoke InitCommonControls
endif
	invoke LoadLibrary, CStr("SHDOC401")	;someone needs that

;;	invoke CoInitialize, NULL
	invoke OleInitialize, NULL

	invoke LoadParms

	invoke SetArguments
	mov ecx, g_argc
	mov edx, g_argv
	.while (ecx)
		pushad
		mov ecx, [edx]
		mov al,[ecx]
		.if ((al == '-') || (al == '/'))
;--------------------------- handle options
		.else
			.if (edx != g_argv)
				invoke lstrcpy, addr g_szPath, ecx
			.endif
		.endif
		popad
		add edx, 4
		dec ecx
	.endw

	mov g_bCreateView, TRUE

	invoke SHGetDesktopFolder, addr pShellFolder
	.if (eax == S_OK)
		invoke Create@CShellBrowser, pShellFolder
		.if (eax)
			mov pShellBrowser, eax
			invoke Show@CShellBrowser, eax, NULL
			.if (eax)
				mov hWndMain, eax
				mov g_hWndDlg, eax
				invoke SetWindowLong, hWndMain, GWL_WNDPROC, offset mywndproc
				mov g_oldwndproc, eax
if ?MULTIWINDOW
				invoke SetClassLong, hWndMain, GCL_WNDPROC, offset mywndproc
endif
			.else
				invoke MessageBox, 0, CStr("couldn't create main window"), 0, MB_OK
			.endif
		.else
			invoke MessageBox, 0, CStr("couldn't create CShellBrowser object"), 0, MB_OK
		.endif
		invoke vf(pShellFolder, IShellFolder, Release)
	.endif


	.while (1)							;main message loop
		invoke GetMessage, addr msg, NULL, 0, 0
		.break .if (eax == 0)
		.if ((msg.message == WM_KEYDOWN) || (msg.message == WM_SYSKEYDOWN))
			invoke TranslateAccelerator@CShellBrowser, pShellBrowser, addr msg
			.continue .if (eax == S_OK)
		.endif
		invoke IsDialogMessage, g_hWndDlg, addr msg
		.continue .if (eax)
		invoke TranslateMessage, addr msg
		invoke DispatchMessage, addr msg
	.endw

	invoke SaveParms
exit:
;;	invoke CoUninitialize
	invoke OleUninitialize
	ret

WinMain endp

start:
	invoke GetModuleHandle,0
	mov g_hInstance, eax
	invoke WinMain,eax,0,0,0
if 0
	invoke GetCurrentProcess
	invoke TerminateProcess, eax, 0
endif
	invoke ExitProcess,eax

end start
