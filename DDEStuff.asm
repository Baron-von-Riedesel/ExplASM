
;--- all the routines needed for this DDE communication
;--- that takes place if user clicks "open", "explore" ...
;--- in the right panel (view window)

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

	include CShellBrowser.inc

NO_COMMAND      equ 0
VIEW_COMMAND    equ 1
EXPLORE_COMMAND equ 2
FIND_COMMAND    equ 3

	.data

g_aApplication	DWORD 0
g_aTopic		DWORD 0

protoSHLockShared	typedef proto :HANDLE, :DWORD
protoSHUnlockShared typedef proto :LPVOID
protoSHFreeShared	typedef proto :LPVOID, :DWORD
protoFree			typedef proto :LPVOID

LPFNSHLOCKSHARED	typedef ptr protoSHLockShared
LPFNSHUNLOCKSHARED	typedef ptr protoSHUnlockShared
LPFNSHFREESHARED	typedef ptr protoSHFreeShared
LPFNFREE			typedef ptr protoFree

g_lpfnSHLockShared		LPFNSHLOCKSHARED	NULL
g_lpfnSHUnlockShared	LPFNSHUNLOCKSHARED	NULL
g_lpfnSHFreeShared		LPFNSHFREESHARED	NULL
g_lpfnFree				LPFNFREE			NULL

	.data?

g_szFoldersApp		BYTE MAX_PATH dup (?)
g_szFoldersTopic	BYTE MAX_PATH dup (?)
g_szOpenFolder		BYTE MAX_PATH dup (?)
g_szExploreFolder	BYTE MAX_PATH dup (?)
g_szFindFolder		BYTE MAX_PATH dup (?)

	.code

__this	textequ <ebx>
_this	textequ <[__this].CShellBrowser>

	MEMBER hWndTV 
	MEMBER pShellFolder, pMalloc
	MEMBER bDontRespond

;--- calc size of a pidl

Pidl_GetSize proc uses esi pidl:LPITEMIDLIST

	mov esi, 0
	mov eax, pidl
	.if (eax)
		.while ([eax].SHITEMID.cb)
			movzx ecx,[eax].SHITEMID.cb
			add esi, ecx
			add eax, ecx
		.endw
		add esi, sizeof WORD
	.endif
	return esi

Pidl_GetSize endp

;--- copy a pidl to a newly allocated one

Pidl_Copy proc public uses esi edi pidlSource:LPITEMIDLIST

	mov esi, pidlSource
	.if (!esi)
		return NULL
	.endif

	invoke Pidl_GetSize, esi
	push eax
	invoke vf(m_pMalloc, IMalloc, Alloc), eax
	pop ecx
	.if (eax)
		mov edi, eax
		rep movsb
	.endif
	ret
Pidl_Copy endp

;--- copy an item of a pidl

Pidl_GetItem proc uses esi edi pidl:LPITEMIDLIST, item:DWORD

	mov eax, pidl
	.if (eax)
		mov esi, item
		.while (esi)
			.if ([eax].SHITEMID.cb)
				movzx edx,[eax].SHITEMID.cb
				add eax, edx
			.else
				return 0
			.endif
			dec esi
		.endw
		mov esi, eax
		movzx ecx,[esi].SHITEMID.cb
		.if (!ecx)
			return 0
		.endif
		add ecx, sizeof WORD
		push ecx
		invoke vf(m_pMalloc, IMalloc, Alloc), ecx
		pop ecx
		mov edi, eax
		rep movsb
		mov [edi-2], cx
	.endif
	ret
Pidl_GetItem endp

;--- copy a pidl, but skip the last item

Pidl_SkipLastItem proc public pidlSource:LPITEMIDLIST, pLastItem:ptr LPITEMIDLIST

	invoke Pidl_Copy, pidlSource
	.if (eax)
		push eax
		mov edx, eax
		.while ([eax].SHITEMID.cb)
			movzx ecx,[eax].SHITEMID.cb
			mov edx, eax
			add eax, ecx
		.endw
		.if (pLastItem)
			push edx
			invoke Pidl_Copy, edx
			mov ecx, pLastItem
			mov [ecx], eax
			pop edx
		.endif
		mov [edx].SHITEMID.cb, 0
		pop eax
	.endif
	ret

Pidl_SkipLastItem endp

;--- get last item

Pidl_GetLastItem proc public pidlSource:LPITEMIDLIST

	mov eax, pidlSource
	.if (eax)
		mov edx, eax
		.while ([eax].SHITEMID.cb)
			movzx ecx,[eax].SHITEMID.cb
			mov edx, eax
			add eax, ecx
		.endw
		invoke Pidl_Copy, edx
	.endif
	ret

Pidl_GetLastItem endp

;--- get default value of a key in HKEY_CLASSES_ROOT

GetValue proc pszKey:LPSTR, pszValue:LPSTR, dwSize:DWORD

local hKey:HANDLE
local dwType:DWORD

	invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, pszKey,
					0, KEY_READ, addr hKey
	.if (eax != ERROR_SUCCESS)
		return FALSE
	.endif
	invoke RegQueryValueEx, hKey, NULL, NULL, addr dwType, pszValue, addr dwSize
	push eax
	invoke RegCloseKey, hKey
	pop eax
	.if (eax == ERROR_SUCCESS)
		return TRUE
	.else
		return FALSE
	.endif
GetValue endp

;--- get Command part of a DDE command string in registry

GetCmd proc pszKey:LPSTR, pszValue:LPSTR, dwSize:DWORD

local szTemp[MAX_PATH]:BYTE

	invoke GetValue, pszKey, addr szTemp, sizeof szTemp
	mov ecx, pszValue
	mov byte ptr [ecx],0
	.if (eax)
		lea edx, szTemp
		mov ah,00
		.while (byte ptr [edx])
			mov al,[edx]
			.if (ah)
				.break .if (al == '(')
				mov [ecx],al
				inc ecx
			.elseif (al == '[')
				mov ah, 1
			.endif
			inc edx
		.endw
		mov byte ptr [ecx],0
		mov eax, 1
	.endif
	ret
GetCmd endp

;--- read in some registry variables needed for DDE stuff

GetDDEVariables proc public

local hinstLib:HINSTANCE   

	invoke GetValue, CStr("Folder\shell\explore\ddeexec\application"),
		addr g_szFoldersApp, sizeof g_szFoldersApp

	invoke GetValue, CStr("Folder\shell\explore\ddeexec\topic"),
		addr g_szFoldersTopic, sizeof g_szFoldersTopic

	invoke GetCmd, CStr("Folder\shell\open\ddeexec"),
		addr g_szOpenFolder, sizeof g_szOpenFolder

	invoke GetCmd, CStr("Folder\shell\explore\ddeexec"),
		addr g_szExploreFolder, sizeof g_szExploreFolder

	invoke GetCmd, CStr("Directory\shell\find\ddeexec"),
		addr g_szFindFolder, sizeof g_szFindFolder

	invoke GlobalAddAtom, addr g_szFoldersApp
	movzx eax, ax
	mov g_aApplication, eax
	invoke GlobalAddAtom, addr g_szFoldersTopic
	movzx eax, ax
	mov g_aTopic, eax
	
	invoke LoadLibrary, CStr("shell32.dll")
	mov hinstLib, eax
	.if (eax)
		invoke GetProcAddress, hinstLib, 521
		mov g_lpfnSHLockShared, eax
		invoke GetProcAddress, hinstLib, 522
		mov g_lpfnSHUnlockShared, eax
		invoke GetProcAddress, hinstLib, 523
		mov g_lpfnSHFreeShared, eax
		invoke FreeLibrary, hinstLib
	.endif
	invoke LoadLibrary, CStr("comctl32.dll")
	mov hinstLib, eax
	.if (eax)
		invoke GetProcAddress, hinstLib, 73
		mov g_lpfnFree, eax
		invoke FreeLibrary, hinstLib
	.endif
	
	DebugOut "GetDDEVariables, Application=%X, Topic=%X", g_aApplication, g_aTopic

	ret

GetDDEVariables endp

;--- check if we can handle command in a DDE command string 

GetCommandTypeFromDDEString proc pszCommand:LPSTR

local pszCopy:LPSTR
local dwChars:DWORD
local nCommand:DWORD
local pszTemp:LPSTR

	mov nCommand, NO_COMMAND
	invoke lstrlen, pszCommand
	inc eax
	mov dwChars, eax
	invoke LocalAlloc, LPTR, dwChars
	mov pszCopy, eax
	.if (eax)
		invoke lstrcpy, pszCopy, pszCommand

		mov ecx, pszCopy
		.while (byte ptr [ecx])
			mov al, [ecx]
			.if (al == '(')
				mov byte ptr [ecx],0
				.break
			.endif
			inc ecx
		.endw

;----  Find the beginning of the command portion by getting the first character 
;----   after the first '['.

		mov ecx, pszCopy
		.while (byte ptr [ecx])
			mov al, [ecx]
			.if (al == '[')
				inc ecx
				.break
			.endif
			inc ecx
		.endw   
		mov pszTemp, ecx

;------------------------------- check the command

		invoke lstrcmpi, pszTemp, addr g_szOpenFolder
		.if (!eax)
			mov nCommand, VIEW_COMMAND
		.else
			invoke lstrcmpi, pszTemp, addr g_szExploreFolder
			.if (!eax)
				mov nCommand, EXPLORE_COMMAND
			.else
				invoke lstrcmpi, pszTemp, addr g_szFindFolder
				.if (!eax)
					mov nCommand, FIND_COMMAND
				.endif
			.endif
		.endif

		invoke LocalFree, pszCopy

	.endif

	return nCommand

GetCommandTypeFromDDEString endp

;--- get a parameter from a DDE command string

GetParameter proc pszCommand:LPSTR, uParm:DWORD, pszParameter:LPSTR, dwMax:DWORD

	mov ecx, pszParameter
	mov edx, pszCommand
	mov ah,0
	.while (byte ptr [edx])
		mov al,[edx]
		.if (ah)
			.if (ah == 1)
				.break .if ((al == ')') || (al == ','))
			.endif
			.if (al == '"')
;				.if (byte ptr [edx+1] == '"')
;					mov [ecx], al
;					inc edx
;					inc ecx
;				.else
					xor ah,2
;				.endif
			.else
				mov [ecx], al
				inc ecx
			.endif
		.elseif ((al == '(') || (al == ','))
			dec uParm
			.if (uParm == 0)
				mov ah, 01
				.if (byte ptr [edx+1] == ' ')
					inc edx
				.endif
			.endif
		.endif
		inc edx
	.endw
	mov byte ptr [ecx],0
	ret
GetParameter endp

;--- convert string to DWORD

_StrToLong proc public pStr:LPSTR

local bNegative:BOOL

	mov bNegative, FALSE
	xor edx, edx
	mov ecx, pStr
	mov al,[ecx]
	.if ((al == '-') || (al == '+'))
		inc ecx
		.if (al == '-')
			mov bNegative, TRUE
		.endif
	.endif
	.while (byte ptr [ecx])
		mov al, [ecx]
		.break .if (!al)
		sub al, '0'
		movzx eax, al
		push ecx
		shl edx, 1
		mov ecx, edx
		shl edx, 2
		add edx, ecx
		pop ecx
		add edx, eax
		inc ecx
	.endw
	mov eax, edx
	.if (bNegative)
		neg eax
	.endif
	ret
_StrToLong endp

;--- get path (1. parameter) from a DDE command string

GetPathFromDDEString proc pszCommand:LPSTR , pszFolder:LPSTR , dwSize:DWORD

	invoke GetParameter, pszCommand, 1, pszFolder, dwSize
	ret

GetPathFromDDEString endp


;--- get 2. parameter from DDE execute string


GetPidlFromDDEString proc pszCommand:LPSTR

local pidl:LPITEMIDLIST
local pidlShared:LPITEMIDLIST
local pidlGlobal:LPITEMIDLIST
local pszHandle:LPSTR
local pszProcessId:LPSTR
local pszEnd:LPSTR
local hShared:HANDLE
local dwProcessId:DWORD
local szTemp[256]:byte

	invoke GetParameter, pszCommand, 2, addr szTemp, sizeof szTemp
	lea eax, szTemp
	mov pszHandle, eax

	mov ecx, pszHandle
	.while (byte ptr [ecx])
		mov al,[ecx]
		.break .if ((al == '-') || (al == '+') || ((al >= '0') && (al <= '9')))
		inc ecx
	.endw
	mov pszHandle, ecx

	mov ecx, pszHandle
	.while (byte ptr [ecx])
		mov al,[ecx]
		.if (al == ':')
			mov byte ptr [ecx],0
			inc ecx
			.break
		.endif
		inc ecx
	.endw
	mov pszProcessId, ecx      
	.while (byte ptr [ecx])
		mov al,[ecx]
		.if (al == ':')
			mov byte ptr [ecx],0
			inc ecx
			.break
		.endif
		inc ecx
	.endw
	mov ecx, pszProcessId
	.if (byte ptr [ecx] == 0)
		mov pszProcessId, NULL
	.endif

	.if (pszProcessId)
		invoke _StrToLong, pszHandle
		mov hShared, eax
		invoke _StrToLong, pszProcessId
		mov dwProcessId, eax

		.if (g_lpfnSHLockShared && g_lpfnSHUnlockShared && g_lpfnSHFreeShared)
			invoke g_lpfnSHLockShared, hShared, dwProcessId
			mov pidlShared, eax
			.if (eax)
;------------------------------------ make a local copy of the PIDL
				invoke Pidl_Copy, pidlShared
				mov pidl, eax
				invoke g_lpfnSHUnlockShared, pidlShared
			.endif
			invoke g_lpfnSHFreeShared, hShared, dwProcessId
		.endif
	.else

		invoke _StrToLong, pszHandle
		mov pidlGlobal, eax

		invoke Pidl_Copy, pidlGlobal
		mov pidl, eax

;---   The shared PIDL was allocated by the shell using a heap that the shell 
;---   maintains. The only way to free this memory is to call the Free 
;---   function in COMCTL32.DLL at ordinal 73.

		.if (g_lpfnFree)
			invoke g_lpfnFree, pidlGlobal
		.endif
	.endif

	return pidl

GetPidlFromDDEString endp

;--- get 3. parameter from DDE execute string

GetShowCmdFromDDEString proc pszCommand:LPSTR

local nShow:DWORD
local szTemp[256]:byte

	mov nShow, SW_SHOWDEFAULT
	invoke GetParameter, pszCommand, 3, addr szTemp, sizeof szTemp
	.if (szTemp)
		invoke _StrToLong, addr szTemp
		mov nShow, eax
	.endif

	return nShow

GetShowCmdFromDDEString endp


;--- get command, path, pidl and cmdShow from DDE execute string


ParseDDECommand proc pszCommand:LPSTR, pszFolder:LPSTR, dwSize:DWORD , ppidl:ptr LPITEMIDLIST,pnShow:ptr DWORD

local uCommand:DWORD

	mov ecx, pszFolder
	mov byte ptr [ecx],0
	mov ecx,ppidl
	mov dword ptr [ecx], NULL
	mov ecx,pnShow
	mov dword ptr [ecx],0

	invoke GetCommandTypeFromDDEString, pszCommand
	mov uCommand, eax

;--- If the command was not recognized, then don't try to free any memory because 
;--- we have no idea what is actually contained on the command line.

	.if (uCommand != NO_COMMAND)
		invoke GetPathFromDDEString, pszCommand, pszFolder, dwSize
		invoke GetPidlFromDDEString, pszCommand
		mov ecx, ppidl
		mov [ecx],eax
		invoke GetShowCmdFromDDEString, pszCommand
		mov ecx, pnShow
		mov [ecx],eax
	.endif

	return uCommand

ParseDDECommand endp

;--- find a pidl in treeview
;--- dont expand treeview for that

FindPidl proc public uses esi pidl:LPITEMIDLIST, bExpand:BOOL

local pShellFolder:LPSHELLFOLDER
local hItem:HANDLE
local pidl2:LPITEMIDLIST
local tvi:TV_ITEM

	invoke TreeView_GetRoot( m_hWndTV)
	mov tvi.hItem, eax
	xor esi, esi
	.while (eax)
;-------------------------------- get next item of the pidl
		invoke Pidl_GetItem, pidl, esi 
		.if (!eax)
			mov eax, tvi.hItem
			.break
		.endif
		mov pidl2, eax
		invoke GetFolder, tvi.hItem
		mov pShellFolder, eax
		.if (!eax)
			invoke vf(m_pMalloc, IMalloc, Free), pidl2
			xor eax, eax
			.break
		.endif
;-------------------------------- make sure the item's children are inserted
		.if (bExpand)
			mov tvi.mask_, TVIF_STATE
			mov tvi.stateMask, TVIS_EXPANDEDONCE 
			invoke TreeView_GetItem( m_hWndTV, addr tvi)
			.if (!(tvi.state & TVIS_EXPANDEDONCE))
				invoke InsertChildItems, tvi.hItem
			.endif
		.endif
;-------------------------------- now scan all children's lParam (is a pidl)
		invoke TreeView_GetChild( m_hWndTV, tvi.hItem)
		.while (eax)
			mov tvi.hItem, eax
			mov tvi.mask_,TVIF_PARAM
			invoke TreeView_GetItem( m_hWndTV, addr tvi)
			.if (eax)
				invoke vf(pShellFolder, IShellFolder, CompareIDs), 0, pidl2, tvi.lParam
				.if (!eax)
					inc eax
					.break
				.endif
			.endif
			invoke TreeView_GetNextSibling( m_hWndTV, tvi.hItem)
		.endw
		push eax
		invoke vf(pShellFolder, IShellFolder, Release)
		invoke vf(m_pMalloc, IMalloc, Free), pidl2
		pop eax
		inc esi
	.endw
	ret
FindPidl endp

;--- navigate to pidl in treeview
;--- returns TRUE if navigation succeeded


NavigateToPidl proc public uses esi pidl:LPITEMIDLIST

	invoke FindPidl, pidl, TRUE
	.if (eax)
		invoke TreeView_SelectItem( m_hWndTV, eax)
	.endif
	ret

NavigateToPidl endp


;--- handle WM_DDE_EXECUTE


OnDDEExecute proc public pszCommand:LPSTR

local lpTemp:LPSTR
local szFolder[MAX_PATH]:BYTE
local pidl:LPITEMIDLIST
local nShow:DWORD
local uCommand:DWORD
local dwRC:DWORD
local sei:SHELLEXECUTEINFO

	mov dwRC, FALSE
	DebugOut "%s", pszCommand

	.if (pszCommand)
		invoke ParseDDECommand, pszCommand, addr szFolder, MAX_PATH, addr pidl, addr nShow
		mov uCommand, eax

		.if (eax == EXPLORE_COMMAND)

			DebugOut "Explore"
			invoke NavigateToPidl, pidl
			mov dwRC, TRUE

		.elseif ((eax == VIEW_COMMAND) || (eax == FIND_COMMAND))

;---------------------------------- we dont handle "view" and "find"
;---------------------------------- and hence must call ShellExecuteEx 
			mov m_bDontRespond, TRUE

			invoke RtlZeroMemory, addr sei, sizeof SHELLEXECUTEINFO
			mov sei.cbSize, sizeof SHELLEXECUTEINFO

			movzx eax, word ptr szFolder
			.if (szFolder)
				mov sei.fMask, 0
				lea eax, szFolder
				mov sei.lpFile, eax
			.else
				mov sei.fMask, SEE_MASK_IDLIST
				mov eax, pidl
				mov sei.lpIDList, eax
			.endif

			mov sei.hwnd, NULL
			mov eax, nShow
			mov sei.nShow, eax
			.if (uCommand == VIEW_COMMAND)
				or sei.fMask, SEE_MASK_CLASSNAME
				mov sei.lpClass, CStr("folder")
				mov sei.lpVerb, CStr("open")
			.else
				mov sei.lpVerb, CStr("find")
			.endif

			invoke ShellExecuteEx, addr sei

			mov m_bDontRespond, FALSE

			mov dwRC, TRUE

		.endif

		.if (pidl)
			invoke vf(m_pMalloc, IMalloc, Free), pidl
		.endif

	.endif
	return dwRC

OnDDEExecute endp

	end
