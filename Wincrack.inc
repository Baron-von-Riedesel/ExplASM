
;--- this file was previously named windowsx.inc
;--- but windowsx.inc is now generated by h2incx.exe
;--- what's remaining here are statusbar message cracker macros

;--- ListView (now in h2incx generated commctrl.inc)

if 0
ListView_GetItemText macro hwndLV,i,iSubItem_,pszText_,cchTextMax_
	sub esp,sizeof LV_ITEM
	mov [esp].LV_ITEM.iSubItem,iSubItem_
	mov [esp].LV_ITEM.pszText,pszText_
	mov [esp].LV_ITEM.cchTextMax,cchTextMax_
	invoke SendMessage,hwndLV,LVM_GETITEMTEXT,i,esp
	add esp,sizeof LV_ITEM
	endm
ListView_SetItemText macro hwndLV,i,iSubItem_,pszText_
	sub esp,sizeof LV_ITEM
	mov [esp].LV_ITEM.iSubItem,iSubItem_
	mov [esp].LV_ITEM.pszText,pszText_
	invoke SendMessage,hwndLV,LVM_SETITEMTEXT,i,esp
	add esp,sizeof LV_ITEM
	endm
ListView_SetItemState macro hwndLV,i,data,mask_
	sub esp,sizeof LV_ITEM
	mov [esp].LV_ITEM.state,data
	mov [esp].LV_ITEM.stateMask,mask_
	invoke SendMessage,hwndLV,LVM_SETITEMSTATE,i,esp
	add esp,sizeof LV_ITEM
	endm
endif    
ListView_SetCheckState macro hwndLV,i,fCheck
	mov eax,fCheck
	inc eax
	shl eax,12
	ListView_SetItemState hwndLV,i,eax,0F000h
	endm


;--- StatusBar

if 1
;*** StatusBar macros (not found in MS C)

SB_ISSIMPLE		equ (WM_USER+14)
SB_SETTIPTEXT	equ (WM_USER+16)

StatusBar_GetText macro hWnd, iPart, pText
	invoke SendMessage, hWnd, SB_GETTEXT, iPart, pText
	endm 
StatusBar_SetText macro hWnd, iPart, pText
	invoke SendMessage, hWnd, SB_SETTEXT, iPart, pText
	endm 
StatusBar_GetTextLength macro hWnd, iPart
	invoke SendMessage, hWnd, SB_GETTEXTLENGTH, iPart, 0
	endm 
StatusBar_SetSimpleMode macro hWnd, bFlag
	invoke SendMessage, hWnd, SB_SIMPLE, bFlag, 0
	endm 
StatusBar_SetParts macro hWnd, iParts, pdwWidths
	invoke SendMessage, hWnd, SB_SETPARTS, iParts, pdwWidths
	endm 
StatusBar_IsSimple macro hWnd
    invoke SendMessage, hWnd, SB_ISSIMPLE, 0, 0
    endm
StatusBar_SetTipText macro hWnd, iPart, pszText
    invoke SendMessage, hWnd, SB_SETTIPTEXT, iPart, pszText
    endm
endif

;*** EOF 