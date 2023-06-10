
;--- define macros for debugging output
;--- DebugOutNoLF: printf format OutputDebugStringA
;--- DebugOut: printf format OutputDebugStringA
;--- @trace: output a string via OutputDebugStringA
;--- @tracew: output a wide string via OutputDebugStringW
;--- @tracedw: output a dword (use __dw2aDebug)

ifndef CStr

CStr macro y:req
local sym,xxx
xxx	textequ @CurSeg
	.const
ifidni <y>,<"">
sym db 0
else
sym db y,0
endif
ifidni xxx,<_TEXT>
	.code
else
	.data
endif
	exitm <offset sym>
	endm

endif

ifndef wsprintf
wsprintfAproto typedef PROTO C :DWORD,:DWORD,:VARARG
externdef c _imp__wsprintfA:ptr wsprintfAproto
wsprintf equ <_imp__wsprintfA>
endif

ifndef _DEBUG
DEBUG = 0
else
DEBUG = 1
endif

DebugOutNoLF Macro x:req,y:VARARG
local sym,ii
if DEBUG
	.const
sym db x,0
	.code
	pushad
  if 0;def DEBUGPREFIX
	invoke OutputDebugString,DEBUGPREFIX
	popad
	pushad
  endif
  ifnb <y>
	sub esp,256
    ii = 0
    for parname,<y>
      ii = ii + 4
	endm
	push esp
	invoke wsprintf,[esp+ii+4],addr sym,y
	pop esp
	invoke OutputDebugString,esp
	add esp,256
  else
	invoke OutputDebugString,addr sym
  endif
	popad
endif
	endm

DebugOut macro x:req,y:VARARG
ifnb <y>
	DebugOutNoLF <x,0dh,0ah>,y
else
	DebugOutNoLF <x,0dh,0ah>
endif
	endm

;--- simple string output if wsprintf is not available

@trace	macro x
local y, defConst
ifdef _DEBUG
	defConst = 1
	for operand,<x>
		if (OPATTR(operand)) and 10010y
			defConst = 0
		endif
	endm
	pushad
	if defConst
		.const
		y	db x, 0
		.code
		invoke OutputDebugString, offset y
	else
		invoke OutputDebugString, x
	endif
	popad
endif
	endm

;--- same for wide chars (OutputDebugStringW doesnt work for win9x)

@tracew	macro x
ifdef _DEBUG
OutputDebugStringW proto stdcall :ptr WORD
	defConst = 1
	for operand,<x>
		if (OPATTR(operand)) and 10010y
			defConst = 0
		endif
	endm
	pushad
	if defConst
		.const
		y	dw L(x)
		.code
		invoke OutputDebugStringW, offset y
	else
		invoke OutputDebugStringW, x
	endif
	popad
endif
	endm

;--- simple number output if wsprintf is not available

@tracedw	macro x
ifdef _DEBUG
__dw2aDebug	proto stdcall
	pushad
	mov eax, x
	call __dw2aDebug
	popad
endif
	endm

