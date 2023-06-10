
;--- define interface IInterfaceViewer

BEGIN_INTERFACE IInterfaceViewer, IUnknown
	STDMETHOD View	,hwndParent:HWND, riid:REFIID, punk:LPUNKNOWN
END_INTERFACE

LPINTERFACEVIEWER typedef ptr IInterfaceViewer


sCLSID_CVViewer textequ <{ 500bc6a0h, 04268h, 011d7h, { 0b5h, 023h, 00h, 050h, 0fch, 04ah, 093h, 073h }}>

;--- define CVViewer coclass structure

ifndef INSIDE_CVViewer
CVViewer struct
CVViewer ends
endif

;--- externals + prototypes

externdef IID_IInterfaceViewer: IID

Create@CVViewer			PROTO :ptr ObjectEntry, :LPUNKNOWN
Destroy@CVViewer		PROTO this_:ptr CVViewer

