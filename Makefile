
# nmake Makefile to create ExplASM.exe and CVViewer.dll
# tools used:
# - JWasm
# - MS Link
# - MS RC
#
# adjust the WinInc paths for includes/libs before running it

!ifndef DEBUG
DEBUG = 0
!endif

!ifndef MASM
MASM=0
!endif

WININC=\wininc

!if $(DEBUG)
AOPTD=-Zi -D_DEBUG
LOPTD=/DEBUG
!endif

SRCMODS = \
!include modules.inc
OBJNAMES = $(SRCMODS:.ASM=.OBJ)
!if $(DEBUG)
OBJMODS = $(OBJNAMES:.\=DEBUG\)
!else
OBJMODS = $(OBJNAMES:.\=RELEASE\)
!endif

NAMEDLL = CVViewer
NAMEEXE = ExplASM

AOPT=-nologo -c -coff -Sg $(AOPTD) -Fl$* -Sg -Fo$* -I$(WININC)\Include
!if $(MASM)
ASM = ml.exe $(AOPT)
!else
ASM = jwasm.exe $(AOPT)
!endif
LINK = link.exe
RC = rc.exe

DEPS = CVViewer.inc CShellBrowser.inc

!if $(DEBUG)
OUTDIR=DEBUG
!else
OUTDIR=RELEASE
!endif

.SUFFIXES: .asm .obj

.asm{$(OUTDIR)}.obj:
    @$(ASM) $<

LIBS=kernel32.lib advapi32.lib user32.lib gdi32.lib uuid.lib ole32.lib oleaut32.lib shell32.lib comctl32.lib

ALL: $(OUTDIR) $(OUTDIR)\$(NAMEDLL).dll $(OUTDIR)\$(NAMEEXE).exe

$(OUTDIR):
	@mkdir $(OUTDIR)

$(OUTDIR)\$(NAMEDLL).dll: $*.obj $(OUTDIR)\rsrc.res $(OBJMODS) Makefile
    @$(LINK) @<< /NOLOGO $(LOPTD)
$*.obj $(OBJMODS) $(OUTDIR)\rsrc.res
/OUT:$*.dll /DLL /MAP:$*.map /DEF:$(NAMEDLL).def /SUBSYSTEM:windows
/LIBPATH:$(WININC)\Lib $(LIBS)
<<

$(OUTDIR)\$(NAMEEXE).exe: $*.obj $(OUTDIR)\rsrc.res $(OBJMODS) Makefile
    @$(LINK) @<< /NOLOGO $(LOPTD)
$*.obj $(OBJMODS) $(OUTDIR)\rsrc.res
/OUT:$*.exe /MAP:$*.map /SUBSYSTEM:windows /OPT:NOWIN98
/LIBPATH:$(WININC)\Lib 
<<

$(OUTDIR)\rsrc.res: rsrc.rc
	@$(RC) -i$(WININC)\Include -fo$*.res rsrc.rc

$(OBJMODS): $(DEBS)

clean:
	erase $(OUTDIR)\*.obj
	erase $(OUTDIR)\*.map
	erase $(OUTDIR)\*.exe
	erase $(OUTDIR)\*.lst
	erase $(OUTDIR)\*.res
