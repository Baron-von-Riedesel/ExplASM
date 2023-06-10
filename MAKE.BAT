@echo off
rem adjust include/lib paths if necessary
jwasm -c -coff -nologo -I\wininc\Include -Fo Release\ -Fl=Release\ *.asm
rc -i \wininc\Include -fo Release\rsrc.res rsrc.rc
set LIB=\wininc\Lib
cd Release
link /NOLOGO /SUBSYSTEM:WINDOWS /OUT:CVViewer.dll CVViewer CShellBrowser DDEStuff CServiceProvider COleCommandTarget CDropTarget rsrc.res /DEF:..\CVViewer.def /DLL 
link /NOLOGO /SUBSYSTEM:WINDOWS /OUT:ExplASM.exe  ExplASM  CShellBrowser DDEStuff CServiceProvider COleCommandTarget CDropTarget COleInPlaceFrame.obj rsrc.res /OPT:NOWIN98
cd ..
