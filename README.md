# vbsqlite
Staticly compiled sqlite into a VB6 ActiveX dll

Requires a custom linker that can swap VB6 `.obj` files for C/C++ replacement (surrogate) `.cobj` files.

First compile the included linker project from `lib/linker` directory, then locate `LINK.EXE` in `C:\Program Files (x86)\Microsoft Visual Studio\VB98` and rename it to `vbLINK.exe`. After this copy `link.exe` from surrogate linker project to `C:\Program Files (x86)\Microsoft Visual Studio\VB98`.

`VbSqlite.vbp` is instrumented with custom linker switches
```
[VBCompiler]
LinkSwitches=KERNEL32.LIB /OPT:NOREF
```
These are required as sqlite sources in `lib` are compiled with `/MD` (link with MSVCRT.LIB) option and need `KERNEL32.LIB` to find winapi functions. `/OPT:NOREF` is needed for the final `VbSqlite.dll` to import all run-time functions from `MSVBVM60.DLL`, otherwise the linker will skip `VbSqlite.obj` as containing only unreferenced symbols.
