# VbSqlite
Staticly compiled sqlite into a VB6 ActiveX dll

This project requires a custom linker that can selectively swap VB6 `.obj` files for C/C++ replacement (surrogate) `.cobj` files before linking final executable.

First compile the included linker project from `lib/linker` directory (just 115 LOC), then locate `LINK.EXE` in `C:\Program Files (x86)\Microsoft Visual Studio\VB98` and rename it to `vbLINK.exe`. After this copy `link.exe` from surrogate linker project to `C:\Program Files (x86)\Microsoft Visual Studio\VB98`.

Note that `VbSqlite.vbp` project is instrumented with custom linker switches (using notepad)
```
[VBCompiler]
LinkSwitches=KERNEL32.LIB /OPT:NOREF
```
These are required as `build.bat` in `lib` compiles sqlite sources with `/MD` option (link with MSVCRT.LIB) and then linker needs `KERNEL32.LIB` to find all used winapi functions. Option `/OPT:NOREF` is needed for the final `VbSqlite.dll` to import VB6 run-time functions from `MSVBVM60.DLL`, otherwise the linker will skip `VbSqlite.obj` as containing only unreferenced symbols and produce invalid executable.

Note that you have to compile `VbSqlite.dll` to `bin` directory only as the replacement linker needs to find surrogate `.cobj` files in executable target folder for the swap-out to occur. Currenly only `mdSqlite.obj` and `mdSqliteHelper.obj` are swapped with C/C++ implementation.

The original `mdSqlite.bas` contains stub function implementations that call into helper `debug_sqlite3.dll` -- stdcall compiled from sqlite's amalgamated sources. The latter `.dll` is only needed in VB6 IDE during debugging sessions as the final `VbSqlite.dll` will have everything needed staticly linked.

The surrogate `mdSqlite.cpp` contains naked redirecting functions that `jmp` into the corresponding original sqlite function -- only name mangling here is important to match VB6 generated function names. The actual amalgamated sqlite source is included in `mdSqliteHelper.c` after setting target `#define`s and `#pragma`s. When compiled this replaces `mdSqliteHelper.bas` module which can be left as an empty placeholder in VB project or as in this case contain constants only as no functions will be available for static linking from the surrogate `mdSqliteHelper.cobj`.

In `lib/lua` directory there are [`LPeg.re`](http://www.inf.puc-rio.br/~roberto/lpeg/re.html) based parsers that produce `mdSqlite.bas` and `mdSqliteHelper.bas` from the amalgamated `sqlite3.c` source file. When `sed` scripts fail the power of PEG grammars and Lua come to rescue!
