@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set lua_exe=luajit.exe
set bin_dir=..\bin
set src_dir=..\src

:: might want to set INCLUDE/LIB explicitly for VC6 if not already registered by default
::set INCLUDE=C:\Program Files (x86)\Microsoft Visual Studio\VC98\atl\include;C:\Program Files (x86)\Microsoft Visual Studio\VC98\mfc\include;C:\Program Files (x86)\Microsoft Visual Studio\VC98\include
::set LIB=C:\Program Files (x86)\Microsoft Visual Studio\VC98\mfc\lib;C:\Program Files (x86)\Microsoft Visual Studio\VC98\lib

pushd %~dp0

:: prepare debug_sqlite3.dll for IDE debugging and .cobj files for final static linking
%cl_exe% /LD mdSqlite.cpp mdSqliteHelper.c /Fedebug_sqlite3.dll /link /DEF:debug_sqlite3.def
copy debug_sqlite3.dll %bin_dir% > nul
copy mdSqlite.obj %bin_dir%\*.cobj > nul
copy mdSqliteHelper.obj %bin_dir%\*.cobj > nul

:: prepare mdSqlite.bas and mdSqliteHelper.bas
echo Generating mdSqlite.bas...
%lua_exe% lua\extract.lua < sqlite\sqlite3.c > %src_dir%\mdSqlite.bas
echo Generating mdSqliteHelper.bas...
%lua_exe% lua\consts.lua sqlite\sqlite3.c > %src_dir%\mdSqliteHelper.bas

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd
