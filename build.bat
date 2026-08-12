@echo off
setlocal

cd /d "%~dp0"
if errorlevel 1 exit /b 1

set "VS_VCVARS=C:\Program Files\Microsoft Visual Studio 10.0\VC\vcvarsall.bat"
if not exist "%VS_VCVARS%" (
  echo ERROR: Visual C++ environment script not found: %VS_VCVARS%
  exit /b 1
)

call "%VS_VCVARS%" x86 >nul
if errorlevel 1 (
  echo ERROR: vcvarsall.bat failed.
  exit /b 1
)

if not exist build mkdir build
if errorlevel 1 exit /b 1

cl /nologo /EHsc /W4 /MT /DUNICODE /D_UNICODE ^
  /D_WIN32_WINNT=0x0501 /DWINVER=0x0501 ^
  /Fo"build\WindowWatch.obj" /Fe"build\WindowWatch.exe" ^
  "src\tools\WindowWatch.cpp" ^
  user32.lib /link /SUBSYSTEM:WINDOWS,5.01
if errorlevel 1 (
  echo ERROR: WindowWatch build failed.
  exit /b 1
)

cl /nologo /EHsc /W4 /MT /D_WIN32_WINNT=0x0501 /DWINVER=0x0501 ^
  /Fo"build\ClippyProbe.obj" /Fe"build\ClippyProbe.exe" ^
  "src\tools\ClippyProbe.cpp" ^
  ole32.lib oleaut32.lib /link /SUBSYSTEM:CONSOLE,5.01
if errorlevel 1 (
  echo ERROR: build failed.
  exit /b 1
)

cl /nologo /EHsc /W4 /MT /LD /DUNICODE /D_UNICODE ^
  /D_WIN32_WINNT=0x0501 /DWINVER=0x0501 ^
  /Fo"build\ClippyShim.obj" /Fe"build\ClippyShim.dll" ^
  "src\addin\ClippyShim.cpp" ^
  ole32.lib oleaut32.lib user32.lib advapi32.lib shlwapi.lib ws2_32.lib ^
  "src\addin\ClippyShim.def" ^
  /link /SUBSYSTEM:WINDOWS,5.01
if errorlevel 1 (
  echo ERROR: ClippyShim build failed.
  exit /b 1
)

echo OK: built C:\clippy\build\WindowWatch.exe
echo OK: built C:\clippy\build\ClippyProbe.exe
echo OK: built C:\clippy\build\ClippyShim.dll
exit /b 0
