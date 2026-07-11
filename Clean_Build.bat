@echo off
setlocal
echo ============================================
echo  Clean Delphi / C++Builder / VB6 artifacts
echo ============================================

:: --- Xóa folder build cache ---
FOR /D /R %%X IN (__history)   DO RD /S /Q "%%X" 2>nul
FOR /D /R %%X IN (__recovery)  DO RD /S /Q "%%X" 2>nul
FOR /D /R %%X IN (dcu)         DO RD /S /Q "%%X" 2>nul
FOR /D /R %%X IN (dcu.win32)   DO RD /S /Q "%%X" 2>nul
FOR /D /R %%X IN (dcu.android) DO RD /S /Q "%%X" 2>nul
FOR /D /R %%X IN (logs)        DO RD /S /Q "%%X" 2>nul

:: --- Delphi / C++Builder ---
del /f /q /s *.dcu
del /f /q /s *.dcp
del /f /q /s *.drc
del /f /q /s *.map
del /f /q /s *.ddp
del /f /q /s *.local
del /f /q /s *.identcache
del /f /q /s *.tvsconfig
del /f /q /s *.rsm
del /f /q /s *.stat
del /f /q /s *.skincfg
del /f /q /s *.dof

:: Backup tự động của IDE (~pas, ~dfm, ~dpr, ~bpg...)
del /f /q /s *.~*
del /f /q /s /A:H *.~*

:: --- C++ Builder object files ---
del /f /q /s *.o
del /f /q /s *.or
del /f /q /s *.obj
del /f /q /s *.ppu
del /f /q /s *.compiled

:: --- VB6 ---
del /f /q /s *.vbw
del /f /q /s *.DCA
del /f /q /s *.SCC
del /f /q /s *.exp
del /f /q /s *.lib

:: --- File rác chung ---
del /f /q /s *.tmp
del /f /q /s *.log
del /f /q /s *.bak
del /f /q /s thumbs.db
del /f /q /s descript.ion
del /f /q /s bugreport.txt

echo Done.
endlocal