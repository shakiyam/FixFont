@echo off
powershell Invoke-ScriptAnalyzer FixFont.ps1
powershell -ExecutionPolicy Bypass -command "&ps2exe.ps1 -inputFile FixFont.ps1 -outputFile FixFont.exe"
del FixFont.exe.config
