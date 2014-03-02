@echo off
cd %~dp0

copy /Y .\Templates\Normal.dot .\bin\Normal.dot
copy /Y .\AddIns\TanabeMacros.xla  .\bin\TanabeMacros.xla

cscript //nologo vbac.wsf decombine
