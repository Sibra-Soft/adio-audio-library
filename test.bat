@echo off

setlocal EnableExtensions EnableDelayedExpansion

"./tests/AdioLibraryTests.exe"

IF ERRORLEVEL 1 (
	EXIT /B 1
) ELSE (
	EXIT /B 0
)