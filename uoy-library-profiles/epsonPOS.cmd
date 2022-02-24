@echo off
title %~n0 - Alma Printing With Powershell
powershell -NoLogo -WindowStyle Minimized -command "& { . ..\FetchAlmaPrint.ps1;Fetch-Jobs -printerId \""848838010001381\"" -localPrinterName \""EPSON TM-T88III Receipt\"" }"





