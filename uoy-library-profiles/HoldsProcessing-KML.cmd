@echo off
title %~n0 - Alma Printing With Powershell - leave this window open or minimised to continue printing
powershell -NoLogo -WindowStyle Minimized -command "& { . ..\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId \""993537480001381\"" -localPrinterName \""EPSON TM-T88III Receipt\"" }"





