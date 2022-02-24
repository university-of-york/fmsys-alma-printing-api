@echo off
title %~n0 - Alma Printing With Powershell
powershell -NoLogo -WindowStyle Minimized -command "& { . ..\FetchAlmaPrint.ps1;Fetch-Jobs -printerId \""19195349880001381\"" -localPrinterName \""PUSH_ITSPRN0705 [Harry Fairhurst - Information Services LFA/ LFA023](Mobility)\"" }"