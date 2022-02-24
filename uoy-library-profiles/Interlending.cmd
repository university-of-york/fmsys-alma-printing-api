@echo off
title %~n0 - Alma Printing With Powershell - leave this window open or minimised to continue printing
powershell -NoLogo -WindowStyle Minimized -command "& { . ..\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId \""19195349880001381\"" -localPrinterName \""PUSH_ITSPRN0705 [Harry Fairhurst - Information Services LFA/ LFA023](Mobility)\"" }"