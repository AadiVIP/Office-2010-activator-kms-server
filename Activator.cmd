@echo off
title KMS Activator for Any Version of MS Office 2010/365 &cls&echo ************************************************* &echo Copyright: Youtube: 2013Electronics and Computers&echo ************************************************* &echo.&echo Supported MS Office 2010 products:&echo Microsoft Office 365 &echo Home and Student &echo Home and Business &echo Standard &echo Professional &echo Professional Plus &echo.&echo Microsoft Office activation...
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
cscript //nologo ospp.vbs /unpkey:B9HB6 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&
cscript //nologo ospp.vbs /inpkey:2KKDC-67TT9-4XT2F-2MG99-B9HB6 >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms.digiboy.ir
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms.chinancce.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ************************************************* &echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ************************************************* &echo.&choice /n /c YN /m "Do you want to restart your PC now [Y,N]?" & if errorlevel 2 exit) || (echo The connection to the server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
shutdown.exe /r /t 00
:notsupported
echo.&echo ************************************************* &echo Incorrect version of MS Office &echo Make sure that you use MS Office 2010/365 version.
:halt
pause >nul
