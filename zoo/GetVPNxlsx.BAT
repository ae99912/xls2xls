  echo off
rem ��ࠡ�⪠ 䠩��� ���⮢ VPN
rem (C)2023 Aleksey Eremin
rem

SET SrcDIR=D:\MyProg\PFO_Forma_otcheta_Android_VPN
set MYJAR=D:\Users\AE\workspace\xls2xls\out\artifacts\xls2xls.jar
set KARTA=D:\Users\AE\workspace\xls2xls\zoo\karta
set OUTXLSX=D:\TMP\out.xlsx

set d=%DATE%
set m=%d:~3,2%
set y=%d:~6,4%
rem set y=%d:~8,4%
set d=%d:~0,2%
set dt=%y%-%m%-%d%

rem ��娢�� ��⠫��
SET ArchDir=%SrcDir%\Archive\%dt%
IF NOT exist %ArchDir% mkdir %ArchDir%


FOR %%I IN (02 12 13 16 18 21 43 52 56 58 59 63 64 73)  DO CALL :mkwork %%I%
rem FOR %%I IN (02 13)  DO CALL :mkwork %%I%

GOTO :EOF

:mkwork
echo ========================================================
rem echo Region %1
set FIL="%SrcDIR%\%1_��� ��ଠ ���� Android VPN"
set XLSX=0
for %%I in (%FIL%*.xlsx) do set XLSX=%%I%
rem for %%I in (%FIL%*.xlsx) do echo %%I%
rem  echo %XLSX%
rem echo %XLSX%
IF exist "%XLSX%" goto w1
echo No file for Region %1
goto :EOF
:w1
echo Region %1 file exist - %XLSX%
echo Begin work
java -jar %MYJAR% %KARTA%%1.txt "%XLSX%"  "%OUTXLSX%"

move "%XLSX%" %ArchDir%\

goto :EOF

