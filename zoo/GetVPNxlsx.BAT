 echo off
rem Обработка файлов отчетов VPN
rem (C)2023 Aleksey Eremin
rem

SET OutDIR=D:\MyProg\PFO_Forma_otcheta_Android_VPN
set MYJAR=D:\Users\AE\workspace\xls2xls\out\artifacts\xls2xls.jar

set MyDir=D:\Users\AE\workspace\xls2xls\zoo

set KARTA=%MyDir%\karta

set srcXlsx=%MyDir%\Out.xlsx
set srcIxlsx=%MyDir%\OutI.xlsx

set Outxlsx=D:\TMP\out.xlsx
set OutIxlsx=D:\TMP\outI.xlsx

rem Отчетное имя файла
set REPXLSX="D:\TMP\!Форма отчета Android v4 +7 VPN_"

set tmplog=%temp%\xls2xls.tmp

set d=%DATE%
set m=%d:~3,2%
set y=%d:~6,4%
rem set y=%d:~8,4%
set d=%d:~0,2%
set dt=%y%-%m%-%d%

rem архивный каталог сегодняшнего дня
SET ArchDir=%OutDIR%\Archive\%dt%
IF NOT exist %ArchDir% (
  rem если нет, значит первый запуск делаем архивный каталог
  mkdir %ArchDir%
)

IF NOT exist %Outxlsx% (
  rem если нет, значит первый запуск сегодня
  rem копируем исходный файл
  copy  %srcXlsx%   %Outxlsx%
  copy  %srcIxlsx%  %OutIxlsx%
)

rem сделаем бакап выходного файла
rem call :makeBak

FOR %%I IN (02 12 13 16 18 21 43 52 56 58 59 63 64 73)  DO CALL :mkwork %%I%
rem FOR %%I IN (02 13)  DO CALL :mkwork %%I%

FC /b %SRCXLSX% %Outxlsx% >>%tmplog%

IF %ERRORLEVEL% GTR 0 (
  echo Итоговый файл изменился
  copy %Outxlsx% %REPXLSX%%dt%.xlsx
)

GOTO :EOF


:mkwork
rem обработка регионального файла
echo ========================================================
for %%I in (%OutDIR%\%1_*.xlsx) do call :wowo %1 "%%I%"
goto :eof


:wowo
echo Region %1 file exist - %2
java -jar %MYJAR% %KARTA%%1.txt  %2  "%Outxlsx%"
java -jar %MYJAR% %KARTA%%1.txt  %2  "%OutIxlsx%" 
rem   move "%%I%"  %ArchDir%\   
goto :eof


:makeBak
rem определим размер выходного файла и если он большой - сделаем копию
set RF=0
for %%I in (%OUTXLSX%) do  set RF=%%~zI%
IF %RF% GTR 20000 copy %OUTXLSX% %OUTXLSX%.bak
goto :EOF


