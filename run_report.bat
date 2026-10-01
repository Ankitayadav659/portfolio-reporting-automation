@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo Looking for files in this folder...
echo.

set MASTER=
set INCOME=
set SCHEMEBM=
set BROADBM=
set SIP=
set XSIPFILE=
set MYSIPFILE=

for %%f in ("PortfolioValuation*.xlsx") do set MASTER=%%~nxf
for %%f in ("Income*.xlsx") do set INCOME=%%~nxf
for %%f in ("Scheme*enchmark*.pdf") do set SCHEMEBM=%%~nxf
for %%f in ("Broad*enchmark*.pdf") do set BROADBM=%%~nxf

rem ---- SIP file ----
rem  XSIP file   (XSIP_REG_RPT_...)  -> only Status = Active SIPs are used
rem  My SIP file (My SIP's-NAME-...) -> current month amount (if 0 -> last month)
rem  ALL XSIP and My SIP files in this folder are used together (one per family
rem  member is fine). Same SIP in both: XSIP wins; else My SIP is used.

for %%f in ("XSIP*.xls*" "XSIP*.csv" "*XSIP*.xls*" "*XSIP*.csv") do if "!XSIPFILE!"=="" set XSIPFILE=%%~nxf
for %%f in ("My SIP*.xls*" "My SIP*.csv" "My_SIP*.xls*" "MySIP*.xls*" "My-SIP*.xls*") do if "!MYSIPFILE!"=="" set MYSIPFILE=%%~nxf

if not "!XSIPFILE!!MYSIPFILE!"=="" (
    set SIP=!XSIPFILE!!MYSIPFILE!
    set SIPRULE=XSIP = Active only, My SIP = current month ^(if 0, last month^), matched per member
) else (
    for %%f in ("*SIP*Registration*.xls*" "*SIP*Report*.xls*" "*SIP*.xls*" "*SIP*.csv") do if "!SIP!"=="" set SIP=%%~nxf
    if not "!SIP!"=="" set SIPRULE=Other SIP file - only Status = Active SIPs
)

echo Master (Portfolio Valuation) file : %MASTER%
echo Income / Brokerage file           : %INCOME%
echo Scheme Benchmark PDF              : %SCHEMEBM%
echo Broad Benchmark PDF               : %BROADBM%
if "!SIP!"=="" (
    echo SIP file                          : none found - report will say NO SIP
) else (
    if not "!XSIPFILE!!MYSIPFILE!"=="" (
        echo SIP file^(s^)                       : every XSIP / My SIP file in this folder:
        for %%f in ("*XSIP*.xls*" "*XSIP*.csv" "My SIP*.xls*" "My SIP*.csv" "My_SIP*.xls*" "MySIP*.xls*" "My-SIP*.xls*") do echo                                        %%~nxf
    ) else (
        echo SIP file                          : !SIP!
    )
    echo SIP rule                          : !SIPRULE!
)
echo.

if "%MASTER%"=="" (
    echo ERROR: Could not find a PortfolioValuation*.xlsx file in this folder.
    pause
    exit /b
)
if "%INCOME%"=="" (
    echo ERROR: Could not find an Income*.xlsx file in this folder.
    pause
    exit /b
)
if "%SCHEMEBM%"=="" (
    echo ERROR: Could not find a Scheme benchmark PDF in this folder.
    pause
    exit /b
)
if "%BROADBM%"=="" (
    echo ERROR: Could not find a Broad benchmark PDF in this folder.
    pause
    exit /b
)

echo Running report builder...
echo.

if "!SIP!"=="" (
    py build_report.py --master "%MASTER%" --income "%INCOME%" --scheme-bm "%SCHEMEBM%" --broad-bm "%BROADBM%" --no-sip --out "Portfolio_Report_OUTPUT.xlsx"
) else if not "!XSIPFILE!!MYSIPFILE!"=="" (
    rem no --sip: build_report.py picks up EVERY XSIP and My SIP file in this folder by itself
    py build_report.py --master "%MASTER%" --income "%INCOME%" --scheme-bm "%SCHEMEBM%" --broad-bm "%BROADBM%" --out "Portfolio_Report_OUTPUT.xlsx"
) else (
    py build_report.py --master "%MASTER%" --income "%INCOME%" --scheme-bm "%SCHEMEBM%" --broad-bm "%BROADBM%" --sip "!SIP!" --out "Portfolio_Report_OUTPUT.xlsx"
)

echo.
echo ============================================
echo Finished. Check Portfolio_Report_OUTPUT.xlsx
echo in this same folder.
echo ============================================
pause
