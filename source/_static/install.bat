@echo off
setlocal enabledelayedexpansion

REM ����Ƿ���� Excel ��������
tasklist | findstr /I "EXCEL.EXE" >nul
if %errorlevel%==0 (
echo.
echo ��⵽ Excel �������У����ȹر� Excel ���ټ�����װ������֪�Ĳ����
pause
exit /b 1
)

REM ��� Excel �汾��ʹ��ע���
set EXCEL_OK=0
set BITNESS=32
set version=δ֪
set EXCEL_DETECTED=0

REM ʹ�ø��ɿ��ķ������ Excel �汾
reg query "HKCR\Excel.Application\CurVer" > tmp_excel_curver.txt 2>nul
if %errorlevel%==0 (
    for /f "tokens=3" %%i in ('type tmp_excel_curver.txt') do (
        set excel_ver=%%i
        set excel_ver=!excel_ver:Excel.Application.=!
        set EXCEL_DETECTED=1
    )
    del tmp_excel_curver.txt
)

REM ��� Excel λ��
if "!EXCEL_DETECTED!"=="1" (
    REM ���64λ Excel
    reg query "HKLM\SOFTWARE\Microsoft\Office\!excel_ver!.0\Excel\InstallRoot" /v Path >nul 2>nul
    if !errorlevel!==0 (
        set BITNESS=64
    ) else (
        REM ���32λ Excel
        reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\!excel_ver!.0\Excel\InstallRoot" /v Path >nul 2>nul
        if !errorlevel!==0 (
            set BITNESS=32
        )
    )
    
    REM ��ȡ��ϸ�汾��
    reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ClientVersionToReport > tmp_excel_version.txt 2>nul
    findstr /R "^.*Version.*" tmp_excel_version.txt >nul
    if !errorlevel!==0 (
        for /f "tokens=3" %%i in ('findstr "Version" tmp_excel_version.txt') do (
            set version=%%i
        )
        del tmp_excel_version.txt
    )
    
    REM ����Ƿ���ڵ��� 2013 (�汾�� 15)
    if !excel_ver! GEQ 15 (
        set EXCEL_OK=1
    )
)

echo.
echo �������
if "!EXCEL_DETECTED!"=="1" (
    echo ��⵽ Excel !BITNESS! λ�汾: !version! ^(Office �汾��: !excel_ver!^)
) else (
    echo δ�ܼ�⵽ Excel �汾��Ϣ
)

if "!EXCEL_OK!" NEQ "1" (
    echo.
    if "!EXCEL_DETECTED!"=="1" (
        echo Excel �汾���� 2013 �汾Ҫ�󣬲�֧�ְ�װ֪�Ĳ����
    ) else (
        echo �޷�ȷ�� Excel �汾����֧�ְ�װ֪�Ĳ����
    )
    echo ��ȷ���Ѱ�װ Office 2013 �����ϰ汾��
    pause
    exit /b 1
)

if "!BITNESS!" NEQ "64" (
    echo.
    echo ��⵽ Excel Ϊ 32 λ�汾��֪�Ĳ����Ҫ 64 λ Excel��
    echo �밲װ 64 λ Office��
    pause
    exit /b 1
)

REM ��� Python �汾
set PYTHON_VERSION=δ֪
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
python --version 2>nul | findstr /R "^Python 3\.1[12]" >nul
if %errorlevel% neq 0 (
echo.
echo ��⵽ Python �汾: %PYTHON_VERSION%
echo ���Ȱ�װ Python 3.11 �� 3.12����ȷ�� python ������á�
pause
exit /b 1
) else (
echo ��⵽ Python �汾: %PYTHON_VERSION% - ����Ҫ��
)

REM �������⻷�������δ���ڣ�
if not exist "venv" (
echo ���ڴ������⻷��...
python -m venv venv
) else (
echo �Ѽ�⵽ venv ���⻷��������������
)

REM �����װ������ zhixinpy
echo ���ڼ������⻷������װ/���� ֪�Ĳ��...
call venv\Scripts\activate.bat
pip install --upgrade zhixinpy -i https://mirrors.aliyun.com/pypi/simple/
zhixinpy install

REM ��ȡ����ʾ��ǰ��װ��zhixinpy�汾
for /f "tokens=2" %%i in ('pip show zhixinpy ^| findstr Version') do set ZHIXIN_VERSION=%%i

echo.
echo ֪�Ĳ����װ��������ɣ���ǰ�汾��%ZHIXIN_VERSION%
pause