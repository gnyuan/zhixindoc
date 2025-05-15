@echo off
setlocal enabledelayedexpansion

REM 检查是否存在 Excel 正在运行
tasklist | findstr /I "EXCEL.EXE" >nul
if %errorlevel%==0 (
echo.
echo 检测到 Excel 正在运行，请先关闭 Excel 后再继续安装或升级知心插件。
pause
exit /b 1
)

REM 检查 Excel 版本（使用注册表）
set EXCEL_OK=0
set BITNESS=32
set version=未知
set EXCEL_DETECTED=0

REM 使用更可靠的方法检测 Excel 版本
reg query "HKCR\Excel.Application\CurVer" > tmp_excel_curver.txt 2>nul
if %errorlevel%==0 (
    for /f "tokens=3" %%i in ('type tmp_excel_curver.txt') do (
        set excel_ver=%%i
        set excel_ver=!excel_ver:Excel.Application.=!
        set EXCEL_DETECTED=1
    )
    del tmp_excel_curver.txt
)

REM 检测 Excel 位数
if "!EXCEL_DETECTED!"=="1" (
    REM 检查64位 Excel
    reg query "HKLM\SOFTWARE\Microsoft\Office\!excel_ver!.0\Excel\InstallRoot" /v Path >nul 2>nul
    if !errorlevel!==0 (
        set BITNESS=64
    ) else (
        REM 检查32位 Excel
        reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\!excel_ver!.0\Excel\InstallRoot" /v Path >nul 2>nul
        if !errorlevel!==0 (
            set BITNESS=32
        )
    )
    
    REM 获取详细版本号
    reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ClientVersionToReport > tmp_excel_version.txt 2>nul
    findstr /R "^.*Version.*" tmp_excel_version.txt >nul
    if !errorlevel!==0 (
        for /f "tokens=3" %%i in ('findstr "Version" tmp_excel_version.txt') do (
            set version=%%i
        )
        del tmp_excel_version.txt
    )
    
    REM 检查是否大于等于 2013 (版本号 15)
    if !excel_ver! GEQ 15 (
        set EXCEL_OK=1
    )
)

echo.
echo 检测结果：
if "!EXCEL_DETECTED!"=="1" (
    echo 检测到 Excel !BITNESS! 位版本: !version! ^(Office 版本号: !excel_ver!^)
) else (
    echo 未能检测到 Excel 版本信息
)

if "!EXCEL_OK!" NEQ "1" (
    echo.
    if "!EXCEL_DETECTED!"=="1" (
        echo Excel 版本低于 2013 版本要求，不支持安装知心插件。
    ) else (
        echo 无法确定 Excel 版本，不支持安装知心插件。
    )
    echo 请确保已安装 Office 2013 或以上版本。
    pause
    exit /b 1
)

if "!BITNESS!" NEQ "64" (
    echo.
    echo 检测到 Excel 为 32 位版本，知心插件需要 64 位 Excel。
    echo 请安装 64 位 Office。
    pause
    exit /b 1
)

REM 检查 Python 版本
set PYTHON_VERSION=未知
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
python --version 2>nul | findstr /R "^Python 3\.1[12]" >nul
if %errorlevel% neq 0 (
echo.
echo 检测到 Python 版本: %PYTHON_VERSION%
echo 请先安装 Python 3.11 或 3.12，并确保 python 命令可用。
pause
exit /b 1
) else (
echo 检测到 Python 版本: %PYTHON_VERSION% - 符合要求
)

REM 创建虚拟环境（如果未存在）
if not exist "venv" (
echo 正在创建虚拟环境...
python -m venv venv
) else (
echo 已检测到 venv 虚拟环境，跳过创建。
)

REM 激活并安装或升级 zhixinpy
echo 正在激活虚拟环境并安装/升级 知心插件...
call venv\Scripts\activate.bat
pip install --upgrade zhixinpy -i https://mirrors.aliyun.com/pypi/simple/
zhixinpy install

REM 获取并显示当前安装的zhixinpy版本
for /f "tokens=2" %%i in ('pip show zhixinpy ^| findstr Version') do set ZHIXIN_VERSION=%%i

echo.
echo 知心插件安装或升级完成！当前版本：%ZHIXIN_VERSION%
pause