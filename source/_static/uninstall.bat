@echo off
setlocal

if not exist "venv" (
    echo 未检测到虚拟环境 venv，无法执行卸载。
    pause
    exit /b 1
)

echo 正在激活虚拟环境并卸载知心插件...
call venv\Scripts\activate.bat
zhixinpy uninstall

echo.
echo 卸载完成！
pause