@echo off
setlocal

if not exist "venv" (
    echo δ��⵽���⻷�� venv���޷�ִ��ж�ء�
    pause
    exit /b 1
)

echo ���ڼ������⻷����ж��֪�Ĳ��...
call venv\Scripts\activate.bat
zhixinpy uninstall

echo.
echo ж����ɣ�
pause