@echo off
:: 检查管理员权限
NET SESSION >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo 已以管理员权限运行
) ELSE (
    echo 请求管理员权限...
    powershell -Command "Start-Process -FilePath '%~dpnx0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo             软件著作权源代码生成器 - Nuitka打包工具
echo ============================================================

:: 检查Python环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python
    echo 请安装Python并确保其添加到PATH中
    pause
    exit /b 1
)

:: 检查Nuitka
python -c "import nuitka" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装Nuitka...
    pip install nuitka
    if %errorlevel% neq 0 (
        echo 安装Nuitka失败
        pause
        exit /b 1
    )
)

:: 运行构建脚本
echo 开始构建...
python nuitka_build.py

:: 显示完成消息
if %errorlevel% equ 0 (
    echo ============================================================
    echo 构建完成！请查看dist目录
    echo ============================================================
) else (
    echo ============================================================
    echo 构建失败，请检查错误信息
    echo ============================================================
)

pause 