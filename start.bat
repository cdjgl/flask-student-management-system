@echo off
chcp 65001 >nul
title 学生管理系统
cd /d "%~dp0"

echo ============================================
echo   Flask + MySQL 学生管理系统 启动脚本
echo ============================================
echo.

rem ---- 1. 检查 Python ----
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+ 并勾选 "Add to PATH"。
    echo        下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

rem ---- 2. 检查 .env 配置（不含任何密钥，密码请填写在 .env 中）----
if not exist ".env" (
    if exist ".env.example" (
        copy ".env.example" ".env" >nul
        echo [提示] 已根据 .env.example 生成 .env。
        echo        请在打开的编辑器中填写 DB_PASSWORD 后保存、关闭，再重新运行本脚本。
        notepad .env
        pause
        exit /b 1
    ) else (
        echo [错误] 缺少 .env 与 .env.example，请确认文件完整。
        pause
        exit /b 1
    )
)

rem ---- 3. 检查依赖，缺失则自动安装 ----
python -c "import flask, pymysql, dbutils, dotenv, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [提示] 首次运行，正在安装依赖（requirements.txt）...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络后重试，或手动执行: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

rem ---- 4. 启动应用 ----
echo.
echo [启动] 访问地址: http://127.0.0.1:5000
echo [账号] 默认管理员 admin / admin123（首次登录后请立即修改密码）
echo [说明] 请确保本机 MySQL 服务已启动、.env 中数据库配置正确。
echo.
python app.py
pause
