@echo off
chcp 65001 >nul
echo xlsx2json 一键转换工具
echo.

REM 检查Node.js
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Node.js
    exit /b 1
)

REM 检查xlsx目录
if not exist "xlsx\" (
    echo 错误: 未找到xlsx目录
    exit /b 1
)

REM 检查xlsx文件
dir /b "xlsx\*.xlsx" >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: xlsx目录中没有找到.xlsx文件
    exit /b 1
)

REM 安装依赖（如果需要）
if not exist "node_modules\" (
    echo 正在安装依赖...
    npm install
)

REM 执行转换
echo 开始转换...
node index.js batch xlsx/

echo 转换完成！输出文件在 json\ 目录中
