@echo off
chcp 65001 >nul
echo xlsx2json 一键转换工具
echo.

REM 显示当前配置
echo 📋 当前配置:
node index.js config
echo.

REM 检查Node.js
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Node.js
    exit /b 1
)

REM 检查输入目录（从配置文件读取）
node -e "const config = JSON.parse(require('fs').readFileSync('config.json', 'utf8')); console.log(config.inputDir);" > temp_input_dir.txt
set /p INPUT_DIR=<temp_input_dir.txt
del temp_input_dir.txt

if not exist "%INPUT_DIR%" (
    echo 错误: 未找到输入目录: %INPUT_DIR%
    exit /b 1
)

REM 检查xlsx文件
dir /b "%INPUT_DIR%\*.xlsx" >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 输入目录中没有找到.xlsx文件: %INPUT_DIR%
    exit /b 1
)

REM 安装依赖（如果需要）
if not exist "node_modules\" (
    echo 正在安装依赖...
    npm install
)

REM 执行转换
echo 开始转换...
node index.js auto

echo 转换完成！输出文件在 json\ 目录中
