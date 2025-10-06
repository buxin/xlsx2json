#!/bin/bash

# xlsx2json macOS 启动脚本
# 双击运行此脚本即可启动程序

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$SCRIPT_DIR"

# 切换到项目目录
cd "$PROJECT_DIR"

# 检查 Node.js 是否安装
if ! command -v node &> /dev/null; then
    echo "❌ 错误: 未找到 Node.js"
    echo "请先安装 Node.js: https://nodejs.org/"
    echo "按任意键退出..."
    read -n 1
    exit 1
fi

# 检查依赖是否安装
if [ ! -d "node_modules" ]; then
    echo "📦 正在安装依赖..."
    npm install
    if [ $? -ne 0 ]; then
        echo "❌ 依赖安装失败"
        echo "按任意键退出..."
        read -n 1
        exit 1
    fi
fi

# 启动程序
echo "🚀 启动 xlsx2json..."
echo "=================================="
node index.js "$@"

# 保持窗口打开
echo ""
echo "=================================="
echo "程序执行完成"
echo "按任意键退出..."
read -n 1
