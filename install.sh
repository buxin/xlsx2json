#!/bin/bash

# xlsx2json macOS 安装脚本
# 自动安装和配置可执行程序

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_NAME="xlsx2json"

echo "🚀 xlsx2json macOS 安装程序"
echo "=================================="

# 检查 Node.js
if ! command -v node &> /dev/null; then
    echo "❌ 未找到 Node.js"
    echo "请先安装 Node.js: https://nodejs.org/"
    echo "或者使用 Homebrew: brew install node"
    exit 1
fi

echo "✅ Node.js 版本: $(node --version)"

# 检查 npm
if ! command -v npm &> /dev/null; then
    echo "❌ 未找到 npm"
    exit 1
fi

echo "✅ npm 版本: $(npm --version)"

# 安装依赖
echo ""
echo "📦 正在安装依赖..."
cd "$PROJECT_DIR"
npm install

if [ $? -ne 0 ]; then
    echo "❌ 依赖安装失败"
    exit 1
fi

echo "✅ 依赖安装完成"

# 设置可执行权限
echo ""
echo "🔧 设置可执行权限..."
chmod +x "$PROJECT_DIR/bin/xlsx2json"
chmod +x "$PROJECT_DIR/start.command"
chmod +x "$PROJECT_DIR/create-app.sh"

echo "✅ 权限设置完成"

# 创建全局链接（可选）
echo ""
echo "🔗 是否创建全局命令链接？(y/n)"
read -r response
if [[ "$response" =~ ^[Yy]$ ]]; then
    # 检查是否有权限创建全局链接
    if [ -w "/usr/local/bin" ]; then
        ln -sf "$PROJECT_DIR/bin/xlsx2json" "/usr/local/bin/xlsx2json"
        echo "✅ 全局命令创建成功: xlsx2json"
    else
        echo "⚠️  需要管理员权限创建全局命令"
        echo "请运行: sudo ln -sf \"$PROJECT_DIR/bin/xlsx2json\" /usr/local/bin/xlsx2json"
    fi
fi

echo ""
echo "🎉 安装完成！"
echo ""
echo "使用方法:"
echo "1. 命令行使用:"
echo "   ./bin/xlsx2json --help"
echo "   ./bin/xlsx2json convert <文件路径>"
echo "   ./bin/xlsx2json batch <目录路径>"
echo ""
echo "2. 图形界面使用:"
echo "   双击 start.command 文件"
echo ""
echo "3. 创建 macOS 应用程序:"
echo "   ./create-app.sh"
echo ""
echo "4. 查看配置:"
echo "   ./bin/xlsx2json config"
