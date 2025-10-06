#!/bin/bash

# xlsx2json macOS 应用程序启动脚本
# 创建 macOS 应用程序包

APP_NAME="xlsx2json"
APP_DIR="$HOME/Applications/${APP_NAME}.app"
CONTENTS_DIR="${APP_DIR}/Contents"
MACOS_DIR="${CONTENTS_DIR}/MacOS"
RESOURCES_DIR="${CONTENTS_DIR}/Resources"

# 获取当前脚本目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$SCRIPT_DIR"

echo "🔧 创建 macOS 应用程序包..."

# 创建应用程序目录结构
mkdir -p "$MACOS_DIR"
mkdir -p "$RESOURCES_DIR"

# 复制项目文件到应用程序包
cp -r "$PROJECT_DIR"/* "$RESOURCES_DIR/"

# 创建 Info.plist
cat > "${CONTENTS_DIR}/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>${APP_NAME}</string>
    <key>CFBundleIdentifier</key>
    <string>com.xlsx2json.app</string>
    <key>CFBundleName</key>
    <string>${APP_NAME}</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleSignature</key>
    <string>????</string>
    <key>CFBundleDisplayName</key>
    <string>Excel to JSON Converter</string>
    <key>CFBundleIconFile</key>
    <string>icon</string>
</dict>
</plist>
EOF

# 创建可执行文件
cat > "${MACOS_DIR}/${APP_NAME}" << 'EOF'
#!/bin/bash

# 获取应用程序包路径
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
RESOURCES_DIR="${APP_DIR}/Contents/Resources"

# 切换到资源目录
cd "$RESOURCES_DIR"

# 检查 Node.js - 使用完整路径
NODE_PATH=$(which node)
if [ -z "$NODE_PATH" ]; then
    osascript -e 'display dialog "未找到 Node.js，请先安装 Node.js" with title "xlsx2json" buttons {"确定"} default button "确定"'
    exit 1
fi

# 安装依赖（如果需要）
if [ ! -d "node_modules" ]; then
    osascript -e 'display dialog "正在安装依赖，请稍候..." with title "xlsx2json" buttons {"确定"} default button "确定"'
    npm install --silent
fi

# 启动程序 - 使用完整路径
"$NODE_PATH" index.js "$@"
EOF

# 设置可执行权限
chmod +x "${MACOS_DIR}/${APP_NAME}"

echo "✅ macOS 应用程序包创建完成！"
echo "📁 位置: $APP_DIR"
echo ""
echo "使用方法:"
echo "1. 在 Finder 中打开 $HOME/Applications/"
echo "2. 双击 ${APP_NAME}.app 运行"
echo "3. 或者将应用程序拖拽到 Dock 中"
