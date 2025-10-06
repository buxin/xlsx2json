# xlsx2json - macOS 可执行程序

一个将 Excel 文件转换为 JSON 格式的 macOS 可执行程序。

## 🚀 快速开始

### 1. 安装程序
```bash
# 运行安装脚本
./install.sh
```

### 2. 使用方法

#### 命令行使用
```bash
# 查看帮助
./bin/xlsx2json --help

# 转换单个文件
./bin/xlsx2json convert <xlsx文件路径>

# 批量转换
./bin/xlsx2json batch <目录路径>

# 使用配置文件自动转换
./bin/xlsx2json auto

# 查看当前配置
./bin/xlsx2json config
```

#### 图形界面使用
```bash
# 双击运行
open start.command
```

#### 创建 macOS 应用程序
```bash
# 创建 .app 应用程序包
./create-app.sh
```

## 📁 文件说明

- `bin/xlsx2json` - 可执行程序入口
- `start.command` - macOS 双击启动脚本
- `install.sh` - 自动安装脚本
- `create-app.sh` - 创建 macOS 应用程序包脚本
- `config.json` - 配置文件
- `index.js` - 主程序文件

## ⚙️ 配置

编辑 `config.json` 文件来配置程序：

```json
{
  "inputDir": "xlsx",
  "outputDir": "json", 
  "skipFirstRow": true,
  "outputMapping": {
    "建筑表.xlsx": "building.json",
    "建筑列表.xlsx": "store.json",
    "怪物属性.xlsx": "monster.json"
  }
}
```

## 🔧 系统要求

- macOS 10.12 或更高版本
- Node.js 12.0 或更高版本
- npm 6.0 或更高版本

## 📦 安装依赖

程序会自动安装以下依赖：
- `xlsx` - Excel 文件处理
- `commander` - 命令行参数解析

## 🎯 功能特性

- ✅ 支持单个文件转换
- ✅ 支持批量目录转换
- ✅ 配置文件支持
- ✅ 文件名映射
- ✅ 跳过第一行（中文注释）
- ✅ macOS 原生支持
- ✅ 图形界面启动
- ✅ 可创建 .app 应用程序包

## 🚨 故障排除

### 权限问题
```bash
# 如果遇到权限问题，运行：
chmod +x bin/xlsx2json
chmod +x start.command
chmod +x install.sh
chmod +x create-app.sh
```

### Node.js 未找到
```bash
# 安装 Node.js
brew install node
# 或者从官网下载：https://nodejs.org/
```

### 依赖安装失败
```bash
# 手动安装依赖
npm install
```

## 📞 支持

如果遇到问题，请检查：
1. Node.js 是否正确安装
2. 文件权限是否正确设置
3. 配置文件格式是否正确
4. 输入文件是否存在

## 📄 许可证

MIT License
