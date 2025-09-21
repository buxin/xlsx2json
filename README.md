# xlsx2json

一个简单易用的Excel文件转JSON工具，使用Node.js开发。

## 功能特点

- ✅ 支持单个xlsx文件转换
- ✅ 支持批量转换目录中的所有xlsx文件
- ✅ 自动跳过第一行（中文注释）
- ✅ 支持自定义输出路径
- ✅ 命令行友好界面
- ✅ 错误处理和进度显示

## 安装

```bash
npm install
```

## 使用方法

### 🚀 一键转换（推荐）

**Windows批处理文件：**
```cmd
# 双击运行或在命令行执行
convert.bat
```

**PowerShell脚本：**
```powershell
# 在PowerShell中执行
.\convert.ps1
```

### 📝 命令行使用

**1. 转换单个文件**
```bash
# 基本用法（输出到默认位置）
node index.js convert xlsx/建筑表.xlsx

# 指定输出文件路径
node index.js convert xlsx/建筑表.xlsx -o output/result.json

# 不跳过第一行
node index.js convert xlsx/建筑表.xlsx --no-skip-header
```

**2. 批量转换**
```bash
# 转换目录中所有xlsx文件
node index.js batch xlsx/

# 指定输出目录
node index.js batch xlsx/ -o output/
```

**3. 查看帮助**
```bash
node index.js --help
node index.js convert --help
node index.js batch --help
```

## 输出格式

转换后的JSON文件格式：

```json
[
  {
    "列名1": "值1",
    "列名2": "值2",
    "列名3": "值3"
  },
  {
    "列名1": "值4",
    "列名2": "值5",
    "列名3": "值6"
  }
]
```

## 注意事项

- 默认跳过第一行作为中文注释
- 空单元格会被转换为空字符串
- 输出目录会自动创建
- 支持中文文件名和路径

## 项目结构

```
xlsx2json/
├── package.json          # 项目配置
├── index.js             # 主程序
├── convert.bat          # Windows一键转换脚本
├── convert.ps1          # PowerShell一键转换脚本
├── convert_all.bat      # 详细版批处理脚本
├── README.md            # 说明文档
├── xlsx/                # 输入xlsx文件目录
│   └── 建筑表.xlsx
└── json/                # 输出JSON文件目录
    └── 建筑表.json
```
