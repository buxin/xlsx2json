# xlsx2json

一个简单易用的Excel文件转JSON工具，使用Node.js开发。

## 功能特点

- ✅ 支持单个xlsx文件转换
- ✅ 支持批量转换目录中的所有xlsx文件
- ✅ 自动跳过第一行（中文注释）
- ✅ 支持自定义输出路径
- ✅ 本地配置文件支持
- ✅ 输入目录配置
- ✅ 输出文件名映射
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
# 使用配置文件自动转换（推荐）
node index.js auto

# 转换指定目录中所有xlsx文件
node index.js batch xlsx/

# 指定输出目录
node index.js batch xlsx/ -o output/
```

**3. 配置管理**
```bash
# 查看当前配置
node index.js config

# 修改配置文件 config.json
# 可以设置输出目录和是否跳过第一行
```

**4. 查看帮助**
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

## 配置文件

项目根目录下的 `config.json` 文件用于配置转换参数：

```json
{
  "inputDir": "xlsx",
  "outputDir": "json",
  "skipFirstRow": true,
  "outputMapping": {
    "建筑表.xlsx": "building.json",
    "用户表.xlsx": "user.json",
    "商品表.xlsx": "product.json"
  },
  "description": "xlsx2json 配置文件",
  "settings": {
    "inputDir": "输入目录路径，相对于项目根目录",
    "outputDir": "输出目录路径，相对于项目根目录",
    "skipFirstRow": "是否跳过第一行（中文注释），true/false",
    "outputMapping": "输出文件名映射，键为xlsx文件名，值为json文件名"
  }
}
```

### 配置说明：
- `inputDir`: 输入目录路径，默认为 "xlsx"
- `outputDir`: 输出目录路径，默认为 "json"
- `skipFirstRow`: 是否跳过第一行（中文注释），默认为 true
- `outputMapping`: 输出文件名映射，可以自定义xlsx文件对应的json文件名

## 注意事项

- 默认跳过第一行作为中文注释
- 空单元格会被转换为空字符串
- 输出目录会自动创建
- 支持中文文件名和路径
- 可通过配置文件自定义输入/输出目录
- 支持文件名映射，如：建筑表.xlsx → building.json

## 项目结构

```
xlsx2json/
├── package.json          # 项目配置
├── index.js             # 主程序
├── config.json          # 配置文件
├── convert.bat          # Windows一键转换脚本
├── README.md            # 说明文档
├── xlsx/                # 输入xlsx文件目录
│   └── building.xlsx
└── json/                # 输出JSON文件目录
    └── building.json
```
