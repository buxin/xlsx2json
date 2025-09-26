const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { program } = require('commander');

// 读取配置文件
function loadConfig() {
    const configPath = path.join(__dirname, 'config.json');
    const defaultConfig = {
        inputDir: 'xlsx',
        outputDir: 'json',
        skipFirstRow: true,
        outputMapping: {}
    };

    try {
        if (fs.existsSync(configPath)) {
            const configData = fs.readFileSync(configPath, 'utf8');
            const config = JSON.parse(configData);
            return { ...defaultConfig, ...config };
        }
    } catch (error) {
        console.warn('⚠️  配置文件读取失败，使用默认配置:', error.message);
    }

    return defaultConfig;
}

// 全局配置
const config = loadConfig();

// 获取输出文件名
function getOutputFileName(inputFileName) {
    const fileName = path.basename(inputFileName);
    
    // 检查是否有映射配置
    if (config.outputMapping && config.outputMapping[fileName]) {
        return config.outputMapping[fileName];
    }
    
    // 默认使用原文件名，扩展名改为.json
    const nameWithoutExt = path.basename(inputFileName, path.extname(inputFileName));
    return `${nameWithoutExt}.json`;
}

/**
 * 将xlsx文件转换为JSON
 * @param {string} inputPath - 输入xlsx文件路径
 * @param {string} outputPath - 输出JSON文件路径（可选）
 * @param {boolean} skipFirstRow - 是否跳过第一行（中文注释）
 */
function convertXlsxToJson(inputPath, outputPath = null, skipFirstRow = null) {
    // 使用配置文件中的设置，如果参数为null则使用配置
    if (skipFirstRow === null) {
        skipFirstRow = config.skipFirstRow;
    }
    try {
        // 检查输入文件是否存在
        if (!fs.existsSync(inputPath)) {
            throw new Error(`输入文件不存在: ${inputPath}`);
        }

        // 读取xlsx文件
        const workbook = XLSX.readFile(inputPath);
        const sheetName = workbook.SheetNames[0]; // 使用第一个工作表
        const worksheet = workbook.Sheets[sheetName];

        // 将工作表转换为JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1, // 使用数组格式，第一行作为标题
            defval: '' // 空单元格的默认值
        });

        let result = jsonData;

        // 如果跳过第一行（中文注释）
        if (skipFirstRow && jsonData.length > 0) {
            result = jsonData.slice(1);
        }

        // 将数组转换为对象数组
        if (result.length > 0) {
            let headers, dataRows;
            
            if (skipFirstRow) {
                // 第一行是注释，第二行是字段名，从第三行开始是数据
                if (jsonData.length < 2) {
                    throw new Error('文件至少需要2行：第一行注释，第二行字段名');
                }
                headers = jsonData[1]; // 第二行作为字段名
                dataRows = jsonData.slice(2); // 从第三行开始是数据
            } else {
                // 不跳过第一行，第一行就是字段名
                headers = jsonData[0];
                dataRows = jsonData.slice(1);
            }
            
            const jsonObjects = dataRows.map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = row[index] || '';
                });
                return obj;
            });

            result = jsonObjects;
        }

        // 确定输出文件路径
        if (!outputPath) {
            const inputDir = path.dirname(inputPath);
            // 使用文件名映射功能
            const outputFileName = getOutputFileName(inputPath);
            // 使用配置文件中的输出目录
            if (path.isAbsolute(config.outputDir)) {
                outputPath = path.join(config.outputDir, outputFileName);
            } else {
                const parentDir = path.dirname(inputDir);
                outputPath = path.join(parentDir, config.outputDir, outputFileName);
            }
        }

        // 确保输出目录存在
        const outputDir = path.dirname(outputPath);
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // 写入JSON文件
        fs.writeFileSync(outputPath, JSON.stringify(result, null, 2), 'utf8');
        
        console.log(`✅ 转换成功！`);
        console.log(`📁 输入文件: ${inputPath}`);
        console.log(`📁 输出文件: ${outputPath}`);
        console.log(`📊 数据行数: ${Array.isArray(result) ? result.length : 0}`);

        return result;

    } catch (error) {
        console.error('❌ 转换失败:', error.message);
        throw error;
    }
}

/**
 * 批量转换目录中的所有xlsx文件
 * @param {string} inputDir - 输入目录路径
 * @param {string} outputDir - 输出目录路径
 */
function convertDirectory(inputDir = null, outputDir = null) {
    try {
        // 使用配置文件中的输入目录
        if (!inputDir) {
            inputDir = config.inputDir;
        }

        if (!fs.existsSync(inputDir)) {
            throw new Error(`输入目录不存在: ${inputDir}`);
        }

        // 使用配置文件中的输出目录
        if (!outputDir) {
            // 如果配置的outputDir是绝对路径，直接使用
            if (path.isAbsolute(config.outputDir)) {
                outputDir = config.outputDir;
            } else {
                // 如果是相对路径，相对于输入目录的父目录
                const parentDir = path.dirname(inputDir);
                outputDir = path.join(parentDir, config.outputDir);
            }
        }

        const files = fs.readdirSync(inputDir);
        const xlsxFiles = files.filter(file => {
            const ext = path.extname(file).toLowerCase();
            const fileName = path.basename(file);
            // 过滤掉临时文件和隐藏文件
            return ext === '.xlsx' && !fileName.startsWith('~$') && !fileName.startsWith('.');
        });

        if (xlsxFiles.length === 0) {
            console.log('📁 目录中没有找到xlsx文件');
            return;
        }

        console.log(`📁 输入目录: ${inputDir}`);
        console.log(`📁 找到 ${xlsxFiles.length} 个xlsx文件`);
        console.log(`📁 输出目录: ${outputDir}`);

        xlsxFiles.forEach((file, index) => {
            const inputPath = path.join(inputDir, file);
            // 使用文件名映射功能
            const outputFileName = getOutputFileName(file);
            const outputPath = path.join(outputDir, outputFileName);
            
            console.log(`\n[${index + 1}/${xlsxFiles.length}] 处理文件: ${file}`);
            console.log(`📄 输出文件: ${outputFileName}`);
            convertXlsxToJson(inputPath, outputPath);
        });

        console.log(`\n🎉 批量转换完成！共处理 ${xlsxFiles.length} 个文件`);

    } catch (error) {
        console.error('❌ 批量转换失败:', error.message);
        throw error;
    }
}

// 命令行参数配置
program
    .name('xlsx2json')
    .description('将xlsx文件转换为JSON格式')
    .version('1.0.0');

program
    .command('convert <input>')
    .description('转换单个xlsx文件')
    .option('-o, --output <path>', '输出JSON文件路径')
    .option('--no-skip-header', '不跳过第一行（默认跳过第一行作为中文注释）')
    .action((input, options) => {
        convertXlsxToJson(input, options.output, options.skipHeader);
    });

program
    .command('batch [inputDir]')
    .description('批量转换目录中的所有xlsx文件')
    .option('-o, --output <dir>', '输出目录路径')
    .action((inputDir, options) => {
        convertDirectory(inputDir, options.output);
    });

program
    .command('auto')
    .description('使用配置文件自动转换（推荐）')
    .action(() => {
        console.log('🚀 使用配置文件自动转换...');
        convertDirectory();
    });

program
    .command('config')
    .description('显示当前配置信息')
    .action(() => {
        console.log('📋 当前配置信息:');
        console.log(`📁 输入目录: ${config.inputDir}`);
        console.log(`📁 输出目录: ${config.outputDir}`);
        console.log(`⏭️  跳过第一行: ${config.skipFirstRow ? '是' : '否'}`);
        console.log(`📄 配置文件: ${path.join(__dirname, 'config.json')}`);
        
        if (config.outputMapping && Object.keys(config.outputMapping).length > 0) {
            console.log('\n📝 文件名映射:');
            Object.entries(config.outputMapping).forEach(([input, output]) => {
                console.log(`  ${input} → ${output}`);
            });
        }
    });

// 如果没有参数，显示帮助信息
if (process.argv.length <= 2) {
    console.log('🔧 xlsx2json - Excel转JSON工具\n');
    console.log('使用方法:');
    console.log('  node index.js convert <xlsx文件路径> [选项]');
    console.log('  node index.js batch <目录路径> [选项]');
    console.log('\n示例:');
    console.log('  node index.js convert xlsx/建筑表.xlsx');
    console.log('  node index.js convert xlsx/建筑表.xlsx -o output/result.json');
    console.log('  node index.js batch xlsx/');
    console.log('\n选项:');
    console.log('  -o, --output <path>    指定输出文件或目录路径');
    console.log('  --no-skip-header       不跳过第一行（默认跳过第一行作为中文注释）');
    console.log('  -h, --help             显示帮助信息');
    console.log('  -V, --version          显示版本号');
}

// 解析命令行参数
program.parse();
