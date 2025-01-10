import xlsx from 'xlsx';
import fs from 'fs';

// 读取 Excel 文件并返回数据
function readExcel(filePath: string): any[] {
    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const sheetNames = workbook.SheetNames;
    const sheet = workbook.Sheets[sheetNames[0]]; // 读取第一个工作表
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    return data;
}

// 合并多个 Excel 文件的数据
function mergeExcelFiles(filePaths: string[]): any[] {
    const mergedData: any[] = [];

    filePaths.forEach(filePath => {
        if (fs.existsSync(filePath)) {
            const data = readExcel(filePath);
            mergedData.push(...data); // 合并数据
        } else {
            console.error(`文件不存在: ${filePath}`);
        }
    });

    return mergedData;
}

// 保存数据到新的 Excel 文件
function saveToExcel(data: any[], filePath: string) {
    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
    xlsx.writeFile(newWorkbook, filePath);
}

// 主函数
function main() {
    const filesToMerge = [
        '解析数据.xlsx',
        '解析数据2.xlsx',
        '解析数据3.xlsx',
        '解析数据4.xlsx'
    ];

    const mergedData = mergeExcelFiles(filesToMerge);
    saveToExcel(mergedData, '爬虫数据.xlsx');
    console.log('数据已合并并保存到 爬虫数据.xlsx');
}

// 执行主函数
main(); 