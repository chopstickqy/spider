import xlsx from 'xlsx';

// 保存数据到新的 Excel 文件
function saveToExcel(data: any[], filePath: string) {
  const newWorkbook = xlsx.utils.book_new();
  const newWorksheet = xlsx.utils.aoa_to_sheet(data);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
  xlsx.writeFile(newWorkbook, filePath);
}


function main() {
  const sort = [
    "A0A2U1NVU3",
    "A0A2U1P6A4",
    "A0A2U1P2K1",
    "A0A2U1PM18",
    "A0A2U1LSG0",
    "A0A2U1KNL7",
    "A0A2U1L7B3",
    "A0A2U1NHF6",
    "A0A2U1NSL9",
    "A0A2U1KER2",
    "A0A2U1L7A9",
    "A0A2U1PFF0",
    "A0A2U1P244",
    "A0A2U1Q4A9",
    "A0A2U1KEU5",
    "A0A2U1LPQ4",
    "A0A2U1LXT4",
    "A0A2U1MXS7",
    "A0A2U1L8J0",
    "A0A2U1N9U0",
    "A0A2U1M6H6",
    "A0A2U1KQY2",
    "A0A2U1PEX7",
    "A0A2U1QJF1",
    "A0A2U1NVE9",
    "A0A2U1QFM4",
    "A0A2U1NZW5",
    "A0A2U1PK94",
    "A0A1W6C7U8",
    "A0A2U1NDW7",
    "A0A2U1LX58",
    "A0A2U1LMU6",
    "A0A2U1Q104",
    "A0A2U1PW94",
    "A0A2U1LQA4",
    "A0A2U1PM50",
    "A0A2U1PA41",
    "A0A2U1MU09",
    "A0A2U1KTX4",
    "A0A2U1Q9Y1",
    "A0A2U1PNK0",
    "A0A2U1NJM8",
    "A0A2U1LEV3",
    "A0A2U1MJC3",
    "A0A2U1LDK5",
    "A0A2U1N2L9",
    "A0A2U1PUA3",
    "A0A2U1PX72",
    "A0A2U1NLL7",
    "A0A2U1KRD8",
    "A0A2U1KEM1",
    "A0A2U1PHH0",
    "A0A2U1PV09",
    "A0A2U1PPC5",
    "A0A2U1L0E4",
    "A0A2U1L6J8",
    "A0A2U1QLA2",
    "A0A2U1MEW8",
    "A0A2U1PJW5",
    "A0A2U1PFW8",
    "A0A1W6C7V2",
    "A0A2U1N5P2",
    "A0A2U1M9B5",
    "A0A2U1MPC8",
    "A0A2U1Q970",
    "A0A2U1MZI9",
    "A0A2U1L5A0",
    "A0A2U1QM52",
    "A0A2U1Q2E8",
    "A0A2U1QHT5",
    "A0A2U1MD88",
    "A0A2U1M5Y9",
    "A0A2U1Q8I6",
    "A0A2U1KSM7",
    "A0A2U1M5D9",
    "A0A2U1NF81",
    "A0A2U1NU44",
    "A0A2U1PPX1",
    "A0A2U1LTW7",
    "A0A2U1NV12",
    "A0A2U1LDD5",
    "A0A2U1PQP5",
    "A0A2U1NED1",
    "A0A2U1L7Z0",
    "A0A2U1LWI4",
    "A0A2U1Q1N7",
    "A0A2U1NUG0",
    "A0A2U1M929",
    "A0A2U1LW17",
    "A0A2U1PD29",
    "A0A2U1NK55",
    "A0A2U1MPV5",
    "A0A2U1NW26",
    "A0A2U1NM70",
    "A0A2U1N4Z8",
    "A0A2U1KG05",
    "A0A2U1MA54",
    "A0A2U1KI03",
    "A0A2U1NNQ3",
    "A0A2U1M0S0",
    "A0A2U1Q435",
    "A0A2U1MB35",
    "A0A2U1N7H5",
    "A0A2U1P0C7",
    "A0A2U1MCM8",
    "A0A2U1M0H6",
    "A0A2U1PFP0",
    "A0A2U1MZL9",
    "A0A2U1L648",
    "A0A2U1LPS3",
    "A0A2U1NW93",
    "A0A2U1L5R0",
    "A0A2U1NFW4",
    "A0A2U1MWA7",
    "A0A2U1PRN8",
    "A0A2U1MTD1",
    "A0A2U1M9W1",
    "A0A2U1L5S8",
    "A0A2U1KBT4",
    "A0A2U1LR45",
    "A0A2U1NSQ4",
    "A0A2U1PPD8",
    "A0A1W6C7N7",
    "A0A2U1MXI0",
    "A0A2U1PY73",
    "A0A2U1QAB6",
    "A0A2U1PIA9",
    "A0A2U1Q0E6",
    "A0A2U1LQE1",
    "A0A2U1QMF1"
  ]

  const workbook = xlsx.readFile('爬虫数据.xlsx'); // 读取Excel文件
  const sheetName = workbook.SheetNames[0]; // 获取第一个工作表
  const worksheet = workbook.Sheets[sheetName]; // 获取工作表
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // 将工作表转换为JSON格式

  // 根据sort数组对数据进行排序
  data.sort((a, b) => {
      const indexA = sort.indexOf(a[0]); // 获取a的索引
      const indexB = sort.indexOf(b[0]); // 获取b的索引
      return indexA - indexB; // 根据索引进行排序
  });

  saveToExcel(data,'排序后的数据.xlsx')
}

// 调用main函数
main();