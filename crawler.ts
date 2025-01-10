// import * as puppeteer from 'puppeteer';
import xlsx from 'xlsx';
import axios from 'axios';
import * as cheerio from 'cheerio';
import fs from 'fs';

// 封装请求方法
async function fetchData(key: string) {
    const url = `https://rest.uniprot.org/uniprotkb/${key}`;
    const headers = {
        'sec-ch-ua-platform': 'macOS',
        'Referer': 'https://www.uniprot.org/',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'Accept': 'application/json',
        'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0'
    };

    try {
        const response = await axios.get(url, { headers });
        return response.data; // 返回响应数据
    } catch (error) {
        console.error('请求失败:', error);
        throw error; // 抛出错误以便后续处理
    }
}


// 封装新的 POST 请求方法
async function blastQuery(sequenceValue: string) {
    const url = 'https://tools.arabidopsis.org/cgi-bin/Blast/TAIRblast.pl';
    // const proxy = getRandomProxy(); // 获取随机代理
    const headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Sec-Fetch-Site': 'same-site',
        'Accept-Language': 'zh-CN,zh-Hans;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Sec-Fetch-Mode': 'navigate',
        'Origin': 'https://www.arabidopsis.org',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.1.1 Safari/605.1.15',
        'Referer': 'https://www.arabidopsis.org/',
        'Sec-Fetch-Dest': 'document',
        'Cookie': '_ga_S17ZS9ZPHD=GS1.1.1736416034.1.0.1736416065.0.0.0; _ga=GA1.1.936239597.1736416034', // 示例 Cookie
    };

    const params = new URLSearchParams();
    params.append('Algorithm', 'blastp');
    params.append('default_db', 'Araport11_pep_20220914');
    params.append('BlastTargetSet', 'Araport11_pep_20220914');
    params.append('textbox', 'seq');
    params.append('QueryText', sequenceValue); // 使用 sequenceValue 作为 QueryText
    params.append('QueryFilter', 'T');
    params.append('ReplyFormat', 'HTML');
    params.append('Comment', 'optional, will be added to output for your use');
    params.append('ReplyTo', '');
    params.append('GappedAlignment', 'T');
    params.append('ReplyVia', 'BROWSER');
    params.append('Matrix', 'blosum62');
    params.append('MaxScores', '100');
    params.append('MaxAlignments', '50');
    params.append('NucleicMismatch', '-3');
    params.append('NucleicMatch', '2');
    params.append('OpenPenalty', '0 (use default)');
    params.append('ExtendPenalty', '0 (use default)');
    params.append('ExtensionThreshold', '0 (use default)');
    params.append('WordSize', '0 (use default)');
    params.append('Expectation', '10');
    params.append('QueryGeneticCode', '1');

    try {
        const response = await axios.post(url, params, {
            headers,
            proxy: {
                host: '127.0.0.1', // 这里用的VPN代理，防止IP被限制
                port: 7890,
                // 启用 SSL
                protocol: 'http', // 或 'https'，根据你的代理类型
            },
        });
        return response.data; // 返回响应数据
    } catch (error) {
        console.error('BLAST 请求失败:', error);
        throw error; // 抛出错误以便后续处理
    }
}

// 解析 HTML 数据，提取第一个 <area> 元素的 href 值
function parseBlastResult(html: string): string | null {
    const $ = cheerio.load(html);
    const areaHref = $('MAP[name="imap"] area').first().attr('href');
    if(!areaHref) {
      console.error('parseBlastResult 失败');
    }
    return areaHref ? areaHref.replace('#', '') : null; // 去掉 # 号
}

// 解析 Excel 文件
function readExcel(filePath: string, sheetName: string): any[] {
    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const sheetNames = workbook.SheetNames;
    console.log(`总共有 ${sheetNames.length} 个工作表`, sheetNames);
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    return data;
}

// 保存数据到 Excel 文件
function saveToExcel(data: any[], filePath: string) {
    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
    xlsx.writeFile(newWorkbook, filePath);
}

// 封装新的 POST 请求方法以获取基因信息
async function fetchGeneData(searchText: string) {
    const url = 'https://www.arabidopsis.org/api/search/gene';
    const headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,vi;q=0.7,zh-TW;q=0.6',
        'content-type': 'application/json;charset=UTF-8',
        'origin': 'https://www.arabidopsis.org',
        'referer': `https://www.arabidopsis.org/results?mainType=general&searchText=${searchText}&category=genes`,
        'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': 'macOS',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
    };

    const data = {
        searchText: searchText
    };

    try {
        const response = await axios.post(url, data, { headers });
        return response.data; // 返回响应数据
    } catch (error) {
        console.error('获取基因数据请求失败:', error);
        throw error; // 抛出错误以便后续处理
    }
}

// 爬虫主函数
async function runCrawler() {
    const data = readExcel('data.xlsx', 'Sheet2');

    const results: any[] = []; // 用于存储结果
    const tasks: Promise<void>[] = []; // 用于存储所有的任务

    // 遍历数据，提取第四个元素作为 key
    for (const row of data) {
        const key = row[3]; // 假设第四个元素在索引 3
        if (key) {
            const task = (async () => {
                console.log('开始解析：' + key);
                try {
                    const apiData = await fetchData(key);
                    const sequenceValue = apiData?.sequence?.value; // 获取 sequence.value
                    if (sequenceValue) {
                        // 使用 sequenceValue 进行 BLAST 查询
                        const blastResultHtml = await blastQuery(sequenceValue);
                        const areaHref = parseBlastResult(blastResultHtml); // 解析 href 值
                        if (areaHref) {
                            // 使用 areaHref 进行基因数据请求
                            const geneData = await fetchGeneData(areaHref);
                            const geneName = geneData?.docs?.[0]?.gene_name?.join('、');
                            const descArr = (geneData?.docs || []).map((item) => item?.description?.join('、') || '');
                            results.push([key, sequenceValue, areaHref, geneName, ...descArr]); // 将 key, sequence.value, href 和基因数据组合成数组
                        }
                    }
                } catch (error) {
                    console.error('解析失败：' + key, error);
                }
                console.log('结束解析：' + key);
            })();
            tasks.push(task); // 将任务添加到任务数组中
        }
    }

    // 等待所有任务完成
    const resultsSettled = await Promise.allSettled(tasks);

    // 处理成功和失败的结果
    resultsSettled.forEach((result, index) => {
        if (result.status === 'rejected') {
            console.error(`任务 ${index} 失败:`, result.reason);
        }
    });

    // 保存结果到新的 Excel 文件
    saveToExcel(results, 'newData.xlsx');
    console.log('数据已保存到 newData.xlsx');
}

// 启动爬虫
runCrawler().catch(console.error);