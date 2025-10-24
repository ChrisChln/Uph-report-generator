// 全局变量
let pickingData = null;
let packingData = null;
let savedEmployees = []; // 保存的员工列表

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    // 设置默认日期为今天 (YYYY-MM-DD 格式用于 date input)
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    document.getElementById('dateInput').value = `${year}-${month}-${day}`;
    
    // 加载保存的员工列表
    loadEmployees();
    
    // 初始化 Preshipment 输入框（默认显示 1 个）
    updatePreshipmentInputs();
    
    // 文件选择事件
    document.getElementById('pickingFile').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            document.getElementById('pickingFileName').textContent = file.name;
            readExcelFile(file, 'picking');
        }
    });
    
    document.getElementById('packingFile').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            document.getElementById('packingFileName').textContent = file.name;
            readExcelFile(file, 'packing');
        }
    });
});

// 格式化日期
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
}

// 读取 Excel 文件
function readExcelFile(file, type) {
    log(`正在读取 ${type === 'picking' ? 'Picking' : 'Packing'} 表格: ${file.name}`, 'info');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (type === 'picking') {
                pickingData = processPickingData(jsonData);
                log(`Picking 数据读取完成，共 ${pickingData.length} 行`, 'success');
            } else {
                packingData = processPackingData(jsonData);
                log(`Packing 数据读取完成，共 ${packingData.length} 行`, 'success');
            }
        } catch (error) {
            log(`读取文件失败: ${error.message}`, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// 处理 Picking 数据
function processPickingData(data) {
    if (data.length < 2) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    const result = [];
    
    rows.forEach((row, index) => {
        if (row.length > 12) {
            const quantity = row[8];     // I列 (索引8) - 拣货数量
            const singleCount = row[9];  // J列 (索引9) - 一分数（单品）
            const multiCount = row[10];  // K列 (索引10) - 二分数（多品）
            const worker = row[11];      // L列 (索引11) - 拣货员
            const time = row[12];        // M列 (索引12) - 操作时间
            
            if (worker && time) {
                const qty = parseInt(quantity) || 1;
                const parsedTime = parseExcelDate(time);
                
                // 判断是单品还是多品
                const single = parseInt(singleCount) || 0;
                const multi = parseInt(multiCount) || 0;
                
                let itemType = 'unknown';
                if (single > 0) {
                    itemType = 'single';  // 单品
                } else if (multi > 0) {
                    itemType = 'multi';   // 多品
                }
                
                result.push({
                    worker: String(worker),
                    time: parsedTime,
                    quantity: qty,
                    type: 'picking',
                    itemType: itemType  // 单品/多品标识
                });
            }
        }
    });
    
    // 显示数据日期信息
    if (result.length > 0) {
        // 过滤出有效的时间数据
        const validTimes = result
            .map(r => r.time)
            .filter(t => t && !isNaN(t.getTime()))
            .sort((a, b) => a - b);  // 按时间排序
        
        if (validTimes.length > 0) {
            const firstDate = validTimes[0];
            const lastDate = validTimes[validTimes.length - 1];
            const dateStr = formatDate(firstDate);
            const firstTime = firstDate.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
            const lastTime = lastDate.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
            log(`Picking 数据日期: ${dateStr}`, 'info');
            log(`时间范围: ${firstTime} - ${lastTime}`, 'info');
        } else {
            log('Picking 数据时间格式无法识别，请检查数据', 'warning');
            log('请检查 M 列的数据格式是否正确', 'warning');
        }
    }
    
    return result;
}

// 处理 Packing 数据
function processPackingData(data) {
    if (data.length < 2) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    const result = [];
    
    rows.forEach(row => {
        if (row.length > 23) {
            const worker = row[21];    // V列 (索引21)
            const quantity = row[7];   // H列 (索引7) - 扫描件数
            const time = row[23];      // X列 (索引23)
            
            if (worker && time) {
                const parsedTime = parseExcelDate(time);
                result.push({
                    worker: String(worker),
                    time: parsedTime,
                    quantity: parseInt(quantity) || 1,
                    type: 'packing'
                });
            }
        }
    });
    
    // 显示数据日期信息
    if (result.length > 0) {
        // 过滤出有效的时间数据
        const validTimes = result
            .map(r => r.time)
            .filter(t => t && !isNaN(t.getTime()))
            .sort((a, b) => a - b);  // 按时间排序
        
        if (validTimes.length > 0) {
            const firstDate = validTimes[0];
            const lastDate = validTimes[validTimes.length - 1];
            const dateStr = formatDate(firstDate);
            const firstTime = firstDate.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
            const lastTime = lastDate.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
            log(`Packing 数据日期: ${dateStr}`, 'info');
            log(`时间范围: ${firstTime} - ${lastTime}`, 'info');
        } else {
            log('Packing 数据时间格式无法识别，请检查数据', 'warning');
        }
    }
    
    return result;
}

// 解析 Excel 日期
function parseExcelDate(excelDate) {
    if (typeof excelDate === 'number') {
        // Excel 日期是从 1900-01-01 开始的天数
        const date = new Date((excelDate - 25569) * 86400 * 1000);
        return date;
    } else if (typeof excelDate === 'string') {
        // 尝试多种日期格式解析
        // 格式1: "2025-10-24 08:30:45"
        // 格式2: "10/24/2025 8:30:45 AM"
        // 格式3: "2025/10/24 08:30:45"
        const date = new Date(excelDate);
        if (!isNaN(date.getTime())) {
            return date;
        }
        
        // 尝试解析其他格式
        const datePatterns = [
            /(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/,  // 2025-10-24 08:30:45
            /(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/,  // 10/24/2025 8:30:45
            /(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/   // 2025/10/24 08:30:45
        ];
        
        for (let pattern of datePatterns) {
            const match = excelDate.match(pattern);
            if (match) {
                let year, month, day, hour, minute, second;
                if (pattern.source.startsWith('(\\d{4})')) {
                    // YYYY-MM-DD or YYYY/MM/DD
                    [, year, month, day, hour, minute, second] = match;
                } else {
                    // MM/DD/YYYY
                    [, month, day, year, hour, minute, second] = match;
                }
                return new Date(year, month - 1, day, hour, minute, second);
            }
        }
    } else if (excelDate instanceof Date) {
        return excelDate;
    }
    return null;
}

// 计算 EWH (基于时间戳，5分钟阈值)
function calculateEWH(timestamps, thresholdMinutes = 5) {
    if (!timestamps || timestamps.length < 2) return 0;
    
    // 过滤并排序时间戳
    const times = timestamps
        .filter(t => t instanceof Date && !isNaN(t.getTime()))
        .sort((a, b) => a - b);
    
    if (times.length < 2) return 0;
    
    const threshold = thresholdMinutes * 60 * 1000; // 转换为毫秒
    let segments = [];
    let start = times[0];
    let prev = times[0];
    
    for (let i = 1; i < times.length; i++) {
        const cur = times[i];
        if (cur - prev > threshold) {
            // 间隔超过阈值，记录当前段
            segments.push([start, prev]);
            start = cur;
        }
        prev = cur;
    }
    segments.push([start, prev]);
    
    // 计算总时长
    let totalMs = 0;
    segments.forEach(([s, e]) => {
        totalMs += e - s;
    });
    
    // 转换为小时
    return totalMs / (3600 * 1000);
}

// 生成报表
function generateReport() {
    if (!pickingData && !packingData) {
        log('请至少上传一个表格文件', 'error');
        alert('请至少上传一个表格文件！');
        return;
    }
    
    log('\n' + '─'.repeat(60), 'info');
    log('开始生成每日工作报表', 'info');
    log('─'.repeat(60) + '\n', 'info');
    
    // 收集用户输入
    const dateInputValue = document.getElementById('dateInput').value;
    // 将 YYYY-MM-DD 转换为 YYYY/MM/DD
    const date = dateInputValue ? dateInputValue.replace(/-/g, '/') : '';
    const time = document.getElementById('timeInput').value;
    const preshipmentCount = parseInt(document.getElementById('preshipmentCount').value) || 0;
    
    const preshipmentWorkers = [];
    for (let i = 0; i < preshipmentCount; i++) {
        const nameSelect = document.getElementById(`ps_name_${i}`);
        const name = nameSelect ? nameSelect.value.trim() : '';
        const quantity = parseInt(document.getElementById(`ps_qty_${i}`).value) || 0;
        const ewh = parseFloat(document.getElementById(`ps_ewh_${i}`).value) || 0;
        if (name && quantity > 0) {
            preshipmentWorkers.push({ name, quantity, ewh });
        }
    }
    
    log('用户输入信息:');
    log(`  日期: ${date}`);
    log(`  工作时间: ${time}`);
    log(`  Preshipment 员工数: ${preshipmentWorkers.length}`);
    preshipmentWorkers.forEach((w, i) => {
        log(`  Preshipment ${i+1}: ${w.name} - ${w.quantity} 件, EWH: ${w.ewh}`);
    });
    log('');
    
    // 禁用按钮
    const btn = document.getElementById('generateBtn');
    btn.disabled = true;
    btn.innerHTML = '正在生成中...';
    
    setTimeout(() => {
        try {
            // 汇总数据
            const report = aggregateData(pickingData, packingData, date, time, preshipmentWorkers);
            
            // 生成 Excel
            exportToExcel(report);
            
            log('\n报表生成成功', 'success');
            log(`统计记录数: ${report.length} 条`, 'success');
            log('─'.repeat(60) + '\n', 'info');
            
        } catch (error) {
            log(`\n生成报表失败: ${error.message}`, 'error');
            console.error(error);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '生成每日工作报表';
        }
    }, 100);
}

// 汇总数据
function aggregateData(pickingData, packingData, date, time, preshipmentWorkers) {
    log('正在汇总数据\n', 'info');
    
    const workerMap = new Map();
    
    // 处理 Picking 数据（分单品和多品）
    if (pickingData && pickingData.length > 0) {
        const totalPickQuantity = pickingData.reduce((sum, item) => sum + (item.quantity || 0), 0);
        const singleItems = pickingData.filter(item => item.itemType === 'single');
        const multiItems = pickingData.filter(item => item.itemType === 'multi');
        const singleCount = singleItems.reduce((sum, item) => sum + (item.quantity || 0), 0);
        const multiCount = multiItems.reduce((sum, item) => sum + (item.quantity || 0), 0);
        
        log(`  处理 Picking 数据 (${pickingData.length} 条记录, 总数量: ${totalPickQuantity})`);
        log(`    - 单品: ${singleCount} 件 (${singleItems.length} 条)`);
        log(`    - 多品: ${multiCount} 件 (${multiItems.length} 条)`);
        
        pickingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingTimes: [],
                    pickingSingleTimes: [],  // 单品时间
                    pickingMultiTimes: [],   // 多品时间
                    packingTimes: [],
                    pickCount: 0,
                    pickSingleCount: 0,      // 单品件数
                    pickMultiCount: 0,       // 多品件数
                    packCount: 0
                });
            }
            const worker = workerMap.get(item.worker);
            worker.pickingTimes.push(item.time);
            worker.pickCount += item.quantity || 1;
            
            // 分类统计单品和多品
            if (item.itemType === 'single') {
                worker.pickingSingleTimes.push(item.time);
                worker.pickSingleCount += item.quantity || 1;
            } else if (item.itemType === 'multi') {
                worker.pickingMultiTimes.push(item.time);
                worker.pickMultiCount += item.quantity || 1;
            }
        });
        log(`  Picking 数据处理完成`);
    }
    
    // 处理 Packing 数据
    if (packingData && packingData.length > 0) {
        const totalPackQuantity = packingData.reduce((sum, item) => sum + (item.quantity || 0), 0);
        log(`  处理 Packing 数据 (${packingData.length} 条记录, 总数量: ${totalPackQuantity})`);
        
        packingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingTimes: [],
                    packingTimes: [],
                    pickCount: 0,
                    packCount: 0
                });
            }
            const worker = workerMap.get(item.worker);
            worker.packingTimes.push(item.time);
            worker.packCount += item.quantity;
        });
        log(`  Packing 数据处理完成`);
    }
    
    // 计算 EWH 和 UPH
    log('\n  正在计算 EWH 和 UPH\n');
    const report = [];
    
    workerMap.forEach((data, worker) => {
        // 计算 Picking 单品和多品 EWH（详细算法 + 5% 补偿）
        const pickingSingleEWH = data.pickingSingleTimes.length > 0 
            ? calculateEWH(data.pickingSingleTimes) * 1.05  // 5% 补偿
            : 0;
        const pickingMultiEWH = data.pickingMultiTimes.length > 0 
            ? calculateEWH(data.pickingMultiTimes) * 1.05   // 5% 补偿
            : 0;
        
        // 计算总 Picking EWH
        const pickingEWH = pickingSingleEWH + pickingMultiEWH;
        
        // 计算 Packing EWH
        const packingEWH = calculateEWH(data.packingTimes);
        
        // 计算总 EWH
        const totalEWH = pickingEWH + packingEWH;
        
        // 计算 UPH
        const pickingSingleUPH = pickingSingleEWH > 0 ? (data.pickSingleCount / pickingSingleEWH) : 0;
        const pickingMultiUPH = pickingMultiEWH > 0 ? (data.pickMultiCount / pickingMultiEWH) : 0;
        const packingUPH = packingEWH > 0 ? (data.packCount / packingEWH) : 0;
        
        // 输出详细日志
        if (data.pickSingleCount > 0 || data.pickMultiCount > 0) {
            log(`  ${worker}:`);
            if (data.pickSingleCount > 0) {
                log(`    单品: ${data.pickSingleCount}件, EWH: ${round(pickingSingleEWH, 2)}h (含5%补偿), UPH: ${round(pickingSingleUPH, 2)}`);
            }
            if (data.pickMultiCount > 0) {
                log(`    多品: ${data.pickMultiCount}件, EWH: ${round(pickingMultiEWH, 2)}h (含5%补偿), UPH: ${round(pickingMultiUPH, 2)}`);
            }
            if (data.packCount > 0) {
                log(`    打包: ${data.packCount}件, EWH: ${round(packingEWH, 2)}h, UPH: ${round(packingUPH, 2)}`);
            }
        }
        
        report.push({
            Date: date,
            TIME: time,
            EWH: round(totalEWH, 2),
            department: 'OB',
            worker: worker,
            '拣货单品': data.pickSingleCount > 0 ? data.pickSingleCount : '',
            '拣货多品': data.pickMultiCount > 0 ? data.pickMultiCount : '',
            pack: data.packCount > 0 ? data.packCount : '',
            box: '',
            Preshipment: '',
            '单品UPH': pickingSingleUPH > 0 ? round(pickingSingleUPH, 2) : '',
            '多品UPH': pickingMultiUPH > 0 ? round(pickingMultiUPH, 2) : '',
            'Packing UPH': packingUPH > 0 ? round(packingUPH, 2) : '',
            'Preship UPH': ''
        });
    });
    
    log(`  EWH 和 UPH 计算完成`);
    
    // 处理 Preshipment 数据
    if (preshipmentWorkers.length > 0) {
        log(`\n  正在处理 Preshipment 数据\n`);
        preshipmentWorkers.forEach(ps => {
            const existing = report.find(r => r.worker === ps.name);
            if (existing) {
                // 员工已存在，更新 Preshipment 和 Preship UPH
                existing.Preshipment = ps.quantity;
                
                // 使用用户填写的 Preshipment EWH 来计算 UPH
                if (ps.ewh > 0) {
                    // 将 Preshipment EWH 加入总 EWH
                    existing.EWH = round(existing.EWH + ps.ewh, 2);
                    // 计算 Preship UPH = Preshipment件数 / Preshipment EWH
                    existing['Preship UPH'] = round(ps.quantity / ps.ewh, 2);
                    log(`  ${ps.name}: Preshipment=${ps.quantity}, EWH=${ps.ewh}, UPH=${existing['Preship UPH']}`);
                } else {
                    log(`  ${ps.name}: Preshipment=${ps.quantity}, UPH=0 (未填写EWH)`);
                }
            } else {
                // 新员工，添加新行（只有 Preshipment）
                const preshipUPH = ps.ewh > 0 ? round(ps.quantity / ps.ewh, 2) : '';
                report.push({
                    Date: date,
                    TIME: time,
                    EWH: ps.ewh,
                    department: 'OB',
                    worker: ps.name,
                    '拣货单品': '',
                    '拣货多品': '',
                    pack: '',
                    box: '',
                    Preshipment: ps.quantity,
                    '单品UPH': '',
                    '多品UPH': '',
                    'Packing UPH': '',
                    'Preship UPH': preshipUPH
                });
                log(`  ${ps.name}: Preshipment=${ps.quantity}, EWH=${ps.ewh}, UPH=${preshipUPH || '0'} (新增员工)`);
            }
        });
        log(`  Preshipment 数据处理完成`);
    }
    
    // 按 EWH 降序排序
    report.sort((a, b) => (b.EWH || 0) - (a.EWH || 0));
    
    return report;
}

// 导出到 Excel
function exportToExcel(data) {
    log('\n正在生成 Excel 文件\n', 'info');
    
    // 创建工作簿
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    
    // 获取数据范围
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    // 表头样式（白字黑底，居中，加粗）
    const headerStyle = {
        fill: { fgColor: { rgb: "000000" } },
        font: { color: { rgb: "FFFFFF" }, bold: true, sz: 12 },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };
    
    // 数据单元格样式（全部居中，边框）
    const cellStyle = {
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "D3D3D3" } },
            bottom: { style: "thin", color: { rgb: "D3D3D3" } },
            left: { style: "thin", color: { rgb: "D3D3D3" } },
            right: { style: "thin", color: { rgb: "D3D3D3" } }
        }
    };
    
    // 空白单元格样式（浅灰色背景）
    const emptyCellStyle = {
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "F5F5F5" } },
        border: {
            top: { style: "thin", color: { rgb: "D3D3D3" } },
            bottom: { style: "thin", color: { rgb: "D3D3D3" } },
            left: { style: "thin", color: { rgb: "D3D3D3" } },
            right: { style: "thin", color: { rgb: "D3D3D3" } }
        }
    };
    
    // 应用表头样式
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_col(C) + "1";
        if (!ws[address]) continue;
        ws[address].s = headerStyle;
    }
    
    // 应用数据单元格样式（全部居中，空白单元格为灰色）
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[address]) continue;
            
            // 检查单元格值，如果为空或0则使用灰色背景
            const cellValue = ws[address].v;
            const isEmpty = cellValue === '' || cellValue === null || cellValue === undefined || cellValue === 0;
            
            ws[address].s = isEmpty ? emptyCellStyle : cellStyle;
        }
    }
    
    // 自动调整列宽
    const colWidths = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
        let maxWidth = 10; // 最小宽度
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const address = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[address]) continue;
            
            const cellValue = ws[address].v;
            if (cellValue) {
                // 计算字符串长度（中文算2个字符，英文算1个）
                const len = String(cellValue).replace(/[\u4e00-\u9fa5]/g, "aa").length;
                maxWidth = Math.max(maxWidth, len);
            }
        }
        
        // 设置列宽（加一些边距）
        colWidths.push({ wch: Math.min(maxWidth + 2, 50) }); // 最大宽度50
    }
    ws['!cols'] = colWidths;
    
    // 设置行高
    const rowHeights = [];
    for (let R = range.s.r; R <= range.e.r; ++R) {
        rowHeights.push({ hpt: R === 0 ? 25 : 20 }); // 表头行高25，数据行高20
    }
    ws['!rows'] = rowHeights;
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, '每日工作报表');
    
    // 生成文件名（工作量表+日期）
    const date = data.length > 0 && data[0].Date ? data[0].Date : formatDate(new Date());
    const dateStr = date.replace(/\//g, ''); // 移除斜杠，例如 20251024
    const filename = `工作量表${dateStr}.xlsx`;
    
    // 下载文件
    XLSX.writeFile(wb, filename, { cellStyles: true });
    
    log(`  Excel 文件已生成: ${filename}`, 'success');
}

// 更新 Preshipment 输入框
function updatePreshipmentInputs() {
    const count = parseInt(document.getElementById('preshipmentCount').value) || 0;
    const container = document.getElementById('preshipmentContainer');
    
    container.innerHTML = '';
    
    if (count === 0) {
        container.innerHTML = '<p style="color: var(--text-tertiary); text-align: center; padding: 12px; font-size: 13px;">暂无 Preshipment 员工</p>';
        return;
    }
    
    for (let i = 0; i < count; i++) {
        const item = document.createElement('div');
        item.className = 'preshipment-item';
        
        // 创建员工选择下拉框
        let optionsHTML = '<option value="">选择员工</option>';
        savedEmployees.forEach(emp => {
            optionsHTML += `<option value="${emp}">${emp}</option>`;
        });
        
        item.innerHTML = `
            <label>#${i + 1}</label>
            <select id="ps_name_${i}">
                ${optionsHTML}
            </select>
            <input type="number" id="ps_qty_${i}" min="0" placeholder="件数" />
            <input type="number" id="ps_ewh_${i}" min="0" step="0.01" placeholder="EWH" />
        `;
        container.appendChild(item);
    }
}

// ========== 员工管理功能 ==========

// 加载保存的员工列表
function loadEmployees() {
    const saved = localStorage.getItem('savedEmployees');
    if (saved) {
        savedEmployees = JSON.parse(saved);
    }
}

// 保存员工列表
function saveEmployees() {
    localStorage.setItem('savedEmployees', JSON.stringify(savedEmployees));
}

// 打开员工管理弹窗
function openEmployeeManager() {
    document.getElementById('employeeModal').style.display = 'block';
    renderEmployeeList();
}

// 关闭员工管理弹窗
function closeEmployeeManager() {
    document.getElementById('employeeModal').style.display = 'none';
    updatePreshipmentInputs(); // 更新下拉选择框
}

// 渲染员工列表
function renderEmployeeList() {
    const listContainer = document.getElementById('employeeList');
    
    if (savedEmployees.length === 0) {
        listContainer.innerHTML = '<p style="color: var(--text-tertiary); text-align: center; padding: 20px;">暂无员工</p>';
        return;
    }
    
    listContainer.innerHTML = '';
    savedEmployees.forEach((employee, index) => {
        const item = document.createElement('div');
        item.className = 'employee-item';
        item.innerHTML = `
            <span>${employee}</span>
            <button onclick="deleteEmployee(${index})">删除</button>
        `;
        listContainer.appendChild(item);
    });
}

// 添加员工
function addEmployee() {
    const input = document.getElementById('newEmployeeName');
    const name = input.value.trim();
    
    if (!name) {
        alert('请输入员工姓名！');
        return;
    }
    
    if (savedEmployees.includes(name)) {
        alert('该员工已存在！');
        return;
    }
    
    savedEmployees.push(name);
    saveEmployees();
    renderEmployeeList();
    input.value = '';
    log(`添加员工: ${name}`, 'success');
}

// 删除员工
function deleteEmployee(index) {
    const name = savedEmployees[index];
    if (confirm(`确定要删除员工 "${name}" 吗？`)) {
        savedEmployees.splice(index, 1);
        saveEmployees();
        renderEmployeeList();
        log(`删除员工: ${name}`, 'info');
    }
}

// 点击弹窗外部关闭
window.onclick = function(event) {
    const modal = document.getElementById('employeeModal');
    if (event.target === modal) {
        closeEmployeeManager();
    }
}

// 日志输出
function log(message, type = '') {
    const container = document.getElementById('logContainer');
    const line = document.createElement('div');
    line.className = 'log-line' + (type ? ` ${type}` : '');
    line.textContent = message;
    container.appendChild(line);
    container.scrollTop = container.scrollHeight;
}

// 四舍五入
function round(num, decimals) {
    return Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
}

// ==================== 人效表功能 ====================

// 生成人效表
function generateEfficiencyReport() {
    log('\n', 'info');
    log('='.repeat(50), 'info');
    log('开始生成人效表 (00:00-12:00)', 'info');
    log('='.repeat(50), 'info');
    
    // 验证文件
    if (!pickingData && !packingData) {
        log('请先上传 Picking 或 Packing 表格文件！', 'error');
        alert('请先上传 Picking 或 Packing 表格文件！');
        return;
    }
    
    try {
        // 筛选 00:00-12:00 的数据
        const morningPickingData = filterMorningData(pickingData);
        const morningPackingData = filterMorningData(packingData);
        
        log(`筛选出上午时段数据：Picking ${morningPickingData.length} 条，Packing ${morningPackingData.length} 条`);
        
        // 聚合员工数据
        const efficiencyData = aggregateEfficiencyData(morningPickingData, morningPackingData);
        
        if (efficiencyData.length === 0) {
            log('没有找到 00:00-12:00 时段的数据！', 'error');
            alert('没有找到 00:00-12:00 时段的数据！');
            return;
        }
        
        // 导出 Excel
        exportEfficiencyReport(efficiencyData);
        
    } catch (error) {
        log(`生成人效表失败: ${error.message}`, 'error');
        console.error(error);
    }
}

// 筛选上午时段数据（00:00-12:00）
function filterMorningData(data) {
    if (!data || data.length === 0) return [];
    
    return data.filter(item => {
        if (!item.time) return false;
        const hour = item.time.getHours();
        return hour >= 0 && hour < 12;
    });
}

// 聚合人效数据
function aggregateEfficiencyData(pickingData, packingData) {
    log('正在聚合人效数据...', 'info');
    
    const workerMap = new Map();
    
    // 处理 Picking 数据
    if (pickingData && pickingData.length > 0) {
        pickingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingTimes: [],
                    packingTimes: [],
                    pickingQuantity: 0,
                    packingQuantity: 0
                });
            }
            const worker = workerMap.get(item.worker);
            worker.pickingTimes.push(item.time);
            worker.pickingQuantity += item.quantity || 1;
        });
    }
    
    // 处理 Packing 数据
    if (packingData && packingData.length > 0) {
        packingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingTimes: [],
                    packingTimes: [],
                    pickingQuantity: 0,
                    packingQuantity: 0
                });
            }
            const worker = workerMap.get(item.worker);
            worker.packingTimes.push(item.time);
            worker.packingQuantity += item.quantity;
        });
    }
    
    // 计算每个员工的 EWH
    const result = [];
    workerMap.forEach((data, workerName) => {
        // 计算 Picking EWH
        let pickingEwh = 0;
        let pickingWarning = false;
        let pickingDetailedEwh = 0;
        let pickingSegments = [];
        
        if (data.pickingTimes.length > 0) {
            data.pickingTimes.sort((a, b) => a - b);
            const pickingResult = calculateEfficiencyEWH(data.pickingTimes);
            pickingEwh = pickingResult.ewh;
            pickingWarning = pickingResult.warning;
            pickingDetailedEwh = pickingResult.detailedEwh;
            pickingSegments = pickingResult.segments;
        }
        
        // 计算 Packing EWH
        let packingEwh = 0;
        let packingWarning = false;
        let packingDetailedEwh = 0;
        let packingSegments = [];
        
        if (data.packingTimes.length > 0) {
            data.packingTimes.sort((a, b) => a - b);
            const packingResult = calculateEfficiencyEWH(data.packingTimes);
            packingEwh = packingResult.ewh;
            packingWarning = packingResult.warning;
            packingDetailedEwh = packingResult.detailedEwh;
            packingSegments = packingResult.segments;
        }
        
        // 只添加有数据的员工
        if (data.pickingQuantity > 0 || data.packingQuantity > 0) {
            result.push({
                worker: workerName,
                pickingQuantity: data.pickingQuantity,
                packingQuantity: data.packingQuantity,
                pickingEwh: pickingEwh,
                packingEwh: packingEwh,
                pickingWarning: pickingWarning,
                packingWarning: packingWarning,
                pickingDetailedEwh: pickingDetailedEwh,
                packingDetailedEwh: packingDetailedEwh,
                pickingSegments: pickingSegments,
                packingSegments: packingSegments
            });
        }
    });
    
    log(`人效数据聚合完成，共 ${result.length} 名员工`, 'success');
    return result;
}

// 计算人效表的特殊 EWH
// 默认：首尾时间差
// 异常检测：如果中间有超过阈值的间隔，使用精确计算
function calculateEfficiencyEWH(times, thresholdMinutes = 15) {
    if (times.length === 0) return { ewh: 0, warning: false, detailedEwh: 0, segments: [] };
    if (times.length === 1) return { ewh: 0, warning: false, detailedEwh: 0, segments: [] };
    
    // 计算首尾 EWH
    const firstTime = times[0];
    const lastTime = times[times.length - 1];
    const simpleEwh = (lastTime - firstTime) / 3600000; // 转换为小时
    
    // 检测异常并识别工作段
    const threshold = thresholdMinutes * 60 * 1000;
    const segments = [];
    let segmentStart = times[0];
    let hasLongGap = false;
    
    for (let i = 1; i < times.length; i++) {
        const gap = times[i] - times[i - 1];
        if (gap > threshold) {
            // 发现间隔，结束当前工作段
            segments.push({
                start: segmentStart,
                end: times[i - 1]
            });
            segmentStart = times[i];
            hasLongGap = true;
        }
    }
    // 添加最后一个工作段
    segments.push({
        start: segmentStart,
        end: times[times.length - 1]
    });
    
    // 如果有异常，使用精确计算（5分钟阈值）
    let detailedEwh = 0;
    if (hasLongGap) {
        detailedEwh = calculateEWH(times, 5); // 使用现有的精确计算方法
    }
    
    return {
        ewh: round(simpleEwh, 2),
        warning: hasLongGap,
        detailedEwh: hasLongGap ? round(detailedEwh, 2) : 0,
        segments: segments
    };
}

// 格式化时间（HH:MM）
function formatTime(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

// 应用组合表格样式
function applyCombinedSheetStyle(ws, pickingCount, packingCount, pickingWarningRows, packingWarningRows) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    // 拣货标题样式（蓝色背景）
    const pickingTitleStyle = {
        fill: { fgColor: { rgb: "0070C0" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };
    
    // 打包标题样式（红色背景）
    const packingTitleStyle = {
        fill: { fgColor: { rgb: "C00000" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };
    
    // 表头样式（黑底白字）
    const headerStyle = {
        fill: { fgColor: { rgb: "000000" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };
    
    // 数据单元格样式
    const cellStyle = {
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "CCCCCC" } },
            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
            left: { style: "thin", color: { rgb: "CCCCCC" } },
            right: { style: "thin", color: { rgb: "CCCCCC" } }
        }
    };
    
    // 空白单元格样式（浅灰色背景）
    const emptyCellStyle = {
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "F5F5F5" } },
        border: {
            top: { style: "thin", color: { rgb: "CCCCCC" } },
            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
            left: { style: "thin", color: { rgb: "CCCCCC" } },
            right: { style: "thin", color: { rgb: "CCCCCC" } }
        }
    };
    
    // 警告行样式（黄色背景）
    const warningStyle = {
        fill: { fgColor: { rgb: "FFF3CD" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "CCCCCC" } },
            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
            left: { style: "thin", color: { rgb: "CCCCCC" } },
            right: { style: "thin", color: { rgb: "CCCCCC" } }
        }
    };
    
    // 转换警告行号为 Set 方便查找
    const warningRowsSet = new Set([...pickingWarningRows, ...packingWarningRows]);
    
    // 计算行位置
    const pickingTitleRow = 0;
    const pickingHeaderRow = 1;
    const pickingDataStart = 2;
    const pickingDataEnd = pickingDataStart + pickingCount - 1;
    const emptyRow = pickingDataEnd + 1;
    const packingTitleRow = emptyRow + 1;
    const packingHeaderRow = packingTitleRow + 1;
    const packingDataStart = packingHeaderRow + 1;
    
    // 应用样式
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;
            
            // 拣货标题行
            if (R === pickingTitleRow) {
                ws[cellAddress].s = pickingTitleStyle;
            }
            // 拣货表头行
            else if (R === pickingHeaderRow) {
                ws[cellAddress].s = headerStyle;
            }
            // 拣货数据行
            else if (R >= pickingDataStart && R <= pickingDataEnd) {
                // 检查单元格值是否为空
                const cellValue = ws[cellAddress].v;
                const isEmpty = cellValue === '' || cellValue === null || cellValue === undefined || cellValue === 0;
                
                if (warningRowsSet.has(R)) {
                    ws[cellAddress].s = warningStyle;
                } else {
                    ws[cellAddress].s = isEmpty ? emptyCellStyle : cellStyle;
                }
            }
            // 空行
            else if (R === emptyRow) {
                // 不应用样式
            }
            // 打包标题行
            else if (R === packingTitleRow) {
                ws[cellAddress].s = packingTitleStyle;
            }
            // 打包表头行
            else if (R === packingHeaderRow) {
                ws[cellAddress].s = headerStyle;
            }
            // 打包数据行
            else if (R >= packingDataStart) {
                // 检查单元格值是否为空
                const cellValue = ws[cellAddress].v;
                const isEmpty = cellValue === '' || cellValue === null || cellValue === undefined || cellValue === 0;
                
                if (warningRowsSet.has(R)) {
                    ws[cellAddress].s = warningStyle;
                } else {
                    ws[cellAddress].s = isEmpty ? emptyCellStyle : cellStyle;
                }
            }
        }
    }
    
    // 合并标题单元格
    ws['!merges'] = [
        { s: { r: pickingTitleRow, c: 0 }, e: { r: pickingTitleRow, c: 3 } }, // 拣货数据标题
        { s: { r: packingTitleRow, c: 0 }, e: { r: packingTitleRow, c: 3 } }  // 打包数据标题
    ];
    
    // 设置行高
    ws['!rows'] = [];
    for (let i = 0; i <= range.e.r; i++) {
        if (i === pickingTitleRow || i === packingTitleRow) {
            ws['!rows'][i] = { hpt: 30 }; // 标题行
        } else if (i === pickingHeaderRow || i === packingHeaderRow) {
            ws['!rows'][i] = { hpt: 25 }; // 表头行
        } else if (i === emptyRow) {
            ws['!rows'][i] = { hpt: 10 }; // 空行
        } else {
            ws['!rows'][i] = { hpt: 20 }; // 数据行
        }
    }
}

// 导出人效表
function exportEfficiencyReport(data) {
    log('\n正在生成 Excel 文件\n', 'info');
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 准备完整的表格数据
    const tableData = [];
    const pickingWarningRows = [];  // 记录有警告的 Picking 行号
    const packingWarningRows = [];  // 记录有警告的 Packing 行号
    
    // ========== 拣货数据部分 ==========
    tableData.push(['拣货数据', '', '', '']); // 标题行
    tableData.push(['员工', '件数', 'EWH', '备注']); // 表头
    
            let pickingCount = 0;
    data.forEach(item => {
        if (item.pickingQuantity > 0) {
            let ewh = item.pickingEwh;
            let remark = '';
            
            // 生成 Picking 工作时间段信息
            if (item.pickingSegments && item.pickingSegments.length > 0) {
                const timeRanges = item.pickingSegments.map(seg => 
                    `${formatTime(seg.start)}-${formatTime(seg.end)}`
                ).join(' ');
                
                if (item.pickingWarning) {
                    // 如果有异常，EWH 列显示精确 EWH + 10% 补偿
                    const compensatedEwh = round(item.pickingDetailedEwh * 1.1, 2);
                    ewh = compensatedEwh;
                    remark = `工作时间段: ${timeRanges} (原始: ${item.pickingEwh}h, 精确: ${item.pickingDetailedEwh}h, 已补偿10%)`;
                    pickingWarningRows.push(tableData.length); // 记录警告行号
                    log(`${item.worker} [Picking]: 检测到工作中断，精确EWH ${item.pickingDetailedEwh}h → 补偿后 ${compensatedEwh}h`, 'warning');
                } else {
                    remark = `工作时间段: ${timeRanges}`;
                }
            }
            
            tableData.push([
                item.worker,
                item.pickingQuantity,
                ewh,
                remark
            ]);
            pickingCount++;
        }
    });
    
    // 添加空行分隔
    tableData.push(['', '', '', '']);
    
    // ========== 打包数据部分 ==========
    tableData.push(['打包数据', '', '', '']); // 标题行
    tableData.push(['员工', '件数', 'EWH', '备注']); // 表头
    
            let packingCount = 0;
    data.forEach(item => {
        if (item.packingQuantity > 0) {
            let ewh = item.packingEwh;
            let remark = '';
            
            // 生成 Packing 工作时间段信息
            if (item.packingSegments && item.packingSegments.length > 0) {
                const timeRanges = item.packingSegments.map(seg => 
                    `${formatTime(seg.start)}-${formatTime(seg.end)}`
                ).join(' ');
                
                if (item.packingWarning) {
                    // 如果有异常，EWH 列显示精确 EWH + 10% 补偿
                    const compensatedEwh = round(item.packingDetailedEwh * 1.1, 2);
                    ewh = compensatedEwh;
                    remark = `工作时间段: ${timeRanges} (原始: ${item.packingEwh}h, 精确: ${item.packingDetailedEwh}h, 已补偿10%)`;
                    packingWarningRows.push(tableData.length); // 记录警告行号
                    log(`${item.worker} [Packing]: 检测到工作中断，精确EWH ${item.packingDetailedEwh}h → 补偿后 ${compensatedEwh}h`, 'warning');
                } else {
                    remark = `工作时间段: ${timeRanges}`;
                }
            }
            
            tableData.push([
                item.worker,
                item.packingQuantity,
                ewh,
                remark
            ]);
            packingCount++;
        }
    });
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(tableData);
    
    // 设置列宽
    ws['!cols'] = [
        { wch: 25 },  // 员工
        { wch: 10 },  // 件数
        { wch: 10 },  // EWH
        { wch: 70 }   // 备注
    ];
    
    // 应用样式
    applyCombinedSheetStyle(ws, pickingCount, packingCount, pickingWarningRows, packingWarningRows);
    
    XLSX.utils.book_append_sheet(wb, ws, '人效表');
    
    // 生成文件名
    const today = new Date();
    const dateStr = formatDate(today).replace(/\//g, '');
    const fileName = `人效表${dateStr}.xlsx`;
    
    // 下载文件
    XLSX.writeFile(wb, fileName);
    
    log(`\nExcel 文件已生成: ${fileName}`, 'success');
    log(`拣货数据: ${pickingCount} 名员工`, 'info');
    log(`打包数据: ${packingCount} 名员工`, 'info');
    log('黄色标记行表示检测到工作中断，EWH已修正为精确值', 'info');
    log('\n报表生成成功！', 'success');
}

