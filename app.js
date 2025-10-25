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
    // 添加更严格的数据验证
    if (!timestamps || !Array.isArray(timestamps) || timestamps.length < 2) {
        return 0;
    }
    
    // 过滤并排序时间戳
    const times = timestamps
        .filter(t => t && t instanceof Date && !isNaN(t.getTime()))
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
        if (s && e && s instanceof Date && e instanceof Date) {
            totalMs += e - s;
        }
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
    
    let preshipmentWorkers = [];
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
        if (w.ewh > 0) {
            const calculatedUPH = round(w.quantity / w.ewh, 2);
            log(`    预期 UPH: ${calculatedUPH}`);
        }
    });
    log('');
    
    // 禁用按钮
    const btn = document.getElementById('generateBtn');
    btn.disabled = true;
    btn.innerHTML = '正在生成中...';
    
    setTimeout(() => {
        try {
            // 验证数据完整性
            if (!pickingData && !packingData) {
                throw new Error('没有可用的数据文件');
            }
            
            // 验证 preshipmentWorkers 数据
            if (!Array.isArray(preshipmentWorkers)) {
                log('警告: preshipmentWorkers 不是数组，将使用空数组', 'warning');
                preshipmentWorkers = [];
            }
            
            // 汇总数据
            const report = aggregateData(pickingData, packingData, date, time, preshipmentWorkers);
            
            // 验证报告数据
            if (!Array.isArray(report)) {
                throw new Error('数据汇总失败，报告格式错误');
            }
            
            if (report.length === 0) {
                log('警告: 没有生成任何员工数据，请检查文件内容', 'warning');
            }
            
            // 生成 Excel
            exportToExcel(report);
            
            log('\n报表生成成功', 'success');
            log(`统计记录数: ${report.length} 条`, 'success');
            log('─'.repeat(60) + '\n', 'info');
            
        } catch (error) {
            log(`\n生成报表失败: ${error.message}`, 'error');
            log(`错误详情: ${error.stack || '无详细信息'}`, 'error');
            console.error('报表生成错误:', error);
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
    if (pickingData && Array.isArray(pickingData) && pickingData.length > 0) {
        const totalPickQuantity = pickingData.reduce((sum, item) => sum + (item.quantity || 0), 0);
        const singleItems = pickingData.filter(item => item && item.itemType === 'single');
        const multiItems = pickingData.filter(item => item && item.itemType === 'multi');
        const singleCount = singleItems.reduce((sum, item) => sum + (item.quantity || 0), 0);
        const multiCount = multiItems.reduce((sum, item) => sum + (item.quantity || 0), 0);
        
        log(`  处理 Picking 数据 (${pickingData.length} 条记录, 总数量: ${totalPickQuantity})`);
        log(`    - 单品: ${singleCount} 件 (${singleItems.length} 条)`);
        log(`    - 多品: ${multiCount} 件 (${multiItems.length} 条)`);
        
        pickingData.forEach(item => {
            // 添加数据验证
            if (!item || !item.worker) {
                log(`  警告: 跳过无效的 Picking 记录`, 'warning');
                return;
            }
            
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
            
            // 确保时间数据有效
            if (item.time && item.time instanceof Date) {
                worker.pickingTimes.push(item.time);
            }
            worker.pickCount += item.quantity || 1;
            
            // 分类统计单品和多品
            if (item.itemType === 'single') {
                if (item.time && item.time instanceof Date) {
                    worker.pickingSingleTimes.push(item.time);
                }
                worker.pickSingleCount += item.quantity || 1;
            } else if (item.itemType === 'multi') {
                if (item.time && item.time instanceof Date) {
                    worker.pickingMultiTimes.push(item.time);
                }
                worker.pickMultiCount += item.quantity || 1;
            }
        });
        log(`  Picking 数据处理完成`);
    }
    
    // 处理 Packing 数据
    if (packingData && Array.isArray(packingData) && packingData.length > 0) {
        const totalPackQuantity = packingData.reduce((sum, item) => sum + (item.quantity || 0), 0);
        log(`  处理 Packing 数据 (${packingData.length} 条记录, 总数量: ${totalPackQuantity})`);
        
        packingData.forEach(item => {
            // 添加数据验证
            if (!item || !item.worker) {
                log(`  警告: 跳过无效的 Packing 记录`, 'warning');
                return;
            }
            
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
            
            // 确保时间数据有效
            if (item.time && item.time instanceof Date) {
                worker.packingTimes.push(item.time);
            }
            worker.packCount += item.quantity || 1;
        });
        log(`  Packing 数据处理完成`);
    }
    
    // 计算 EWH 和 UPH
    log('\n  正在计算 EWH 和 UPH\n');
    const report = [];
    
    workerMap.forEach((data, worker) => {
        // 确保所有时间数组都存在且为数组
        const pickingSingleTimes = Array.isArray(data.pickingSingleTimes) ? data.pickingSingleTimes : [];
        const pickingMultiTimes = Array.isArray(data.pickingMultiTimes) ? data.pickingMultiTimes : [];
        const packingTimes = Array.isArray(data.packingTimes) ? data.packingTimes : [];
        
        // 计算 Picking 单品和多品 EWH（精准计算，无补偿）
        const pickingSingleEWH = pickingSingleTimes.length > 0 
            ? calculateEWH(pickingSingleTimes)  // 精准计算，无补偿
            : 0;
        const pickingMultiEWH = pickingMultiTimes.length > 0 
            ? calculateEWH(pickingMultiTimes)   // 精准计算，无补偿
            : 0;
        
        // 计算总 Picking EWH
        const pickingEWH = pickingSingleEWH + pickingMultiEWH;
        
        // 计算 Packing EWH
        const packingEWH = packingTimes.length > 0 ? calculateEWH(packingTimes) : 0;
        
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
            pick: data.pickCount > 0 ? data.pickCount : '',
            pack: data.packCount > 0 ? data.packCount : '',
            box: '',
            Preshipment: '',
            'Packing UPH': packingUPH > 0 ? round(packingUPH, 2) : '',
            'Picking UPH': (pickingSingleUPH + pickingMultiUPH) > 0 ? round((data.pickCount / pickingEWH), 2) : '',
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
                    existing['Preship UPH'] = '';
                    log(`  ${ps.name}: Preshipment=${ps.quantity}, UPH=空 (未填写EWH)`);
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
                    pick: '',
                    pack: '',
                    box: '',
                    Preshipment: ps.quantity,
                    'Packing UPH': '',
                    'Picking UPH': '',
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

// 功能选择弹窗相关函数
function openFunctionModal() {
    document.getElementById('functionModal').style.display = 'block';
}

function closeFunctionModal() {
    document.getElementById('functionModal').style.display = 'none';
}

function executeSelectedFunctions() {
    const dailyReport = document.getElementById('dailyReport').checked;
    const efficiencyReport = document.getElementById('efficiencyReport').checked;
    
    if (!dailyReport && !efficiencyReport) {
        alert('请至少选择一个功能！');
        return;
    }
    
    // 关闭弹窗
    closeFunctionModal();
    
    // 执行选中的功能
    if (dailyReport) {
        log('\n开始执行：每日工作报表', 'info');
        generateReport();
    }
    
    if (efficiencyReport) {
        // 延迟执行人效表，避免同时执行造成冲突
        setTimeout(() => {
            log('\n开始执行：人效表', 'info');
            generateEfficiencyReport();
        }, dailyReport ? 2000 : 0);
    }
}

// 点击弹窗外部关闭
window.onclick = function(event) {
    const employeeModal = document.getElementById('employeeModal');
    const functionModal = document.getElementById('functionModal');
    
    if (event.target === employeeModal) {
        closeEmployeeManager();
    }
    
    if (event.target === functionModal) {
        closeFunctionModal();
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
    
    // 处理 Picking 数据（分单品和多品）
    if (pickingData && pickingData.length > 0) {
        pickingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingSingleTimes: [],
                    pickingMultiTimes: [],
                    packingTimes: [],
                    pickingSingleQuantity: 0,
                    pickingMultiQuantity: 0,
                    packingQuantity: 0
                });
            }
            const worker = workerMap.get(item.worker);
            
            // 根据itemType分类
            if (item.itemType === 'single') {
                worker.pickingSingleTimes.push(item.time);
                worker.pickingSingleQuantity += item.quantity || 1;
            } else if (item.itemType === 'multi') {
                worker.pickingMultiTimes.push(item.time);
                worker.pickingMultiQuantity += item.quantity || 1;
            }
        });
    }
    
    // 处理 Packing 数据
    if (packingData && packingData.length > 0) {
        packingData.forEach(item => {
            if (!workerMap.has(item.worker)) {
                workerMap.set(item.worker, {
                    worker: item.worker,
                    pickingSingleTimes: [],
                    pickingMultiTimes: [],
                    packingTimes: [],
                    pickingSingleQuantity: 0,
                    pickingMultiQuantity: 0,
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
        // 计算 Picking 单品 EWH（只使用精准计算）
        let pickingSingleEwh = 0;
        let pickingSingleSegments = [];
        
        if (data.pickingSingleTimes.length > 0) {
            data.pickingSingleTimes.sort((a, b) => a - b);
            const singleResult = calculateEfficiencyEWH(data.pickingSingleTimes);
            pickingSingleEwh = singleResult.detailedEwh; // 直接使用精准EWH
            pickingSingleSegments = singleResult.segments;
        }
        
        // 计算 Picking 多品 EWH（只使用精准计算）
        let pickingMultiEwh = 0;
        let pickingMultiSegments = [];
        
        if (data.pickingMultiTimes.length > 0) {
            data.pickingMultiTimes.sort((a, b) => a - b);
            const multiResult = calculateEfficiencyEWH(data.pickingMultiTimes);
            pickingMultiEwh = multiResult.detailedEwh; // 直接使用精准EWH
            pickingMultiSegments = multiResult.segments;
        }
        
        // 计算 Packing EWH（只使用精准计算）
        let packingEwh = 0;
        let packingSegments = [];
        
        if (data.packingTimes.length > 0) {
            data.packingTimes.sort((a, b) => a - b);
            const packingResult = calculateEfficiencyEWH(data.packingTimes);
            packingEwh = packingResult.detailedEwh; // 直接使用精准EWH
            packingSegments = packingResult.segments;
        }
        
        // 只添加有数据的员工
        if (data.pickingSingleQuantity > 0 || data.pickingMultiQuantity > 0 || data.packingQuantity > 0) {
            result.push({
                worker: workerName,
                pickingSingleQuantity: data.pickingSingleQuantity,
                pickingMultiQuantity: data.pickingMultiQuantity,
                packingQuantity: data.packingQuantity,
                pickingSingleEwh: pickingSingleEwh,
                pickingMultiEwh: pickingMultiEwh,
                packingEwh: packingEwh,
                pickingSingleSegments: pickingSingleSegments,
                pickingMultiSegments: pickingMultiSegments,
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
function calculateEfficiencyEWH(times, thresholdMinutes = 5) {
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
    
    // 始终计算精确 EWH（5分钟阈值）
    const detailedEwh = calculateEWH(times, 5); // 使用现有的精确计算方法
    
    return {
        ewh: round(simpleEwh, 2),
        warning: hasLongGap,
        detailedEwh: round(detailedEwh, 2), // 始终返回详细EWH
        segments: segments
    };
}

// 格式化时间（HH:MM）
function formatTime(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

// 生成工作时间段的可视化图标（用于备注列）
function generateTimeSegmentVisual(segments) {
    if (!segments || segments.length === 0) return '';
    
    // 创建可视化图标 - 使用柱状图表示工作时间段
    const result = [];
    
    segments.forEach((seg, index) => {
        const start = seg.start;
        const end = seg.end;
        const duration = Math.round((end - start) / (1000 * 60)); // 分钟数
        
        // 生成每个时间段的详细信息
        const timeStr = `${formatTime(start)}-${formatTime(end)}`;
        result.push(`${timeStr} (${Math.round(duration)}分钟)`);
    });
    
    return result.join('\n');
}

// 应用组合表格样式
function applyCombinedSheetStyle(ws, pickingCount, packingCount) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    // PICKING 标题样式（蓝色背景）
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
    
    // PACKING 标题样式（红色背景）
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
    
    // 备注单元格样式（支持自动换行和多行显示）
    const remarkCellStyle = {
        alignment: { horizontal: "left", vertical: "top", wrapText: true },
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
            
            // PICKING 标题行
            if (R === pickingTitleRow) {
                ws[cellAddress].s = pickingTitleStyle;
            }
            // PICKING 表头行
            else if (R === pickingHeaderRow) {
                ws[cellAddress].s = headerStyle;
            }
            // PICKING 数据行
            else if (R >= pickingDataStart && R <= pickingDataEnd) {
                const cellValue = ws[cellAddress].v;
                const isEmpty = cellValue === '' || cellValue === null || cellValue === undefined || cellValue === 0;
                
                // 备注列（第7列，索引为7）使用特殊样式
                if (C === 7) {
                    ws[cellAddress].s = remarkCellStyle;
                } else {
                    ws[cellAddress].s = isEmpty ? emptyCellStyle : cellStyle;
                }
            }
            // 空行
            else if (R === emptyRow) {
                // 不应用样式
            }
            // PACKING 标题行
            else if (R === packingTitleRow) {
                ws[cellAddress].s = packingTitleStyle;
            }
            // PACKING 表头行
            else if (R === packingHeaderRow) {
                ws[cellAddress].s = headerStyle;
            }
            // PACKING 数据行
            else if (R >= packingDataStart) {
                const cellValue = ws[cellAddress].v;
                const isEmpty = cellValue === '' || cellValue === null || cellValue === undefined || cellValue === 0;
                
                // 备注列（第7列，索引为7）使用特殊样式
                if (C === 7) {
                    ws[cellAddress].s = remarkCellStyle;
                } else {
                    ws[cellAddress].s = isEmpty ? emptyCellStyle : cellStyle;
                }
            }
        }
    }
    
    // 合并标题单元格
    ws['!merges'] = [
        { s: { r: pickingTitleRow, c: 0 }, e: { r: pickingTitleRow, c: 7 } },   // PICKING 标题
        { s: { r: packingTitleRow, c: 0 }, e: { r: packingTitleRow, c: 7 } }     // PACKING 标题
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
    
    // 从数据中提取日期
    let currentDate = '';
    if (data.length > 0) {
        // 尝试从第一个有效数据中提取日期
        const firstItem = data.find(item => 
            (item.pickingSingleSegments && item.pickingSingleSegments.length > 0) ||
            (item.pickingMultiSegments && item.pickingMultiSegments.length > 0) ||
            (item.packingSegments && item.packingSegments.length > 0)
        );
        
        if (firstItem) {
            let firstSegment = null;
            if (firstItem.pickingSingleSegments && firstItem.pickingSingleSegments.length > 0) {
                firstSegment = firstItem.pickingSingleSegments[0];
            } else if (firstItem.pickingMultiSegments && firstItem.pickingMultiSegments.length > 0) {
                firstSegment = firstItem.pickingMultiSegments[0];
            } else if (firstItem.packingSegments && firstItem.packingSegments.length > 0) {
                firstSegment = firstItem.packingSegments[0];
            }
            
            if (firstSegment && firstSegment.start) {
                currentDate = formatDate(firstSegment.start);
            }
        }
    }
    
    // 如果没有找到日期，使用当前日期作为后备
    if (!currentDate) {
        const today = new Date();
        currentDate = formatDate(today);
    }
    
    // ========== PICKING 数据部分 ==========
    tableData.push(['PICKING', '', '', '', '', '', '', '']); // 标题行
    tableData.push(['日期', '', '员工', '件数', 'EWH', 'UPH', '多品', '备注']); // 表头
    
    let pickingCount = 0;
    data.forEach(item => {
        // 处理单品数据（只使用精准计算）
        if (item.pickingSingleQuantity > 0) {
            const ewh = round(item.pickingSingleEwh, 2);
            let remark = '';
            
            // 生成工作时间段信息（含可视化图标）
            if (item.pickingSingleSegments && item.pickingSingleSegments.length > 0) {
                remark = generateTimeSegmentVisual(item.pickingSingleSegments);
            }
            
            // 计算 UPH
            const uph = ewh > 0 ? round(item.pickingSingleQuantity / ewh, 2) : '';
            
            tableData.push([
                currentDate, // A栏：日期
                '', // B栏留空
                item.worker, // C栏：员工
                item.pickingSingleQuantity, // D栏：件数
                ewh, // E栏：EWH
                uph, // F栏：UPH
                'No', // G栏：多品标记
                remark // H栏：备注
            ]);
            pickingCount++;
        }
        
        // 处理多品数据（只使用精准计算）
        if (item.pickingMultiQuantity > 0) {
            const ewh = round(item.pickingMultiEwh, 2);
            let remark = '';
            
            // 生成工作时间段信息（含可视化图标）
            if (item.pickingMultiSegments && item.pickingMultiSegments.length > 0) {
                remark = generateTimeSegmentVisual(item.pickingMultiSegments);
            }
            
            // 计算 UPH
            const uph = ewh > 0 ? round(item.pickingMultiQuantity / ewh, 2) : '';
            
            tableData.push([
                currentDate, // A栏：日期
                '', // B栏留空
                item.worker, // C栏：员工
                item.pickingMultiQuantity, // D栏：件数
                ewh, // E栏：EWH
                uph, // F栏：UPH
                'Yes', // G栏：多品标记
                remark // H栏：备注
            ]);
            pickingCount++;
        }
    });
    
    // 添加空行分隔
    tableData.push(['', '', '', '', '', '', '', '']);
    
    // ========== PACKING 数据部分 ==========
    tableData.push(['PACKING', '', '', '', '', '', '', '']); // 标题行
    tableData.push(['日期', '', '员工', '件数', 'EWH', 'UPH', '多品', '备注']); // 表头
    
    let packingCount = 0;
    data.forEach(item => {
        if (item.packingQuantity > 0) {
            const ewh = round(item.packingEwh, 2);
            let remark = '';
            
            // 生成 Packing 工作时间段信息（含可视化图标）
            if (item.packingSegments && item.packingSegments.length > 0) {
                remark = generateTimeSegmentVisual(item.packingSegments);
            }
            
            // 计算 UPH
            const uph = ewh > 0 ? round(item.packingQuantity / ewh, 2) : '';
            
            tableData.push([
                currentDate, // A栏：日期
                '', // B栏留空
                item.worker, // C栏：员工
                item.packingQuantity, // D栏：件数
                ewh, // E栏：EWH
                uph, // F栏：UPH
                'No', // G栏：Packing 暂时标记为 No，后续可以扩展为多品功能
                remark // H栏：备注
            ]);
            packingCount++;
        }
    });
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(tableData);
    
    // 设置列宽
    ws['!cols'] = [
        { wch: 12 },  // A栏：日期
        { wch: 10 },  // B栏：留空
        { wch: 25 },  // C栏：员工
        { wch: 10 },  // D栏：件数
        { wch: 10 },  // E栏：EWH
        { wch: 10 },  // F栏：UPH
        { wch: 8 },   // G栏：多品
        { wch: 50 }   // H栏：备注（工作时间段）
    ];
    
    // 应用样式
    applyCombinedSheetStyle(ws, pickingCount, packingCount);
    
    XLSX.utils.book_append_sheet(wb, ws, '人效表');
    
    // 生成文件名（使用数据日期）
    const dateStr = currentDate.replace(/\//g, '');
    const fileName = `人效表${dateStr}.xlsx`;
    
    // 下载文件
    XLSX.writeFile(wb, fileName);
    
    log(`\nExcel 文件已生成: ${fileName}`, 'success');
    log(`PICKING: ${pickingCount} 条记录 (单品和多品合并)`, 'info');
    log(`PACKING: ${packingCount} 条记录`, 'info');
    log('多品列: No=单品, Yes=多品', 'info');
    log('所有员工使用精准EWH计算，无补偿机制', 'info');
    log('\n报表生成成功！', 'success');
}

