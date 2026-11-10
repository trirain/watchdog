import { parseLeaveSheet, generateExportFile } from './logic';
import { DEFAULT_TYPES } from './constants';

declare const XLSX: any;

let rawWorkbook: any = null;
let typeList: string[] = DEFAULT_TYPES;

const fileInput = document.getElementById('leaveFile') as HTMLInputElement;
const templateInput = document.getElementById('templateFile') as HTMLInputElement;
const monthContainer = document.getElementById('monthSelectContainer')!;
const monthSelect = document.getElementById('monthSelect') as HTMLSelectElement;
const exportBtn = document.getElementById('exportBtn') as HTMLButtonElement;
const statusText = document.getElementById('statusText')!;

fileInput.addEventListener('change', async (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (!file) return;
    
    statusText.textContent = "正在读取休假表...";
    try {
        const buffer = await file.arrayBuffer();
        // 关键：增加 cellStyles: true 才能读到删除线
        rawWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
        
        const monthSheets = rawWorkbook.SheetNames.filter((name: string) => name.includes("月休假申请"));
        
        monthSelect.innerHTML = '';
        if (monthSheets.length === 0) {
            statusText.textContent = "未在休假表中找到包含“月休假申请”的Sheet";
            return;
        }

        monthSheets.forEach((name: string) => {
            const option = document.createElement('option');
            option.value = name;
            option.text = name.replace('休假申请', ''); 
            monthSelect.appendChild(option);
        });

        monthContainer.hidden = false;
        statusText.textContent = "休假表读取成功，请选择月份";
        checkReady();
    } catch (err: any) {
        console.error(err);
        statusText.textContent = "读取失败：" + err.message;
    }
});

// ... 后面 templateInput 和 导出逻辑保持不变，使用你之前成功的代码即可 ...
templateInput.addEventListener('change', async (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (!file) return;
    statusText.textContent = "正在读取模板...";
    try {
        const buffer = await file.arrayBuffer();
        const templateWb = XLSX.read(buffer, { type: 'array' });
        const typeSheetName = templateWb.SheetNames.find((name: string) => name.includes('类型列表') || name.includes('类型'));
        if (typeSheetName) {
            const typeSheet = templateWb.Sheets[typeSheetName];
            const rows = XLSX.utils.sheet_to_json(typeSheet, { header: 1 }) as any[][];
            typeList = rows.flat().filter(v => v && String(v).trim().length > 0).map(v => String(v).trim());
            statusText.textContent = `已从模板读取 ${typeList.length} 个标准锁班类型`;
        } else {
            typeList = [...DEFAULT_TYPES];
            statusText.textContent = "未找到类型列表Sheet，已使用内置字典";
        }
    } catch (err: any) {
        console.error(err);
        statusText.textContent = "读取模板失败：" + err.message;
        typeList = [...DEFAULT_TYPES];
    }
    checkReady();
});

function checkReady() {
    exportBtn.disabled = !(rawWorkbook && typeList.length > 0);
}

exportBtn.addEventListener('click', () => {
    const selectedSheetName = monthSelect.value;
    if (!selectedSheetName) return;
    statusText.textContent = "正在处理，请稍候...";
    try {
        const sheet = rawWorkbook.Sheets[selectedSheetName];
        const records = parseLeaveSheet(sheet, typeList);
        if (records.length === 0) {
            statusText.textContent = "未在该月份找到有效的排班数据。";
            return;
        }
        const wb = generateExportFile(records);
        XLSX.writeFile(wb, `锁班数据_${selectedSheetName}.xlsx`);
        statusText.textContent = `导出成功！共处理 ${records.length} 条记录。请查看标红和标黄的自动修正项。`;
    } catch (err: any) {
        console.error(err);
        statusText.textContent = "处理失败: " + err.message;
    }
});
