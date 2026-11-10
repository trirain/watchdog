export interface LeaveRecord {
  empId: string;
  name: string;
  rawType: string;
  cleanType: string;
  startDate: string;
  endDate: string;
  isModified: boolean;
}

export function parseLeaveSheet(worksheet: any, typeList: string[]): LeaveRecord[] {
  // 将 sheet 转为二维数组以便遍历
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" }) as any[][];
  const records: LeaveRecord[] = [];
  
  let colIndexes: { empId: number, name: number, type: number, start: number, end: number }[] = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    
    // 寻找表头行 (包含"员工号"的单元格)
    const empIdIdx = row.findIndex((cell: any) => String(cell).includes("员工号"));
    if (empIdIdx !== -1) {
       colIndexes = [];
       row.forEach((cell: any, colIdx: number) => {
           if (String(cell).includes("员工号")) {
               // 标准模板列偏移：员工号, 姓名, 请假类型, 开始日期, 结束日期
               colIndexes.push({
                   empId: colIdx, name: colIdx + 1, type: colIdx + 2, start: colIdx + 3, end: colIdx + 4
               });
           }
       });
       continue;
    }

    // 解析数据行
    if (colIndexes.length > 0 && row.length > 0) {
        for (const cols of colIndexes) {
            const empId = String(row[cols.empId] || "").trim();
            const name = String(row[cols.name] || "").trim();
            const rawType = String(row[cols.type] || "").trim();
            const startDate = String(row[cols.start] || "").trim().split(" ")[0]; // 去掉时间部分
            const endDate = String(row[cols.end] || "").trim().split(" ")[0];

            // 过滤掉空行、示例行、说明行
            if (!empId || empId === "000000" || name === "飞行员" || rawType === "请假类型" || rawType.includes("说明")) {
                continue;
            }
            
            let cleanType = rawType;
            let isModified = false;

            // 清洗逻辑
            if (!typeList.includes(rawType) && rawType !== "") {
                const matchedType = matchType(rawType, typeList);
                if (matchedType) {
                    cleanType = matchedType;
                    isModified = true; 
                } else {
                    cleanType = rawType; 
                    isModified = true; // 未匹配上的也标红提醒人工核对
                }
            }

            records.push({ empId, name, rawType, cleanType, startDate, endDate, isModified });
        }
    }
  }
  return records;
}

// 模糊匹配清洗规则（例如把 "健康疗养" 匹配成 "RECU_LVE-健康疗养"）
function matchType(raw: string, typeList: string[]): string | null {
    if (typeList.includes(raw)) return raw;
    for (const t of typeList) {
        // 取 "-" 后面的中文部分进行匹配
        const parts = t.split('-');
        const chinesePart = parts.length > 1 ? parts[parts.length - 1] : t;
        if (raw.includes(chinesePart) || chinesePart.includes(raw)) {
            return t;
        }
    }
    return null;
}

// 生成带样式的 Excel
export function generateExportFile(records: LeaveRecord[]) {
    const data = [["员工号", "姓名", "锁班类型", "开始日期", "结束日期"]];
    records.forEach(r => data.push([r.empId, r.name, r.cleanType, r.startDate, r.endDate]));

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];

    // 应用样式 (xlsx-js-style 语法)
    const range = XLSX.utils.decode_range(ws['!ref']!);
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[XLSX.utils.encode_cell({ r, c })];
            if (!cell) continue;
            
            cell.s = {
                font: { name: "宋体", sz: 11 },
                alignment: { vertical: "center", horizontal: "center" },
                border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} }
            };
            
            if (r === 0) {
                cell.s.font.bold = true;
                cell.s.fill = { fgColor: { rgb: "E0E0E0" } };
            } else if (records[r - 1].isModified) {
                // 如果进行了自动清洗，加粗、标红、加浅红底色
                cell.s.font.bold = true;
                cell.s.font.color = { rgb: "FF0000" };
                cell.s.fill = { fgColor: { rgb: "FFF2F2" } };
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "锁班数据");
    return wb;
}
