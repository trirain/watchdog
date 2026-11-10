declare const XLSX: any;

export interface LeaveRecord {
  empId: string;
  name: string;
  rawType: string;
  cleanType: string;
  startDate: string;
  endDate: string;
  isModified: boolean;
  isBirthdaySplit: boolean; // 新增：是否为生日拆分出的行
}

// 辅助函数：将各种日期格式统一转换为 YYYY/M/D
function formatDate(dateVal: any): string {
  if (!dateVal) return "";
  // 如果是 Excel 的 Date 对象
  if (dateVal instanceof Date) {
    return `${dateVal.getFullYear()}/${dateVal.getMonth() + 1}/${dateVal.getDate()}`;
  }
  // 如果是字符串（如 "2026-11-05 00:00:00" 或 "2026/11/5"）
  const str = String(dateVal).trim();
  if (!str) return "";
  // 尝试匹配 YYYY-MM-DD 或 YYYY/M/D
  const match = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    return `${match[1]}/${parseInt(match[2], 10)}/${parseInt(match[3], 10)}`;
  }
  return str;
}

// 辅助函数：解析备注中的生日（如 07birth, 10、13birth）
function extractBirthdays(remark: string): number[] {
  if (!remark) return [];
  const days: number[] = [];
  // 匹配类似于 07birth, 7bir, 10、13birth
  const regex = /(\d+(?:[、,，]\d+)*)\s*bir(?:th)?/gi;
  let match;
  while ((match = regex.exec(remark)) !== null) {
    // 提取数字部分，用逗号或顿号分隔
    const nums = match[1].split(/[、,，]/);
    nums.forEach(n => {
      const day = parseInt(n, 10);
      if (day >= 1 && day <= 31) days.push(day);
    });
  }
  return days;
}

// 辅助函数：根据生日拆分日期区间
function splitDatesByBirthday(startStr: string, endStr: string, birthdays: number[]): { start: string, end: string }[] {
  if (birthdays.length === 0) return [{ start: startStr, end: endStr }];
  
  const startParts = startStr.split('/').map(Number);
  const endParts = endStr.split('/').map(Number);
  if (startParts.length !== 3 || endParts.length !== 3) return [{ start: startStr, end: endStr }];

  const startDate = new Date(startParts[0], startParts[1] - 1, startParts[2]);
  const endDate = new Date(endParts[0], endParts[1] - 1, endParts[2]);
  
  // 如果格式解析失败，直接返回原值
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    return [{ start: startStr, end: endStr }];
  }

  const result: { start: string, end: string }[] = [];
  let currentStart = new Date(startDate);

  // 遍历区间的每一天
  while (currentStart <= endDate) {
    // 如果当前天是生日，说明上一段要在这里断开
    if (birthdays.includes(currentStart.getDate())) {
      currentStart.setDate(currentStart.getDate() + 1);
      continue;
    }

    // 寻找连续的非生日区间
    let currentEnd = new Date(currentStart);
    while (currentEnd < endDate) {
      const nextDay = new Date(currentEnd);
      nextDay.setDate(nextDay.getDate() + 1);
      if (birthdays.includes(nextDay.getDate())) break;
      currentEnd = nextDay;
    }

    result.push({
      start: `${currentStart.getFullYear()}/${currentStart.getMonth() + 1}/${currentStart.getDate()}`,
      end: `${endDate.getFullYear() === currentEnd.getFullYear() ? '' : currentEnd.getFullYear() + '/'}${currentEnd.getFullYear() === currentStart.getFullYear() ? '' : ''}${currentEnd.getMonth() + 1}/${currentEnd.getDate()}`,
    });
    // 上面这行结束日期的格式化有点绕，直接写死 YYYY/M/D 最稳
    result[result.length - 1].end = `${currentEnd.getFullYear()}/${currentEnd.getMonth() + 1}/${currentEnd.getDate()}`;

    currentStart = new Date(currentEnd);
    currentStart.setDate(currentStart.getDate() + 1);
  }

  return result.length > 0 ? result : [{ start: startStr, end: endStr }];
}

export function parseLeaveSheet(worksheet: any, typeList: string[]): LeaveRecord[] {
  const records: LeaveRecord[] = [];
  const merges = worksheet['!merges'] || [];
  
  // 1. 构建合并单元格映射：记录哪些坐标被合并了，它们的值应该去哪里取
  const mergeMap = new Map<string, { r: number, c: number }>();
  merges.forEach((m: any) => {
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        mergeMap.set(`${r},${c}`, { r: m.s.r, c: m.s.c });
      }
    }
  });

  // 获取有效的数据范围
  const range = XLSX.utils.decode_range(worksheet['!ref'] || "A1:A1");
  let colIndexes: { empId: number, name: number, type: number, start: number, end: number, remark: number }[] = [];
  let headerRow = -1;

  // 2. 扫描并定位表头
  for (let r = range.s.r; r <= range.e.r; r++) {
    let found = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && String(cell.v).includes("员工号")) {
        // 找到表头行，提取各列偏移
        colIndexes = [];
        for (let cc = c; cc <= range.e.c; cc++) {
          const headerCell = worksheet[XLSX.utils.encode_cell({ r, c: cc })];
          if (headerCell && String(headerCell.v).includes("员工号")) {
            colIndexes.push({
              empId: cc, name: cc + 1, type: cc + 2, start: cc + 3, end: cc + 4, remark: cc + 8 // 假设备注在往后第8列
            });
          }
        }
        if (colIndexes.length > 0) {
          headerRow = r;
          found = true;
          break;
        }
      }
    }
    if (found) break;
  }

  if (headerRow === -1 || colIndexes.length === 0) return records;

  // 3. 遍历数据行
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    for (const cols of colIndexes) {
      // 获取真实单元格（考虑合并单元格的情况）
      const getRealCell = (row: number, col: number) => {
        const mapped = mergeMap.get(`${row},${col}`);
        if (mapped && (mapped.r !== row || mapped.c !== col)) {
          return worksheet[XLSX.utils.encode_cell({ r: mapped.r, c: mapped.c })];
        }
        return worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
      };

      const empIdCell = getRealCell(r, cols.empId);
      const nameCell = getRealCell(r, cols.name);
      const typeCell = getRealCell(r, cols.type);
      const startCell = getRealCell(r, cols.start);
      const endCell = getRealCell(r, cols.end);
      const remarkCell = getRealCell(r, cols.remark);

      const empId = empIdCell ? String(empIdCell.v).trim() : "";
      const name = nameCell ? String(nameCell.v).trim() : "";
      const rawType = typeCell ? String(typeCell.v).trim() : "";
      
      // 过滤掉示例行、空行和说明行
      if (!empId || empId === "000000" || name === "飞行员" || rawType === "请假类型" || rawType.includes("说明")) {
        continue;
      }

      // 需求2：判断该行是否被划掉 (strike)
      // 只要员工号、姓名、类型的任意一个被划掉，就视为整行被划掉
      const isStruck = (cell: any) => cell && cell.s && cell.s.font && cell.s.font.strike;
      if (isStruck(empIdCell) || isStruck(nameCell) || isStruck(typeCell)) {
        continue; // 跳过被划掉的行
      }

      // 解析日期和备注
      const startStr = formatDate(startCell ? startCell.v : "");
      const endStr = formatDate(endCell ? endCell.v : "");
      const remark = remarkCell ? String(remarkCell.v).trim() : "";

      // 需求1：清洗锁班类型
      let cleanType = rawType;
      let isModified = false;
      if (!typeList.includes(rawType) && rawType !== "") {
        const matchedType = matchType(rawType, typeList);
        if (matchedType) {
          cleanType = matchedType;
          isModified = true;
        } else {
          cleanType = rawType;
          isModified = true;
        }
      }

      // 需求4：处理生日拆分
      const birthdays = extractBirthdays(remark);
      if (birthdays.length > 0 && startStr && endStr) {
        const dateSegments = splitDatesByBirthday(startStr, endStr, birthdays);
        dateSegments.forEach(seg => {
          records.push({
            empId, name, rawType, cleanType, 
            startDate: seg.start, endDate: seg.end, 
            isModified, 
            isBirthdaySplit: true // 标记为拆分行，后续需要标黄
          });
        });
      } else {
        // 正常情况，不拆分
        if (startStr || endStr) {
          records.push({
            empId, name, rawType, cleanType, 
            startDate: startStr, endDate: endStr, 
            isModified, 
            isBirthdaySplit: false
          });
        }
      }
    }
  }
  return records;
}

function matchType(raw: string, typeList: string[]): string | null {
    if (typeList.includes(raw)) return raw;
    for (const t of typeList) {
        const parts = t.split('-');
        const chinesePart = parts.length > 1 ? parts[parts.length - 1] : t;
        if (raw.includes(chinesePart) || chinesePart.includes(raw)) {
            return t;
        }
    }
    return null;
}

export function generateExportFile(records: LeaveRecord[]) {
    const data = [["员工号", "姓名", "锁班类型", "开始日期", "结束日期"]];
    records.forEach(r => data.push([r.empId, r.name, r.cleanType, r.startDate, r.endDate]));

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];

    const range = XLSX.utils.decode_range(ws['!ref']!);
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[XLSX.utils.encode_cell({ r, c })];
            if (!cell) continue;
            
            // 基础样式
            cell.s = {
                font: { name: "宋体", sz: 11 },
                alignment: { vertical: "center", horizontal: "center" },
                border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} }
            };
            
            if (r === 0) {
                // 表头样式
                cell.s.font.bold = true;
                cell.s.fill = { fgColor: { rgb: "E0E0E0" } };
            } else {
                const record = records[r - 1];
                if (record.isModified) {
                    // 类型被清洗，标红
                    cell.s.font.bold = true;
                    cell.s.font.color = { rgb: "FF0000" };
                    cell.s.fill = { fgColor: { rgb: "FFF2F2" } };
                }
                if (record.isBirthdaySplit) {
                    // 生日拆分，标黄（为了醒目，全部单元格都标黄）
                    cell.s.fill = { fgColor: { rgb: "FFFF00" } }; 
                    cell.s.font.bold = true;
                }
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "锁班数据");
    return wb;
}
