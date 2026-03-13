require('dotenv').config();
const express = require('express');
const crypto = require('@wecom/crypto');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// ---------- 配置（与企业微信、智能表格后台一致）----------
const TOKEN = process.env.WECOM_TOKEN || 'IQCLf5VMl31IenTIoPk6953';
const AES_KEY = process.env.WECOM_ENCODING_AES_KEY || 'x62C4zUbz8kGWunRHkN8m3t9nyDkzO8zELS3AtcWQ7f';
const SHEET_WEBHOOK =
  process.env.SHEET_WEBHOOK_URL ||
  'https://qyapi.weixin.qq.com/cgi-bin/wedoc/smartsheet/webhook?key=1V8oeRYw2EhwSrkhRkY5LkAbJuQRIGO96pMrj7CKbtrsVBkyfdmfhtvgtpoif0YjEKcACNe4ukyDmxgpIHSWH0vkP7wpz5aRbqNJQ51iK8fI';

// 汇总表 Webhook（可选）：用于「按角色合计」的明细，每条消息写一行，表格内用分组汇总得合计
const SUMMARY_WEBHOOK = process.env.SUMMARY_WEBHOOK_URL || '';
// 汇总表字段 ID：在智能表格创建「汇总明细」表并配置 Webhook 后，从示例 schema 里复制 key 填到下面（或用环境变量）
// 管理员密码：设置后仅带正确 X-Admin-Token 的请求可「分析表格」「保存合并规则」；未设置则不限制
const ADMIN_SECRET = process.env.ADMIN_SECRET || '';
// 智能表格自动接入：若设置 INGEST_SECRET，则 POST /api/ingest-from-smart-table 时需带请求头 X-Ingest-Secret
const INGEST_SECRET = process.env.INGEST_SECRET || '';

const SUMMARY_FIELDS = {
  role: process.env.SUMMARY_FIELD_ROLE || 'fRole',
  diff: process.env.SUMMARY_FIELD_DIFF || 'fDiff',
  salary: process.env.SUMMARY_FIELD_SALARY || 'fSalary',
  date: process.env.SUMMARY_FIELD_DATE || 'fDate',
};

// ---------- 1~200 级升级所需经验 ----------
const EXP_TABLE = {
  1: 15, 2: 34, 3: 57, 4: 92, 5: 135, 6: 372, 7: 560, 8: 840, 9: 1242, 10: 1716,
  11: 2360, 12: 3216, 13: 4200, 14: 5460, 15: 7050, 16: 8840, 17: 11040, 18: 13716, 19: 16680, 20: 20216,
  21: 24402, 22: 28980, 23: 34320, 24: 40512, 25: 47216, 26: 54900, 27: 63666, 28: 73080, 29: 83720, 30: 95700,
  31: 108480, 32: 122760, 33: 138666, 34: 155540, 35: 174216, 36: 194832, 37: 216600, 38: 240500, 39: 266682, 40: 294216,
  41: 324240, 42: 356916, 43: 391160, 44: 428280, 45: 468450, 46: 510420, 47: 555680, 48: 604416, 49: 655200, 50: 709716,
  51: 748608, 52: 789631, 53: 832902, 54: 878545, 55: 926689, 56: 977471, 57: 1031036, 58: 1087536, 59: 1147132, 60: 1209994,
  61: 1276301, 62: 1346242, 63: 1420016, 64: 1497832, 65: 1579913, 66: 1666492, 67: 1757815, 68: 1854143, 69: 1955750, 70: 2062925,
  71: 2175973, 72: 2295216, 73: 2420993, 74: 2553663, 75: 2693603, 76: 2841212, 77: 2996910, 78: 3161140, 79: 3334370, 80: 3517093,
  81: 3709829, 82: 3913127, 83: 4127566, 84: 4353756, 85: 4592341, 86: 4844001, 87: 5109452, 88: 5389449, 89: 5684790, 90: 5996316,
  91: 6324914, 92: 6671519, 93: 7037118, 94: 7422752, 95: 7829518, 96: 8258575, 97: 8711144, 98: 9188514, 99: 9692044, 100: 10223168,
  101: 10783397, 102: 11374327, 103: 11997640, 104: 12655110, 105: 13348610, 106: 14080113, 107: 14851703, 108: 15665576, 109: 16524049, 110: 17429566,
  111: 18384706, 112: 19392187, 113: 20454878, 114: 21575805, 115: 22758159, 116: 24005306, 117: 25320796, 118: 26708375, 119: 28171993, 120: 29715818,
  121: 31344244, 122: 33061908, 123: 34873700, 124: 36784778, 125: 38800583, 126: 40926854, 127: 43169645, 128: 45535341, 129: 48030677, 130: 50662758,
  131: 53439077, 132: 56367538, 133: 59456479, 134: 62714694, 135: 66151459, 136: 69776558, 137: 73600313, 138: 77633610, 139: 81887931, 140: 86375389,
  141: 91108760, 142: 96101520, 143: 101367883, 144: 106922842, 145: 112782213, 146: 118962678, 147: 125481832, 148: 132358236, 149: 139611467, 150: 147262175,
  151: 155332142, 152: 163844343, 153: 172823012, 154: 182293713, 155: 192283408, 156: 202820538, 157: 213935103, 158: 225658746, 159: 238024845, 160: 251068606,
  161: 264827165, 162: 279339693, 163: 294647508, 164: 310794191, 165: 327825712, 166: 345790561, 167: 364739883, 168: 384727628, 169: 405810702, 170: 428049128,
  171: 451506220, 172: 476248760, 173: 502347192, 174: 529875818, 175: 558913012, 176: 589541445, 177: 621848316, 178: 655925603, 179: 691870326, 180: 729784819,
  181: 769777027, 182: 811960808, 183: 856456260, 184: 903390063, 185: 952895838, 186: 1005114529, 187: 1060194805, 188: 1118293480, 189: 1179575962, 190: 1244216724,
  191: 1312399800, 192: 1384319309, 193: 1460180007, 194: 1540197871, 195: 1624600714, 196: 1713628833, 197: 1807535693, 198: 1906588648, 199: 2011069705, 200: 2121276324,
};

function expForLevel(level) {
  return EXP_TABLE[level] ?? null;
}

function parseNum(v) {
  if (v === undefined || v === null) return NaN;
  const s = String(v).replace(/[\s,，\u200B-\u200D\uFEFF]/g, '').trim();
  let n = parseInt(s, 10);
  if (Number.isNaN(n)) {
    const m = s.match(/\d+/);
    n = m ? parseInt(m[0], 10) : NaN;
  }
  return n;
}

/** 根据等级、开始经验、结束经验、角色名计算差值和工资（与 calc-diff 一致，用于分析表格） */
function calcDiffAndSalary(level, expStart, expEndStr, roleName) {
  const levelNum = parseNum(level);
  const hasLevel = !Number.isNaN(levelNum);
  const startNum = parseNum(expStart);
  const endStr = String(expEndStr ?? '').trim();
  const role = (roleName || '').trim();
  if (Number.isNaN(startNum)) return { diff: null, salary: null, skip: true };
  const upgradeMatch = endStr.match(/升[级級]\s*[\+＋]?\s*(\d+)/);
  let diff;
  if (upgradeMatch) {
    if (!hasLevel) return { diff: null, salary: null, skip: true };
    const extra = parseInt(upgradeMatch[1], 10);
    const need = expForLevel(levelNum);
    if (need == null) return { diff: null, salary: null, skip: true };
    diff = Math.round(need / 10000 - startNum + extra);
  } else {
    const endNum = parseNum(expEndStr);
    if (Number.isNaN(endNum)) return { diff: null, salary: null, skip: true };
    diff = Math.max(0, endNum - startNum);
  }
  let salary;
  const roleLower = role.toLowerCase();
  if (roleLower === 'mao' || role === '米露' || role === '咪露') salary = (diff / 1440) * 10;
  else if (!hasLevel || levelNum < 160) salary = (diff / 1020) * 10;
  else salary = (diff / 1200) * 10;
  salary = Math.round(salary * 100) / 100;
  return { diff, salary, skip: false };
}

/** 解析表格文本：逗号或 Tab 分隔，去 BOM、去首尾空格 */
function parseTable(text) {
  let s = String(text || '').replace(/^\uFEFF/, '').trim();
  const lines = s.split(/\r?\n/).filter((l) => l.length > 0);
  if (lines.length === 0) return { headers: [], rows: [] };
  const first = lines[0];
  const delim = first.includes('\t') ? '\t' : ',';
  const headers = first.split(delim).map((h) => String(h).replace(/\s+/g, ' ').trim());
  const rows = lines.slice(1).map((line) => line.split(delim).map((c) => String(c).replace(/\s+/g, ' ').trim()));
  return { headers, rows };
}

/** 将年月日的“日”钳到该月有效范围，返回 YYYY-MM-DD（避免 2-30、3-32 等无效日期） */
function clampToValidDate(y, month, day) {
  const m = Math.max(1, Math.min(12, month));
  const d = new Date(y, m - 1, 1);
  const maxDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  const dayClamped = Math.max(1, Math.min(day, maxDay));
  return `${y}-${String(m).padStart(2, '0')}-${String(dayClamped).padStart(2, '0')}`;
}

/** 若 dateStr 为 YYYY-MM-DD 且无效（如 2026-02-30），则规范为当月最后一天 */
function clampToValidDateStr(dateStr) {
  const m = String(dateStr || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return dateStr;
  const y = parseInt(m[1], 10), month = parseInt(m[2], 10), day = parseInt(m[3], 10);
  return clampToValidDate(y, month, day);
}

/** 中文星期 → JS getDay()：周日=0, 周一=1, …, 周六=6 */
function parseWeekday(str) {
  const s = String(str || '').trim();
  if (/^周?日$/.test(s)) return 0;
  if (/^周?一$/.test(s)) return 1;
  if (/^周?二$/.test(s)) return 2;
  if (/^周?三$/.test(s)) return 3;
  if (/^周?四$/.test(s)) return 4;
  if (/^周?五$/.test(s)) return 5;
  if (/^周?六$/.test(s)) return 6;
  return null;
}

/** 某年某月中具有指定星期几的日期列表（1-31） */
function daysInMonthWithWeekday(year, month, weekdayNum) {
  const days = [];
  const last = new Date(year, month, 0).getDate();
  for (let d = 1; d <= last; d++) {
    if (new Date(year, month - 1, d).getDay() === weekdayNum) days.push(d);
  }
  return days;
}

/** 用星期反推正确日期：3 月里 13+周五 → 2026-03-13；dayHint 为日期列的数字，weekdayStr 为星期列 */
function resolveDateByWeekday(year, month, dayHint, weekdayStr) {
  const w = parseWeekday(weekdayStr);
  if (w == null) return null;
  const candidates = daysInMonthWithWeekday(year, month, w);
  if (candidates.length === 0) return null;
  const day = (dayHint >= 1 && dayHint <= 31 && candidates.includes(dayHint))
    ? dayHint
    : candidates[0];
  return clampToValidDate(year, month, day);
}

/** 仅用于导入：根据日期列+星期列得到 YYYY-MM-DD，不返回 Dxx；无星期则用 normalizeDateCell 再转成完整日期 */
function resolveImportRowDate(rawDate, rawWeekday, defaultYear, defaultMonth) {
  const s = String(rawDate ?? '').trim();
  let year = defaultYear, month = defaultMonth, dayHint = null;
  // 先解析「月.日」「月/日」「月-日」（如 1.3 = 1月3日），按该月的星期对应分析
  const dotOrSlash = s.match(/^(\d{1,2})[./](\d{1,2})$/);
  if (dotOrSlash) {
    month = parseInt(dotOrSlash[1], 10);
    dayHint = parseInt(dotOrSlash[2], 10);
    if (month >= 1 && month <= 12 && dayHint >= 1 && dayHint <= 31) year = defaultYear;
  } else {
    const part = s.match(/^(\d{1,2})[-/](\d{1,2})$/);
    if (part) {
      month = parseInt(part[1], 10);
      dayHint = parseInt(part[2], 10);
      if (month >= 1 && month <= 12 && dayHint >= 1 && dayHint <= 31) year = defaultYear;
    } else {
      const full = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
      if (full) {
        year = parseInt(full[1], 10);
        month = parseInt(full[2], 10);
        dayHint = parseInt(full[3], 10);
      } else {
        const num = parseInt(s, 10);
        if (!Number.isNaN(num) && num >= 1 && num <= 31) dayHint = num;
      }
    }
  }
  const hasWeekday = rawWeekday != null && String(rawWeekday).trim();
  if (hasWeekday) {
    const resolved = resolveDateByWeekday(year, month, dayHint || 1, String(rawWeekday).trim());
    if (resolved) return resolved;
  }
  if (dayHint != null && month >= 1 && month <= 12)
    return clampToValidDate(year, month, dayHint);
  const normalized = normalizeDateCell(rawDate);
  if (!normalized) return null;
  if (/^D(\d{2})$/.test(normalized)) {
    const day = parseInt(normalized.slice(1), 10);
    return clampToValidDate(defaultYear, defaultMonth, day);
  }
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(normalized)) return clampToValidDateStr(normalized);
  return normalized;
}

/** 将表格单元格中的日期规范为 YYYY-MM-DD；仅数字 1-31 时视为“仅日”存为 D01..D31；无效日期钳到当月最后一天 */
function normalizeDateCell(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  // 先匹配 月.日 / 月/日（如 3.1、3.2、12.31）避免被下面当成纯数字
  const dotOrSlash = s.match(/^(\d{1,2})[./](\d{1,2})$/);
  if (dotOrSlash) {
    const month = parseInt(dotOrSlash[1], 10);
    const day = parseInt(dotOrSlash[2], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const y = new Date().getFullYear();
      return clampToValidDate(y, month, day);
    }
  }
  // 月-日 / 月/日（如 3-1、3-2）必须在 parseInt 前处理，否则 "3-1" 会被当成 3 变成 D03
  const part = s.match(/^(\d{1,2})[-/](\d{1,2})$/);
  if (part) {
    const month = parseInt(part[1], 10);
    const day = parseInt(part[2], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const y = new Date().getFullYear();
      return clampToValidDate(y, month, day);
    }
  }
  let num = parseInt(s, 10);
  if (Number.isNaN(num)) num = Math.floor(parseFloat(s)); // 支持 45390.0 这类 Excel 导出
  if (!Number.isNaN(num)) {
    if (num >= 1 && num <= 31) return 'D' + String(num).padStart(2, '0');
    // Excel 日期序列号：1 = 1900-01-01，大数字按天换算为真实日期，避免多天被合并成“当天”
    if (num >= 1 && num <= 2958465) {
      const d = new Date(1900, 0, 1);
      d.setDate(d.getDate() + (num - 1));
      const y = d.getFullYear(), m = d.getMonth() + 1, day = d.getDate();
      if (y >= 1900 && y <= 2100) return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  const full = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (full) return clampToValidDate(parseInt(full[1], 10), parseInt(full[2], 10), parseInt(full[3], 10));
  return null;
}

/**
 * 导入专用：解析粘贴/上传的表格，得到记录列表。
 * 日期一律按「日期列 + 星期列」解析为 YYYY-MM-DD（有星期则按星期定 3-13/3-3/3-30），不输出 Dxx。
 * 表头需含：等级、开始/结束经验、角色名称；可选 日期、星期、差值。
 */
function analyzeTableContent(text) {
  const { headers, rows } = parseTable(text);
  const idxLevel = headers.findIndex((h) => /等级/.test(h));
  let idxExpStart = headers.findIndex((h) => /经验.*开始|开始.*经验|经验值\s*[（(]?\s*开始|起始/.test(h));
  if (idxExpStart < 0) idxExpStart = headers.findIndex((h) => /^开始$|^开始经验$/.test(h) || /开始|起始/.test(h));
  let idxExpEnd = headers.findIndex((h) => /经验.*结束|结束.*经验|经验值\s*[（(]?\s*结束/.test(h));
  if (idxExpEnd < 0) idxExpEnd = headers.findIndex((h) => /^结束$|^结束经验$/.test(h) || /结束/.test(h));
  const idxRole = headers.findIndex((h) => /角色|昵称/.test(h));
  const idxDiff = headers.findIndex((h) => /差值/.test(h));
  const idxDate = headers.findIndex((h) => /日期/.test(h));
  const idxWeekday = headers.findIndex((h) => /星期/.test(h));

  const now = new Date();
  const defaultYear = now.getFullYear();
  const defaultMonth = now.getMonth() + 1;
  const todayStr = new Date().toISOString().slice(0, 10);
  const result = [];

  for (const row of rows) {
    if (row.every((c) => !String(c).trim())) continue;
    const level = row[idxLevel];
    const expStart = row[idxExpStart] ?? row[idxExpStart - 1] ?? row[idxExpStart + 1];
    const expEnd = row[idxExpEnd] ?? row[idxExpEnd - 1] ?? row[idxExpEnd + 1];
    const roleName = idxRole >= 0 ? (row[idxRole] || '').trim() : '';

    const rawDate = idxDate >= 0 ? row[idxDate] : null;
    const rawWeekday = idxWeekday >= 0 ? row[idxWeekday] : null;
    const date = resolveImportRowDate(rawDate, rawWeekday, defaultYear, defaultMonth) || todayStr;

    let diff = null;
    if (idxDiff >= 0 && row[idxDiff] !== undefined && row[idxDiff] !== '') {
      const d = parseNum(row[idxDiff]);
      if (!Number.isNaN(d)) diff = d;
    }
    if (diff === null && level !== undefined && expStart !== undefined && expEnd !== undefined) {
      const calc = calcDiffAndSalary(level, expStart, expEnd, roleName);
      if (!calc.skip) diff = calc.diff;
    }
    if (diff === null || Number.isNaN(diff)) continue;
    const expStartNum = parseNum(expStart);
    result.push({
      date,
      roleName: roleName || '未知',
      diff,
      expStart: Number.isNaN(expStartNum) ? '' : expStartNum,
      expEnd: (expEnd !== undefined && expEnd !== null) ? String(expEnd).trim() : '',
    });
  }

  let totalDiff = 0;
  const byDate = {};
  const byRole = {};
  for (const r of result) {
    totalDiff += r.diff;
    byDate[r.date] = (byDate[r.date] || 0) + r.diff;
    byRole[r.roleName] = (byRole[r.roleName] || 0) + r.diff;
  }
  return {
    rows: result,
    totalDiff,
    byDate,
    byRole,
    detectedHeaders: headers,
    rawRowCount: rows.length,
  };
}

// 解析：账号xxx 等级n 开始经验xxx 结束经验xxx（或 结束升级+xxx）
// 角色名只取「等级」前面的部分，避免把整句当成昵称
function parseMessage(text, sender) {
  const raw = (text || '').replace(/@财务账号/g, '').replace(/@财务/g, '').replace(/@总结/g, '').trim();
  const accountMatch = raw.match(/账号\s*(\S+?)(?=\s*等级|等级|$)/) || raw.match(/账号\s*(\S+)/);
  const account = accountMatch ? accountMatch[1].trim() : raw.match(/账号\s*(\S+)/)?.[1];
  const level = parseInt(raw.match(/等级\s*(\d+)/)?.[1], 10);
  const startStr = raw.match(/(?:开始|经验开始|开始经验)\s*(\d+)/)?.[1];
  const endStr = raw.match(/(?:结束|经验结束|结束经验)\s*(\d+|升级\+?\d+)/)?.[1];

  if (!account || !level || !startStr || !endStr) {
    throw new Error('格式需包含：账号xxx 等级n 开始经验n 结束经验n（或结束升级+n）');
  }

  const expStart = parseInt(startStr, 10);
  const upgradeMatch = endStr.match(/升级\+?(\d+)/);
  let expEnd, diff;

  if (upgradeMatch) {
    const extra = parseInt(upgradeMatch[1], 10);
    const need = expForLevel(level);
    if (need == null) throw new Error(`无等级${level}经验数据`);
    diff = Math.round(need / 10000 - expStart + extra);
    expEnd = `升级+${extra}`;
  } else {
    expEnd = parseInt(endStr, 10);
    if (Number.isNaN(expEnd) || expEnd <= expStart) throw new Error('结束经验需大于开始经验');
    diff = expEnd - expStart;
  }

  let salary;
  if (account === 'mao' || account === '米露') salary = (diff / 1440) * 10;
  else if (level < 160) salary = (diff / 1020) * 10;
  else salary = (diff / 1200) * 10;
  salary = Math.round(salary * 100) / 100;

  // 使用中国时区，保证日期和星期与用户看到的一致并随每条消息实时更新
  const now = new Date();
  const chinaTz = 'Asia/Shanghai';
  const dateNum = parseInt(new Intl.DateTimeFormat('zh-CN', { timeZone: chinaTz, day: 'numeric' }).format(now), 10);
  const weekdayStr = new Intl.DateTimeFormat('zh-CN', { timeZone: chinaTz, weekday: 'short' }).format(now);
  return {
    date: dateNum,
    weekday: weekdayStr,
    wechatName: sender,
    roleName: account,
    level,
    expStart,
    expEnd: String(expEnd),
    diff,
    salary,
    note: '',
    photoTime: '正常拍照',
  };
}

async function appendSheet(row) {
  const body = {
    schema: {
      f1PSmO: '日期', f3MlEe: '星期', fAfOBp: '微信名称', fDe867: '角色名称',
      fR7LNY: '等级', fXUN1W: '经验值（开始）', fXifUI: '经验值（结束）',
      fYOy6d: '差值', fcAmkj: '工资', fpsEdy: '备注', fxB4Oz: '文本11',
    },
    add_records: [{
      values: {
        f1PSmO: row.date,
        f3MlEe: [{ text: row.weekday }],
        fAfOBp: row.wechatName,
        fDe867: row.roleName,
        fR7LNY: row.level,
        fXUN1W: row.expStart,
        fXifUI: row.expEnd,
        fYOy6d: row.diff,
        fcAmkj: row.salary,
        fpsEdy: row.note || '',
        fxB4Oz: [{ text: row.photoTime }],
      },
    }],
  };
  const { data } = await axios.post(SHEET_WEBHOOK, body, {
    headers: { 'Content-Type': 'application/json' },
    timeout: 10000,
  });
  console.log('[表格接口返回]', JSON.stringify(data));
  return data;
}

// ---------- 本地数据存一份（与智能表格同步，供网页统计）----------
const DATA_DIR = path.join(__dirname, 'data');
const RECORDS_FILE = path.join(DATA_DIR, 'records.json');
const ANALYZED_FILE = path.join(DATA_DIR, 'analyzed.json');
const ROLE_ALIASES_FILE = path.join(DATA_DIR, 'role-aliases.json');
/** 服务器内建智能表格：机器人通过 URL 回调写入，实时更新后驱动可分析表格 */
const SMART_TABLE_FILE = path.join(DATA_DIR, 'smart-table.json');

function loadRecords() {
  try {
    if (fs.existsSync(RECORDS_FILE)) {
      const raw = fs.readFileSync(RECORDS_FILE, 'utf8');
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    }
  } catch (e) {
    console.error('[store] load error', e.message);
  }
  return [];
}

function loadAnalyzed() {
  try {
    if (fs.existsSync(ANALYZED_FILE)) {
      const raw = fs.readFileSync(ANALYZED_FILE, 'utf8');
      const data = JSON.parse(raw);
      if (data && Array.isArray(data.rows) && data.rows.length > 0) return data;
    }
  } catch (e) {
    console.error('[analyzed] load error', e.message);
  }
  return null;
}

function saveAnalyzed(data) {
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    const out = { ...data, updatedAt: new Date().toISOString() };
    fs.writeFileSync(ANALYZED_FILE, JSON.stringify(out, null, 0), 'utf8');
  } catch (e) {
    console.error('[analyzed] save error', e.message);
  }
}

const DEFAULT_HEADERS = ['日期', '角色名称', '开始经验', '结束经验', '差值'];

/** 按表头名匹配列索引（用于从「完整行对象」里取逻辑列） */
function getHeaderIndex(headers, pattern) {
  if (!headers || !headers.length) return -1;
  const i = headers.findIndex((h) => pattern.test(String(h)));
  return i;
}

/** 从一行对象 + 表头 提取用于统计的 5 个字段（日期、角色、差值、开始经验、结束经验） */
function getLogicalFromRow(row, headers) {
  if (!row || typeof row !== 'object') return null;
  const idxDate = getHeaderIndex(headers, /日期/);
  const idxRole = getHeaderIndex(headers, /角色|昵称/);
  const idxDiff = getHeaderIndex(headers, /差值/);
  const idxStart = getHeaderIndex(headers, /开始|起始|经验.*开始|经验值\s*[（(]?\s*开始/);
  const idxEnd = getHeaderIndex(headers, /结束|经验.*结束|经验值\s*[（(]?\s*结束/);
  const date = idxDate >= 0 ? (row[headers[idxDate]] ?? '') : '';
  const roleName = idxRole >= 0 ? (row[headers[idxRole]] ?? '').trim() : '';
  const expStart = idxStart >= 0 ? (row[headers[idxStart]] ?? '') : '';
  const expEnd = idxEnd >= 0 ? (row[headers[idxEnd]] ?? '') : '';
  let diff = null;
  if (idxDiff >= 0 && row[headers[idxDiff]] !== undefined && row[headers[idxDiff]] !== '') {
    const n = parseNum(row[headers[idxDiff]]);
    if (!Number.isNaN(n)) diff = n;
  }
  const idxLevel = getHeaderIndex(headers, /等级/);
  if (diff === null && idxLevel >= 0 && expStart !== '' && expEnd !== '') {
    const level = row[headers[idxLevel]];
    const calc = calcDiffAndSalary(level, expStart, expEnd, roleName);
    if (!calc.skip) diff = calc.diff;
  }
  if (diff === null || Number.isNaN(diff)) return null;
  return { date, roleName, diff, expStart, expEnd };
}

function loadSmartTable() {
  try {
    if (fs.existsSync(SMART_TABLE_FILE)) {
      const raw = fs.readFileSync(SMART_TABLE_FILE, 'utf8');
      const data = JSON.parse(raw);
      if (data.headers && Array.isArray(data.rows)) return { headers: data.headers, rows: data.rows };
      if (Array.isArray(data.rows) && data.rows.length > 0 && data.rows[0] && typeof data.rows[0] === 'object' && !Array.isArray(data.rows[0])) {
        const legacy = data.rows;
        const rows = legacy.map((r) => ({
          '日期': r.date ?? r['日期'] ?? '',
          '角色名称': r.roleName ?? r['角色名称'] ?? '',
          '开始经验': r.expStart ?? r['开始经验'] ?? '',
          '结束经验': r.expEnd ?? r['结束经验'] ?? '',
          '差值': r.diff ?? r['差值'] ?? '',
        }));
        return { headers: DEFAULT_HEADERS.slice(), rows };
      }
    }
  } catch (e) {
    console.error('[smart-table] load error', e.message);
  }
  return { headers: DEFAULT_HEADERS.slice(), rows: [] };
}

function saveSmartTable(data) {
  const { headers, rows } = typeof data === 'object' && data && Array.isArray(data.rows) ? data : { headers: DEFAULT_HEADERS, rows: data || [] };
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(SMART_TABLE_FILE, JSON.stringify({ headers, rows, updatedAt: new Date().toISOString() }, null, 0), 'utf8');
  } catch (e) {
    console.error('[smart-table] save error', e.message);
  }
}

/** YYYY-MM-DD → 月.日（如 2026-03-13→3.13），用于智能表格存储与网页展示一致 */
function toMonthDotDay(dateStr) {
  const s = String(dateStr || '').trim();
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return s;
  const month = parseInt(m[2], 10);
  const day = parseInt(m[3], 10);
  if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return `${month}.${day}`;
  return s;
}

/** 任意日期格式 → YYYY-MM-DD，用于去重 key 与按日期筛选比较 */
function dateToComparable(dateStr) {
  return dateToDedupeKey(dateStr) || '';
}

/** 去重时统一日期：D12、3.13、2026-03-12 视为同一天；结果做有效日期钳位 */
function dateToDedupeKey(dateStr) {
  const s = String(dateStr || '').trim();
  if (!s) return '';
  const dMatch = s.match(/^D(\d{2})$/);
  if (dMatch) {
    const day = parseInt(dMatch[1], 10);
    const now = new Date();
    const y = now.getFullYear(), m = now.getMonth() + 1;
    return clampToValidDate(y, m, day);
  }
  const monthDotDay = s.match(/^(\d{1,2})\.(\d{1,2})$/);
  if (monthDotDay) {
    const month = parseInt(monthDotDay[1], 10);
    const day = parseInt(monthDotDay[2], 10);
    const y = new Date().getFullYear();
    return clampToValidDate(y, month, day);
  }
  const full = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (full) return clampToValidDateStr(s);
  const part = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (part) return clampToValidDate(parseInt(part[1], 10), parseInt(part[2], 10), parseInt(part[3], 10));
  return s;
}

/** 完全一样的行只保留第一条（自动去重）；保留的行日期统一存为 月.日（与智能表格展示一致） */
function dedupeSmartTableRows(rows) {
  const seen = new Set();
  return rows.filter((r) => {
    const dateKey = dateToDedupeKey(r.date);
    const key = dateKey + '|' + (r.roleName || '') + '|' + (r.diff ?? '') + '|' + (r.expStart ?? '') + '|' + (r.expEnd ?? '');
    if (seen.has(key)) return false;
    seen.add(key);
    const storeDate = dateKey ? toMonthDotDay(dateKey) || dateKey : (r.date != null ? String(r.date).trim() : '');
    if (storeDate && storeDate !== (r.date != null ? String(r.date).trim() : '')) r.date = storeDate;
    return true;
  });
}

/**
 * 根据智能表格重算分析结果并落盘（供按角色/按日期查询）。
 * 从「完整行」里按表头提取 日期/角色/差值 等，去重后只写入 analyzed，不改写智能表格。
 */
function rebuildAnalyzedFromSmartTable() {
  const { headers, rows } = loadSmartTable();
  const logicalList = [];
  for (const row of rows) {
    const logical = getLogicalFromRow(row, headers);
    if (logical) logicalList.push(logical);
  }
  const deduped = dedupeSmartTableRows(logicalList);
  let totalDiff = 0;
  const byDate = {};
  const byRole = {};
  for (const r of deduped) {
    totalDiff += r.diff || 0;
    const d = r.date || '';
    const name = r.roleName || '未知';
    byDate[d] = (byDate[d] || 0) + (r.diff || 0);
    byRole[name] = (byRole[name] || 0) + (r.diff || 0);
  }
  const roles = Object.keys(byRole).sort();
  const key = (s) => String(s).toLowerCase().trim();
  const byKey = {};
  for (const name of roles) {
    const k = key(name);
    if (!byKey[k]) byKey[k] = [];
    byKey[k].push(name);
  }
  const suggestMerge = Object.entries(byKey)
    .filter(([, variants]) => variants.length > 1)
    .map(([canonicalKey, variants]) => ({ key: canonicalKey, variants: variants.sort() }));
  lastAnalyzedRolesCache = { roles, suggestMerge };
  saveAnalyzed({
    rows: deduped,
    totalDiff,
    byDate,
    byRole,
    detectedHeaders: headers.slice(),
    rawRowCount: rows.length,
  });
}

/** 供统计与查询使用的数据：优先用「表格分析结果」，否则用机器人写入的本地记录 */
function getDataForStats() {
  const analyzed = loadAnalyzed();
  if (analyzed && analyzed.rows && analyzed.rows.length > 0) return analyzed.rows;
  const records = loadRecords();
  return records.map((r) => ({
    date: r.date || '',
    roleName: r.roleName || '未知',
    diff: r.diff || 0,
    expStart: r.expStart !== undefined ? r.expStart : '',
    expEnd: r.expEnd !== undefined ? r.expEnd : '',
  }));
}

/** 最近一次分析成功的角色列表缓存（解决「刷新列表」读不到文件时仍能显示） */
let lastAnalyzedRolesCache = null;

// ---------- 角色名称合并（大小写/相似视为同一角色）----------
function loadRoleAliases() {
  try {
    if (fs.existsSync(ROLE_ALIASES_FILE)) {
      const raw = fs.readFileSync(ROLE_ALIASES_FILE, 'utf8');
      const data = JSON.parse(raw);
      return data.aliases || {};
    }
  } catch (e) {
    console.error('[role-aliases] load error', e.message);
  }
  return {};
}

function parseRoleRules(rulesText) {
  const aliases = {};
  const lines = String(rulesText || '').split(/\r?\n/);
  for (const line of lines) {
    const t = line.trim();
    if (!t || t.startsWith('#')) continue;
    const eq = t.indexOf('=');
    if (eq <= 0) continue;
    const canonical = t.slice(0, eq).trim();
    if (!canonical) continue;
    aliases[canonical] = canonical;
    const right = t.slice(eq + 1).split(',').map((s) => s.trim()).filter(Boolean);
    for (const a of right) aliases[a] = canonical;
  }
  return aliases;
}

function saveRoleAliases(rulesText) {
  const aliases = parseRoleRules(rulesText);
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(ROLE_ALIASES_FILE, JSON.stringify({ rulesText: rulesText || '', aliases }, null, 2), 'utf8');
  } catch (e) {
    console.error('[role-aliases] save error', e.message);
  }
  return aliases;
}

function getCanonicalRoleName(name) {
  const aliases = loadRoleAliases();
  return aliases[name] ?? name;
}

/** 某规范名下的所有合并名称（规范名 + 所有映射到该规范名的别名） */
function getVariantsForCanonical(canonical) {
  const aliases = loadRoleAliases();
  const list = [];
  const seen = new Set();
  if (canonical && !seen.has(canonical)) {
    list.push(canonical);
    seen.add(canonical);
  }
  for (const [name, can] of Object.entries(aliases)) {
    if (can === canonical && !seen.has(name)) {
      list.push(name);
      seen.add(name);
    }
  }
  return list.sort();
}

/** 当前查询使用的数据来源（供前端展示「实时数据来自哪里」） */
function getDataSource() {
  const analyzed = loadAnalyzed();
  if (analyzed && analyzed.rows && analyzed.rows.length > 0) {
    return { source: 'analyzed', recordCount: analyzed.rows.length, updatedAt: analyzed.updatedAt || null };
  }
  const records = loadRecords();
  return { source: 'records', recordCount: records.length, updatedAt: null };
}

function saveRecords(records) {
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(RECORDS_FILE, JSON.stringify(records, null, 0), 'utf8');
  } catch (e) {
    console.error('[store] save error', e.message);
  }
}

function saveRecord(row) {
  const dateStr = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  const records = loadRecords();
  records.push({
    date: dateStr,
    roleName: row.roleName,
    diff: row.diff,
    salary: row.salary,
    wechatName: row.wechatName,
    level: row.level,
    expStart: row.expStart,
    expEnd: row.expEnd,
  });
  saveRecords(records);
}

/** 写入汇总明细表：每条消息一行（角色、本次差值、本次工资、日期），表格内用「分组汇总」按角色合计 */
async function appendSummarySheet(row) {
  if (!SUMMARY_WEBHOOK) return null;
  const body = {
    schema: {
      [SUMMARY_FIELDS.role]: '角色名称',
      [SUMMARY_FIELDS.diff]: '本次差值',
      [SUMMARY_FIELDS.salary]: '本次工资',
      [SUMMARY_FIELDS.date]: '日期',
    },
    add_records: [{
      values: {
        [SUMMARY_FIELDS.role]: row.roleName,
        [SUMMARY_FIELDS.diff]: row.diff,
        [SUMMARY_FIELDS.salary]: row.salary,
        [SUMMARY_FIELDS.date]: `${row.date}日`,
      },
    }],
  };
  try {
    const { data } = await axios.post(SUMMARY_WEBHOOK, body, {
      headers: { 'Content-Type': 'application/json' },
      timeout: 10000,
    });
    if (data && data.errcode && data.errcode !== 0) console.error('[汇总表]', data.errcode, data.errmsg);
    else console.log('[汇总表] 已写入 角色=', row.roleName, '差值=', row.diff);
    return data;
  } catch (e) {
    console.error('[汇总表] 请求失败', e.message);
    return null;
  }
}

// ---------- 服务 ----------
const app = express();

app.use((req, res, next) => {
  console.log('[请求]', req.method, req.url);
  next();
});

/** 管理员校验：未设置 ADMIN_SECRET 则不限制；设置了则要求请求头 X-Admin-Token 与 ADMIN_SECRET 一致 */
function requireAdmin(req, res, next) {
  if (!ADMIN_SECRET) return next();
  const token = req.headers['x-admin-token'];
  if (token === ADMIN_SECRET) return next();
  res.status(403).setHeader('Content-Type', 'application/json; charset=utf-8');
  res.send(JSON.stringify({ error: '需要管理员权限，请先登录' }));
}

/** 前端用于判断是否需管理员登录（未配置 ADMIN_SECRET 则所有人可操作） */
app.get('/api/admin-required', (req, res) => {
  res.json({ adminRequired: !!ADMIN_SECRET });
});

/** 校验管理员密码，返回可用于后续请求的 token */
app.post('/api/admin-check', express.json({ limit: '1kb' }), (req, res) => {
  try {
    const password = (req.body && req.body.password != null) ? String(req.body.password) : '';
    if (!ADMIN_SECRET) {
      return res.json({ ok: true, token: '', message: '未配置管理员密码，所有人可操作' });
    }
    if (password === ADMIN_SECRET) {
      return res.json({ ok: true, token: ADMIN_SECRET });
    }
    res.json({ ok: false, error: '密码错误' });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/test', (req, res) => {
  res.json({ ok: true, msg: 'wechat bot running' });
});

// 统计页：优先读 public/index.html，没有则发内联页面（避免 Render 未带上 public 时报错）
const INDEX_HTML_PATH = path.join(__dirname, 'public', 'index.html');
const INDEX_HTML_INLINE = `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>角色经验统计</title><style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"PingFang SC",sans-serif;margin:0;padding:20px;background:#f5f6f8}h1{font-size:1.35rem;color:#1a1a2e}.sub{color:#666;font-size:.9rem}section{background:#fff;border-radius:10px;padding:16px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.08)}section h2{font-size:1rem;color:#333;margin:0 0 12px 0}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:10px 12px;border-bottom:1px solid #eee}th{color:#666;font-weight:600}.num{text-align:right;font-variant-numeric:tabular-nums}.query-row{display:flex;gap:10px;align-items:center;margin-bottom:12px}.query-row input[type=date]{padding:8px 12px;border:1px solid #ddd;border-radius:6px;font-size:1rem}.query-row button{padding:8px 16px;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}.empty{color:#999;padding:16px 0}.total{font-weight:600;color:#1a1a2e}</style></head><body><h1>角色经验统计</h1><p class="sub">数据与智能表格同步，从今天起按角色汇总差值总和；可查询任意日期当天练了多少经验。</p><section><h2>从今天起 · 各角色差值总和</h2><p class="sub" style="margin-bottom:12px">统计日期 ≥ <span id="sinceDate">-</span></p><table><thead><tr><th>角色名称</th><th class="num">差值总和</th></tr></thead><tbody id="rolesBody"></tbody><tfoot><tr><td class="total">合计</td><td class="num total" id="rolesTotal">0</td></tr></tfoot></table><p id="rolesEmpty" class="empty" style="display:none">暂无数据（今天起还没有记录）</p></section><section><h2>查询某天练了多少经验（差值）</h2><div class="query-row"><input type="date" id="queryDate"/><button type="button" id="queryBtn">查询</button></div><table id="dayTable" style="display:none"><thead><tr><th>角色名称</th><th class="num">当日差值</th></tr></thead><tbody id="dayBody"></tbody><tfoot><tr><td class="total">当日合计</td><td class="num total" id="dayTotal">0</td></tr></tfoot></table><p id="dayEmpty" class="empty" style="display:none">该日暂无记录</p></section><script>const base=window.location.origin;function todayStr(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}async function loadRoleStats(){const from=todayStr();document.getElementById('sinceDate').textContent=from;const res=await fetch(base+'/api/stats/roles?from='+encodeURIComponent(from));const data=await res.json();const tbody=document.getElementById('rolesBody');const totalEl=document.getElementById('rolesTotal');const emptyEl=document.getElementById('rolesEmpty');tbody.innerHTML='';const entries=Object.entries(data.roles||{}).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);let total=0;entries.forEach(([name,sum])=>{total+=sum;const tr=document.createElement('tr');tr.innerHTML='<td>'+name.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</td><td class="num">'+sum.toLocaleString()+'</td>';tbody.appendChild(tr)});totalEl.textContent=total.toLocaleString();emptyEl.style.display=entries.length?'none':'block'}async function queryDay(){const date=document.getElementById('queryDate').value;if(!date)return;const res=await fetch(base+'/api/query?date='+encodeURIComponent(date));const data=await res.json();const table=document.getElementById('dayTable');const tbody=document.getElementById('dayBody');const totalEl=document.getElementById('dayTotal');const emptyEl=document.getElementById('dayEmpty');tbody.innerHTML='';const byRole=data.byRole||{};const entries=Object.entries(byRole).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);let total=0;entries.forEach(([name,sum])=>{total+=sum;const tr=document.createElement('tr');tr.innerHTML='<td>'+name.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</td><td class="num">'+sum.toLocaleString()+'</td>';tbody.appendChild(tr)});totalEl.textContent=total.toLocaleString();table.style.display=entries.length?'table':'none';emptyEl.style.display=entries.length?'none':'block'}document.getElementById('queryBtn').addEventListener('click',queryDay);document.getElementById('queryDate').value=todayStr();loadRoleStats();</script></body></html>`;

app.get('/', (req, res) => {
  try {
    if (fs.existsSync(INDEX_HTML_PATH)) {
      return res.type('html').send(fs.readFileSync(INDEX_HTML_PATH, 'utf8'));
    }
  } catch (_) {}
  res.type('html').send(INDEX_HTML_INLINE);
});

app.use(express.static(path.join(__dirname, 'public')));

// 从某日起各角色差值总和（按规范名合并：大小写/别名算同一角色）
app.get('/api/stats/roles', (req, res) => {
  try {
    const from = req.query.from || '';
    const records = getDataForStats();
    const filtered = from ? records.filter((r) => (r.date || '') >= from) : records;
    const byRole = {};
    for (const r of filtered) {
      const name = getCanonicalRoleName(r.roleName || '未知');
      byRole[name] = (byRole[name] || 0) + (r.diff || 0);
    }
    res.json({ since: from || null, roles: byRole });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 查询某天或日期范围内所有角色合计（按规范名合并）；date=单日 或 dateFrom+dateTo=范围
app.get('/api/query', (req, res) => {
  try {
    const date = (req.query.date || '').trim();
    const dateFrom = (req.query.dateFrom || '').trim();
    const dateTo = (req.query.dateTo || '').trim();
    let records = getDataForStats();
    if (date) {
      records = records.filter((r) => dateToComparable(r.date) === date);
    } else if (dateFrom && dateTo) {
      records = records.filter((r) => {
        const d = dateToComparable(r.date);
        return d && d >= dateFrom && d <= dateTo;
      });
    } else {
      return res.status(400).json({ error: '请传 date=单日 或 dateFrom 与 dateTo=范围' });
    }
    const byRole = {};
    let total = 0;
    for (const r of records) {
      total += r.diff || 0;
      const name = getCanonicalRoleName(r.roleName || '未知');
      byRole[name] = (byRole[name] || 0) + (r.diff || 0);
    }
    res.json({ date: date || null, dateFrom: dateFrom || null, dateTo: dateTo || null, totalDiff: total, byRole, records });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 按角色名称搜索（返回规范名、合并名称列表、各名称的差值、按日差值、差值之和）
// 可选 date=单日 或 dateFrom+dateTo=范围，不传则返回该角色全部
app.get('/api/role', (req, res) => {
  try {
    const name = (req.query.name || '').trim();
    if (!name) return res.status(400).json({ error: '请传 name=角色名称' });
    const date = (req.query.date || '').trim();
    const dateFrom = (req.query.dateFrom || '').trim();
    const dateTo = (req.query.dateTo || '').trim();
    const canonicalSearch = getCanonicalRoleName(name);
    let records = getDataForStats().filter((r) => {
      const canonical = getCanonicalRoleName(r.roleName || '');
      return canonical === canonicalSearch || (r.roleName || '').includes(name) || canonical.includes(name);
    });
    if (date) {
      records = records.filter((r) => dateToComparable(r.date) === date);
    } else if (dateFrom && dateTo) {
      records = records.filter((r) => {
        const d = dateToComparable(r.date);
        return d && d >= dateFrom && d <= dateTo;
      });
    }
    const byDate = {};
    const byVariant = {};
    let totalDiff = 0;
    for (const r of records) {
      const d = r.date || '';
      const rawName = r.roleName || '未知';
      byDate[d] = (byDate[d] || 0) + (r.diff || 0);
      byVariant[rawName] = (byVariant[rawName] || 0) + (r.diff || 0);
      totalDiff += r.diff || 0;
    }
    const byDateList = Object.entries(byDate)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([dt, diff]) => ({ date: dt, diff }));
    const variants = getVariantsForCanonical(canonicalSearch);
    const recordsList = records.map((r) => ({
      date: r.date != null ? String(r.date) : '',
      roleName: (r.roleName != null && r.roleName !== '') ? String(r.roleName) : '未知',
      diff: Number(r.diff) || 0,
      expStart: r.expStart !== undefined && r.expStart !== null ? r.expStart : '',
      expEnd: r.expEnd !== undefined && r.expEnd !== null ? r.expEnd : '',
    })).sort((a, b) => (b.date || '').localeCompare(a.date || '') || (a.roleName || '').localeCompare(b.roleName || ''));
    res.json({
      roleName: canonicalSearch,
      variants,
      byVariant,
      totalDiff,
      byDate: byDateList,
      recordCount: recordsList.length,
      records: recordsList,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 当前数据来源（实时）：分析结果 或 机器人记录，供前端展示
app.get('/api/data-source', (req, res) => {
  try {
    res.json(getDataSource());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 列出当前数据中出现的所有角色名称；空时尝试直读 analyzed 文件或用最近分析缓存
app.get('/api/roles', (req, res) => {
  try {
    const rows = getDataForStats();
    const set = new Set();
    for (const r of rows) {
      const n = (r.roleName || '').trim();
      if (n) set.add(n);
    }
    let roles = Array.from(set).sort();
    let suggestMerge = [];
    if (roles.length === 0 && fs.existsSync(ANALYZED_FILE)) {
      try {
        const raw = fs.readFileSync(ANALYZED_FILE, 'utf8');
        const data = JSON.parse(raw);
        if (data && Array.isArray(data.rows)) {
          for (const r of data.rows) {
            const n = (r.roleName || '').trim();
            if (n) set.add(n);
          }
          roles = Array.from(set).sort();
        }
      } catch (_) {}
    }
    if (roles.length > 0) {
      const key = (s) => String(s).toLowerCase().trim();
      const byKey = {};
      for (const name of roles) {
        const k = key(name);
        if (!byKey[k]) byKey[k] = [];
        byKey[k].push(name);
      }
      suggestMerge = Object.entries(byKey)
        .filter(([, variants]) => variants.length > 1)
        .map(([canonicalKey, variants]) => ({ key: canonicalKey, variants: variants.sort() }));
    } else if (lastAnalyzedRolesCache && lastAnalyzedRolesCache.roles && lastAnalyzedRolesCache.roles.length > 0) {
      roles = lastAnalyzedRolesCache.roles;
      suggestMerge = lastAnalyzedRolesCache.suggestMerge || [];
    }
    res.json({ roles, suggestMerge });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 获取当前角色合并规则（文本，一行一条：规范名=别名1,别名2）
app.get('/api/role-aliases', (req, res) => {
  try {
    let rulesText = '';
    if (fs.existsSync(ROLE_ALIASES_FILE)) {
      const data = JSON.parse(fs.readFileSync(ROLE_ALIASES_FILE, 'utf8'));
      rulesText = data.rulesText || '';
    }
    res.json({ rulesText });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 保存角色合并规则（body.rulesText，同上格式）；需管理员
app.post('/api/role-aliases', requireAdmin, express.json({ limit: '64kb' }), (req, res) => {
  try {
    const rulesText = req.body && req.body.rulesText !== undefined ? String(req.body.rulesText) : '';
    saveRoleAliases(rulesText);
    res.json({ ok: true, message: '已保存，统计与搜索将按规范名合并' });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ---------- 服务器内建智能表格：机器人通过 URL 回调写入一行，实时更新可分析表格 ----------
const SMART_TABLE_SECRET = process.env.SMART_TABLE_SECRET || '';

function requireSmartTableSecret(req, res, next) {
  if (!SMART_TABLE_SECRET) return next();
  const token = req.headers['x-smart-table-secret'];
  if (token === SMART_TABLE_SECRET) return next();
  res.status(403).setHeader('Content-Type', 'application/json; charset=utf-8');
  res.send(JSON.stringify({ error: 'X-Smart-Table-Secret 无效' }));
}

/** 实时查看服务器智能表格（只读）；返回表头 + 行，支持完整导入后的多列 */
app.get('/api/smart-table', (req, res) => {
  try {
    const data = loadSmartTable();
    let updatedAt = '';
    try {
      if (fs.existsSync(SMART_TABLE_FILE)) {
        const raw = fs.readFileSync(SMART_TABLE_FILE, 'utf8');
        const parsed = JSON.parse(raw);
        updatedAt = parsed.updatedAt || '';
      }
    } catch (_) {}
    res.json({ headers: data.headers, rows: data.rows, updatedAt });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/** 管理员编辑智能表格某一行（按行号 0 起）；body 可传任意表头对应字段 */
app.patch('/api/smart-table/row/:index', requireAdmin, express.json({ limit: '16kb' }), (req, res) => {
  try {
    const index = parseInt(req.params.index, 10);
    if (Number.isNaN(index) || index < 0) return res.status(400).json({ error: '行号无效' });
    const data = loadSmartTable();
    if (index >= data.rows.length) return res.status(404).json({ error: '该行不存在' });
    const b = req.body || {};
    const row = data.rows[index];
    data.headers.forEach((h) => {
      if (b[h] !== undefined) row[h] = b[h] !== null ? String(b[h]).trim() : '';
    });
    if (b.date !== undefined) {
      let d = String(b.date).trim();
      if (d) { const n = normalizeDateCell(d); if (n) d = n; }
      const idx = getHeaderIndex(data.headers, /日期/);
      if (idx >= 0) row[data.headers[idx]] = d || row[data.headers[idx]];
    }
    if (b.roleName !== undefined) {
      const idx = getHeaderIndex(data.headers, /角色|昵称/);
      if (idx >= 0) row[data.headers[idx]] = String(b.roleName).trim() || row[data.headers[idx]];
    }
    saveSmartTable(data);
    rebuildAnalyzedFromSmartTable();
    res.json({ ok: true, message: '已更新', row: data.rows[index] });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/** 管理员删除智能表格某一行（按行号 0 起） */
app.delete('/api/smart-table/row/:index', requireAdmin, (req, res) => {
  try {
    const index = parseInt(req.params.index, 10);
    if (Number.isNaN(index) || index < 0) return res.status(400).json({ error: '行号无效' });
    const data = loadSmartTable();
    if (index >= data.rows.length) return res.status(404).json({ error: '该行不存在' });
    data.rows.splice(index, 1);
    saveSmartTable(data);
    rebuildAnalyzedFromSmartTable();
    res.json({ ok: true, message: '已删除', recordCount: data.rows.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/** 机器人回调：写入一行到服务器智能表格（按当前表头补齐列） */
app.post('/api/smart-table/row', requireSmartTableSecret, express.json({ limit: '64kb' }), (req, res) => {
  try {
    const b = req.body || {};
    let date = (b.date != null) ? String(b.date).trim() : '';
    if (date) {
      const normalized = normalizeDateCell(date);
      if (normalized) date = toMonthDotDay(normalized) || normalized;
    }
    if (!date) date = new Date().toISOString().slice(0, 10);
    const roleName = (b.roleName != null ? String(b.roleName).trim() : '') || (b.角色名称 != null ? String(b.角色名称).trim() : '') || '未知';
    const diff = parseNum(b.diff ?? b.差值);
    const expStart = b.expStart ?? b['经验值(开)'] ?? b.开始经验 ?? '';
    const expEnd = b.expEnd ?? b['经验值(结束)'] ?? b.结束经验 ?? '';
    if (roleName === '未知' && (diff === undefined || Number.isNaN(diff))) {
      return res.status(400).json({ error: '请提供 roleName（或角色名称）与 diff（或差值）' });
    }
    const data = loadSmartTable();
    const row = {};
    data.headers.forEach((h) => { row[h] = ''; });
    const idxDate = getHeaderIndex(data.headers, /日期/);
    const idxRole = getHeaderIndex(data.headers, /角色|昵称/);
    const idxDiff = getHeaderIndex(data.headers, /差值/);
    const idxStart = getHeaderIndex(data.headers, /开始|起始|经验.*开始|经验值\s*[（(]?\s*开始/);
    const idxEnd = getHeaderIndex(data.headers, /结束|经验.*结束|经验值\s*[（(]?\s*结束/);
    if (idxDate >= 0) row[data.headers[idxDate]] = date;
    if (idxRole >= 0) row[data.headers[idxRole]] = roleName;
    if (idxDiff >= 0) row[data.headers[idxDiff]] = Number.isFinite(diff) ? diff : 0;
    if (idxStart >= 0) row[data.headers[idxStart]] = expStart !== undefined && expStart !== null ? expStart : '';
    if (idxEnd >= 0) row[data.headers[idxEnd]] = expEnd !== undefined && expEnd !== null ? String(expEnd).trim() : '';
    data.rows.push(row);
    saveSmartTable(data);
    rebuildAnalyzedFromSmartTable();
    console.log('[smart-table] 新增 1 条，共', data.rows.length, '条');
    res.json({ ok: true, message: '已写入智能表格并更新分析', recordCount: data.rows.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/** 管理员导入：粘贴 CSV/TSV 时完整保留表头与所有列行（与文件一致）；body.rows 时为追加行 */
app.post('/api/smart-table/import', requireAdmin, express.json({ limit: '6mb' }), (req, res) => {
  try {
    const body = req.body || {};
    if (body.text && typeof body.text === 'string') {
      const { headers, rows: rawRows } = parseTable(body.text.trim());
      if (!headers.length) return res.status(400).json({ error: '表头为空' });
      const rows = rawRows.map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]).trim() : ''; });
        return obj;
      });
      saveSmartTable({ headers, rows });
      rebuildAnalyzedFromSmartTable();
      console.log('[smart-table] 粘贴导入（全部保留）', rows.length, '条，', headers.length, '列');
      return res.json({
        ok: true,
        message: '已导入到智能表格（与粘贴内容完全一致）',
        imported: rows.length,
        recordCount: rows.length,
        columns: headers.length,
        headers,
        rows,
      });
    }
    if (Array.isArray(body.rows) && body.rows.length > 0) {
      const data = loadSmartTable();
      const todayStr = new Date().toISOString().slice(0, 10);
      for (const b of body.rows) {
        const row = {};
        data.headers.forEach((h) => { row[h] = ''; });
        let date = (b.date != null) ? String(b.date).trim() : '';
        if (date) { const n = normalizeDateCell(date); if (n) date = toMonthDotDay(n) || n; }
        if (!date) date = todayStr;
        const roleName = (b.roleName != null ? String(b.roleName).trim() : '') || (b.角色名称 != null ? String(b.角色名称).trim() : '') || '未知';
        const diff = parseNum(b.diff ?? b.差值);
        const idxDate = getHeaderIndex(data.headers, /日期/);
        const idxRole = getHeaderIndex(data.headers, /角色|昵称/);
        const idxDiff = getHeaderIndex(data.headers, /差值/);
        const idxStart = getHeaderIndex(data.headers, /开始|起始|经验.*开始|经验值\s*[（(]?\s*开始/);
        const idxEnd = getHeaderIndex(data.headers, /结束|经验.*结束|经验值\s*[（(]?\s*结束/);
        if (idxDate >= 0) row[data.headers[idxDate]] = date;
        if (idxRole >= 0) row[data.headers[idxRole]] = roleName;
        if (idxDiff >= 0) row[data.headers[idxDiff]] = Number.isFinite(diff) ? diff : 0;
        if (idxStart >= 0) row[data.headers[idxStart]] = b.expStart ?? b['经验值(开)'] ?? '';
        if (idxEnd >= 0) row[data.headers[idxEnd]] = (b.expEnd ?? b['经验值(结束)'] ?? '') !== undefined && (b.expEnd ?? b['经验值(结束)']) !== null ? String(b.expEnd ?? b['经验值(结束)']).trim() : '';
        data.rows.push(row);
      }
      saveSmartTable(data);
      rebuildAnalyzedFromSmartTable();
      console.log('[smart-table] 追加', body.rows.length, '条，共', data.rows.length, '条');
      return res.json({ ok: true, message: '已导入到智能表格并更新分析', imported: body.rows.length, recordCount: data.rows.length });
    }
    return res.status(400).json({ error: '请提供 body.text（粘贴的 CSV/TSV）或 body.rows（数组）' });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/** 管理员上传文件：完整保留表头与所有列、所有行，与文件完全一致（替换当前智能表格） */
app.post('/api/smart-table/import-upload', requireAdmin, (req, res, next) => {
  upload.single('file')(req, res, (err) => {
    if (err) return next(err);
    try {
      if (!req.file || !req.file.buffer) return sendJson(res, 400, { error: '请选择要上传的表格文件（支持 .csv、.txt、.xlsx）' });
      const name = (req.file.originalname || '').toLowerCase();
      let text;
      if (name.endsWith('.xlsx') || name.endsWith('.xls') || (req.file.mimetype && req.file.mimetype.includes('spreadsheet'))) {
        const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
        if (!wb.SheetNames || !wb.SheetNames.length) return sendJson(res, 400, { error: 'Excel 文件中没有工作表' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        text = XLSX.utils.sheet_to_csv(sheet);
      } else {
        text = req.file.buffer.toString('utf8');
      }
      if (!text || !text.trim()) return sendJson(res, 400, { error: '文件内容为空' });
      const { headers, rows: rawRows } = parseTable(text.trim());
      if (!headers.length) return sendJson(res, 400, { error: '表头为空' });
      const rows = rawRows.map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]).trim() : ''; });
        return obj;
      });
      saveSmartTable({ headers, rows });
      rebuildAnalyzedFromSmartTable();
      console.log('[smart-table] 文件导入（全部保留）', rows.length, '条，', headers.length, '列');
      sendJson(res, 200, {
        ok: true,
        message: '已导入到智能表格（与文件完全一致）',
        imported: rows.length,
        recordCount: rows.length,
        columns: headers.length,
        headers,
        rows,
      });
    } catch (e) {
      sendJson(res, 500, { error: '解析失败: ' + (e.message || String(e)) });
    }
  });
});

/** 一次性写入多行；body.rows 为数组，按当前表头补齐列 */
app.post('/api/smart-table/rows', requireSmartTableSecret, express.json({ limit: '10mb' }), (req, res) => {
  try {
    const list = req.body && Array.isArray(req.body.rows) ? req.body.rows : [];
    const data = loadSmartTable();
    const todayStr = new Date().toISOString().slice(0, 10);
    for (const b of list) {
      let date = (b.date != null) ? String(b.date).trim() : '';
      if (date) { const n = normalizeDateCell(date); if (n) date = toMonthDotDay(n) || n; }
      if (!date) date = todayStr;
      const roleName = (b.roleName != null ? String(b.roleName).trim() : '') || (b.角色名称 != null ? String(b.角色名称).trim() : '') || '未知';
      const diff = parseNum(b.diff ?? b.差值);
      const expStart = b.expStart ?? b['经验值(开)'] ?? b.开始经验 ?? '';
      const expEnd = b.expEnd ?? b['经验值(结束)'] ?? b.结束经验 ?? '';
      const row = {};
      data.headers.forEach((h) => { row[h] = ''; });
      const idxDate = getHeaderIndex(data.headers, /日期/);
      const idxRole = getHeaderIndex(data.headers, /角色|昵称/);
      const idxDiff = getHeaderIndex(data.headers, /差值/);
      const idxStart = getHeaderIndex(data.headers, /开始|起始|经验.*开始|经验值\s*[（(]?\s*开始/);
      const idxEnd = getHeaderIndex(data.headers, /结束|经验.*结束|经验值\s*[（(]?\s*结束/);
      if (idxDate >= 0) row[data.headers[idxDate]] = date;
      if (idxRole >= 0) row[data.headers[idxRole]] = roleName;
      if (idxDiff >= 0) row[data.headers[idxDiff]] = Number.isFinite(diff) ? diff : 0;
      if (idxStart >= 0) row[data.headers[idxStart]] = expStart !== undefined && expStart !== null ? expStart : '';
      if (idxEnd >= 0) row[data.headers[idxEnd]] = expEnd !== undefined && expEnd !== null ? String(expEnd).trim() : '';
      data.rows.push(row);
    }
    saveSmartTable(data);
    rebuildAnalyzedFromSmartTable();
    console.log('[smart-table] 批量写入', list.length, '条，共', data.rows.length, '条');
    res.json({ ok: true, message: '已写入并更新分析', recordCount: data.rows.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

function runAnalyzeAndRespond(text, res) {
  const analyzed = analyzeTableContent(text);
  const debug = { detectedHeaders: analyzed.detectedHeaders || [], rawRowCount: analyzed.rawRowCount || 0 };
  if (!analyzed.rows.length) {
    return res.status(400).json({
      error: '未能解析出有效记录。请确认表头含：等级、经验开始/经验值（开始）、经验结束/经验值（结束）、角色名称；或有「差值」列。',
      ...debug,
    });
  }
  saveAnalyzed(analyzed);
  const roles = Object.keys(analyzed.byRole || {}).sort();
  const key = (s) => String(s).toLowerCase().trim();
  const byKey = {};
  for (const name of roles) {
    const k = key(name);
    if (!byKey[k]) byKey[k] = [];
    byKey[k].push(name);
  }
  const suggestMerge = Object.entries(byKey)
    .filter(([, variants]) => variants.length > 1)
    .map(([canonicalKey, variants]) => ({ key: canonicalKey, variants: variants.sort() }));
  lastAnalyzedRolesCache = { roles, suggestMerge };
  res.json({
    ok: true,
    message: '已分析并保存，下方可实时按角色搜索、按日期查询',
    recordCount: analyzed.rows.length,
    totalDiff: analyzed.totalDiff,
    byDate: analyzed.byDate,
    byRole: analyzed.byRole,
    roles,
    suggestMerge,
    ...debug,
  });
}

// 智能表格自动接入 Webhook：数据变更时由智能表格或定时任务调用，自动更新分析结果（无需管理员）
app.post('/api/ingest-from-smart-table', express.json({ limit: '6mb' }), (req, res) => {
  if (INGEST_SECRET) {
    const secret = req.headers['x-ingest-secret'];
    if (secret !== INGEST_SECRET) {
      return res.status(403).setHeader('Content-Type', 'application/json; charset=utf-8')
        .send(JSON.stringify({ error: 'X-Ingest-Secret 无效' }));
    }
  }
  const body = req.body || {};
  const url = (body.url != null) ? String(body.url).trim() : '';
  const text = (body.text != null ? body.text : body.csv != null ? body.csv : body.data != null ? body.data : '');
  if (url && /^https?:\/\//i.test(url)) {
    axios.get(url, { timeout: 30000, responseType: 'text', maxContentLength: 10 * 1024 * 1024 })
      .then((r) => {
        const t = (r.data != null && typeof r.data === 'string') ? r.data : String(r.data || '');
        if (!t.trim()) return res.status(400).json({ error: '该地址返回内容为空' });
        runAnalyzeAndRespond(t, res);
      })
      .catch((err) => {
        const msg = err.response ? `状态 ${err.response.status}` : (err.message || String(err));
        res.status(500).json({ error: '获取数据失败: ' + msg });
      });
    return;
  }
  const raw = (typeof text === 'string' ? text : (text != null ? String(text) : '')).trim();
  if (!raw) return res.status(400).json({ error: '请提供 body.url（数据地址）或 body.text/body.csv（表格 CSV/TSV 文本）' });
  runAnalyzeAndRespond(raw, res);
});

// 从智能表格地址拉取 CSV/TSV 并分析（需管理员）；用于关联智能表格，表格更新后点「同步」即可更新分析结果
app.post('/api/sync-from-url', requireAdmin, express.json({ limit: '2kb' }), (req, res) => {
  try {
    const url = (req.body && req.body.url != null) ? String(req.body.url).trim() : '';
    if (!url) return res.status(400).json({ error: '请提供 body.url（智能表格导出的 CSV 地址或公开可访问的数据链接）' });
    if (!/^https?:\/\//i.test(url)) return res.status(400).json({ error: 'url 需为 http(s) 开头' });
    axios.get(url, { timeout: 30000, responseType: 'text', maxContentLength: 10 * 1024 * 1024 })
      .then((r) => {
        const text = (r.data != null && typeof r.data === 'string') ? r.data : String(r.data || '');
        if (!text.trim()) return res.status(400).json({ error: '该地址返回内容为空' });
        runAnalyzeAndRespond(text, res);
      })
      .catch((err) => {
        const msg = err.response ? `状态 ${err.response.status}` : (err.message || String(err));
        res.status(500).json({ error: '获取智能表格数据失败: ' + msg });
      });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 分析表格：上传/粘贴 CSV、TSV 或 xlsx；需管理员
app.post('/api/analyze-table', requireAdmin, express.json({ limit: '5mb' }), (req, res) => {
  try {
    const text = (req.body && (req.body.text || req.body.csv || req.body.data)) || '';
    if (!text.trim()) return res.status(400).json({ error: '请提供表格内容：body.text 或 body.csv（支持 CSV/TSV 粘贴）' });
    runAnalyzeAndRespond(text, res);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 分析表格（文件上传）：支持 .csv / .txt / .xlsx；保证始终返回 JSON，避免前端收到 HTML
function sendJson(res, status, body) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.status(status).send(JSON.stringify(body));
}

app.post('/api/analyze-table-upload', requireAdmin, (req, res, next) => {
  upload.single('file')(req, res, (err) => {
    if (err) return next(err);
    try {
      if (!req.file || !req.file.buffer) return sendJson(res, 400, { error: '请选择要上传的表格文件（支持 .csv、.txt、.xlsx）' });
      const name = (req.file.originalname || '').toLowerCase();
      let text;
      if (name.endsWith('.xlsx') || name.endsWith('.xls') || (req.file.mimetype && req.file.mimetype.includes('spreadsheet'))) {
        const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
        if (!wb.SheetNames || !wb.SheetNames.length) return sendJson(res, 400, { error: 'Excel 文件中没有工作表' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        text = XLSX.utils.sheet_to_csv(sheet);
      } else {
        text = req.file.buffer.toString('utf8');
      }
      if (!text || !text.trim()) return sendJson(res, 400, { error: '文件内容为空' });
      runAnalyzeAndRespond(text, res);
    } catch (e) {
      sendJson(res, 500, { error: '解析失败: ' + (e.message || String(e)) });
    }
  });
});

// API 错误统一返回 JSON，避免返回 HTML 导致前端报错
app.use((err, req, res, next) => {
  if (req.path.startsWith('/api/')) {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    const msg = err.code === 'LIMIT_FILE_SIZE' ? '文件过大（最大 10MB）' : (err.message || String(err));
    return res.status(err.status || 500).send(JSON.stringify({ error: msg }));
  }
  next(err);
});

// 未匹配的 /api/* 一律返回 JSON 404，避免被 static 或其它中间件返回 HTML
app.use('/api/*', (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.status(404).send(JSON.stringify({ error: '接口不存在: ' + req.method + ' ' + req.path }));
});

// 回调只收 JSON body（智能机器人标准格式）
app.use('/callback', express.json({ limit: '2mb' }));

app.get('/callback', (req, res) => {
  try {
    const q = req.query;
    // 官方要求：对请求参数做 Urldecode，否则验证不成功
    const msgSig = q.msg_signature ? decodeURIComponent(String(q.msg_signature)) : '';
    const ts = q.timestamp ? decodeURIComponent(String(q.timestamp)) : '';
    const nonce = q.nonce ? decodeURIComponent(String(q.nonce)) : '';
    let echostr = q.echostr ? decodeURIComponent(String(q.echostr)) : '';
    if (!msgSig || !ts || !nonce || !echostr) {
      return res.status(400).send('missing params');
    }
    const sig = crypto.getSignature(TOKEN, ts, nonce, echostr);
    if (sig !== msgSig) {
      console.error('[GET] 签名不符', {
        tokenLen: TOKEN.length,
        tokenPrefix: TOKEN.slice(0, 4) + '...',
        echostrLen: echostr.length,
        computedSig: sig,
        receivedSig: msgSig,
      });
      return res.status(401).send('bad signature');
    }
    const { message } = crypto.decrypt(AES_KEY, echostr);
    console.log('[GET] 验证通过, echostr明文:', message);
    // 响应必须为明文，不能加引号、不能带 BOM、不能带换行符
    res.set('Content-Type', 'text/plain; charset=utf-8');
    res.set('Cache-Control', 'no-cache');
    return res.end(message, 'utf8');
  } catch (e) {
    console.error('[GET] 错误:', e.message);
    return res.status(500).send('error');
  }
});

app.post('/callback', async (req, res) => {
  try {
    const q = req.query;
    const msgSig = q.msg_signature;
    const ts = q.timestamp;
    const nonce = q.nonce;
    if (!msgSig || !ts || !nonce) {
      return res.status(400).send('missing params');
    }
    const body = req.body;
    const encrypt = body && (body.encrypt || body.Encrypt);
    if (!encrypt) {
      console.error('[POST] 无 encrypt');
      return res.status(400).send('no encrypt');
    }
    const sig = crypto.getSignature(TOKEN, ts, nonce, encrypt);
    if (sig !== msgSig) {
      console.error('[POST] 签名不符');
      return res.status(401).send('bad signature');
    }
    const { message } = crypto.decrypt(AES_KEY, encrypt);
    console.log('[POST] 解密成功');

    const msg = (function () {
      try {
        return JSON.parse(message);
      } catch (_) {
        return null;
      }
    })();
    if (!msg || !msg.text || !msg.text.content) {
      console.log('[POST] 非文本或无 content，跳过');
      return res.send('success');
    }
    const content = msg.text.content;
    // 优先用成员姓名（避免 userid 显示成乱码），没有再用 userid
    const sender = (msg.from && (msg.from.name || msg.from.username || msg.from.userid)) || '未知';
    console.log('[POST] 内容:', content, '发件人:', sender);

    // 只处理符合「账号 等级 开始经验 结束经验」格式的消息；其余群消息忽略（支持「接收所有群消息」模式）
    try {
      const row = parseMessage(content, sender);
      console.log('[POST] 解析结果:', row);

      const chinaTz = 'Asia/Shanghai';
      const now = new Date();
      const m = parseInt(new Intl.DateTimeFormat('zh-CN', { timeZone: chinaTz, month: 'numeric' }).format(now), 10);
      const d = row.date != null ? row.date : parseInt(new Intl.DateTimeFormat('zh-CN', { timeZone: chinaTz, day: 'numeric' }).format(now), 10);
      const dateDisplay = `${m}.${d}`;
      const data = loadSmartTable();
      const smartRow = {};
      data.headers.forEach((h) => { smartRow[h] = ''; });
      const idxDate = getHeaderIndex(data.headers, /日期/);
      const idxRole = getHeaderIndex(data.headers, /角色|昵称/);
      const idxDiff = getHeaderIndex(data.headers, /差值/);
      const idxStart = getHeaderIndex(data.headers, /开始|起始|经验.*开始|经验值\s*[（(]?\s*开始/);
      const idxEnd = getHeaderIndex(data.headers, /结束|经验.*结束|经验值\s*[（(]?\s*结束/);
      if (idxDate >= 0) smartRow[data.headers[idxDate]] = dateDisplay;
      if (idxRole >= 0) smartRow[data.headers[idxRole]] = row.roleName || '未知';
      if (idxDiff >= 0) smartRow[data.headers[idxDiff]] = row.diff != null ? row.diff : 0;
      if (idxStart >= 0) smartRow[data.headers[idxStart]] = row.expStart !== undefined && row.expStart !== null ? row.expStart : '';
      if (idxEnd >= 0) smartRow[data.headers[idxEnd]] = row.expEnd !== undefined && row.expEnd !== null ? String(row.expEnd) : '';
      data.rows.push(smartRow);
      saveSmartTable(data);
      rebuildAnalyzedFromSmartTable();
      console.log('[callback] 已写入服务器智能表格，共', data.rows.length, '条');

      const rowWithDate = { ...row, date: dateDisplay };
      const sheetRes = await appendSheet(rowWithDate);
      if (sheetRes && sheetRes.errcode && sheetRes.errcode !== 0) {
        console.error('[POST] 表格写入失败', sheetRes.errcode, sheetRes.errmsg);
      } else {
        console.log('[POST] 表格已写入');
        saveRecord(row);
      }
      await appendSummarySheet(rowWithDate);
    } catch (parseErr) {
      // 非经验记录格式（如普通聊天），直接跳过，不报错
      console.log('[POST] 非经验记录格式，跳过:', parseErr.message);
    }
    return res.send('success');
  } catch (e) {
    console.error('[POST] 错误:', e.message);
    return res.status(500).send('error');
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log('服务已启动 端口:', port);
  try {
    const data = loadSmartTable();
    if (data.rows.length > 0) rebuildAnalyzedFromSmartTable();
  } catch (e) {
    console.error('[smart-table] 启动重建分析失败', e.message);
  }
  console.log('WECOM_TOKEN 来自环境变量:', process.env.WECOM_TOKEN ? '是' : '否');
  console.log('回调: /callback  主表Webhook:', SHEET_WEBHOOK ? '已配置' : '未配置');
  if (SUMMARY_WEBHOOK) console.log('汇总表Webhook: 已配置（按角色合计明细）');
  const syncUrl = process.env.SYNC_URL || '';
  const syncMins = parseInt(process.env.SYNC_INTERVAL_MINUTES || '0', 10);
  if (syncUrl && /^https?:\/\//i.test(syncUrl) && syncMins > 0) {
    const doSync = () => {
      axios.get(syncUrl, { timeout: 30000, responseType: 'text', maxContentLength: 10 * 1024 * 1024 })
        .then((r) => {
          const text = (r.data != null && typeof r.data === 'string') ? r.data : String(r.data || '');
          if (text.trim()) {
            const analyzed = analyzeTableContent(text);
            if (analyzed.rows && analyzed.rows.length > 0) {
              saveAnalyzed(analyzed);
              console.log('[智能表格同步] 已更新，共', analyzed.rows.length, '条');
            }
          }
        })
        .catch((e) => console.error('[智能表格同步] 失败:', e.message || e));
    };
    doSync();
    setInterval(doSync, syncMins * 60 * 1000);
    console.log('智能表格自动同步: 已开启，每', syncMins, '分钟从 SYNC_URL 拉取并更新');
  }
});
