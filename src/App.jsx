import { useState, useEffect, useCallback, useRef, useMemo } from “react”;
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, AreaChart, Area, CartesianGrid } from “recharts”;
import * as XLSX from “xlsx”;

// ─────────────────────────────────────────────────────────────────────────────
// КОНСТАНТАЛАР
// ─────────────────────────────────────────────────────────────────────────────

const APP_VERSION = “1.0.0”;
const STORAGE_KEYS = {
USERS_INDEX: “ad_users_v1”,
PRO_CODES:   “ad_pro_v1”,
};

const PRO_CODES = [
“PRO-2847”,“PRO-3K9X”,“PRO-7M4Z”,“PRO-9B2W”,“PRO-4T8R”,
“PRO-6N5V”,“PRO-1Q7F”,“PRO-8C3D”,“PRO-5H2J”,“PRO-2X9L”,
“PRO-7G4S”,“PRO-3P8K”,“PRO-9W1M”,“PRO-4Y6B”,“PRO-6R3T”,
“PRO-1Z8N”,“PRO-8V5Q”,“PRO-5F2H”,“PRO-3E7C”,“PRO-9U4W”,
];

const CATEGORIES = [
{ id: “food”,       label: { mn: “🍔 Хоол”,        kz: “🍔 Тамақ”    }, color: “#ef4444”, needs: true  },
{ id: “transport”,  label: { mn: “🚗 Тээвэр”,       kz: “🚗 Көлік”    }, color: “#f97316”, needs: true  },
{ id: “housing”,    label: { mn: “🏠 Орон сууц”,    kz: “🏠 Үй”       }, color: “#3b82f6”, needs: true  },
{ id: “health”,     label: { mn: “💊 Эрүүл мэнд”,   kz: “💊 Денсаулық”}, color: “#10b981”, needs: true  },
{ id: “education”,  label: { mn: “📚 Боловсрол”,    kz: “📚 Білім”    }, color: “#06b6d4”, needs: true  },
{ id: “entertain”,  label: { mn: “🎮 Зугаа”,        kz: “🎮 Ойын”     }, color: “#8b5cf6”, wants: true  },
{ id: “clothes”,    label: { mn: “👕 Хувцас”,       kz: “👕 Киім”     }, color: “#ec4899”, wants: true  },
{ id: “phone”,      label: { mn: “📱 Утас/Net”,     kz: “📱 Байланыс” }, color: “#6366f1”, wants: true  },
{ id: “cafe”,       label: { mn: “☕ Кафе”,          kz: “☕ Кафе”     }, color: “#d97706”, wants: true  },
{ id: “savings”,    label: { mn: “💎 Хадгаламж”,    kz: “💎 Жинақ”    }, color: “#f59e0b”, saving: true },
{ id: “invest”,     label: { mn: “📈 Хөрөнгө”,      kz: “📈 Инвестиция”}, color: “#84cc16”, saving: true },
{ id: “fee”,        label: { mn: “🏦 Хураамж”,      kz: “🏦 Комиссия” }, color: “#94a3b8”               },
{ id: “transfer”,   label: { mn: “💸 Шилжүүлэг”,   kz: “💸 Аударым”  }, color: “#a78bfa”               },
{ id: “income”,     label: { mn: “💰 Орлого”,       kz: “💰 Кіріс”    }, color: “#22c55e”, isIncome: true},
{ id: “other”,      label: { mn: “📦 Бусад”,        kz: “📦 Басқа”    }, color: “#64748b”               },
];

const TEXTS = {
mn: {
appName: “Мөнгөний Дэвтэр”,
login: “Нэвтрэх”, register: “Бүртгүүлэх”,
yourName: “Таны нэр”, pin4: “4 оронтой PIN”,
pinAgain: “PIN давтана уу”,
nameTooShort: “Нэр хэт богино (2+ тэмдэгт)”,
pinInvalid: “PIN яг 4 тоо байх ёстой”,
pinMismatch: “PIN таарахгүй байна”,
userNotFound: “Хэрэглэгч олдсонгүй”,
pinWrong: “PIN буруу”,
noAccount: “Шинэ хэрэглэгч?”, hasAccount: “Бүртгэлтэй юу?”,
secureNote: “Таны өгөгдөл зөвхөн танд харагдана 🔐”,
logout: “Гарах”,
tabs: { dash: “📊 Нүүр”, bank: “🏦 Банк”, add: “➕ Нэмэх”, history: “📅 Түүх”, math: “🧮 Тооцоо”, tips: “💡 Зөвлөгөө” },
income: “Орлого”, expense: “Зарлага”, balance: “Үлдэгдэл”,
healthScore: “Санхүүгийн эрүүл мэнд”,
catLabel: “Ангилал”, descLabel: “Тайлбар”, amountLabel: “Дүн (₮)”, dateLabel: “Огноо”,
saveBtn: “Хадгалах”, deleteBtn: “Устгах”,
uploadXlsx: “Банкны хуулга оруулах (.xlsx)”,
supportedBanks: “Хаан · Голомт · Хас · Ард банк”,
txCount: “Гүйлгээний тоо”,
noData: “Өгөгдөл алга”,
proCode: “PRO код”, proActivate: “Идэвхжүүлэх”, proClose: “Хаах”,
proCodeError: “Код буруу эсвэл ашиглагдсан”,
proUnlock: “PRO горим идэвхжүүлэх”,
installApp: “Апп болгон суулгах”,
installIos: “Safari → Хуваалцах → Нүүр дэлгэцэнд”,
installAndroid: “Chrome → Цэс → Нүүр дэлгэцэнд”,
clearAndStart: “Бүгдийг устгаж эхлэх”,
demoStart: “Демо өгөгдлөөр эхлэх”,
monthlySum: “Сарын хураангуй”,
savingsRate: “Хадгаламжийн хувь”,
spendRate: “Зарлагын хувь”,
dailyAvg: “Өдөрт дундаж”,
weeklyAvg: “7 хоногт дундаж”,
emergencyFund: “Яаралтай сан”,
monthsToFund: “Яаралтай санд хүрэх”,
annualSavings: “Жилийн хадгаламж”,
annualSpend: “Жилийн зарлага”,
investForecast: “Хөрөнгө оруулалтын таамаглал”,
investNote: “Сар бүр хадгаламж · 8% жилийн хүү (нийлмэл)”,
formulaNote: “FV = PMT × ((1+r)ⁿ − 1) / r”,
rule5030: “50/30/20 дүрэм”,
needs: “Хэрэгцээ”, wants: “Хүсэл”, savingsW: “Хадгаламж”,
excellent: “Маш сайн ✅”, average: “Дундаж ⚠️”, poor: “Муу 🔴”,
aiTitle: “🤖 Ухаалаг зөвлөгч”,
aiAnalyze: “Миний санхүүг шинжилж зөвлөгөө өг”,
aiLoading: “Тооцоолж байна…”,
filterThisMonth: “Энэ сар”, filterLast3: “3 сар”, filterYear: “Жил”, filterAll: “Бүгд”,
chartMonthly: “Сарын орлого / зарлага”,
recent: “Сүүлийн бичлэгүүд”,
bankSummary: “Банкны хуулгын хураангуй”,
processing: “Уншиж байна…”,
bankError: “Файл уншихад алдаа гарлаа. Монгол банкны Excel хуулга оруулна уу.”,
tip1t: “50/30/20 дүрэм”, tip1: “Орлогын 50%-ийг хэрэгцээ, 30%-ийг хүсэл, 20%-ийг хадгаламжид зарцуулах нь дэлхийн алтан стандарт.”,
tip2t: “Жижиг зардал”, tip2: “Өдөр бүр нэг кофе = сард 27,000₮ = жилд 324,000₮. Жижиг зардлуудыг хянах нь маш чухал.”,
tip3t: “Яаралтай сан”, tip3: “3-6 сарын зарлагатай тэнцэх яаралтай сан байлгах нь ямар ч эрсдэлийг даван туулахад тусална.”,
tip4t: “Автомат хадгаламж”, tip4: “Цалин авсан даруй хадгаламждаа шилжүүл. «Өөртөө эхлэж төл» — Роберт Кийосаки.”,
tip5t: “Банкны хураамж”, tip5: “Ухаалаг банкны хураамжийг хянаарай. Жилд мөнгийг нэмж чухал сан болно.”,
tip6t: “Зээлийн карт”, tip6: “Зээлийн картыг бүрэн төлөөрэй. 20-25% хүүтэй зээл хамгийн үнэтэй мөнгө.”,
tip7t: “24 цагийн дүрэм”, tip7: “Том худалдан авалтын өмнө 24 цаг хүлээ. 80% тохиолдолд хүсэл өөрөө арилна.”,
tip8t: “Хөрөнгө оруулалт”, tip8: “Орлогын 10%-ийг жил бүр хөрөнгө оруулбал 30 жилд 8% жилийн өгөөжтэй 10 дахин өснө.”,
tip9t: “Захиалга шалгах”, tip9: “Сар бүр ашиглагдаагүй захиалгуудаа шалгаарай. Дундаж хүн сард 3-4 буруу захиалгад мөнгө алддаг.”,
tip10t: “Хоол хүнс”, tip10: “Жагсаалтгүй дэлгүүр орвол 35% илүү зарцуулна. Өлсөж байхдаа битгий яв.”,
},
kz: {
appName: “Ақша Дәптер”,
login: “Кіру”, register: “Тіркелу”,
yourName: “Атыңыз”, pin4: “4 санды PIN”,
pinAgain: “PIN қайталаңыз”,
nameTooShort: “Ат тым қысқа (2+ таңба)”,
pinInvalid: “PIN дәл 4 сан болуы керек”,
pinMismatch: “PIN сәйкес емес”,
userNotFound: “Пайдаланушы табылмады”,
pinWrong: “PIN қате”,
noAccount: “Жаңа пайдаланушы?”, hasAccount: “Тіркелгенсіз бе?”,
secureNote: “Деректеріңіз тек сізге көрінеді 🔐”,
logout: “Шығу”,
tabs: { dash: “📊 Басты”, bank: “🏦 Банк”, add: “➕ Қосу”, history: “📅 Тарих”, math: “🧮 Математика”, tips: “💡 Кеңестер” },
income: “Кіріс”, expense: “Шығыс”, balance: “Қалды”,
healthScore: “Қаржылық денсаулық”,
catLabel: “Категория”, descLabel: “Сипаттама”, amountLabel: “Сома (₮)”, dateLabel: “Күні”,
saveBtn: “Сақтау”, deleteBtn: “Өшіру”,
uploadXlsx: “Банк анықтамасын жүктеу (.xlsx)”,
supportedBanks: “Хаан · Голомт · Хас · Ард банк”,
txCount: “Мәміле саны”,
noData: “Деректер жоқ”,
proCode: “PRO код”, proActivate: “Белсендіру”, proClose: “Жабу”,
proCodeError: “Код қате немесе пайдаланылған”,
proUnlock: “PRO режимді ашу”,
installApp: “Қосымша ретінде орнату”,
installIos: “Safari → Бөлісу → Басты экранға”,
installAndroid: “Chrome → Мәзір → Басты экранға”,
clearAndStart: “Барлығын тазартып бастау”,
demoStart: “Демо деректермен бастау”,
monthlySum: “Ай қорытындысы”,
savingsRate: “Жинақ пайызы”,
spendRate: “Шығыс пайызы”,
dailyAvg: “Күніне орта”,
weeklyAvg: “Аптасына орта”,
emergencyFund: “Апат қоры”,
monthsToFund: “Апат қорына дейін”,
annualSavings: “Жылдық жинақ”,
annualSpend: “Жылдық шығыс”,
investForecast: “Инвестиция болжамы”,
investNote: “Ай сайынғы жинақ · 8% жылдық пайыз (күрделі)”,
formulaNote: “FV = PMT × ((1+r)ⁿ − 1) / r”,
rule5030: “50/30/20 ережесі”,
needs: “Қажеттілік”, wants: “Тілек”, savingsW: “Жинақ”,
excellent: “Өте жақсы ✅”, average: “Орташа ⚠️”, poor: “Нашар 🔴”,
aiTitle: “🤖 Ақылды кеңесші”,
aiAnalyze: “Менің қаржымды талдап кеңес бер”,
aiLoading: “Есептелуде…”,
filterThisMonth: “Бұл ай”, filterLast3: “3 ай”, filterYear: “Жыл”, filterAll: “Барлығы”,
chartMonthly: “Ай кірісі / шығысы”,
recent: “Соңғы жазбалар”,
bankSummary: “Банк анықтамасының қорытындысы”,
processing: “Оқылуда…”,
bankError: “Файл оқылмады. Монгол банктің Excel анықтамасын жүктеңіз.”,
tip1t: “50/30/20 ережесі”, tip1: “Кірістің 50% — қажеттілік, 30% — тілек, 20% — жинақ. Бұл әлемдік алтын стандарт.”,
tip2t: “Ұсақ шығын”, tip2: “Күнде 1 кофе = айына 9,000₮ = жылына 108,000₮. Ұсақ шығынды бақылау маңызды.”,
tip3t: “Апат қоры”, tip3: “3-6 айлық шығын мөлшерінде апат қоры болуы кез келген тәуекелді жеңуге көмектеседі.”,
tip4t: “Автожинақ”, tip4: “Жалақы түскен күні бірден жинаққа аударыңыз. «Өзіңізге алдымен төлеңіз» — Р. Кийосаки.”,
tip5t: “Банк комиссиясы”, tip5: “Банк комиссияларын бақылаңыз. Жылдық жиынтық үлкен соманы құрайды.”,
tip6t: “Несие картасы”, tip6: “Несие картасын толық төлеңіз. 20-25% пайыздық зайым — ең қымбат ақша.”,
tip7t: “24 сағат ережесі”, tip7: “Үлкен сатып алу алдында 24 сағат күтіңіз. 80% жағдайда тілек өзі өтеді.”,
tip8t: “Инвестиция”, tip8: “Кірістің 10% жыл сайын инвестицияласаңыз, 30 жылда 8% жылдық кірісте 10x өседі.”,
tip9t: “Жазылымдар”, tip9: “Ай сайын пайдаланылмаған жазылымдарды тексеріңіз. Орта есеп 3-4 бос жазылым бар.”,
tip10t: “Азық-түлік”, tip10: “Тізімсіз дүкенге барсаңыз 35% артық жұмсайсыз. Аш қарынмен бармаңыз.”,
},
};

// ─────────────────────────────────────────────────────────────────────────────
// AUTH UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

function getUsersIndex() {
try { return JSON.parse(localStorage.getItem(STORAGE_KEYS.USERS_INDEX) || “{}”); }
catch { return {}; }
}

function saveUsersIndex(data) {
localStorage.setItem(STORAGE_KEYS.USERS_INDEX, JSON.stringify(data));
}

async function sha256(text) {
const buf = await crypto.subtle.digest(“SHA-256”, new TextEncoder().encode(text));
return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, “0”)).join(””);
}

async function hashPin(userId, pin) {
return sha256(`${userId}:${pin}:aksha_v1`);
}

function getUserDataKey(userId) {
return `ad_data_${userId}`;
}

function saveUserData(userId, data) {
try { localStorage.setItem(getUserDataKey(userId), JSON.stringify(data)); }
catch {}
}

function loadUserData(userId) {
try { return JSON.parse(localStorage.getItem(getUserDataKey(userId)) || “null”); }
catch { return null; }
}

// ─────────────────────────────────────────────────────────────────────────────
// PRO CODE UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

function getDeviceId() {
try {
const raw = [navigator.userAgent, screen.width, screen.height, navigator.language].join(”|”);
let h = 5381;
for (let i = 0; i < raw.length; i++) h = (h * 33) ^ raw.charCodeAt(i);
return Math.abs(h >>> 0).toString(36).toUpperCase();
} catch { return “DEVICE”; }
}

function checkProCode(code, userId) {
try {
const db = JSON.parse(localStorage.getItem(STORAGE_KEYS.PRO_CODES) || “{}”);
const c = code.trim().toUpperCase();
const dev = getDeviceId();
if (!PRO_CODES.includes(c)) return { ok: false };
if (db[c] && db[c] !== `${userId}:${dev}`) return { ok: false, taken: true };
db[c] = `${userId}:${dev}`;
localStorage.setItem(STORAGE_KEYS.PRO_CODES, JSON.stringify(db));
return { ok: true };
} catch { return { ok: false }; }
}

function isProActive(userId) {
try {
const db = JSON.parse(localStorage.getItem(STORAGE_KEYS.PRO_CODES) || “{}”);
const dev = getDeviceId();
return PRO_CODES.some(k => db[k] === `${userId}:${dev}`);
} catch { return false; }
}

// ─────────────────────────────────────────────────────────────────────────────
// BANK XLSX PARSER
// ─────────────────────────────────────────────────────────────────────────────

const COL_HINTS = {
date:    [“огноо”, “date”, “гүйлгээний огноо”, “он сар өдөр”],
credit:  [“кредит гүйлгээ”, “кредит”, “орлого”, “credit”, “нэмэгдэл”],
debit:   [“дебит гүйлгээ”, “дебит”, “зарлага”, “debit”, “хасагдал”],
balance: [“эцсийн үлдэгдэл”, “үлдэгдэл”, “balance”],
desc:    [“гүйлгээний утга”, “утга”, “description”, “тайлбар”],
};

function detectColumns(headerRow) {
const map = {};
(headerRow || []).forEach((cell, i) => {
const v = String(cell || “”).toLowerCase().trim();
for (const [key, hints] of Object.entries(COL_HINTS)) {
if (!map[key] && hints.some(h => v.includes(h))) map[key] = i;
}
});
return map;
}

function parseBankXlsx(rawData) {
// Find header row
let headerRowIdx = -1;
let colMap = {};
for (let i = 0; i < Math.min(rawData.length, 20); i++) {
const m = detectColumns(rawData[i]);
if (Object.keys(m).length >= 3) { headerRowIdx = i; colMap = m; break; }
}
if (headerRowIdx === -1) return null;

// Extract metadata from rows before header
let owner = “”;
let period = “”;
for (let i = 0; i < headerRowIdx; i++) {
const row = rawData[i] || [];
row.forEach((cell, j) => {
const v = String(cell || “”).toLowerCase();
if (v.includes(“хэрэглэгч”) || v.includes(“нэр”)) {
const next = row[j + 3] || row[j + 2] || row[j + 1];
if (next && String(next).trim()) owner = String(next).trim();
}
if (v.includes(“интервал”) || v.includes(“хугацаа”)) {
const next = row[j + 3] || row[j + 2] || row[j + 1];
if (next) period = String(next).trim();
}
});
}

// Parse transactions
const transactions = [];
let totalCredit = 0;
let totalDebit = 0;

for (let i = headerRowIdx + 1; i < rawData.length; i++) {
const row = rawData[i];
if (!row || row.every(c => !c)) continue;

```
const firstCell = String(row[0] || "").toLowerCase();
if (firstCell.includes("нийт") || firstCell.includes("total")) continue;

const parseNum = val => Math.abs(parseFloat(String(val || "0").replace(/[^\d.-]/g, "")) || 0);

const credit  = colMap.credit  !== undefined ? parseNum(row[colMap.credit])  : 0;
const debit   = colMap.debit   !== undefined ? parseNum(row[colMap.debit])   : 0;
const balance = colMap.balance !== undefined ? parseNum(row[colMap.balance]) : 0;
const desc    = colMap.desc    !== undefined ? String(row[colMap.desc] || "").trim() : "";

let date = "";
if (colMap.date !== undefined && row[colMap.date]) {
  const d = row[colMap.date];
  date = d instanceof Date ? d.toISOString().slice(0, 10) : String(d).slice(0, 10);
}

if (credit > 0 || debit > 0) {
  totalCredit += credit;
  totalDebit  += debit;
  transactions.push({ date, credit, debit, balance, desc, id: i });
}
```

}

return { owner, period, totalCredit, totalDebit, net: totalCredit - totalDebit, transactions };
}

function autoCatFromDesc(desc) {
const d = (desc || “”).toLowerCase();
if (d.includes(“хураамж”) || d.includes(“fee”)) return “fee”;
if (d.includes(“шилжүүлэг”) || d.includes(“transfer”)) return “transfer”;
if (d.includes(“хоол”) || d.includes(“food”) || d.includes(“ресторан”)) return “food”;
if (d.includes(“такси”) || d.includes(“ubер”) || d.includes(“тээвэр”)) return “transport”;
if (d.includes(“эмнэлэг”) || d.includes(“эм “)) return “health”;
if (d.includes(“хадгаламж”) || d.includes(“депозит”)) return “savings”;
if (d.includes(“кафе”) || d.includes(“cafe”) || d.includes(“coffee”)) return “cafe”;
if (d.includes(“интернет”) || d.includes(“утас”) || d.includes(“мобайл”)) return “phone”;
return “other”;
}

// ─────────────────────────────────────────────────────────────────────────────
// MATH ENGINE — Математикалық логикаға негізделген ақылды кеңесші
// ─────────────────────────────────────────────────────────────────────────────

function calcMetrics(income, expenses) {
const spending = expenses.filter(e => !e.isIncome && !CATEGORIES.find(c => c.id === e.cat)?.saving);
const savingExp = expenses.filter(e => CATEGORIES.find(c => c.id === e.cat)?.saving);
const needsExp  = expenses.filter(e => CATEGORIES.find(c => c.id === e.cat)?.needs && !e.isIncome);
const wantsExp  = expenses.filter(e => CATEGORIES.find(c => c.id === e.cat)?.wants);

const totalSpent  = spending.reduce((s, e) => s + e.amount, 0);
const totalSaved  = savingExp.reduce((s, e) => s + e.amount, 0);
const totalNeeds  = needsExp.reduce((s, e) => s + e.amount, 0);
const totalWants  = wantsExp.reduce((s, e) => s + e.amount, 0);

const savingsRate = income > 0 ? totalSaved / income : 0;
const spendRate   = income > 0 ? totalSpent / income : 0;

// 50/30/20 skorlar
const ideal50 = income * 0.5;
const ideal30 = income * 0.3;
const ideal20 = income * 0.2;
const score50 = totalNeeds <= ideal50 ? 100 : Math.max(0, Math.round(100 - ((totalNeeds - ideal50) / ideal50) * 100));
const score30 = totalWants <= ideal30 ? 100 : Math.max(0, Math.round(100 - ((totalWants - ideal30) / ideal30) * 100));
const score20 = totalSaved >= ideal20 ? 100 : Math.round((totalSaved / ideal20) * 100);
const healthScore = Math.round((score50 + score30 + score20) / 3);

const perDay   = Math.round(totalSpent / 30);
const perWeek  = Math.round(totalSpent / 4.33);
const emgTarget = totalSpent * 6;
const monthsToEmg = totalSaved > 0 ? Math.ceil(emgTarget / totalSaved) : null;

// Investment forecasts (compound interest FV = PMT × ((1+r)^n - 1) / r)
const r = 0.08 / 12;
const forecasts = [1, 3, 5, 10, 15, 20, 30].map(years => {
const n = years * 12;
const fv = totalSaved > 0 ? Math.round(totalSaved * ((Math.pow(1 + r, n) - 1) / r)) : 0;
return { years, fv };
});

return {
totalSpent, totalSaved, totalNeeds, totalWants,
savingsRate, spendRate,
ideal50, ideal30, ideal20,
score50, score30, score20, healthScore,
perDay, perWeek, emgTarget, monthsToEmg,
forecasts, balance: income - totalSpent,
};
}

function generateAdvice(metrics, income, lang) {
const T = TEXTS[lang];
const { healthScore, savingsRate, spendRate, totalSaved, totalSpent,
totalNeeds, totalWants, ideal50, ideal30, ideal20,
perDay, monthsToEmg, emgTarget, balance } = metrics;
const advice = [];

// 1. Ерекше денсаулық ұпайы бойынша
if (healthScore >= 85) {
advice.push({
icon: “🏆”, type: “success”,
title: lang === “mn” ? “Гайхалтай санхүүгийн байдал!” : “Тамаша қаржылық жағдай!”,
text: lang === “mn”
? `Таны санхүүгийн эрүүл мэндийн оноо ${healthScore}/100. Та 50/30/20 дүрмийг маш сайн мөрдөж байна.`
: `Сіздің қаржылық денсаулық ұпайыңыз ${healthScore}/100. 50/30/20 ережесін керемет сақтап отырсыз.`,
});
} else if (healthScore < 50) {
advice.push({
icon: “🚨”, type: “danger”,
title: lang === “mn” ? “Яаралтай анхаарал шаардлагатай” : “Шұғыл назар аудару керек”,
text: lang === “mn”
? `Оноо ${healthScore}/100. Зардлыг ${(spendRate * 100).toFixed(0)}%-иас 80%-иас доош бууруулах шаардлагатай.`
: `Ұпай ${healthScore}/100. Шығысты ${(spendRate * 100).toFixed(0)}%-дан 80%-дан төмен түсіру керек.`,
});
}

// 2. Жинақ жылдамдығы
if (savingsRate < 0.1) {
advice.push({
icon: “💎”, type: “warning”,
title: lang === “mn” ? “Хадгаламж хэт бага” : “Жинақ тым аз”,
text: lang === “mn”
? `Одоогийн хадгаламж ${(savingsRate * 100).toFixed(1)}%. Зорилго: ${income > 0 ? Math.round(income * 0.2).toLocaleString() : 0}₮/сар (20%). Сар бүр ${Math.round((ideal20 - totalSaved)).toLocaleString()}₮ нэмэх хэрэгтэй.`
: `Қазіргі жинақ ${(savingsRate * 100).toFixed(1)}%. Мақсат: ${income > 0 ? Math.round(income * 0.2).toLocaleString() : 0}₮/ай (20%). Ай сайын ${Math.round((ideal20 - totalSaved)).toLocaleString()}₮ қосу керек.`,
});
} else if (savingsRate >= 0.2) {
advice.push({
icon: “💰”, type: “success”,
title: lang === “mn” ? “Хадгаламж маш сайн!” : “Жинақ өте жақсы!”,
text: lang === “mn”
? `${(savingsRate * 100).toFixed(1)}% хадгалж байна. 30 жилд ${metrics.forecasts.find(f => f.years === 30)?.fv.toLocaleString()}₮ болно (8% хүүгээр).`
: `${(savingsRate * 100).toFixed(1)}% жинақтап жатырсыз. 30 жылда ${metrics.forecasts.find(f => f.years === 30)?.fv.toLocaleString()}₮ болады (8% пайызбен).`,
});
}

// 3. Апат қоры
if (monthsToEmg !== null) {
advice.push({
icon: “🛡️”, type: “info”,
title: lang === “mn” ? “Яаралтай сангийн тооцоо” : “Апат қоры есебі”,
text: lang === “mn”
? `Зорилго: ${emgTarget.toLocaleString()}₮ (6 сарын зарлага). Одоогийн хадгаламж хурдаар ${monthsToEmg} сарын дараа хүрнэ.`
: `Мақсат: ${emgTarget.toLocaleString()}₮ (6 айлық шығыс). Қазіргі жинақпен ${monthsToEmg} айдан кейін жетеді.`,
});
}

// 4. Шығыс артық болса
if (totalNeeds > ideal50) {
const over = Math.round(totalNeeds - ideal50);
advice.push({
icon: “🏠”, type: “warning”,
title: lang === “mn” ? “Хэрэгцээний зардал их байна” : “Қажеттілік шығысы артық”,
text: lang === “mn”
? `Хэрэгцээний зардал ${totalNeeds.toLocaleString()}₮ (${Math.round(totalNeeds / (income || 1) * 100)}%). Зорилго 50%-иас доош. ${over.toLocaleString()}₮ хэтэрсэн.`
: `Қажеттілік шығысы ${totalNeeds.toLocaleString()}₮ (${Math.round(totalNeeds / (income || 1) * 100)}%). Мақсат 50%-дан төмен. ${over.toLocaleString()}₮ артық.`,
});
}

// 5. Ойын-сауық артық болса
if (totalWants > ideal30) {
const over = Math.round(totalWants - ideal30);
advice.push({
icon: “🎮”, type: “warning”,
title: lang === “mn” ? “Хүслийн зардал их байна” : “Тілек шығысы артық”,
text: lang === “mn”
? `Хүслийн зардал ${totalWants.toLocaleString()}₮. Зорилго ${ideal30.toLocaleString()}₮ (30%). ${over.toLocaleString()}₮ хэтэрсэн. Сар бүр ${Math.round(over / 3).toLocaleString()}₮ хэмнэхэд хялбар.`
: `Тілек шығысы ${totalWants.toLocaleString()}₮. Мақсат ${ideal30.toLocaleString()}₮ (30%). ${over.toLocaleString()}₮ артық. Ай сайын ${Math.round(over / 3).toLocaleString()}₮ үнемдеу оңай.`,
});
}

// 6. Инвестиция болжамы
if (totalSaved > 0) {
const f5  = metrics.forecasts.find(f => f.years === 5)?.fv  || 0;
const f10 = metrics.forecasts.find(f => f.years === 10)?.fv || 0;
advice.push({
icon: “📈”, type: “info”,
title: lang === “mn” ? “Хөрөнгө оруулалтын боломж” : “Инвестиция мүмкіндігі”,
text: lang === “mn”
? `Сар бүр ${totalSaved.toLocaleString()}₮ хадгалж, 8% хүүгээр: 5 жилд → ${f5.toLocaleString()}₮, 10 жилд → ${f10.toLocaleString()}₮.`
: `Ай сайын ${totalSaved.toLocaleString()}₮ жинақтап, 8% пайызбен: 5 жылда → ${f5.toLocaleString()}₮, 10 жылда → ${f10.toLocaleString()}₮.`,
});
}

// 7. Баланс теріс болса
if (balance < 0) {
advice.push({
icon: “🚫”, type: “danger”,
title: lang === “mn” ? “Зарлага орлогоос их байна!” : “Шығыс кірістен артық!”,
text: lang === “mn”
? `${Math.abs(balance).toLocaleString()}₮ дутагдалтай байна. Яаралтай зардлыг танах шаардлагатай.`
: `${Math.abs(balance).toLocaleString()}₮ тапшылық бар. Шығысты шұғыл азайту керек.`,
});
}

return advice.slice(0, 6);
}

// ─────────────────────────────────────────────────────────────────────────────
// DEMO DATA
// ─────────────────────────────────────────────────────────────────────────────

const TODAY = new Date().toISOString().slice(0, 10);
const THIS_MONTH = TODAY.slice(0, 7);

const DEMO_EXPENSES = [
{ id: 1,  desc: “Хоол хүнс”,        amount: 85000,  cat: “food”,      date: `${THIS_MONTH}-05`, isIncome: false },
{ id: 2,  desc: “Автобус, такси”,    amount: 35000,  cat: “transport”, date: `${THIS_MONTH}-06`, isIncome: false },
{ id: 3,  desc: “Орон суцны түрээс”, amount: 350000, cat: “housing”,   date: `${THIS_MONTH}-01`, isIncome: false },
{ id: 4,  desc: “Эрүүл мэнд”,        amount: 25000,  cat: “health”,    date: `${THIS_MONTH}-08`, isIncome: false },
{ id: 5,  desc: “Зугаа цэнгэл”,      amount: 45000,  cat: “entertain”, date: `${THIS_MONTH}-10`, isIncome: false },
{ id: 6,  desc: “Хадгаламж”,         amount: 150000, cat: “savings”,   date: `${THIS_MONTH}-01`, isIncome: false },
{ id: 7,  desc: “Банкны хураамж”,    amount: 5000,   cat: “fee”,       date: `${THIS_MONTH}-15`, isIncome: false },
{ id: 8,  desc: “Хувцас”,            amount: 60000,  cat: “clothes”,   date: `${THIS_MONTH}-12`, isIncome: false },
];

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const fmt   = n => (n || 0).toLocaleString() + “ ₮”;
const fmtSh = n => n >= 1e6 ? (n / 1e6).toFixed(1) + “сая” : n >= 1000 ? (n / 1000).toFixed(0) + “к” : String(Math.round(n));
const pct   = (a, b) => b > 0 ? Math.round((a / b) * 100) : 0;

const HEALTH_COLOR = s => s >= 75 ? “#22c55e” : s >= 50 ? “#f59e0b” : “#ef4444”;
const HEALTH_LABEL = (s, lang) => s >= 75 ? TEXTS[lang].excellent : s >= 50 ? TEXTS[lang].average : TEXTS[lang].poor;

// ─────────────────────────────────────────────────────────────────────────────
// STYLES (design tokens)
// ─────────────────────────────────────────────────────────────────────────────

const S = {
bg: “#0d0d14”,
surface: “rgba(255,255,255,0.04)”,
surfaceHover: “rgba(255,255,255,0.07)”,
border: “rgba(255,255,255,0.08)”,
borderStrong: “rgba(255,255,255,0.15)”,
text: “#f1f5f9”,
textMuted: “rgba(255,255,255,0.45)”,
accent: “#6366f1”,
accentLight: “rgba(99,102,241,0.15)”,
success: “#22c55e”,
warning: “#f59e0b”,
danger: “#ef4444”,
radius: “16px”,
radiusSm: “10px”,
};

const card = (extra = {}) => ({
background: S.surface,
borderRadius: S.radius,
border: `1px solid ${S.border}`,
padding: 16,
…extra,
});

const inputStyle = {
width: “100%”,
background: “rgba(255,255,255,0.06)”,
border: `1px solid ${S.borderStrong}`,
borderRadius: S.radiusSm,
padding: “11px 14px”,
color: S.text,
fontSize: 14,
outline: “none”,
boxSizing: “border-box”,
};

// ─────────────────────────────────────────────────────────────────────────────
// APP
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
const [lang, setLang] = useState(“mn”);
const T = TEXTS[lang];

// ── Auth state ─────────────────────────────────────────────────────────────
const [screen, setScreen]   = useState(“welcome”); // welcome | login | app
const [authMode, setAuthMode] = useState(“login”);
const [authName, setAuthName] = useState(””);
const [authPin,  setAuthPin]  = useState(””);
const [authPin2, setAuthPin2] = useState(””);
const [authErr,  setAuthErr]  = useState(””);
const [authBusy, setAuthBusy] = useState(false);
const [user, setUser] = useState(null); // { name, userId }

// ── App state ──────────────────────────────────────────────────────────────
const [tab,      setTab]      = useState(“dash”);
const [isPro,    setIsPro]    = useState(false);
const [showPro,  setShowPro]  = useState(false);
const [proCode,  setProCode]  = useState(””);
const [proErr,   setProErr]   = useState(””);
const [income,   setIncome]   = useState(1200000);
const [expenses, setExpenses] = useState(DEMO_EXPENSES);
const [form,     setForm]     = useState({ desc: “”, amount: “”, cat: “food”, date: TODAY });
const [histFilter, setHistFilter] = useState(THIS_MONTH);
const [bankResult, setBankResult] = useState(null);
const [bankLoading, setBankLoading] = useState(false);
const [bankError,   setBankError]   = useState(””);
const [aiAdvice, setAiAdvice] = useState([]);
const [aiShown,  setAiShown]  = useState(false);
const chatEndRef = useRef(null);

// ── Computed ───────────────────────────────────────────────────────────────
const metrics = useMemo(() => calcMetrics(income, expenses), [income, expenses]);

const catData = useMemo(() =>
CATEGORIES
.filter(c => !c.isIncome)
.map(c => ({
name: c.label[lang] || c.label.mn,
value: expenses.filter(e => e.cat === c.id && !e.isIncome).reduce((s, e) => s + e.amount, 0),
color: c.color,
}))
.filter(d => d.value > 0),
[expenses, lang]);

const monthlyData = useMemo(() => {
const byMonth = {};
expenses.forEach(e => {
const m = (e.date || “”).slice(0, 7);
if (!m) return;
if (!byMonth[m]) byMonth[m] = { month: m, income: 0, expense: 0, savings: 0 };
const cat = CATEGORIES.find(c => c.id === e.cat);
if (e.isIncome || cat?.isIncome) byMonth[m].income += e.amount;
else if (cat?.saving) byMonth[m].savings += e.amount;
else byMonth[m].expense += e.amount;
});
return Object.values(byMonth).sort((a, b) => a.month.localeCompare(b.month)).slice(-12);
}, [expenses]);

const filteredExpenses = useMemo(() => {
if (!histFilter) return expenses;
if (histFilter === “3m”) {
const cutoff = new Date(); cutoff.setMonth(cutoff.getMonth() - 3);
return expenses.filter(e => new Date(e.date) >= cutoff);
}
if (histFilter === “year”) {
return expenses.filter(e => (e.date || “”).startsWith(TODAY.slice(0, 4)));
}
return expenses.filter(e => (e.date || “”).startsWith(histFilter));
}, [expenses, histFilter]);

// ── Persist ────────────────────────────────────────────────────────────────
const persist = useCallback((exp, inc) => {
if (!user) return;
saveUserData(user.userId, { expenses: exp, income: inc });
}, [user]);

const updateExpenses = useCallback((newExp) => {
setExpenses(newExp);
persist(newExp, income);
}, [income, persist]);

const updateIncome = useCallback((val) => {
setIncome(val);
persist(expenses, val);
}, [expenses, persist]);

// ── Auth handlers ──────────────────────────────────────────────────────────
const handleAuth = useCallback(async () => {
setAuthErr(””);
const name = authName.trim();
const pin  = authPin.trim();
if (name.length < 2) { setAuthErr(T.nameTooShort); return; }
if (!/^\d{4}$/.test(pin)) { setAuthErr(T.pinInvalid); return; }

```
setAuthBusy(true);
const userId  = name.toLowerCase().replace(/\s+/g, "_");
const idx     = getUsersIndex();

if (authMode === "register") {
  if (authPin2 !== pin) { setAuthErr(T.pinMismatch); setAuthBusy(false); return; }
  const pinHash = await hashPin(userId, pin);
  idx[userId] = { name, pinHash, createdAt: Date.now() };
  saveUsersIndex(idx);
  setUser({ name, userId });
  setExpenses([]);
  setIncome(1200000);
  saveUserData(userId, { expenses: [], income: 1200000 });
  setIsPro(isProActive(userId));
  setScreen("app");
} else {
  if (!idx[userId]) { setAuthErr(T.userNotFound); setAuthBusy(false); return; }
  const pinHash = await hashPin(userId, pin);
  if (idx[userId].pinHash !== pinHash) { setAuthErr(T.pinWrong); setAuthBusy(false); return; }
  const data = loadUserData(userId);
  setUser({ name: idx[userId].name, userId });
  setExpenses(data?.expenses || []);
  setIncome(data?.income || 1200000);
  setIsPro(isProActive(userId));
  setScreen("app");
}
setAuthBusy(false);
```

}, [authName, authPin, authPin2, authMode, T]);

const handleLogout = () => {
setUser(null); setScreen(“welcome”);
setAuthName(””); setAuthPin(””); setAuthPin2(””); setAuthErr(””);
setExpenses(DEMO_EXPENSES); setBankResult(null); setAiAdvice([]); setAiShown(false);
};

// ── PRO ────────────────────────────────────────────────────────────────────
const activatePro = () => {
if (!user) return;
const result = checkProCode(proCode, user.userId);
if (result.ok) { setIsPro(true); setProErr(””); setShowPro(false); setProCode(””); }
else setProErr(result.taken ? (lang === “mn” ? “Өөр хэрэглэгч ашиглаж байна” : “Басқа пайдаланушы қолданып жатыр”) : T.proCodeError);
};

// ── Bank file ──────────────────────────────────────────────────────────────
const handleBankFile = useCallback((e) => {
const file = e.target.files?.[0];
if (!file) return;
setBankLoading(true); setBankError(””); setBankResult(null);
const reader = new FileReader();
reader.onload = (ev) => {
try {
const wb = XLSX.read(ev.target.result, { type: “array”, cellDates: true });
let best = null;
for (const sheetName of wb.SheetNames) {
const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null, raw: false });
const result = parseBankXlsx(raw);
if (result && result.transactions.length > 0 && (!best || result.transactions.length > best.transactions.length)) {
best = { …result, sheetName };
}
}
if (best) {
setBankResult(best);
const newExp = best.transactions.map(tx => ({
id:       Date.now() + tx.id + Math.random(),
desc:     tx.desc || (tx.credit > 0 ? “Орлого” : “Зарлага”),
amount:   tx.credit > 0 ? tx.credit : tx.debit,
cat:      tx.credit > 0 ? “income” : autoCatFromDesc(tx.desc),
date:     tx.date || TODAY,
isIncome: tx.credit > 0,
fromBank: true,
}));
const merged = […expenses.filter(e => !e.fromBank), …newExp];
updateExpenses(merged);
if (best.totalCredit > 0) updateIncome(best.totalCredit);
} else {
setBankError(T.bankError);
}
} catch (err) {
setBankError(T.bankError + “ (” + err.message + “)”);
}
setBankLoading(false);
};
reader.readAsArrayBuffer(file);
e.target.value = “”;
}, [expenses, T, updateExpenses, updateIncome]);

// ── AI Math Advice ─────────────────────────────────────────────────────────
const runAdvice = useCallback(() => {
const advice = generateAdvice(metrics, income, lang);
setAiAdvice(advice);
setAiShown(true);
}, [metrics, income, lang]);

// ── Add expense ────────────────────────────────────────────────────────────
const addExpense = () => {
if (!form.desc || !form.amount) return;
const cat = CATEGORIES.find(c => c.id === form.cat);
const newExp = {
id:       Date.now(),
desc:     form.desc,
amount:   parseFloat(form.amount) || 0,
cat:      form.cat,
date:     form.date,
isIncome: cat?.isIncome || false,
};
updateExpenses([…expenses, newExp]);
setForm({ desc: “”, amount: “”, cat: “food”, date: TODAY });
};

const deleteExpense = (id) => updateExpenses(expenses.filter(e => e.id !== id));

// ─────────────────────────────────────────────────────────────────────────────
// RENDER HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const Pill = ({ label, active, onClick }) => (
<button onClick={onClick} style={{
padding: “6px 12px”, borderRadius: 20, border: “none”, cursor: “pointer”,
fontSize: 11, fontWeight: 700, whiteSpace: “nowrap”,
background: active ? S.accent : S.surface,
color: active ? “#fff” : S.textMuted,
transition: “all .15s”,
}}>{label}</button>
);

const MetricRow = ({ icon, label, value, accent }) => (
<div style={{ display: “flex”, justifyContent: “space-between”, alignItems: “center”, padding: “9px 0”, borderBottom: `1px solid ${S.border}` }}>
<span style={{ fontSize: 12, color: S.textMuted }}>{icon} {label}</span>
<span style={{ fontSize: 13, fontWeight: 700, color: accent || S.text }}>{value}</span>
</div>
);

const ADVICE_COLORS = { success: S.success, warning: S.warning, danger: S.danger, info: S.accent };

// ─────────────────────────────────────────────────────────────────────────────
// WELCOME SCREEN
// ─────────────────────────────────────────────────────────────────────────────
if (screen === “welcome”) return (
<div style={{ minHeight: “100vh”, background: S.bg, display: “flex”, flexDirection: “column”, alignItems: “center”, justifyContent: “center”, padding: 24, fontFamily: “‘Segoe UI’, system-ui, sans-serif”, color: S.text }}>
{/* Lang */}
<div style={{ display: “flex”, gap: 8, marginBottom: 32 }}>
{[“mn”, “kz”].map(l => (
<button key={l} onClick={() => setLang(l)} style={{ padding: “7px 16px”, borderRadius: 20, border: `1.5px solid ${lang === l ? S.accent : S.border}`, background: lang === l ? S.accentLight : S.surface, color: lang === l ? “#a5b4fc” : S.textMuted, fontSize: 13, fontWeight: 700, cursor: “pointer” }}>
{l === “mn” ? “🇲🇳 MN” : “🇰🇿 KZ”}
</button>
))}
</div>

```
  {/* Logo */}
  <div style={{ fontSize: 56, marginBottom: 12 }}>💰</div>
  <div style={{ fontSize: 22, fontWeight: 900, letterSpacing: 1, marginBottom: 4 }}>{T.appName}</div>
  <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 36, letterSpacing: 2 }}>PERSONAL FINANCE · v{APP_VERSION}</div>

  {/* Clear & start */}
  <div style={{ ...card(), width: "100%", maxWidth: 340, marginBottom: 12 }}>
    <div style={{ fontSize: 12, color: S.textMuted, marginBottom: 10 }}>{T.clearAndStart}</div>
    <button onClick={() => { setExpenses([]); setScreen("login"); }} style={{ width: "100%", padding: 12, borderRadius: S.radiusSm, border: "none", background: "#991b1b", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
      🗑 {T.clearAndStart}
    </button>
  </div>

  {/* Demo */}
  <div style={{ ...card(), width: "100%", maxWidth: 340, marginBottom: 12 }}>
    <button onClick={() => { setExpenses(DEMO_EXPENSES); setScreen("app"); }} style={{ width: "100%", padding: 12, borderRadius: S.radiusSm, border: "none", background: S.accent, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
      ▶ {T.demoStart}
    </button>
  </div>

  {/* Login */}
  <div style={{ ...card(), width: "100%", maxWidth: 340, marginBottom: 12 }}>
    <button onClick={() => setScreen("login")} style={{ width: "100%", padding: 12, borderRadius: S.radiusSm, border: `1px solid ${S.borderStrong}`, background: S.surface, color: S.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
      🔐 {T.login} / {T.register}
    </button>
  </div>

  {/* Install */}
  <div style={{ ...card(), width: "100%", maxWidth: 340 }}>
    <div style={{ fontSize: 12, fontWeight: 700, color: S.textMuted, marginBottom: 8 }}>📲 {T.installApp}</div>
    <div style={{ fontSize: 11, color: S.textMuted, lineHeight: 1.9 }}>
      📱 {T.installIos}<br />
      🤖 {T.installAndroid}
    </div>
  </div>

  <div style={{ fontSize: 9, color: S.textMuted, marginTop: 20, opacity: 0.4 }}>🇲🇳 Mongolia · AES-256 · 2026</div>
</div>
```

);

// ─────────────────────────────────────────────────────────────────────────────
// LOGIN SCREEN
// ─────────────────────────────────────────────────────────────────────────────
if (screen === “login”) {
const idx = getUsersIndex();
const registeredUsers = Object.values(idx);

```
return (
  <div style={{ minHeight: "100vh", background: S.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "'Segoe UI', system-ui, sans-serif", color: S.text }}>
    <div style={{ display: "flex", gap: 8, marginBottom: 28 }}>
      {["mn", "kz"].map(l => (
        <button key={l} onClick={() => setLang(l)} style={{ padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${lang === l ? S.accent : S.border}`, background: lang === l ? S.accentLight : S.surface, color: lang === l ? "#a5b4fc" : S.textMuted, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
          {l === "mn" ? "🇲🇳" : "🇰🇿"}
        </button>
      ))}
    </div>

    <div style={{ fontSize: 40, marginBottom: 8 }}>🔐</div>
    <div style={{ fontSize: 18, fontWeight: 900, marginBottom: 4 }}>{authMode === "login" ? T.login : T.register}</div>
    <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 24 }}>{T.secureNote}</div>

    <div style={{ ...card(), width: "100%", maxWidth: 340 }}>
      {/* Name */}
      <div style={{ marginBottom: 12 }}>
        <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.yourName}</div>
        <input value={authName} onChange={e => setAuthName(e.target.value)} placeholder={T.yourName} style={inputStyle} onKeyDown={e => e.key === "Enter" && handleAuth()} />
      </div>

      {/* PIN */}
      <div style={{ marginBottom: 12 }}>
        <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.pin4}</div>
        <input type="password" inputMode="numeric" maxLength={4} value={authPin} onChange={e => setAuthPin(e.target.value.replace(/\D/g, "").slice(0, 4))} placeholder="••••" style={{ ...inputStyle, letterSpacing: 8, textAlign: "center", fontSize: 18 }} onKeyDown={e => e.key === "Enter" && handleAuth()} />
      </div>

      {/* PIN 2 (register) */}
      {authMode === "register" && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.pinAgain}</div>
          <input type="password" inputMode="numeric" maxLength={4} value={authPin2} onChange={e => setAuthPin2(e.target.value.replace(/\D/g, "").slice(0, 4))} placeholder="••••" style={{ ...inputStyle, letterSpacing: 8, textAlign: "center", fontSize: 18 }} onKeyDown={e => e.key === "Enter" && handleAuth()} />
        </div>
      )}

      {authErr && (
        <div style={{ background: "rgba(239,68,68,0.12)", border: "1px solid rgba(239,68,68,0.3)", borderRadius: S.radiusSm, padding: "9px 12px", fontSize: 12, color: "#fca5a5", marginBottom: 12 }}>
          {authErr}
        </div>
      )}

      <button onClick={handleAuth} disabled={authBusy} style={{ width: "100%", padding: 12, borderRadius: S.radiusSm, border: "none", background: authBusy ? S.surface : S.accent, color: "#fff", fontSize: 13, fontWeight: 700, cursor: authBusy ? "default" : "pointer", marginBottom: 10 }}>
        {authBusy ? "..." : authMode === "login" ? `🔓 ${T.login}` : `✅ ${T.register}`}
      </button>

      <button onClick={() => { setAuthMode(authMode === "login" ? "register" : "login"); setAuthErr(""); }} style={{ width: "100%", padding: 9, borderRadius: S.radiusSm, border: `1px solid ${S.border}`, background: "transparent", color: S.textMuted, fontSize: 12, cursor: "pointer" }}>
        {authMode === "login" ? T.noAccount : T.hasAccount}
      </button>

      {/* Quick login */}
      {registeredUsers.length > 0 && authMode === "login" && (
        <div style={{ marginTop: 16, borderTop: `1px solid ${S.border}`, paddingTop: 12 }}>
          <div style={{ fontSize: 10, color: S.textMuted, marginBottom: 8 }}>{lang === "mn" ? "Бүртгэлтэй хэрэглэгчид" : "Тіркелген пайдаланушылар"}</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {registeredUsers.map((u, i) => (
              <button key={i} onClick={() => setAuthName(u.name)} style={{ padding: "5px 10px", borderRadius: 16, border: `1px solid ${S.border}`, background: S.surface, color: S.textMuted, fontSize: 11, cursor: "pointer" }}>
                👤 {u.name}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>

    <button onClick={() => setScreen("welcome")} style={{ marginTop: 16, padding: "8px 16px", borderRadius: 20, border: `1px solid ${S.border}`, background: "transparent", color: S.textMuted, fontSize: 12, cursor: "pointer" }}>
      ← {lang === "mn" ? "Буцах" : "Артқа"}
    </button>
  </div>
);
```

}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────
const TABS = [
{ id: “dash”,    label: T.tabs.dash },
{ id: “bank”,    label: T.tabs.bank },
{ id: “add”,     label: T.tabs.add  },
{ id: “history”, label: T.tabs.history },
{ id: “math”,    label: T.tabs.math },
{ id: “tips”,    label: T.tabs.tips },
];

return (
<div style={{ minHeight: “100vh”, background: S.bg, fontFamily: “‘Segoe UI’, system-ui, sans-serif”, color: S.text, paddingBottom: 80 }}>

```
  {/* PRO Modal */}
  {showPro && (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ ...card(), width: "100%", maxWidth: 300, padding: 24 }}>
        <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 4 }}>🔑 {T.proCode}</div>
        <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 14 }}>PRO-XXXX</div>
        <input value={proCode} onChange={e => setProCode(e.target.value.toUpperCase())} maxLength={8} placeholder="PRO-XXXX" style={{ ...inputStyle, textAlign: "center", fontSize: 18, fontWeight: 900, letterSpacing: 6, marginBottom: 8 }} />
        {proErr && <div style={{ color: "#fca5a5", fontSize: 11, marginBottom: 8, textAlign: "center" }}>{proErr}</div>}
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={activatePro} style={{ flex: 1, padding: 11, borderRadius: S.radiusSm, border: "none", background: "#d97706", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>{T.proActivate}</button>
          <button onClick={() => { setShowPro(false); setProErr(""); }} style={{ flex: 1, padding: 11, borderRadius: S.radiusSm, border: `1px solid ${S.border}`, background: S.surface, color: S.textMuted, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{T.proClose}</button>
        </div>
      </div>
    </div>
  )}

  {/* Header */}
  <div style={{ background: "rgba(0,0,0,0.6)", backdropFilter: "blur(20px)", borderBottom: `1px solid ${S.border}`, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100 }}>
    <div>
      <div style={{ fontSize: 15, fontWeight: 900 }}>
        💰 {T.appName}
        {isPro && <span style={{ marginLeft: 6, fontSize: 9, background: "#d97706", padding: "2px 6px", borderRadius: 6, fontWeight: 800, verticalAlign: "middle" }}>PRO</span>}
      </div>
      {user && <div style={{ fontSize: 10, color: S.textMuted }}>👤 {user.name} · 🔐</div>}
    </div>
    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
      {["mn", "kz"].map(l => (
        <button key={l} onClick={() => setLang(l)} style={{ padding: "4px 8px", borderRadius: 8, border: `1px solid ${lang === l ? S.accent : S.border}`, background: lang === l ? S.accentLight : "transparent", color: lang === l ? "#a5b4fc" : S.textMuted, fontSize: 10, cursor: "pointer" }}>
          {l === "mn" ? "🇲🇳" : "🇰🇿"}
        </button>
      ))}
      {!isPro && (
        <button onClick={() => setShowPro(true)} style={{ padding: "5px 9px", borderRadius: 9, border: "1px solid rgba(217,119,6,0.4)", background: "rgba(217,119,6,0.1)", color: "#fbbf24", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>🔑</button>
      )}
      <button onClick={user ? handleLogout : () => setScreen("welcome")} style={{ padding: "5px 9px", borderRadius: 9, border: `1px solid ${S.border}`, background: S.surface, color: S.textMuted, fontSize: 10, cursor: "pointer" }}>
        {user ? T.logout : "←"}
      </button>
    </div>
  </div>

  {/* Tabs */}
  <div style={{ display: "flex", gap: 4, padding: "8px 12px", overflowX: "auto", background: "rgba(0,0,0,0.3)", borderBottom: `1px solid ${S.border}` }}>
    {TABS.map(t => <Pill key={t.id} label={t.label} active={tab === t.id} onClick={() => setTab(t.id)} />)}
  </div>

  <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 12 }}>

    {/* ════════ DASHBOARD ════════════════════════════════════════════════ */}
    {tab === "dash" && (
      <>
        {/* Income input */}
        <div style={card()}>
          <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 6 }}>{T.income} (₮)</div>
          <input type="number" value={income} onChange={e => updateIncome(parseInt(e.target.value) || 0)} style={inputStyle} />
        </div>

        {/* 3 stats */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
          {[
            { label: T.income,  value: fmtSh(income),             color: S.success, icon: "💵" },
            { label: T.expense, value: fmtSh(metrics.totalSpent), color: S.danger,  icon: "📤" },
            { label: T.balance, value: fmtSh(metrics.balance),    color: metrics.balance >= 0 ? S.success : S.danger, icon: metrics.balance >= 0 ? "✅" : "⚠️" },
          ].map(x => (
            <div key={x.label} style={{ ...card({ padding: "12px 8px", textAlign: "center" }) }}>
              <div style={{ fontSize: 20 }}>{x.icon}</div>
              <div style={{ fontSize: 9, color: S.textMuted, marginTop: 3 }}>{x.label}</div>
              <div style={{ fontSize: 12, fontWeight: 800, color: x.color, marginTop: 3 }}>{x.value}</div>
            </div>
          ))}
        </div>

        {/* Health score */}
        <div style={card()}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 700 }}>{T.healthScore}</div>
            <div style={{ fontSize: 26, fontWeight: 900, color: HEALTH_COLOR(metrics.healthScore) }}>
              {metrics.healthScore}<span style={{ fontSize: 11, opacity: 0.5 }}>/100</span>
            </div>
          </div>
          <div style={{ background: S.surface, borderRadius: 8, height: 10, overflow: "hidden", marginBottom: 6 }}>
            <div style={{ height: "100%", width: `${metrics.healthScore}%`, background: HEALTH_COLOR(metrics.healthScore), borderRadius: 8, transition: "width 1s ease" }} />
          </div>
          <div style={{ fontSize: 12, color: HEALTH_COLOR(metrics.healthScore), fontWeight: 600, marginBottom: 14 }}>
            {HEALTH_LABEL(metrics.healthScore, lang)}
          </div>
          {/* 50/30/20 */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
            {[
              { label: T.needs,    score: metrics.score50, ideal: fmt(Math.round(income * 0.5)) },
              { label: T.wants,    score: metrics.score30, ideal: fmt(Math.round(income * 0.3)) },
              { label: T.savingsW, score: metrics.score20, ideal: fmt(Math.round(income * 0.2)) },
            ].map(r => (
              <div key={r.label} style={{ background: "rgba(255,255,255,0.03)", borderRadius: S.radiusSm, padding: 10, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: S.textMuted }}>{r.label}</div>
                <div style={{ fontSize: 20, fontWeight: 900, color: HEALTH_COLOR(r.score) }}>{r.score}<span style={{ fontSize: 9 }}>%</span></div>
                <div style={{ fontSize: 9, color: S.textMuted, marginTop: 2 }}>{r.ideal}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Pie chart */}
        {catData.length > 0 && (
          <div style={card()}>
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>📊 {lang === "mn" ? "Ангилал" : "Категориялар"}</div>
            <ResponsiveContainer width="100%" height={160}>
              <PieChart>
                <Pie data={catData} cx="50%" cy="50%" innerRadius={40} outerRadius={65} dataKey="value" paddingAngle={2}>
                  {catData.map((e, i) => <Cell key={i} fill={e.color} />)}
                </Pie>
                <Tooltip formatter={v => fmt(v)} contentStyle={{ background: "#1a1a2e", border: `1px solid ${S.border}`, borderRadius: 10, color: "#fff", fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 6 }}>
              {catData.map((c, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, background: "rgba(255,255,255,0.05)", padding: "3px 8px", borderRadius: 8 }}>
                  <div style={{ width: 6, height: 6, borderRadius: "50%", background: c.color, flexShrink: 0 }} />
                  <span style={{ color: S.textMuted }}>{c.name}</span>
                  <b>{fmtSh(c.value)}</b>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Area chart */}
        {monthlyData.length > 1 && (
          <div style={card()}>
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>📅 {T.chartMonthly}</div>
            <ResponsiveContainer width="100%" height={150}>
              <AreaChart data={monthlyData} margin={{ top: 5, right: 5, bottom: 0, left: 0 }}>
                <defs>
                  <linearGradient id="gi" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={S.success} stopOpacity={0.3} /><stop offset="95%" stopColor={S.success} stopOpacity={0} /></linearGradient>
                  <linearGradient id="ge" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={S.danger} stopOpacity={0.3} /><stop offset="95%" stopColor={S.danger} stopOpacity={0} /></linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                <XAxis dataKey="month" tick={{ fill: S.textMuted, fontSize: 9 }} axisLine={false} tickLine={false} />
                <YAxis hide />
                <Tooltip formatter={v => fmt(v)} contentStyle={{ background: "#1a1a2e", border: `1px solid ${S.border}`, borderRadius: 10, color: "#fff", fontSize: 10 }} />
                <Area type="monotone" dataKey="income"  stroke={S.success} strokeWidth={2} fill="url(#gi)" name={T.income} />
                <Area type="monotone" dataKey="expense" stroke={S.danger}  strokeWidth={2} fill="url(#ge)" name={T.expense} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* AI Math Advice button */}
        <button onClick={runAdvice} style={{ padding: 13, borderRadius: S.radius, border: "none", background: S.accent, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
          {T.aiTitle} — {T.aiAnalyze}
        </button>

        {/* AI Math Advice results */}
        {aiShown && aiAdvice.map((a, i) => (
          <div key={i} style={{ ...card({ background: `${ADVICE_COLORS[a.type]}12`, border: `1px solid ${ADVICE_COLORS[a.type]}30`, padding: 14 }) }}>
            <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
              <span style={{ fontSize: 20, flexShrink: 0 }}>{a.icon}</span>
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: ADVICE_COLORS[a.type], marginBottom: 4 }}>{a.title}</div>
                <div style={{ fontSize: 12, lineHeight: 1.7, color: S.text }}>{a.text}</div>
              </div>
            </div>
          </div>
        ))}

        {/* Recent */}
        <div style={card()}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>🧾 {T.recent}</div>
          {expenses.length === 0 && <div style={{ color: S.textMuted, fontSize: 12, textAlign: "center", padding: 12 }}>{T.noData}</div>}
          {[...expenses].reverse().slice(0, 6).map(e => {
            const cat = CATEGORIES.find(c => c.id === e.cat);
            return (
              <div key={e.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${S.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 9, minWidth: 0 }}>
                  <div style={{ width: 30, height: 30, borderRadius: 9, background: `${cat?.color || "#888"}22`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, flexShrink: 0 }}>
                    {(cat?.label[lang] || cat?.label.mn || "📦").slice(0, 2)}
                  </div>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{e.desc} {e.fromBank && "🏦"}</div>
                    <div style={{ fontSize: 10, color: S.textMuted }}>{e.date}</div>
                  </div>
                </div>
                <div style={{ fontSize: 12, fontWeight: 700, color: e.isIncome ? S.success : S.danger, flexShrink: 0, marginLeft: 8 }}>
                  {e.isIncome ? "+" : "−"}{fmtSh(e.amount)}
                </div>
              </div>
            );
          })}
        </div>
      </>
    )}

    {/* ════════ BANK ═════════════════════════════════════════════════════ */}
    {tab === "bank" && (
      <>
        <div style={{ fontSize: 14, fontWeight: 800 }}>{T.tabs.bank}</div>

        {/* Upload area */}
        <div style={{ ...card({ background: "rgba(99,102,241,0.05)", border: `2px dashed rgba(99,102,241,0.3)` }), textAlign: "center", padding: 28 }}>
          <div style={{ fontSize: 36, marginBottom: 10 }}>📂</div>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>{T.uploadXlsx}</div>
          <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 16 }}>{T.supportedBanks}</div>
          <label style={{ padding: "11px 22px", borderRadius: 12, border: "none", background: S.accent, color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer", display: "inline-block" }}>
            {lang === "mn" ? "📂 Файл сонгох" : "📂 Файл таңдау"}
            <input type="file" accept=".xlsx,.xls" onChange={handleBankFile} style={{ display: "none" }} />
          </label>
        </div>

        {bankLoading && <div style={{ ...card({ textAlign: "center" }), padding: 20 }}><div style={{ fontSize: 22 }}>⏳</div><div style={{ fontSize: 12, color: S.textMuted, marginTop: 6 }}>{T.processing}</div></div>}

        {bankError && <div style={{ ...card({ background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.25)" }) }}><div style={{ fontSize: 12, color: "#fca5a5" }}>⚠️ {bankError}</div></div>}

        {bankResult && (
          <>
            <div style={card()}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>{T.bankSummary}</div>
              {bankResult.owner && <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 4 }}>👤 {bankResult.owner}</div>}
              {bankResult.period && <div style={{ fontSize: 10, color: S.textMuted, marginBottom: 12 }}>📅 {bankResult.period}</div>}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {[
                  { label: lang === "mn" ? "Нийт орлого"    : "Жалпы кіріс",   value: fmt(bankResult.totalCredit), color: S.success },
                  { label: lang === "mn" ? "Нийт зарлага"   : "Жалпы шығыс",   value: fmt(bankResult.totalDebit),  color: S.danger  },
                  { label: lang === "mn" ? "Цэвэр урсгал"   : "Таза ағыс",      value: fmt(bankResult.net),         color: bankResult.net >= 0 ? S.success : S.danger },
                  { label: T.txCount,                                             value: bankResult.transactions.length + (lang === "mn" ? " ш" : " дана"), color: S.accent  },
                ].map(x => (
                  <div key={x.label} style={{ background: "rgba(255,255,255,0.03)", borderRadius: S.radiusSm, padding: 12, textAlign: "center" }}>
                    <div style={{ fontSize: 9, color: S.textMuted }}>{x.label}</div>
                    <div style={{ fontSize: 14, fontWeight: 800, color: x.color, marginTop: 4 }}>{x.value}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Transaction list */}
            <div style={card()}>
              <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 10 }}>{lang === "mn" ? "Гүйлгээнүүд" : "Мәмілелер"} ({bankResult.transactions.length})</div>
              <div style={{ maxHeight: 320, overflowY: "auto" }}>
                {bankResult.transactions.map((tx, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: `1px solid ${S.border}` }}>
                    <div style={{ flex: 1, minWidth: 0, marginRight: 8 }}>
                      <div style={{ fontSize: 11, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{tx.desc}</div>
                      <div style={{ fontSize: 9, color: S.textMuted }}>{tx.date}</div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      {tx.credit > 0 && <div style={{ fontSize: 11, fontWeight: 700, color: S.success }}>+{fmt(tx.credit)}</div>}
                      {tx.debit  > 0 && <div style={{ fontSize: 11, fontWeight: 700, color: S.danger }}>−{fmt(tx.debit)}</div>}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </>
        )}
      </>
    )}

    {/* ════════ ADD ══════════════════════════════════════════════════════ */}
    {tab === "add" && (
      <>
        <div style={{ fontSize: 14, fontWeight: 800 }}>{T.tabs.add}</div>
        <div style={card()}>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.descLabel}</div>
            <input value={form.desc} onChange={e => setForm(p => ({ ...p, desc: e.target.value }))} placeholder="..." style={inputStyle} />
          </div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.amountLabel}</div>
            <input type="number" value={form.amount} onChange={e => setForm(p => ({ ...p, amount: e.target.value }))} placeholder="50000" style={inputStyle} />
          </div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 5 }}>{T.dateLabel}</div>
            <input type="date" value={form.date} onChange={e => setForm(p => ({ ...p, date: e.target.value }))} style={inputStyle} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 11, color: S.textMuted, marginBottom: 7 }}>{T.catLabel}</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 5 }}>
              {CATEGORIES.map(c => (
                <button key={c.id} onClick={() => setForm(p => ({ ...p, cat: c.id }))} style={{ padding: "8px 10px", borderRadius: S.radiusSm, border: `1.5px solid ${form.cat === c.id ? c.color : "transparent"}`, background: form.cat === c.id ? `${c.color}18` : S.surface, color: S.text, fontSize: 11, cursor: "pointer", textAlign: "left", transition: "all .12s" }}>
                  {c.label[lang] || c.label.mn}
                </button>
              ))}
            </div>
          </div>
          <button onClick={addExpense} style={{ width: "100%", padding: 12, borderRadius: S.radiusSm, border: "none", background: S.success, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>✅ {T.saveBtn}</button>
        </div>
      </>
    )}

    {/* ════════ HISTORY ══════════════════════════════════════════════════ */}
    {tab === "history" && (
      <>
        <div style={{ fontSize: 14, fontWeight: 800 }}>{T.tabs.history}</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {[
            { label: T.filterThisMonth, value: THIS_MONTH },
            { label: T.filterLast3,     value: "3m"       },
            { label: T.filterYear,      value: TODAY.slice(0, 4) },
            { label: T.filterAll,       value: ""         },
          ].map(f => <Pill key={f.value} label={f.label} active={histFilter === f.value} onClick={() => setHistFilter(f.value)} />)}
        </div>

        {/* Bar chart */}
        {monthlyData.length > 0 && (
          <div style={card()}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 10 }}>📊 {lang === "mn" ? "Сарын зарлага" : "Ай шығысы"}</div>
            <ResponsiveContainer width="100%" height={140}>
              <BarChart data={monthlyData} margin={{ top: 4, right: 4, bottom: 0, left: 0 }}>
                <XAxis dataKey="month" tick={{ fill: S.textMuted, fontSize: 9 }} axisLine={false} tickLine={false} />
                <YAxis hide />
                <Tooltip formatter={v => fmt(v)} contentStyle={{ background: "#1a1a2e", border: `1px solid ${S.border}`, borderRadius: 10, color: "#fff", fontSize: 10 }} />
                <Bar dataKey="expense" radius={[4, 4, 0, 0]} name={T.expense}>
                  {monthlyData.map((_, i) => <Cell key={i} fill={`hsl(${350 + i * 8},70%,55%)`} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* List */}
        <div style={card()}>
          <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
            {lang === "mn" ? "Бүх бичлэгүүд" : "Барлық жазбалар"} ({filteredExpenses.length})
            {filteredExpenses.length > 0 && (
              <span style={{ marginLeft: 8, fontSize: 10, color: S.danger }}>
                −{fmt(filteredExpenses.filter(e => !e.isIncome).reduce((s, e) => s + e.amount, 0))}
              </span>
            )}
          </div>
          <div style={{ maxHeight: 420, overflowY: "auto" }}>
            {filteredExpenses.length === 0 && <div style={{ color: S.textMuted, fontSize: 12, textAlign: "center", padding: 12 }}>{T.noData}</div>}
            {[...filteredExpenses].reverse().map(e => {
              const cat = CATEGORIES.find(c => c.id === e.cat);
              return (
                <div key={e.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${S.border}` }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{e.desc} {e.fromBank && "🏦"}</div>
                    <div style={{ fontSize: 10, color: S.textMuted }}>{cat?.label[lang] || cat?.label.mn} · {e.date}</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0, marginLeft: 8 }}>
                    <span style={{ fontSize: 12, fontWeight: 700, color: e.isIncome ? S.success : S.danger }}>
                      {e.isIncome ? "+" : "−"}{fmtSh(e.amount)}
                    </span>
                    <button onClick={() => deleteExpense(e.id)} style={{ background: "rgba(239,68,68,0.15)", border: "none", color: S.danger, borderRadius: 7, padding: "3px 7px", cursor: "pointer", fontSize: 11 }}>{T.deleteBtn}</button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </>
    )}

    {/* ════════ MATH ═════════════════════════════════════════════════════ */}
    {tab === "math" && (
      <>
        <div style={{ fontSize: 14, fontWeight: 800 }}>{T.tabs.math}</div>

        {/* Summary */}
        <div style={card()}>
          <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 10 }}>📋 {T.monthlySum}</div>
          {[
            { icon: "📅", label: T.dailyAvg,   value: fmt(metrics.perDay),  accent: S.text },
            { icon: "📆", label: T.weeklyAvg,  value: fmt(metrics.perWeek), accent: S.text },
            { icon: "💎", label: T.savingsRate, value: `${pct(metrics.totalSaved, income)}%`, accent: pct(metrics.totalSaved, income) >= 20 ? S.success : S.warning },
            { icon: "📤", label: T.spendRate,   value: `${Math.min(100, pct(metrics.totalSpent, income))}%`, accent: pct(metrics.totalSpent, income) <= 80 ? S.success : S.danger },
            { icon: "🛡", label: T.emergencyFund, value: fmt(metrics.emgTarget), accent: S.text },
            { icon: "⏳", label: T.monthsToFund, value: metrics.monthsToEmg ? `${metrics.monthsToEmg} ${lang === "mn" ? "сар" : "ай"}` : "—", accent: S.accent },
            { icon: "💰", label: T.annualSavings, value: fmt(metrics.totalSaved * 12), accent: S.success },
            { icon: "📊", label: T.annualSpend,   value: fmt(metrics.totalSpent * 12), accent: S.danger  },
          ].map(r => <MetricRow key={r.label} {...r} />)}
        </div>

        {/* Investment forecast */}
        <div style={card()}>
          <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 4 }}>{T.investForecast}</div>
          <div style={{ fontSize: 10, color: S.textMuted, marginBottom: 12 }}>{T.investNote} · {fmt(metrics.totalSaved)}/ай</div>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={metrics.forecasts.filter(f => [5, 10, 15, 20, 30].includes(f.years)).map(f => ({ name: `${f.years}${lang === "mn" ? "жил" : "жыл"}`, value: f.fv }))} margin={{ top: 4, right: 4, bottom: 0, left: 0 }}>
              <XAxis dataKey="name" tick={{ fill: S.textMuted, fontSize: 9 }} axisLine={false} tickLine={false} />
              <YAxis hide />
              <Tooltip formatter={v => fmt(v)} contentStyle={{ background: "#1a1a2e", border: `1px solid ${S.border}`, borderRadius: 10, color: "#fff", fontSize: 10 }} />
              <Bar dataKey="value" radius={[5, 5, 0, 0]}>
                {metrics.forecasts.filter(f => [5,10,15,20,30].includes(f.years)).map((_, i) => <Cell key={i} fill={`hsl(${155 + i * 20},65%,55%)`} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <div style={{ fontSize: 9, color: S.textMuted, marginTop: 6 }}>📐 {T.formulaNote}</div>
        </div>

        {/* 50/30/20 */}
        <div style={{ ...card({ background: "rgba(99,102,241,0.06)", border: `1px solid rgba(99,102,241,0.2)` }) }}>
          <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 12, color: "#a5b4fc" }}>📐 {T.rule5030}</div>
          {[
            { label: `🏠 ${T.needs} (50%)`,   pct: 50, color: S.success,  actual: metrics.totalNeeds,  ideal: metrics.ideal50 },
            { label: `🎮 ${T.wants} (30%)`,    pct: 30, color: S.warning,  actual: metrics.totalWants,  ideal: metrics.ideal30 },
            { label: `💎 ${T.savingsW} (20%)`, pct: 20, color: "#f59e0b",  actual: metrics.totalSaved,  ideal: metrics.ideal20 },
          ].map(r => (
            <div key={r.label} style={{ marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 5 }}>
                <span>{r.label}</span>
                <span style={{ fontWeight: 700, color: r.color }}>{fmt(Math.round(income * r.pct / 100))}</span>
              </div>
              <div style={{ background: "rgba(255,255,255,0.07)", borderRadius: 7, height: 9 }}>
                <div style={{ width: `${r.pct}%`, height: "100%", borderRadius: 7, background: r.color }} />
              </div>
              <div style={{ fontSize: 10, color: S.textMuted, marginTop: 3 }}>
                {lang === "mn" ? "Одоогийн" : "Қазіргі"}: {fmt(r.actual)} · {lang === "mn" ? "Зорилго" : "Мақсат"}: {fmt(Math.round(r.ideal))}
              </div>
            </div>
          ))}
        </div>
      </>
    )}

    {/* ════════ TIPS ═════════════════════════════════════════════════════ */}
    {tab === "tips" && (
      <>
        <div style={{ fontSize: 14, fontWeight: 800 }}>{T.tabs.tips}</div>
        {Array.from({ length: 10 }, (_, i) => i + 1).map(n => {
          const titleKey = `tip${n}t`;
          const textKey  = `tip${n}`;
          if (!T[titleKey]) return null;
          return (
            <div key={n} style={{ ...card({ padding: 14 }), borderLeft: `3px solid hsl(${140 + n * 18},60%,50%)` }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: `hsl(${140 + n * 18},60%,65%)`, marginBottom: 5 }}>{T[titleKey]}</div>
              <div style={{ fontSize: 12, lineHeight: 1.75, color: S.text }}>{T[textKey]}</div>
            </div>
          );
        })}
      </>
    )}

    {/* PRO unlock */}
    {!isPro && (tab === "dash" || tab === "tips") && (
      <div style={{ ...card({ background: "rgba(217,119,6,0.07)", border: "1px solid rgba(217,119,6,0.25)" }), marginTop: 4 }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: "#fbbf24", marginBottom: 8 }}>🔑 {T.proUnlock}</div>
        <div style={{ fontSize: 11, color: S.textMuted, lineHeight: 1.7, marginBottom: 10 }}>
          🏪 {lang === "mn" ? "Бизнесийн касс" : "Бизнес кассасы"} · 📊 {lang === "mn" ? "Нарийн аналитик" : "Кешенді аналитика"} · 📈 {lang === "mn" ? "Урт хугацааны болжам" : "Ұзақ мерзімді болжам"}
        </div>
        <button onClick={() => setShowPro(true)} style={{ width: "100%", padding: 11, borderRadius: S.radiusSm, border: "none", background: "#d97706", color: "#fff", fontWeight: 800, cursor: "pointer", fontSize: 12 }}>
          🔑 PRO
        </button>
      </div>
    )}

  </div>
</div>
```

);
}
