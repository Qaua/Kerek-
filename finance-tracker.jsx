import { useState, useEffect, useCallback, useRef, useMemo } from “react”;
import {
PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis,
Tooltip, ResponsiveContainer, AreaChart, Area, LineChart, Line, CartesianGrid
} from “recharts”;
import * as XLSX from “xlsx”;

// ════════════════════════════════════════════════════════════════
// 🔑 GEMINI API KEY — кодта тіркелген
// ════════════════════════════════════════════════════════════════
const GEMINI_API_KEY = “YOUR_GEMINI_KEY_HERE”; // ← Осы жерге Gemini кілтін қойыңыз

// ════════════════════════════════════════════════════════════════
// 🔐 AES-256 ШИФРЛАУ — әр пайдаланушы тек өз деректерін көреді
// ════════════════════════════════════════════════════════════════
async function getKey(password) {
const enc = new TextEncoder();
const keyMaterial = await crypto.subtle.importKey(“raw”, enc.encode(password), “PBKDF2”, false, [“deriveKey”]);
return crypto.subtle.deriveKey(
{ name:“PBKDF2”, salt: enc.encode(“aksha_salt_2026”), iterations:100000, hash:“SHA-256” },
keyMaterial, { name:“AES-GCM”, length:256 }, false, [“encrypt”,“decrypt”]
);
}

async function encryptData(data, password) {
try {
const key = await getKey(password);
const iv = crypto.getRandomValues(new Uint8Array(12));
const enc = new TextEncoder();
const encrypted = await crypto.subtle.encrypt({ name:“AES-GCM”, iv }, key, enc.encode(JSON.stringify(data)));
const buf = new Uint8Array(encrypted);
const combined = new Uint8Array(iv.length + buf.length);
combined.set(iv); combined.set(buf, iv.length);
return btoa(String.fromCharCode(…combined));
} catch { return null; }
}

async function decryptData(cipher, password) {
try {
const combined = Uint8Array.from(atob(cipher), c => c.charCodeAt(0));
const iv = combined.slice(0, 12);
const data = combined.slice(12);
const key = await getKey(password);
const dec = await crypto.subtle.decrypt({ name:“AES-GCM”, iv }, key, data);
return JSON.parse(new TextDecoder().decode(dec));
} catch { return null; }
}

// Хэш — PIN тексеру үшін
async function hashPin(pin, userId) {
const enc = new TextEncoder();
const buf = await crypto.subtle.digest(“SHA-256”, enc.encode(userId + “:” + pin + “:aksha2026”));
return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,“0”)).join(””);
}

// ════════════════════════════════════════════════════════════════
// 👤 ПАЙДАЛАНУШЫ ЖҮЙЕСІ
// ════════════════════════════════════════════════════════════════
const USERS_INDEX_KEY = “aksha_users_index”; // тек аттар тізімі (шифрланбаған)

function getUsersIndex() {
try { return JSON.parse(localStorage.getItem(USERS_INDEX_KEY) || “{}”); } catch { return {}; }
}
function saveUsersIndex(idx) {
localStorage.setItem(USERS_INDEX_KEY, JSON.stringify(idx));
}
function getUserDataKey(userId) { return `aksha_data_${userId}`; }

async function saveUserData(userId, pin, data) {
const encrypted = await encryptData(data, userId + “:” + pin);
if (encrypted) localStorage.setItem(getUserDataKey(userId), encrypted);
}

async function loadUserData(userId, pin) {
const cipher = localStorage.getItem(getUserDataKey(userId));
if (!cipher) return null;
return decryptData(cipher, userId + “:” + pin);
}

// ════════════════════════════════════════════════════════════════
// PRO КОД ЖҮЙЕСІ
// ════════════════════════════════════════════════════════════════
const PRO_CODES = [
“PRO-2847”,“PRO-3K9X”,“PRO-7M4Z”,“PRO-9B2W”,“PRO-4T8R”,
“PRO-6N5V”,“PRO-1Q7F”,“PRO-8C3D”,“PRO-5H2J”,“PRO-2X9L”,
“PRO-7G4S”,“PRO-3P8K”,“PRO-9W1M”,“PRO-4Y6B”,“PRO-6R3T”,
“PRO-1Z8N”,“PRO-8V5Q”,“PRO-5F2H”,“PRO-3E7C”,“PRO-9U4W”,
];

function getDeviceId() {
try {
const raw = [navigator.userAgent,navigator.language,screen.width,screen.height,new Date().getTimezoneOffset()].join(”|”);
let h = 0; for (let i=0;i<raw.length;i++) h=(Math.imul(31,h)+raw.charCodeAt(i))|0;
return Math.abs(h).toString(36).toUpperCase();
} catch { return “X”; }
}

function checkProCode(code, userId) {
try {
const KEY = `aksha_pro_${userId}`;
const db = JSON.parse(localStorage.getItem(KEY)||”{}”);
const c = code.trim().toUpperCase();
const dev = getDeviceId();
if (!PRO_CODES.includes(c)) return { ok:false, err:“❌ Код олдсонгүй / Табылмады” };
if (db[c] && db[c]===dev) return { ok:true };
if (db[c] && db[c]!==dev) return { ok:false, err:“🚫 Өөр төхөөрөмжид идэвхжсэн / Басқа құрылғыда” };
db[c]=dev; localStorage.setItem(KEY, JSON.stringify(db));
return { ok:true };
} catch { return { ok:false, err:“⚠️ Алдаа” }; }
}

function isProActive(userId) {
try {
const db = JSON.parse(localStorage.getItem(`aksha_pro_${userId}`)||”{}”);
const dev = getDeviceId();
return PRO_CODES.some(k => db[k]===dev);
} catch { return false; }
}

// ════════════════════════════════════════════════════════════════
// БАНК ФАЙЛЫН ОҚАУ
// ════════════════════════════════════════════════════════════════
const COL_PATTERNS = {
date:    [“огноо”,“гүйлгээний огноо”,“date”,“он сар өдөр”],
credit:  [“кредит гүйлгээ”,“кредит”,“орлого”,“credit”,“нэмэгдэл”],
debit:   [“дебит гүйлгээ”,“дебит”,“зарлага”,“debit”,“хасагдал”],
balance: [“эцсийн үлдэгдэл”,“үлдэгдэл”,“balance”],
desc:    [“гүйлгээний утга”,“утга”,“description”,“тайлбар”],
account: [“харьцсан данс”,“данс”,“account”],
};

function detectCols(headers) {
const r = {};
headers.forEach((h,i) => {
if (!h) return;
const low = String(h).toLowerCase().trim();
for (const [k,pats] of Object.entries(COL_PATTERNS)) {
if (!r[k] && pats.some(p=>low.includes(p))) r[k]=i;
}
});
return r;
}

function parseBankFile(data) {
let headerRow=-1, colMap={};
for (let i=0;i<Math.min(data.length,15);i++) {
const m = detectCols(data[i]);
if (Object.keys(m).length>=3) { headerRow=i; colMap=m; break; }
}
if (headerRow===-1) return {error:true};

let owner=””,iban=””,currency=“MNT”,dateRange=””;
for (let i=0;i<headerRow;i++) {
(data[i]||[]).forEach((cell,j) => {
const v = String(cell||””).toLowerCase();
if (v.includes(“хэрэглэгч”)) { const n=data[i][j+3]||data[i][j+2]||data[i][j+1]; if(n)owner=String(n).trim(); }
if (/mn\d{15,}/i.test(String(cell))) iban=String(cell).trim();
if (v.includes(“iban”)) { const n=data[i][j+3]||data[i][j+2]||data[i][j+1]; if(n&&/MN/i.test(String(n)))iban=String(n).trim(); }
if (v.includes(“интервал”)||v.includes(“period”)) { const n=data[i][j+3]||data[i][j+2]||data[i][j+1]; if(n)dateRange=String(n).trim(); }
if (v.includes(“валют”)) { const n=data[i][j+3]||data[i][j+2]||data[i][j+1]; if(n)currency=String(n).trim(); }
});
}

const txs=[]; let totC=0,totD=0;
for (let i=headerRow+1;i<data.length;i++) {
const row=data[i];
if (!row||row.every(c=>!c)) continue;
const s0=String(row[0]||””).toLowerCase();
if (s0.includes(“нийт”)||s0.includes(“итого”)||s0.includes(“total”)) continue;
const cr=Math.abs(parseFloat(String(row[colMap.credit]||“0”).replace(/[^\d.-]/g,””))||0);
const db=Math.abs(parseFloat(String(row[colMap.debit]||“0”).replace(/[^\d.-]/g,””))||0);
const desc=String(row[colMap.desc]||””).trim();
const acc=String(row[colMap.account]||””).trim();
const bal=parseFloat(String(row[colMap.balance]||“0”).replace(/[^\d.-]/g,””))||0;
let dateStr=””;
const dv=row[colMap.date];
if (dv) dateStr=dv instanceof Date?dv.toISOString().slice(0,10):String(dv).slice(0,10);
if (cr>0||db>0) {
totC+=cr; totD+=db;
txs.push({id:i,date:dateStr,credit:cr,debit:db,desc:desc||(cr>0?“Орлого”:“Зарлага”),account:acc,balance:bal,type:cr>0?“credit”:“debit”});
}
}
return {txs,totC,totD,owner,iban,currency,dateRange,net:totC-totD,count:txs.length};
}

function autoCategory(desc) {
const d=(desc||””).toLowerCase();
if (d.includes(“хураамж”)||d.includes(“fee”)) return “fee”;
if (d.includes(“шилжүүлэг”)||d.includes(“transfer”)) return “transfer”;
if (d.includes(“хоол”)||d.includes(“food”)||d.includes(“ресторан”)||d.includes(“кафе”)) return “food”;
if (d.includes(“такси”)||d.includes(“убер”)||d.includes(“автобус”)||d.includes(“тээвэр”)) return “transport”;
if (d.includes(“эмнэлэг”)||d.includes(“эм “)||d.includes(“эрүүл”)) return “health”;
if (d.includes(“хадгаламж”)||d.includes(“депозит”)) return “savings”;
if (d.includes(“интернет”)||d.includes(“утас”)||d.includes(“мобайл”)) return “phone”;
return “other”;
}

// ════════════════════════════════════════════════════════════════
// САНАТТАР
// ════════════════════════════════════════════════════════════════
const CATS = [
{id:“food”,     nm:{mn:“🍔 Хоол”,    kz:“🍔 Тамақ”,  en:“🍔 Food”},     col:”#FF6B6B”},
{id:“transport”,nm:{mn:“🚗 Тээвэр”,  kz:“🚗 Көлік”,  en:“🚗 Transport”},col:”#4ECDC4”},
{id:“housing”,  nm:{mn:“🏠 Орон сууц”,kz:“🏠 Үй”,    en:“🏠 Housing”},  col:”#45B7D1”},
{id:“health”,   nm:{mn:“💊 Эрүүл мэнд”,kz:“💊 Дене”, en:“💊 Health”},   col:”#96CEB4”},
{id:“entertain”,nm:{mn:“🎮 Зугаа”,   kz:“🎮 Ойын”,   en:“🎮 Fun”},       col:”#FFEAA7”},
{id:“savings”,  nm:{mn:“💎 Хадгаламж”,kz:“💎 Жинақ”, en:“💎 Savings”},  col:”#F7DC6F”,isSav:true},
{id:“fee”,      nm:{mn:“🏦 Хураамж”, kz:“🏦 Комиссия”,en:“🏦 Fee”},     col:”#fb923c”},
{id:“transfer”, nm:{mn:“💸 Шилжүүлэг”,kz:“💸 Аудару”, en:“💸 Transfer”},col:”#a78bfa”},
{id:“phone”,    nm:{mn:“📱 Утас/Net”, kz:“📱 Байланыс”,en:“📱 Phone”},   col:”#f472b6”},
{id:“clothes”,  nm:{mn:“👗 Хувцас”,  kz:“👗 Киім”,   en:“👗 Clothes”},  col:”#DDA0DD”},
{id:“education”,nm:{mn:“📚 Боловсрол”,kz:“📚 Білім”,  en:“📚 Education”},col:”#98D8C8”},
{id:“other”,    nm:{mn:“📦 Бусад”,   kz:“📦 Басқа”,  en:“📦 Other”},    col:”#94a3b8”},
];

const fmt  = n => (n||0).toLocaleString() + “ ₮”;
const fmtM = n => n>=1e6?(n/1e6).toFixed(1)+“сая ₮”:n>=1000?(n/1000).toFixed(0)+“к ₮”:n+” ₮”;
const pct  = (a,b) => b>0?Math.round(a/b*100):0;
const sc   = s => s>=80?”#4ECDC4”:s>=60?”#FFEAA7”:”#FF6B6B”;
const TODAY = new Date().toISOString().slice(0,10);
const THIS_MONTH = TODAY.slice(0,7);

// Аудармалар
const TX = {
mn:{ login:“Нэвтрэх”, register:“Бүртгүүлэх”, name:“Нэр”, pin:“PIN (4 оронтой)”, pinConfirm:“PIN давтана уу”,
nameP:“Нэрээ оруулна уу”, pinP:“1234”, welcome:“Тавтай морил”,
exists:“Хэрэглэгч байна, PIN оруулна уу”, notFound:“Хэрэглэгч олдсонгүй”,
pinWrong:“PIN буруу байна ⚠️”, pinMismatch:“PIN таарахгүй байна”,
nameShort:“Нэр 2-с дээш тэмдэгт байх ёстой”,
pinShort:“PIN яг 4 оронтой байх ёстой”,
newUser:“Шинэ хэрэглэгч? →”, hasAccount:“Бүртгэлтэй бол →”,
myData:“Миний өгөгдөл зөвхөн надад харагдана 🔐”,
logOut:“Гарах”, dashboard:“📊 Нүүр”, bank:“🏦 Банк”,
add:“➕ Нэмэх”, history:“📅 Түүх”, math:“🧮 Математик”,
tips:“🧠 Зөвлөгөө”, chat:“🤖 AI”, biz:“🏪 Бизнес”,
income:“Орлого”, spent:“Зарлага”, left:“Үлдэгдэл”,
health:“💚 Санхүүгийн эрүүл мэнд”,
uploadBtn:“📂 Банкны хуулга оруулах (.xlsx)”,
banks:“Хаан · Голомт · Хас · Ард банк болон бусад”,
analyze:“🔍 AI шинжилгээ”, analyzing:“⏳ Шинжилж байна…”,
desc:“Тайлбар”, amount:“Дүн (₮)”, date:“Огноо”, cat:“Ангилал”,
save:“✅ Хадгалах”, del:“✕”,
monthly:“Сарын хураангуй”, annual:“Жилийн харьцуулалт”,
perDay:“Өдөрт дундаж”, savRate:“Хадгаламжийн хувь”,
spRate:“Зарлагын хувь”, emgFund:“Яаралтай сан”,
emgMonths:“Яаралтай санд хүрэх”, yrSav:“Жилийн хадгаламж”,
invest:“📈 Хөрөнгө оруулалтын таамаглал”, invNote:“8% жилийн хүү (нийлмэл)”,
allRec:“Бүх бичлэгүүд”, bankTx:“Банкны гүйлгээнүүд”,
noData:“Өгөгдөл алга”, proTitle:“🔑 PRO код оруулна уу”,
proPlaceholder:“PRO-XXXX”, proEnter:“✅ Идэвхжүүлэх”, proClose:“Хаах”,
proBtn:“🔑 PRO”, proUnlock:“PRO боломжуудыг нээх”,
back:“⬅”,
q1:“Сарын эцэс гэхэд хэд үлдэх вэ?”, q2:“Хоолонд хэтэрхий их зарцуулж байна уу?”,
q3:“Банкны хураамжийг хэрхэн бууруулах?”, q4:“10 жилд баян болох төлөвлөгөө”,
q5:“Яаралтай санаа хэзээ хуримтлуулах вэ?”,
chatPh:“Асуултаа бичнэ үү…”,
thisMonth:“Энэ сар”, last3:“Сүүлийн 3 сар”, last6:“Сүүлийн 6 сар”, yrView:“Жилийн дүгнэлт”,
},
kz:{ login:“Кіру”, register:“Тіркелу”, name:“Аты-жөні”, pin:“PIN (4 сан)”, pinConfirm:“PIN қайталаңыз”,
nameP:“Атыңызды жазыңыз”, pinP:“1234”, welcome:“Қош келдіңіз”,
exists:“Пайдаланушы бар, PIN енгізіңіз”, notFound:“Пайдаланушы табылмады”,
pinWrong:“PIN қате ⚠️”, pinMismatch:“PIN сәйкес емес”,
nameShort:“Ат 2 таңбадан ұзын болуы керек”,
pinShort:“PIN дәл 4 санды болуы керек”,
newUser:“Жаңа пайдаланушы? →”, hasAccount:“Тіркелгенсіз бе? →”,
myData:“Менің деректерім тек маған көрінеді 🔐”,
logOut:“Шығу”, dashboard:“📊 Басты”, bank:“🏦 Банк”,
add:“➕ Қосу”, history:“📅 Тарих”, math:“🧮 Математика”,
tips:“🧠 Кеңестер”, chat:“🤖 AI”, biz:“🏪 Бизнес”,
income:“Кіріс”, spent:“Шығыс”, left:“Қалды”,
health:“💚 Қаржылық денсаулық”,
uploadBtn:“📂 Банк анықтамасын жүктеу (.xlsx)”,
banks:“Хаан · Голомт · Хас · Ард банк”,
analyze:“🔍 AI талдау”, analyzing:“⏳ Талдануда…”,
desc:“Сипаттама”, amount:“Сома (₮)”, date:“Күні”, cat:“Категория”,
save:“✅ Сақтау”, del:“✕”,
monthly:“Ай қорытындысы”, annual:“Жылдық салыстыру”,
perDay:“Күніне орта”, savRate:“Жинақ %”, spRate:“Шығыс %”,
emgFund:“Апат қоры”, emgMonths:“Апат қорына”, yrSav:“Жылдық жинақ”,
invest:“📈 Инвестиция болжамы”, invNote:“8% жылдық пайыз”,
allRec:“Барлық жазбалар”, bankTx:“Банк мәмілелері”,
noData:“Деректер жоқ”, proTitle:“🔑 PRO кодты енгізіңіз”,
proPlaceholder:“PRO-XXXX”, proEnter:“✅ Белсендіру”, proClose:“Жабу”,
proBtn:“🔑 PRO”, proUnlock:“PRO мүмкіндіктерін ашу”,
back:“⬅”,
q1:“Ай соңына дейін қанша қалады?”, q2:“Тамаққа тым көп жұмсаймын ба?”,
q3:“Банк комиссиясын қалай азайтамын?”, q4:“10 жылда байлы болу жоспары”,
q5:“Апат қорымды қашан жинаймын?”,
chatPh:“Сұрағыңызды жазыңыз…”,
thisMonth:“Бұл ай”, last3:“Соңғы 3 ай”, last6:“Соңғы 6 ай”, yrView:“Жылдық есеп”,
},
};

// ════════════════════════════════════════════════════════════════
// НЕГІЗГІ ҚОСЫМША
// ════════════════════════════════════════════════════════════════
export default function App() {
const [lang, setLang] = useState(“mn”);
const t = TX[lang] || TX.mn;

// ── Аутентификация ───────────────────────────────────────────
const [authMode,    setAuthMode]    = useState(“login”); // login | register | app
const [nameInput,   setNameInput]   = useState(””);
const [pinInput,    setPinInput]    = useState(””);
const [pin2Input,   setPin2Input]   = useState(””);
const [authError,   setAuthError]   = useState(””);
const [authLoading, setAuthLoading] = useState(false);
const [currentUser, setCurrentUser] = useState(null); // {name, userId, pin}

// ── Қосымша деректері ────────────────────────────────────────
const [tab,       setTab]       = useState(“dashboard”);
const [expenses,  setExpenses]  = useState([]);
const [income,    setIncome]    = useState(1200000);
const [isPro,     setIsPro]     = useState(false);
const [showCode,  setShowCode]  = useState(false);
const [codeIn,    setCodeIn]    = useState(””);
const [codeErr,   setCodeErr]   = useState(””);

// ── Банк ─────────────────────────────────────────────────────
const [bankData,  setBankData]  = useState(null);
const [bankErr,   setBankErr]   = useState(””);
const [bankLoad,  setBankLoad]  = useState(false);
const [aiText,    setAiText]    = useState(””);
const [aiLoad,    setAiLoad]    = useState(false);

// ── Форм ─────────────────────────────────────────────────────
const [form, setForm] = useState({desc:””,amount:””,cat:“food”,date:TODAY});

// ── Тарих сүзгі ──────────────────────────────────────────────
const [histFilter, setHistFilter] = useState(THIS_MONTH);

// ── AI чат ───────────────────────────────────────────────────
const [chat,  setChat]  = useState([]);
const [chatQ, setChatQ] = useState(””);
const [chatL, setChatL] = useState(false);
const chatRef = useRef(null);

useEffect(() => { chatRef.current?.scrollIntoView({behavior:“smooth”}); }, [chat]);

// ── Деректерді сақтау/жүктеу ─────────────────────────────────
const saveData = useCallback(async (exp, inc, u=currentUser) => {
if (!u) return;
await saveUserData(u.userId, u.pin, { expenses:exp, income:inc, savedAt:Date.now() });
}, [currentUser]);

const setExpAndSave = useCallback((newExp) => {
setExpenses(newExp);
saveData(newExp, income);
}, [income, saveData]);

const setIncomeAndSave = useCallback((val) => {
setIncome(val);
saveData(expenses, val);
}, [expenses, saveData]);

// ── Кіру / Тіркелу ───────────────────────────────────────────
const handleAuth = useCallback(async () => {
setAuthError(””); setAuthLoading(true);
const name = nameInput.trim();
const pin  = pinInput.trim();

```
if (!name || name.length < 2) { setAuthError(t.nameShort); setAuthLoading(false); return; }
if (!pin || pin.length !== 4 || !/^\d{4}$/.test(pin)) { setAuthError(t.pinShort); setAuthLoading(false); return; }

const userId = name.toLowerCase().replace(/\s+/g,"_");
const idx    = getUsersIndex();

if (authMode === "register") {
  if (pin2Input !== pin) { setAuthError(t.pinMismatch); setAuthLoading(false); return; }
  const pinHash = await hashPin(pin, userId);
  idx[userId] = { name, pinHash, createdAt: Date.now() };
  saveUsersIndex(idx);
  const user = { name, userId, pin };
  setCurrentUser(user);
  setExpenses([]);
  await saveUserData(userId, pin, { expenses:[], income:1200000, savedAt:Date.now() });
  setIsPro(isProActive(userId));
  setAuthMode("app");
} else {
  if (!idx[userId]) { setAuthError(t.notFound); setAuthLoading(false); return; }
  const pinHash = await hashPin(pin, userId);
  if (idx[userId].pinHash !== pinHash) { setAuthError(t.pinWrong); setAuthLoading(false); return; }
  const user = { name: idx[userId].name, userId, pin };
  const data = await loadUserData(userId, pin);
  setCurrentUser(user);
  if (data) { setExpenses(data.expenses||[]); setIncome(data.income||1200000); }
  else { setExpenses([]); }
  setIsPro(isProActive(userId));
  setAuthMode("app");
}
setAuthLoading(false);
```

}, [nameInput, pinInput, pin2Input, authMode, t]);

const handleLogout = () => {
setCurrentUser(null); setAuthMode(“login”);
setNameInput(””); setPinInput(””); setPin2Input(””);
setExpenses([]); setBankData(null); setAiText(””); setChat([]);
};

// ── Математика ────────────────────────────────────────────────
const totalSpent = useMemo(() =>
expenses.filter(e=>!CATS.find(c=>c.id===e.cat)?.isSav && !e.isCredit).reduce((s,e)=>s+e.amount,0),
[expenses]);
const totalSav = useMemo(() =>
expenses.filter(e=>CATS.find(c=>c.id===e.cat)?.isSav).reduce((s,e)=>s+e.amount,0),
[expenses]);
const health = useMemo(() => {
const needs = expenses.filter(e=>[“food”,“transport”,“housing”,“health”].includes(e.cat)).reduce((s,e)=>s+e.amount,0);
const wants = expenses.filter(e=>[“entertain”,“fee”,“phone”,“clothes”].includes(e.cat)).reduce((s,e)=>s+e.amount,0);
const s1 = needs<=income*.5?100:Math.max(0,Math.round(100-(needs-income*.5)/(income*.5)*100));
const s2 = wants<=income*.3?100:Math.max(0,Math.round(100-(wants-income*.3)/(income*.3)*100));
const s3 = totalSav>=income*.2?100:pct(totalSav,income*.2);
return Math.round((s1+s2+s3)/3);
}, [expenses, income, totalSav]);

const catData = useMemo(() =>
CATS.map(c=>({nm:c.nm[lang]||c.nm.mn, val:expenses.filter(e=>e.cat===c.id&&!e.isCredit).reduce((s,e)=>s+e.amount,0), col:c.col})).filter(d=>d.val>0),
[expenses, lang]);

// ── Ай/жыл тарихы ────────────────────────────────────────────
const monthlyData = useMemo(() => {
const byMonth = {};
expenses.forEach(e => {
const m = (e.date||””).slice(0,7);
if (!m) return;
if (!byMonth[m]) byMonth[m] = {month:m, credit:0, debit:0, savings:0};
if (e.isCredit) byMonth[m].credit += e.amount;
else if (CATS.find(c=>c.id===e.cat)?.isSav) byMonth[m].savings += e.amount;
else byMonth[m].debit += e.amount;
});
return Object.values(byMonth).sort((a,b)=>a.month.localeCompare(b.month)).slice(-12);
}, [expenses]);

const filteredExp = useMemo(() => {
if (!histFilter) return expenses;
return expenses.filter(e=>(e.date||””).startsWith(histFilter));
}, [expenses, histFilter]);

// ── Файл жүктеу ───────────────────────────────────────────────
const handleFile = useCallback((e) => {
const file = e.target.files?.[0]; if (!file) return;
setBankLoad(true); setBankErr(””); setBankData(null); setAiText(””);
const reader = new FileReader();
reader.onload = (ev) => {
try {
const wb = XLSX.read(ev.target.result, {type:“array”,cellDates:true});
let best = null;
for (const sh of wb.SheetNames) {
const ws = wb.Sheets[sh];
const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:false});
const r = parseBankFile(raw);
if (!r.error && r.count>0 && (!best||r.count>best.count)) best={…r,sheetName:sh};
}
if (best) {
setBankData(best);
const newExps = best.txs.map(tx=>({
id: Date.now()+tx.id+Math.random(),
desc: tx.desc||(tx.credit>0?“Орлого”:“Зарлага”),
amount: tx.credit>0?tx.credit:tx.debit,
cat: tx.credit>0?“other”:autoCategory(tx.desc),
date: tx.date||TODAY,
isBank:true, isCredit:tx.credit>0,
}));
const merged = […expenses.filter(e=>!e.isBank), …newExps];
setExpenses(merged);
if (best.totC>0) { setIncome(best.totC); saveData(merged, best.totC); }
else saveData(merged, income);
} else {
setBankErr(lang===“mn”?“Баганын толгой мэдээлэл танигдсангүй. Монгол банкны Excel хуулга оруулна уу.”:“Файл форматы танылмады.”);
}
} catch(err) { setBankErr(String(err)); }
setBankLoad(false);
};
reader.readAsArrayBuffer(file);
e.target.value=””;
}, [expenses, income, lang, saveData]);

// ── AI талдауы (Gemini + Claude) ──────────────────────────────
const runAnalysis = useCallback(async () => {
if (!bankData) return;
setAiLoad(true); setAiText(””);
const sum = bankData.txs.slice(0,25).map(t=>`${t.date}|${t.credit>0?"+"+t.credit:"−"+t.debit}₮|${t.desc}`).join(”\n”);
const prompt = `Та ${lang===“mn”?“монгол”:“қазақ”} хэлний санхүүгийн шинжээч юм.

Банкны хуулга:

- Эзэмшигч: ${bankData.owner||“Хэрэглэгч”}
- Хугацаа: ${bankData.dateRange||”—”}
- Нийт орлого: ${bankData.totC.toLocaleString()}₮
- Нийт зарлага: ${bankData.totD.toLocaleString()}₮
- Цэвэр урсгал: ${bankData.net.toLocaleString()}₮
- Гүйлгээний тоо: ${bankData.count}

Гүйлгээнүүд:
${sum}

${lang===“mn”?“Монгол хэлээр дараах зүйлсийг тайлбарла:”:“Қазақ тілінде төмендегілерді түсіндір:”}

1. 💰 Орлого/зарлагын дүн шинжилгээ (тоогоор)
1. 📊 Гол гүйлгээнүүдийн тайлбар
1. ⚠️ Анхааруулга (хураамж, давтагдах шилжүүлэг гэх мэт)
1. 💡 5 нарийн зөвлөгөө (тоогоор)
1. 📈 Цаашдын зөвлөмж

Товч, тодорхой, ойлгомжтой бич.`;

```
try {
  // Gemini API
  if (GEMINI_API_KEY && GEMINI_API_KEY !== "YOUR_GEMINI_KEY_HERE") {
    const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`,{
      method:"POST", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({contents:[{parts:[{text:prompt}]}]})
    });
    const d = await r.json();
    if (!d.error) { setAiText(d.candidates?.[0]?.content?.parts?.[0]?.text||"Хариу алдаа"); setAiLoad(false); return; }
  }
  // Fallback → Claude
  const r = await fetch("https://api.anthropic.com/v1/messages",{
    method:"POST", headers:{"Content-Type":"application/json"},
    body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1500,messages:[{role:"user",content:prompt}]})
  });
  const d = await r.json();
  setAiText(d.content?.[0]?.text||"Алдаа");
} catch(err) { setAiText("⚠️ "+err.message); }
setAiLoad(false);
```

}, [bankData, lang]);

// ── AI чат ────────────────────────────────────────────────────
const sendChat = useCallback(async () => {
if (!chatQ.trim()) return;
const q=chatQ.trim(); setChatQ(””);
const nh=[…chat,{r:“user”,t:q}]; setChat(nh); setChatL(true);
const bk=bankData?`Банкны хуулга: орлого ${bankData.totC.toLocaleString()}₮, зарлага ${bankData.totD.toLocaleString()}₮, цэвэр ${bankData.net.toLocaleString()}₮.`:””;
const sys=`Та ${lang==="mn"?"монгол хэлний":"қазақша"} санхүүгийн зөвлөх. Хэрэглэгч ${currentUser?.name}: орлого ${income.toLocaleString()}₮, зарлага ${totalSpent.toLocaleString()}₮, эрүүл мэнд ${health}/100. ${bk} Тодорхой тоогоор, товч, emoji-тэй хариул.`;
try {
const r=await fetch(“https://api.anthropic.com/v1/messages”,{method:“POST”,headers:{“Content-Type”:“application/json”},
body:JSON.stringify({model:“claude-sonnet-4-20250514”,max_tokens:900,system:sys,messages:nh.map(m=>({role:m.r===“user”?“user”:“assistant”,content:m.t}))})
});
const d=await r.json();
setChat([…nh,{r:“ai”,t:d.content?.[0]?.text||“Алдаа”}]);
} catch { setChat([…nh,{r:“ai”,t:“Алдаа”}]); }
setChatL(false);
}, [chatQ, chat, income, totalSpent, health, bankData, lang, currentUser]);

const activatePro = () => {
if (!currentUser) return;
const r = checkProCode(codeIn, currentUser.userId);
if (r.ok) { setIsPro(true); setCodeErr(””); setShowCode(false); setCodeIn(””); }
else setCodeErr(r.err);
};

const addExp = () => {
if (!form.desc||!form.amount) return;
const e={…form,id:Date.now(),amount:parseFloat(form.amount)||0};
const ne=[…expenses,e]; setExpenses(ne); saveData(ne,income);
setForm({desc:””,amount:””,cat:“food”,date:TODAY});
};

// ── Стильдер ─────────────────────────────────────────────────
const BG   = “linear-gradient(135deg,#0a0a1a,#1a1040,#0d1525)”;
const CARD = {background:“rgba(255,255,255,0.06)”,borderRadius:20,padding:16,backdropFilter:“blur(10px)”,border:“1px solid rgba(255,255,255,0.08)”};
const INP  = {width:“100%”,background:“rgba(255,255,255,0.07)”,border:“1px solid rgba(255,255,255,0.15)”,borderRadius:14,padding:“12px 14px”,color:”#fff”,fontSize:13,boxSizing:“border-box”,outline:“none”};
const LF   = {mn:“🇲🇳”,kz:“🇰🇿”};

const TABS = [
{id:“dashboard”,lbl:t.dashboard},{id:“bank”,lbl:t.bank},
{id:“add”,lbl:t.add},{id:“history”,lbl:t.history},{id:“math”,lbl:t.math},
…(isPro?[{id:“biz”,lbl:t.biz}]:[]),
{id:“tips”,lbl:t.tips},{id:“chat”,lbl:t.chat},
];

// ════════════════════════════════════════════════════════════════
// 🔐 LOGIN ЭКРАНЫ
// ════════════════════════════════════════════════════════════════
if (authMode !== “app”) return (
<div style={{minHeight:“100vh”,background:BG,display:“flex”,flexDirection:“column”,alignItems:“center”,justifyContent:“center”,padding:24,fontFamily:”‘Segoe UI’,sans-serif”,color:”#fff”}}>

```
  {/* Тіл */}
  <div style={{display:"flex",gap:8,marginBottom:32}}>
    {Object.entries(LF).map(([l,f])=>(
      <button key={l} onClick={()=>setLang(l)} style={{padding:"8px 18px",borderRadius:24,border:`2px solid ${lang===l?"#7c3aed":"rgba(255,255,255,0.1)"}`,background:lang===l?"rgba(124,58,237,0.2)":"rgba(255,255,255,0.05)",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",transition:"all .2s"}}>{f} {l.toUpperCase()}</button>
    ))}
  </div>

  {/* Логотип */}
  <div style={{textAlign:"center",marginBottom:36}}>
    <div style={{fontSize:56,filter:"drop-shadow(0 0 20px rgba(124,58,237,0.6))"}}>💰</div>
    <div style={{fontSize:26,fontWeight:900,letterSpacing:2,marginTop:10,background:"linear-gradient(135deg,#7c3aed,#4ECDC4)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>{lang==="mn"?"МӨНГӨНИЙ ДЭВТЭР":"АҚША ДӘПТЕР"}</div>
    <div style={{fontSize:11,opacity:0.4,marginTop:4,letterSpacing:3}}>PERSONAL FINANCE · SECURE · PRO</div>
  </div>

  {/* Форм */}
  <div style={{...CARD,width:"100%",maxWidth:360,padding:24}}>
    <div style={{fontSize:18,fontWeight:800,marginBottom:4}}>
      {authMode==="login"?t.login:t.register}
    </div>
    <div style={{fontSize:11,opacity:0.4,marginBottom:20,display:"flex",alignItems:"center",gap:6}}>
      🔐 {t.myData}
    </div>

    <div style={{marginBottom:12}}>
      <div style={{fontSize:11,opacity:0.6,marginBottom:5}}>{t.name}</div>
      <input value={nameInput} onChange={e=>setNameInput(e.target.value)} placeholder={t.nameP}
        style={INP} onKeyDown={e=>e.key==="Enter"&&handleAuth()}/>
    </div>
    <div style={{marginBottom:12}}>
      <div style={{fontSize:11,opacity:0.6,marginBottom:5}}>{t.pin}</div>
      <input value={pinInput} onChange={e=>setPinInput(e.target.value.replace(/\D/g,"").slice(0,4))}
        placeholder={t.pinP} type="password" inputMode="numeric" maxLength={4}
        style={INP} onKeyDown={e=>e.key==="Enter"&&handleAuth()}/>
    </div>
    {authMode==="register" && (
      <div style={{marginBottom:12}}>
        <div style={{fontSize:11,opacity:0.6,marginBottom:5}}>{t.pinConfirm}</div>
        <input value={pin2Input} onChange={e=>setPin2Input(e.target.value.replace(/\D/g,"").slice(0,4))}
          placeholder={t.pinP} type="password" inputMode="numeric" maxLength={4}
          style={INP} onKeyDown={e=>e.key==="Enter"&&handleAuth()}/>
      </div>
    )}

    {authError && (
      <div style={{background:"rgba(255,80,80,0.15)",border:"1px solid rgba(255,80,80,0.3)",borderRadius:10,padding:"9px 12px",fontSize:12,color:"#ff8080",marginBottom:12}}>
        {authError}
      </div>
    )}

    <button onClick={handleAuth} disabled={authLoading} style={{width:"100%",padding:"13px",borderRadius:14,border:"none",background:authLoading?"rgba(255,255,255,0.1)":"linear-gradient(135deg,#7c3aed,#4ECDC4)",color:"#fff",fontSize:14,fontWeight:800,cursor:authLoading?"default":"pointer",marginBottom:12}}>
      {authLoading?"⏳...":(authMode==="login"?`🔓 ${t.login}`:`✅ ${t.register}`)}
    </button>

    <button onClick={()=>{setAuthMode(authMode==="login"?"register":"login");setAuthError("");}} style={{width:"100%",padding:"9px",borderRadius:12,border:"1px solid rgba(255,255,255,0.1)",background:"transparent",color:"rgba(255,255,255,0.5)",fontSize:12,cursor:"pointer"}}>
      {authMode==="login"?t.newUser:t.hasAccount}
    </button>

    {/* Тіркелген пайдаланушылар */}
    {(() => {
      const idx = getUsersIndex();
      const names = Object.values(idx);
      if (!names.length) return null;
      return (
        <div style={{marginTop:16,borderTop:"1px solid rgba(255,255,255,0.08)",paddingTop:12}}>
          <div style={{fontSize:10,opacity:0.4,marginBottom:8}}>{lang==="mn"?"Бүртгэлтэй хэрэглэгчид":"Тіркелген пайдаланушылар"}</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {names.map((u,i)=>(
              <button key={i} onClick={()=>setNameInput(u.name)} style={{padding:"5px 10px",borderRadius:20,border:"1px solid rgba(255,255,255,0.12)",background:"rgba(255,255,255,0.06)",color:"#fff",fontSize:11,cursor:"pointer"}}>
                👤 {u.name}
              </button>
            ))}
          </div>
        </div>
      );
    })()}
  </div>

  <div style={{fontSize:9,opacity:0.2,marginTop:20}}>AES-256 шифрлау · v3.2 · 🇲🇳 2026</div>
</div>
```

);

// ════════════════════════════════════════════════════════════════
// 🏠 НЕГІЗГІ ҚОСЫМША
// ════════════════════════════════════════════════════════════════
return (
<div style={{minHeight:“100vh”,background:BG,fontFamily:”‘Segoe UI’,sans-serif”,color:”#fff”,paddingBottom:90}}>

```
  {/* PRO modal */}
  {showCode && (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.9)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{...CARD,width:"100%",maxWidth:300,padding:22}}>
        <div style={{fontSize:16,fontWeight:900,marginBottom:4}}>{t.proTitle}</div>
        <div style={{fontSize:11,opacity:0.5,marginBottom:12}}>PRO-XXXX</div>
        <input value={codeIn} onChange={e=>setCodeIn(e.target.value.toUpperCase())} maxLength={8}
          placeholder={t.proPlaceholder} style={{...INP,textAlign:"center",fontSize:20,fontWeight:900,letterSpacing:5,marginBottom:8}}/>
        {codeErr&&<div style={{color:"#FF6B6B",fontSize:11,marginBottom:8,textAlign:"center"}}>{codeErr}</div>}
        <div style={{display:"flex",gap:8}}>
          <button onClick={activatePro} style={{flex:1,padding:11,borderRadius:13,border:"none",background:"linear-gradient(135deg,#f59e0b,#d97706)",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>{t.proEnter}</button>
          <button onClick={()=>{setShowCode(false);setCodeErr("");}} style={{flex:1,padding:11,borderRadius:13,border:"none",background:"rgba(255,255,255,0.1)",color:"#fff",cursor:"pointer",fontWeight:700}}>{t.proClose}</button>
        </div>
      </div>
    </div>
  )}

  {/* Header */}
  <div style={{background:"rgba(0,0,0,0.4)",backdropFilter:"blur(20px)",padding:"12px 14px",borderBottom:"1px solid rgba(255,255,255,0.07)",display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:100}}>
    <div>
      <div style={{fontSize:15,fontWeight:900,background:"linear-gradient(135deg,#a78bfa,#4ECDC4)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>
        💰 {lang==="mn"?"Мөнгөний Дэвтэр":"Ақша Дәптер"}
        {isPro&&<span style={{fontSize:9,background:"linear-gradient(135deg,#f59e0b,#d97706)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginLeft:5,fontWeight:900}}>PRO</span>}
      </div>
      <div style={{fontSize:10,opacity:0.45,marginTop:1}}>👤 {currentUser?.name} · 🔐</div>
    </div>
    <div style={{display:"flex",gap:5,alignItems:"center"}}>
      {Object.entries(LF).map(([l,f])=>(
        <button key={l} onClick={()=>setLang(l)} style={{padding:"4px 7px",borderRadius:8,border:`1px solid ${lang===l?"#7c3aed":"transparent"}`,background:lang===l?"rgba(124,58,237,0.2)":"rgba(255,255,255,0.05)",color:"#fff",fontSize:11,cursor:"pointer"}}>{f}</button>
      ))}
      {!isPro&&<button onClick={()=>setShowCode(true)} style={{padding:"5px 9px",borderRadius:9,border:"1px solid rgba(255,215,0,0.4)",background:"rgba(255,215,0,0.08)",color:"#fcd34d",fontSize:10,fontWeight:700,cursor:"pointer"}}>{t.proBtn}</button>}
      <button onClick={handleLogout} style={{padding:"5px 9px",borderRadius:9,border:"1px solid rgba(255,255,255,0.12)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.6)",fontSize:10,cursor:"pointer"}}>{t.logOut}</button>
    </div>
  </div>

  {/* Tabs */}
  <div style={{display:"flex",gap:3,padding:"8px 10px",overflowX:"auto",background:"rgba(0,0,0,0.3)",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
    {TABS.map(tb=>(
      <button key={tb.id} onClick={()=>setTab(tb.id)} style={{padding:"7px 11px",borderRadius:20,border:"none",cursor:"pointer",fontSize:10,fontWeight:700,whiteSpace:"nowrap",transition:"all .2s",
        background:tab===tb.id?"linear-gradient(135deg,#7c3aed,#4ECDC4)":"rgba(255,255,255,0.07)",
        color:tab===tb.id?"#fff":"rgba(255,255,255,0.6)"}}>
        {tb.lbl}
      </button>
    ))}
  </div>

  <div style={{padding:12,display:"flex",flexDirection:"column",gap:11}}>

    {/* ══ DASHBOARD ══════════════════════════════════════════ */}
    {tab==="dashboard" && <>
      {/* Кіріс */}
      <div style={CARD}>
        <div style={{fontSize:10,opacity:0.5,marginBottom:5}}>{t.income} (₮)</div>
        <input type="number" value={income} onChange={e=>setIncomeAndSave(parseInt(e.target.value)||0)} style={INP}/>
      </div>

      {/* 3 карточка */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
        {[
          {lbl:t.income,  v:fmtM(income),      c:"#4ECDC4",   ic:"💵"},
          {lbl:t.spent,   v:fmtM(totalSpent),   c:"#FF6B6B",   ic:"📤"},
          {lbl:t.left,    v:fmtM(income-totalSpent), c:income-totalSpent>=0?"#96CEB4":"#FF6B6B", ic:income-totalSpent>=0?"✅":"⚠️"},
        ].map(x=>(
          <div key={x.lbl} style={{...CARD,padding:"11px 7px",textAlign:"center"}}>
            <div style={{fontSize:18}}>{x.ic}</div>
            <div style={{fontSize:9,opacity:0.5,marginTop:2}}>{x.lbl}</div>
            <div style={{fontSize:11,fontWeight:900,color:x.c,marginTop:3}}>{x.v}</div>
          </div>
        ))}
      </div>

      {/* Денсаулық ұпайы */}
      <div style={CARD}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:9}}>
          <div style={{fontSize:12,fontWeight:700}}>{t.health}</div>
          <div style={{fontSize:24,fontWeight:900,color:sc(health)}}>{health}<span style={{fontSize:10,opacity:0.6}}>/100</span></div>
        </div>
        <div style={{background:"rgba(255,255,255,0.08)",borderRadius:8,height:10,overflow:"hidden",marginBottom:5}}>
          <div style={{height:"100%",width:`${health}%`,borderRadius:8,background:`linear-gradient(90deg,${sc(health)},${sc(health)}88)`,transition:"width 1s ease"}}/>
        </div>
        <div style={{fontSize:11,color:sc(health),fontWeight:600}}>
          {health>=80?(lang==="mn"?"Маш сайн ✅":"Өте жақсы ✅"):health>=60?(lang==="mn"?"Дундаж ⚠️":"Орташа ⚠️"):(lang==="mn"?"Сайжруулах хэрэгтэй 🔴":"Жақсарту керек 🔴")}
        </div>
      </div>

      {/* Pie chart */}
      {catData.length>0 && (
        <div style={CARD}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:9}}>📊 {lang==="mn"?"Ангилал":"Категориялар"}</div>
          <ResponsiveContainer width="100%" height={155}>
            <PieChart>
              <Pie data={catData.map(c=>({name:c.nm,value:c.val}))} cx="50%" cy="50%" innerRadius={38} outerRadius={62} dataKey="value" paddingAngle={3}>
                {catData.map((e,i)=><Cell key={i} fill={e.col}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#1a1040",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,color:"#fff",fontSize:10}}/>
            </PieChart>
          </ResponsiveContainer>
          <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:4}}>
            {catData.map((c,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:3,fontSize:9,background:"rgba(255,255,255,0.06)",padding:"2px 7px",borderRadius:8}}>
                <div style={{width:6,height:6,borderRadius:"50%",background:c.col,flexShrink:0}}/>
                {c.nm} <b>{fmtM(c.val)}</b>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Ай графигі */}
      {monthlyData.length>1 && (
        <div style={CARD}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:9}}>📅 {lang==="mn"?"Сарын орлого vs зарлага":"Ай кірісі мен шығысы"}</div>
          <ResponsiveContainer width="100%" height={160}>
            <AreaChart data={monthlyData} margin={{top:5,right:5,bottom:0,left:0}}>
              <defs>
                <linearGradient id="gc" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#4ECDC4" stopOpacity={0.3}/><stop offset="95%" stopColor="#4ECDC4" stopOpacity={0}/></linearGradient>
                <linearGradient id="gd" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#FF6B6B" stopOpacity={0.3}/><stop offset="95%" stopColor="#FF6B6B" stopOpacity={0}/></linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)"/>
              <XAxis dataKey="month" tick={{fill:"rgba(255,255,255,0.4)",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis hide/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#1a1040",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,color:"#fff",fontSize:10}}/>
              <Area type="monotone" dataKey="credit" stroke="#4ECDC4" strokeWidth={2} fill="url(#gc)" name={t.income}/>
              <Area type="monotone" dataKey="debit"  stroke="#FF6B6B" strokeWidth={2} fill="url(#gd)" name={t.spent}/>
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* Соңғы 5 жазба */}
      <div style={CARD}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:8}}>🧾 {lang==="mn"?"Сүүлийн бичлэгүүд":"Соңғы жазбалар"}</div>
        {expenses.length===0 && <div style={{opacity:0.3,fontSize:12,textAlign:"center",padding:16}}>{t.noData}</div>}
        {[...expenses].reverse().slice(0,6).map(e=>{
          const c=CATS.find(x=>x.id===e.cat);
          return (
            <div key={e.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
              <div style={{display:"flex",alignItems:"center",gap:7,minWidth:0,flex:1}}>
                <div style={{width:30,height:30,borderRadius:9,background:`${c?.col||"#888"}22`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>
                  {c?.nm[lang]?.slice(0,2)||"📦"}
                </div>
                <div style={{minWidth:0}}>
                  <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.desc} {e.isBank&&"🏦"}</div>
                  <div style={{fontSize:9,opacity:0.4}}>{c?.nm[lang]} · {e.date}</div>
                </div>
              </div>
              <div style={{fontSize:12,fontWeight:700,color:e.isCredit?"#4ECDC4":"#FF6B6B",flexShrink:0,marginLeft:8}}>
                {e.isCredit?"+":"-"}{fmtM(e.amount)}
              </div>
            </div>
          );
        })}
      </div>
    </>}

    {/* ══ БАНК АНЫҚТАМАСЫ ══════════════════════════════════ */}
    {tab==="bank" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.bank}</div>

      {/* Жүктеу аймағы */}
      <div style={{background:"rgba(78,205,196,0.06)",border:"2px dashed rgba(78,205,196,0.35)",borderRadius:20,padding:22,textAlign:"center"}}>
        <div style={{fontSize:32,marginBottom:8}}>📂</div>
        <div style={{fontSize:13,fontWeight:700,marginBottom:4}}>{t.uploadBtn}</div>
        <div style={{fontSize:10,opacity:0.45,marginBottom:14}}>{t.banks}</div>
        <label style={{padding:"11px 22px",borderRadius:14,border:"none",background:"linear-gradient(135deg,#4ECDC4,#45B7D1)",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",display:"inline-block"}}>
          {lang==="mn"?"📂 Файл сонгох":"📂 Файл таңдау"}
          <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
        </label>
      </div>

      {bankLoad && <div style={{...CARD,textAlign:"center",padding:20}}><div style={{fontSize:22}}>⏳</div><div style={{fontSize:12,marginTop:6,opacity:0.6}}>{lang==="mn"?"Уншиж байна...":"Оқылуда..."}</div></div>}
      {bankErr  && <div style={{...CARD,background:"rgba(255,107,107,0.1)",border:"1px solid rgba(255,107,107,0.3)"}}><div style={{fontSize:12,color:"#ff8080"}}>⚠️ {bankErr}</div></div>}

      {bankData && <>
        <div style={CARD}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:4}}>📋 {bankData.sheetName} {bankData.owner&&`— ${bankData.owner}`}</div>
          {bankData.dateRange && <div style={{fontSize:10,opacity:0.4,marginBottom:10}}>📅 {bankData.dateRange}</div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            {[
              {lbl:lang==="mn"?"Нийт орлого":"Жалпы кіріс",   v:fmt(bankData.totC), c:"#4ECDC4"},
              {lbl:lang==="mn"?"Нийт зарлага":"Жалпы шығыс",  v:fmt(bankData.totD), c:"#FF6B6B"},
              {lbl:lang==="mn"?"Цэвэр урсгал":"Таза ағыс",    v:fmt(bankData.net),  c:bankData.net>=0?"#96CEB4":"#FF6B6B"},
              {lbl:lang==="mn"?"Гүйлгээний тоо":"Мәміле саны",v:bankData.count+"шт",c:"#a78bfa"},
            ].map(x=>(
              <div key={x.lbl} style={{background:"rgba(255,255,255,0.05)",borderRadius:12,padding:11,textAlign:"center"}}>
                <div style={{fontSize:9,opacity:0.5}}>{x.lbl}</div>
                <div style={{fontSize:13,fontWeight:900,color:x.c,marginTop:3}}>{x.v}</div>
              </div>
            ))}
          </div>
          <button onClick={runAnalysis} disabled={aiLoad} style={{width:"100%",padding:"12px",borderRadius:13,border:"none",background:aiLoad?"rgba(255,255,255,0.1)":"linear-gradient(135deg,#7c3aed,#4ECDC4)",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            {aiLoad?t.analyzing:t.analyze} {GEMINI_API_KEY!=="YOUR_GEMINI_KEY_HERE"?"(Gemini ✓)":"(Claude AI)"}
          </button>
        </div>

        {/* Мәмілелер тізімі */}
        <div style={CARD}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:8}}>{t.bankTx} ({bankData.count})</div>
          <div style={{maxHeight:320,overflowY:"auto"}}>
            {bankData.txs.map((tx,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid rgba(255,255,255,0.04)"}}>
                <div style={{flex:1,minWidth:0,marginRight:8}}>
                  <div style={{fontSize:11,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{tx.desc}</div>
                  <div style={{fontSize:9,opacity:0.35}}>{tx.date}{tx.account&&` · ${tx.account}`}</div>
                </div>
                <div style={{textAlign:"right",flexShrink:0}}>
                  {tx.credit>0 && <div style={{fontSize:11,fontWeight:700,color:"#4ECDC4"}}>+{fmt(tx.credit)}</div>}
                  {tx.debit>0  && <div style={{fontSize:11,fontWeight:700,color:"#FF6B6B"}}>−{fmt(tx.debit)}</div>}
                  {tx.balance>0 && <div style={{fontSize:9,opacity:0.3}}>{fmtM(tx.balance)}</div>}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* AI талдауы */}
        {aiText && (
          <div style={{...CARD,background:"rgba(124,58,237,0.08)",border:"1px solid rgba(124,58,237,0.25)"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#a78bfa",marginBottom:8}}>
              🤖 AI {lang==="mn"?"Шинжилгээ":"Талдауы"} {GEMINI_API_KEY!=="YOUR_GEMINI_KEY_HERE"?"(Gemini)":"(Claude)"}
            </div>
            <div style={{fontSize:12,lineHeight:1.85,whiteSpace:"pre-wrap"}}>{aiText}</div>
          </div>
        )}
      </>}
    </>}

    {/* ══ ШЫҒЫС ҚОС ════════════════════════════════════════ */}
    {tab==="add" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.add}</div>
      <div style={CARD}>
        {[[t.desc,"desc","text","..."],[t.amount,"amount","number","50000"],[t.date,"date","date",""]].map(([lbl,fld,tp,ph])=>(
          <div key={fld} style={{marginBottom:10}}>
            <div style={{fontSize:10,opacity:0.5,marginBottom:4}}>{lbl}</div>
            <input type={tp} value={form[fld]} placeholder={ph} onChange={e=>setForm(p=>({...p,[fld]:e.target.value}))} style={INP}/>
          </div>
        ))}
        <div style={{marginBottom:12}}>
          <div style={{fontSize:10,opacity:0.5,marginBottom:6}}>{t.cat}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5}}>
            {CATS.map(c=>(
              <button key={c.id} onClick={()=>setForm(p=>({...p,cat:c.id}))} style={{padding:"8px",borderRadius:11,border:`2px solid ${form.cat===c.id?c.col:"transparent"}`,background:form.cat===c.id?`${c.col}18`:"rgba(255,255,255,0.04)",color:"#fff",fontSize:11,cursor:"pointer",textAlign:"left",transition:"all .15s"}}>
                {c.nm[lang]||c.nm.mn}
              </button>
            ))}
          </div>
        </div>
        <button onClick={addExp} style={{width:"100%",padding:"12px",borderRadius:13,border:"none",background:"linear-gradient(135deg,#4ECDC4,#45B7D1)",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>{t.save}</button>
      </div>
    </>}

    {/* ══ ТАРИХ ════════════════════════════════════════════ */}
    {tab==="history" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.history}</div>

      {/* Сүзгі */}
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {[
          {lbl:t.thisMonth, val:THIS_MONTH},
          {lbl:t.last3, val:"3m"},
          {lbl:t.last6, val:"6m"},
          {lbl:t.yrView, val:TODAY.slice(0,4)},
          {lbl:lang==="mn"?"Бүгд":"Барлығы", val:""},
        ].map(f=>(
          <button key={f.val} onClick={()=>setHistFilter(f.val)} style={{padding:"6px 12px",borderRadius:18,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,
            background:histFilter===f.val?"linear-gradient(135deg,#7c3aed,#4ECDC4)":"rgba(255,255,255,0.07)",color:"#fff"}}>
            {f.lbl}
          </button>
        ))}
      </div>

      {/* Ай сайынғы диаграмма */}
      {monthlyData.length>0 && (
        <div style={CARD}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:9}}>📊 {lang==="mn"?"Сарын зарлага":"Ай шығысы"}</div>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={monthlyData} margin={{top:4,right:4,bottom:0,left:0}}>
              <XAxis dataKey="month" tick={{fill:"rgba(255,255,255,0.4)",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis hide/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#1a1040",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,color:"#fff",fontSize:10}}/>
              <Bar dataKey="debit" radius={[5,5,0,0]} name={t.spent}>
                {monthlyData.map((_,i)=><Cell key={i} fill={`hsl(${345+i*10},70%,58%)`}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* Тізім */}
      <div style={CARD}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:8}}>
          {t.allRec} ({filteredExp.length})
          {filteredExp.length>0 && <span style={{marginLeft:8,fontSize:10,color:"#FF6B6B"}}>
            -{fmt(filteredExp.filter(e=>!e.isCredit).reduce((s,e)=>s+e.amount,0))}
          </span>}
        </div>
        <div style={{maxHeight:400,overflowY:"auto"}}>
          {filteredExp.length===0 && <div style={{opacity:0.3,fontSize:12,textAlign:"center",padding:16}}>{t.noData}</div>}
          {[...filteredExp].reverse().map(e=>{
            const c=CATS.find(x=>x.id===e.cat);
            return (
              <div key={e.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid rgba(255,255,255,0.04)"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:11,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.desc} {e.isBank&&"🏦"}</div>
                  <div style={{fontSize:9,opacity:0.35}}>{c?.nm[lang]||c?.nm.mn} · {e.date}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:7,flexShrink:0,marginLeft:8}}>
                  <span style={{fontSize:11,fontWeight:700,color:e.isCredit?"#4ECDC4":"#FF6B6B"}}>{e.isCredit?"+":"-"}{fmtM(e.amount)}</span>
                  <button onClick={()=>{const ne=expenses.filter(x=>x.id!==e.id);setExpenses(ne);saveData(ne,income);}} style={{background:"rgba(255,107,107,0.15)",border:"none",color:"#FF6B6B",borderRadius:7,padding:"2px 6px",cursor:"pointer",fontSize:10}}>{t.del}</button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </>}

    {/* ══ МАТЕМАТИКА ═══════════════════════════════════════ */}
    {tab==="math" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.math}</div>

      <div style={CARD}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:9}}>📋 {t.monthly}</div>
        {[
          {ic:"📅",lbl:t.perDay,    v:fmt(Math.round(totalSpent/30))},
          {ic:"💎",lbl:t.savRate,   v:`${pct(totalSav,income)}%`},
          {ic:"📤",lbl:t.spRate,    v:`${Math.min(100,pct(totalSpent,income))}%`},
          {ic:"🛡",lbl:t.emgFund,   v:fmt(totalSpent*6)},
          {ic:"⏳",lbl:t.emgMonths, v:totalSav>0?Math.ceil(totalSpent*6/totalSav)+(lang==="mn"?" сар":" ай"):"-"},
          {ic:"💰",lbl:t.yrSav,     v:fmt(totalSav*12)},
          {ic:"📊",lbl:lang==="mn"?"Жилийн зарлага":"Жылдық шығын",v:fmt(totalSpent*12)},
        ].map(r=>(
          <div key={r.lbl} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
            <span style={{fontSize:11,opacity:0.65}}>{r.ic} {r.lbl}</span>
            <span style={{fontSize:12,fontWeight:700}}>{r.v}</span>
          </div>
        ))}
      </div>

      <div style={CARD}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:4}}>{t.invest}</div>
        <div style={{fontSize:10,opacity:0.45,marginBottom:10}}>{t.invNote} · {fmt(totalSav)}/ай</div>
        <ResponsiveContainer width="100%" height={145}>
          <BarChart data={Array.from({length:6},(_,i)=>{
            const y=(i+1)*5, r=0.08/12, n=y*12;
            return {y:`${y}жил`,v:Math.round(totalSav*((Math.pow(1+r,n)-1)/r))};
          })}>
            <XAxis dataKey="y" tick={{fill:"rgba(255,255,255,0.5)",fontSize:9}} axisLine={false} tickLine={false}/>
            <YAxis hide/>
            <Tooltip formatter={v=>fmtM(v)} contentStyle={{background:"#1a1040",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,color:"#fff",fontSize:10}}/>
            <Bar dataKey="v" radius={[5,5,0,0]}>
              {Array.from({length:6},(_,i)=><Cell key={i} fill={`hsl(${155+i*20},65%,55%)`}/>)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        <div style={{fontSize:9,opacity:0.35,marginTop:6}}>📐 FV = PMT × ((1+r)ⁿ − 1) / r</div>
      </div>

      {/* 50/30/20 */}
      <div style={{...CARD,background:"rgba(124,58,237,0.07)",border:"1px solid rgba(124,58,237,0.2)"}}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:10,color:"#a78bfa"}}>📐 50/30/20 {lang==="mn"?"дүрэм":"ережесі"}</div>
        {[
          {lbl:lang==="mn"?"🏠 Хэрэгцээ (50%)":"🏠 Қажеттілік (50%)", pct:50, c:"#4ECDC4"},
          {lbl:lang==="mn"?"🎮 Хүсэл (30%)":"🎮 Тілек (30%)",          pct:30, c:"#FFEAA7"},
          {lbl:lang==="mn"?"💎 Хадгаламж (20%)":"💎 Жинақ (20%)",        pct:20, c:"#96CEB4"},
        ].map(r=>(
          <div key={r.lbl} style={{marginBottom:12}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:4}}>
              <span>{r.lbl}</span>
              <span style={{fontWeight:800,color:r.c}}>{fmt(Math.round(income*r.pct/100))}</span>
            </div>
            <div style={{background:"rgba(255,255,255,0.08)",borderRadius:7,height:8}}>
              <div style={{width:`${r.pct}%`,height:"100%",borderRadius:7,background:r.c,transition:"width .5s"}}/>
            </div>
          </div>
        ))}
      </div>
    </>}

    {/* ══ КЕҢЕСТЕР ══════════════════════════════════════════ */}
    {tab==="tips" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.tips}</div>
      {(lang==="mn"?[
        {ic:"💡",cat:"Үндэс",       t:"50/30/20: орлогын 50% хэрэгцээ, 30% хүсэл, 20% хадгаламж — дэлхийн алтан стандарт."},
        {ic:"📊",cat:"Сэтгэл зүй",  t:"Гараар бичсэн хүмүүс 23% бага зарцуулдаг — тархи хариуцлага мэдэрдэг."},
        {ic:"🎯",cat:"Хадгаламж",   t:"3-6 сарын яаралтай сан. Ажил алдвал, өвчилвөл хамгаалдаг."},
        {ic:"☕",cat:"Жижиг зардал",t:"Өдөр бүр кофе = сард ~27,000₮ = жилд ~324,000₮. Жижиг зардал — том мөнгө."},
        {ic:"🏦",cat:"Банкны хураамж",t:"Ухаалаг банкны хураамж гүйлгээ бүрт 50₮ = хэдэн зуун гүйлгээнд мэдэгдэхүйц дүн болно."},
        {ic:"📱",cat:"Захиалга",    t:"Ашиглагдаагүй захиалгуудыг шалгаарай. Дундаж хүн сард 3-4 буруу захиалгад төлдөг."},
        {ic:"💳",cat:"Зээл",        t:"Зээлийн картыг бүрэн төлөөрэй. 20-25% хүү хамгийн үнэтэй мөнгө."},
        {ic:"🌙",cat:"24 цагийн дүрэм",t:"Том худалдан авалтын өмнө 24 цаг хүлээгээрэй. 80% тохиолдолд хүсэл өөрөө арилна."},
        {ic:"📈",cat:"Хөрөнгө оруулалт",t:"Сар бүр орлогын 10% хөрөнгө оруулалт хийвэл 30 жилд 8% жилийн хүүгээр 10 дахин өснө."},
        {ic:"🛒",cat:"Хоол хүнс",   t:"Жагсаалтгүйгээр дэлгүүр орвол 35% илүү зарцуулдаг. Жагсаалт бич, өлсөж байхдаа битгий яв."},
        {ic:"🏠",cat:"Орон сууц",   t:"Орон суцны зардал орлогын 30%-иас хэтрэхгүй байх ёстой — Quicken/Mint стандарт."},
        {ic:"🤝",cat:"Гэр бүл",    t:"Хамтдаа төсөв гаргадаг гэр бүлүүд 40% бага маргалддаг, 2 дахин хурдан хуримтлуулдаг."},
      ]:[
        {ic:"💡",cat:"Негіз",       t:"50/30/20: кірістің 50% қажеттілік, 30% тілек, 20% жинақ — әлемдік алтын стандарт."},
        {ic:"📊",cat:"Психология",  t:"Шығысты қолмен жазған адамдар 23% аз жұмсайды — ми жауапкершілік сезінеді."},
        {ic:"🎯",cat:"Жинақ",       t:"3-6 айлық апат қоры жинаңыз. Жұмыс жоғалтсаңыз немесе ауырсаңыз сізді қорғайды."},
        {ic:"☕",cat:"Ұсақ шығын",  t:"Күнде 1 кофе = айына ~9,000₮ = жылына ~108,000₮. Ұсақ шығын — үлкен ақша."},
        {ic:"🏦",cat:"Банк комиссия",t:"Ухаалаг банк хураамжын – сома шимді болғанымен жинақталады. Мәмілелер санын бақылаңыз."},
        {ic:"💳",cat:"Несие",       t:"Несие картасын толық төлеңіз. 20-25% пайыз — ең қымбат ақша. Минимум тұзақ."},
        {ic:"🌙",cat:"24 сағат",    t:"Үлкен сатып алу алдында 24 сағат күтіңіз. 80% жағдайда тілек өзі өтеді."},
        {ic:"📈",cat:"Инвестиция",  t:"Ай сайын кірістің 10% инвестицияға салсаңыз, 30 жылда 10x өседі (8% жылдық)."},
        {ic:"🛒",cat:"Азық-түлік",  t:"Тізімсіз дүкенге барсаңыз 35% артық жұмсайсыз. Тізім жазыңыз, аш барма."},
      ]).map((tp,i)=>(
        <div key={i} style={{...CARD,borderLeft:`4px solid hsl(${140+i*13},60%,55%)`,padding:13}}>
          <div style={{display:"flex",gap:9,alignItems:"flex-start"}}>
            <div style={{fontSize:19,flexShrink:0}}>{tp.ic}</div>
            <div>
              <div style={{fontSize:9,color:`hsl(${140+i*13},60%,65%)`,fontWeight:700,marginBottom:3,textTransform:"uppercase",letterSpacing:1}}>{tp.cat}</div>
              <div style={{fontSize:12,lineHeight:1.75}}>{tp.t}</div>
            </div>
          </div>
        </div>
      ))}
    </>}

    {/* ══ AI ЧАТ ════════════════════════════════════════════ */}
    {tab==="chat" && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.chat}</div>
      {bankData && (
        <div style={{fontSize:10,background:"rgba(78,205,196,0.08)",border:"1px solid rgba(78,205,196,0.2)",borderRadius:10,padding:"7px 11px",opacity:0.8}}>
          🏦 {lang==="mn"?"Банкны хуулга ачааллагдсан — AI танд хувийн зөвлөгөө өгнө":"Банк анықтамасы жүктелді — AI жеке кеңес береді"}
        </div>
      )}
      <div style={{...CARD,minHeight:250,maxHeight:370,overflowY:"auto",display:"flex",flexDirection:"column",gap:9,padding:11}}>
        {chat.length===0 && <div style={{opacity:0.25,fontSize:12,textAlign:"center",marginTop:40}}>{t.chatPh}</div>}
        {chat.map((m,i)=>(
          <div key={i} style={{display:"flex",justifyContent:m.r==="user"?"flex-end":"flex-start"}}>
            <div style={{maxWidth:"87%",padding:"9px 13px",borderRadius:14,fontSize:12,lineHeight:1.75,
              background:m.r==="user"?"linear-gradient(135deg,#7c3aed,#4ECDC4)":"rgba(255,255,255,0.08)",
              whiteSpace:"pre-wrap"}}>
              {m.t}
            </div>
          </div>
        ))}
        {chatL && <div style={{display:"flex"}}><div style={{background:"rgba(255,255,255,0.08)",padding:"8px 13px",borderRadius:13,fontSize:11,opacity:0.5}}>⏳ ...</div></div>}
        <div ref={chatRef}/>
      </div>
      <div style={{display:"flex",gap:7}}>
        <input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&sendChat()}
          placeholder={t.chatPh} style={{...INP,flex:1}}/>
        <button onClick={sendChat} disabled={chatL||!chatQ.trim()} style={{padding:"11px 14px",borderRadius:13,border:"none",background:"linear-gradient(135deg,#7c3aed,#4ECDC4)",color:"#fff",fontSize:16,cursor:"pointer",flexShrink:0}}>➤</button>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
        {[t.q1,t.q2,t.q3,t.q4,t.q5].map(q=>(
          <button key={q} onClick={()=>{setChatQ(q);setTimeout(sendChat,60);}} style={{padding:"5px 9px",borderRadius:18,border:"1px solid rgba(255,255,255,0.12)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.75)",fontSize:10,cursor:"pointer"}}>{q}</button>
        ))}
      </div>
    </>}

    {/* ══ БИЗНЕС (PRO) ══════════════════════════════════════ */}
    {tab==="biz" && isPro && <>
      <div style={{fontSize:14,fontWeight:900}}>{t.biz}</div>
      <div style={{...CARD,background:"rgba(245,158,11,0.06)",border:"1px solid rgba(245,158,11,0.2)",textAlign:"center",padding:20}}>
        <div style={{fontSize:28,marginBottom:6}}>🏪</div>
        <div style={{fontSize:13,fontWeight:700,marginBottom:4}}>{lang==="mn"?"Бизнесийн модуль":"Бизнес модулі"}</div>
        <div style={{fontSize:11,opacity:0.5}}>{lang==="mn"?"Дараагийн шинэчлэлтэд орно":"Келесі жаңартуда қосылады"}</div>
      </div>
    </>}

    {/* PRO Unlock banner */}
    {!isPro && (tab==="dashboard"||tab==="tips") && (
      <div style={{background:"rgba(245,158,11,0.07)",border:"1px solid rgba(245,158,11,0.25)",borderRadius:18,padding:14}}>
        <div style={{fontSize:13,fontWeight:800,color:"#fcd34d",marginBottom:5}}>{t.proUnlock}</div>
        <div style={{fontSize:11,opacity:0.75,lineHeight:1.7,marginBottom:9}}>
          🏪 {lang==="mn"?"Бизнесийн касс":"Бизнес кассасы"} · 📊 {lang==="mn"?"Нарийн аналитик":"Кешенді аналитика"} · 📈 {lang==="mn"?"Урт хугацааны болжам":"Ұзақ мерзімді болжам"}
        </div>
        <button onClick={()=>setShowCode(true)} style={{width:"100%",padding:"11px",borderRadius:13,border:"none",background:"linear-gradient(135deg,#f59e0b,#d97706)",color:"#fff",fontSize:12,fontWeight:800,cursor:"pointer"}}>{t.proBtn}</button>
      </div>
    )}

  </div>
</div>
```

);
}
