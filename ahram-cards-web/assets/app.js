/* =====================================================================
   فك كروت الاهرام — منطق الاستخراج (يعمل بالكامل داخل المتصفح)
   الإصدار v1.2
   لا يُرفع أي ملف ولا تُحفظ أي بيانات على أي خادم.
   يطابق منطق برنامج سطح المكتب card_extractor_apple.py
   v1.1: اكتشاف مرن لأعمدة الكود/السيريال — يتحمّل تغيّر مكان العمود
         أو اختلاف تسميته (مثل «الرقم السري (PIN)») بين الفواتير.
   v1.2: اكتشاف مرن لاسم الشركة — يتحمّل اختلاف التهجئة واللغة وإعادة
         التسمية والأخطاء الإملائية (مثل «Etisalst» بدل «Etisalat»).
   ===================================================================== */

"use strict";

// أسماء أعمدة رأس الجدول كما تظهر في ملف الفاتورة
const HEADER_CODE = "الكود";
const HEADER_SERIAL = "السيريال";

// أكواد الشركات — مطابقة مرنة تتحمّل اختلاف التهجئة واللغة وإعادة التسمية والأخطاء الإملائية.
// لكل شركة:
//   exact    : أسماء مطابِقة تماماً (إنجليزي/عربي) — الأسرع والأدق
//   contains : نصوص مميّزة إن ظهرت في الاسم دلّت على الشركة (تعالج العربي/الاختصار/إعادة التسمية)
//   fuzzy    : أسماء إنجليزية تُقاس عليها أقرب تهجئة (تعالج الأخطاء الإملائية). أطوال ≥5 فقط.
const COMPANIES = [
  { code: "010", exact: ["vodafone", "فودافون"],
    contains: ["voda", "فودا"], fuzzy: ["vodafone"] },
  { code: "011", exact: ["etisalat", "etsalat", "اتصالات"],
    contains: ["etis", "etsa", "atis", "اتصال", "e&"], fuzzy: ["etisalat"] },
  { code: "012", exact: ["orange", "اورانج", "اورنج", "mobinil", "موبينيل"],
    contains: ["orang", "اوران", "اورن", "mobinil", "موبين"], fuzzy: ["orange", "mobinil"] },
  { code: "015", exact: ["we", "وي", "te data", "tedata", "telecom egypt"],
    contains: ["وي", "te data", "tedata", "telecom"], fuzzy: [] }, // «we» أقصر من أن تُطابَق تقريبياً بأمان
];

// ----------------------------- عناصر الواجهة -----------------------------
const $ = (id) => document.getElementById(id);
const dropzone = $("drop");
const fileInput = $("file");
const fileNameEl = $("fileName");
const runBtn = $("runBtn");
const clearBtn = $("clearBtn");
const optMerge = $("optMerge");
const optConfirm = $("optConfirm");
const logEl = $("log");
const warnBanner = $("warnBanner");

let currentFile = null;

// ----------------------------- أدوات مساعدة -----------------------------
// level: "" | "ok" | "warn" | "err"  (يلوّن السطر في السجل)
function log(msg, level) {
  const line = document.createElement("div");
  line.className = "logline" + (level ? " " + level : "");
  line.textContent = msg;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}
function clearLog() { logEl.innerHTML = ""; }

function showWarnBanner(count) {
  warnBanner.textContent = `⚠️ يوجد ${count} تنبيه — هناك فئة/ورقة فيها بيانات ناقصة. راجع الأسطر المميّزة بالأصفر في السجل.`;
  warnBanner.classList.remove("hidden");
}
function hideWarnBanner() { warnBanner.classList.add("hidden"); }

function normalizeCell(ws, r, c) {
  // تحويل قيمة الخلية إلى نص نظيف مع محاولة الحفاظ على الأصفار الأولى
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr];
  if (!cell) return "";
  if (cell.t === "n" && Number.isInteger(cell.v)) return String(cell.v);
  if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
  if (cell.v === undefined || cell.v === null) return "";
  return String(cell.v).trim();
}

function companyOf(sheetName) {
  // اسم الشركة = أول كلمة في اسم الورقة (Vodafone 15LE -> Vodafone)
  const parts = sheetName.trim().split(/\s+/);
  return parts[0] || sheetName.trim();
}

function levenshtein(a, b) {
  // مسافة التحرير بين نصّين (عدد الإضافات/الحذف/الاستبدال) — لمطابقة التهجئة التقريبية
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  let cur = new Array(n + 1);
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, cur] = [cur, prev];
  }
  return prev[n];
}

function companyCode(company, sheetName) {
  // كود الشركة الرقمي (Vodafone -> 010)، بمطابقة مرنة على ثلاث مراحل.
  // فارغ يبقى فارغاً (مع تنبيه) إذا كانت الشركة غير معروفة فعلاً.
  const w = normHeader(company);                       // أول كلمة (الشركة)
  const full = normHeader(sheetName || company);       // اسم الورقة كاملاً (لتواقيع متعددة الكلمات)
  if (!w) return "";

  // 1) تطابق دقيق
  for (const co of COMPANIES) if (co.exact.includes(w)) return co.code;

  // 2) توقيع مميّز ضمن الاسم (يعالج العربي/الإنجليزي/الاختصار/إعادة التسمية مثل Mobinil)
  for (const co of COMPANIES)
    for (const sig of co.contains)
      if (w.includes(sig) || full.includes(sig)) return co.code;

  // 3) أقرب تهجئة (يعالج الأخطاء الإملائية مثل Etisalst/Vodfone/Ornage)
  //    على الأسماء الطويلة فقط (≥5) لتفادي مطابقة كلمة قصيرة بالخطأ.
  let best = { code: "", dist: Infinity };
  for (const co of COMPANIES)
    for (const name of co.fuzzy) {
      if (name.length < 5) continue;
      const d = levenshtein(w, name);
      const thresh = Math.max(1, Math.floor(name.length / 3)); // 8→2، 6→2، 5→1
      if (d <= thresh && d < best.dist) best = { code: co.code, dist: d };
    }
  return best.code;
}

function categoryCode(sheetName) {
  // كود الفئة = أول رقم في اسم الورقة (Vodafone 200LE -> 200)
  const m = sheetName.match(/\d+/);
  return m ? m[0] : "";
}

function safeName(name) {
  // إزالة الرموز غير المسموح بها في أسماء الملفات والمجلدات
  let out = String(name);
  for (const ch of '<>:"/\\|?*') out = out.split(ch).join("_");
  out = out.trim();
  return out || "sheet";
}

// مطابقة مرنة لأسماء أعمدة الرأس (يتحمّل اختلاف التسمية بين الفواتير)
function normHeader(s) { return String(s).toLowerCase().replace(/\s+/g, " ").trim(); }

function isSerialHeader(s) {
  // عمود السيريال: «السيريال» / «السريال» / Serial …
  const n = normHeader(s);
  return n.includes("سيريال") || n.includes("سريال") || n.includes("serial");
}

function isCodeHeader(s) {
  // عمود الكود/الرقم السري: «الكود» / «الرقم السري (PIN)» / PIN / «كود الشحن» …
  const n = normHeader(s);
  return n === "الكود" || n.includes("الرقم السري") || n.includes("الرقم السرى") ||
         n.includes("pin") || n.includes("كود الشحن") || n.includes("رقم الشحن") || n.includes("كود الكارت");
}

function countNonEmpty(ws, col, fromRow, toRow) {
  // عدد الخلايا غير الفارغة في عمود ضمن نطاق صفوف (لاختيار العمود الذي يحوي بيانات فعلاً)
  let n = 0;
  for (let r = fromRow; r <= toRow; r++) if (normalizeCell(ws, r, col) !== "") n++;
  return n;
}

function findHeader(ws) {
  // البحث المرن عن صف رأس الجدول.
  // يُرجع {row, codeCol, serialCol, codeHeader, serialHeader, codeFromFallback} أو null.
  // عند وجود أكثر من عمود مرشّح للكود: يُفضّل «الكود» إن كان يحوي بيانات،
  // وإلا يعود إلى العمود الذي يحوي بيانات فعلاً (مثل «الرقم السري (PIN)»).
  if (!ws["!ref"]) return null;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const maxScan = Math.min(range.e.r, range.s.r + 24); // حتى 25 صفاً
  for (let r = range.s.r; r <= maxScan; r++) {
    const codeCols = [], serialCols = [];
    let exactCodeCol = null, exactSerialCol = null;
    const headerText = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const s = normalizeCell(ws, r, c);
      if (!s) continue;
      headerText[c] = s;
      if (isCodeHeader(s)) { codeCols.push(c); if (normHeader(s) === "الكود") exactCodeCol = c; }
      if (isSerialHeader(s)) { serialCols.push(c); if (normHeader(s) === "السيريال") exactSerialCol = c; }
    }
    if (codeCols.length === 0 || serialCols.length === 0) continue;

    // يختار العمود: «الدقيق» إن كان به بيانات، وإلا الأكثر امتلاءً، وإلا الدقيق/الأول
    const pick = (cols, exact) => {
      if (exact !== null && countNonEmpty(ws, exact, r + 1, range.e.r) > 0) return exact;
      let best = null, bestN = -1;
      for (const c of cols) { const n = countNonEmpty(ws, c, r + 1, range.e.r); if (n > bestN) { bestN = n; best = c; } }
      if (bestN > 0) return best;
      return exact !== null ? exact : cols[0];
    };
    const codeCol = pick(codeCols, exactCodeCol);
    const serialCol = pick(serialCols, exactSerialCol);
    return {
      row: r, codeCol, serialCol,
      codeHeader: headerText[codeCol] || HEADER_CODE,
      serialHeader: headerText[serialCol] || HEADER_SERIAL,
      codeFromFallback: normHeader(headerText[codeCol] || "") !== "الكود",
    };
  }
  return null;
}

function headerDiagnosis(ws) {
  // هل ظهر عمود كود و/أو سيريال (بأي تسمية) في الورقة؟ (لاكتشاف عمود ناقص)
  let sawCode = false, sawSerial = false;
  if (!ws["!ref"]) return { sawCode, sawSerial };
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const maxScan = Math.min(range.e.r, range.s.r + 24);
  for (let r = range.s.r; r <= maxScan; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const s = normalizeCell(ws, r, c);
      if (!s) continue;
      if (isCodeHeader(s)) sawCode = true;
      if (isSerialHeader(s)) sawSerial = true;
    }
  }
  return { sawCode, sawSerial };
}

function rowsToCSV(rows) {
  // CSV بترميز UTF-8 مع BOM (لدعم العربية في إكسل) ونهايات أسطر CRLF
  const esc = (v) => {
    const s = v == null ? "" : String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return "﻿" + rows.map((r) => r.map(esc).join(",")).join("\r\n");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

// ----------------------------- نافذة تأكيد كود الفئة -----------------------------
let approveAll = false;

function askCategory(sheetTitle, company, ccode, detectedCode) {
  // تعرض النافذة وتُرجع Promise بـ {code, skip, all}
  return new Promise((resolve) => {
    const overlay = $("modalOverlay");
    $("mSheet").textContent = sheetTitle;
    $("mCompany").textContent = `${company}  (كود الشركة: ${ccode || "—"})`;
    const codeInput = $("mCode");
    codeInput.value = detectedCode;

    overlay.classList.remove("hidden");
    codeInput.focus();
    codeInput.select();

    const cleanup = () => {
      overlay.classList.add("hidden");
      $("mApprove").onclick = null;
      $("mApproveAll").onclick = null;
      $("mSkip").onclick = null;
      codeInput.onkeydown = null;
    };

    $("mApprove").onclick = () => { const v = codeInput.value.trim(); cleanup(); resolve({ code: v, skip: false, all: false }); };
    $("mApproveAll").onclick = () => { const v = codeInput.value.trim(); cleanup(); resolve({ code: v, skip: false, all: true }); };
    $("mSkip").onclick = () => { cleanup(); resolve({ code: detectedCode, skip: true, all: false }); };
    codeInput.onkeydown = (e) => { if (e.key === "Enter") $("mApprove").onclick(); };
  });
}

// ----------------------------- الاستخراج -----------------------------
async function extract(file) {
  const t0 = performance.now();       // لإثبات عدم وجود طلبات شبكة أثناء المعالجة
  hideWarnBanner();

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const base = file.name.replace(/\.[^.]+$/, "");
  const merge = optMerge.checked;
  const useConfirm = optConfirm.checked;
  approveAll = false;

  const outputs = [];                 // [{ path, rows }]
  const mergedByCompany = {};         // { company: rows[] } لوضع الدمج
  const summary = {};
  const warnings = [];                // تنبيهات النواقص
  let totalRows = 0;

  // تنبيه: يسجّل سطراً أصفر ويضيفه لقائمة التنبيهات
  const warn = (m) => { log(m, "warn"); warnings.push(m); };

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const header = findHeader(ws);
    if (!header) {
      // قد تكون ورقة منتج لكن ينقصها أحد العمودين
      const d = headerDiagnosis(ws);
      if (d.sawCode !== d.sawSerial) {
        warn(`⚠️ ${sheetName}: تبدو ورقة منتج لكن ينقصها عمود ${d.sawCode ? "«السيريال»" : "«الكود»"} — تم تخطّيها`);
      } else {
        log(`تخطّي ورقة (ليست ورقة منتج): ${sheetName}`);
      }
      continue;
    }

    // شفافية: إذا أُخذ الكود من عمود غير «الكود» (لأن عمود «الكود» فارغ) نُعلم المستخدم
    if (header.codeFromFallback) {
      log(`ℹ️ ${sheetName}: عمود «الكود» فارغ — تم أخذ الكود من العمود «${header.codeHeader}»`, "ok");
    }

    const compName = companyOf(sheetName);
    let ccode = companyCode(compName, sheetName);
    let catcode = categoryCode(sheetName);

    // نافذة تأكيد كود الفئة (العمود الرابع)
    if (useConfirm && !approveAll) {
      const res = await askCategory(sheetName, compName, ccode, catcode);
      if (res.skip) { log(`تخطّي الورقة بطلب المستخدم: ${sheetName}`); continue; }
      catcode = res.code;
      if (res.all) approveAll = true;
    }

    if (!ccode) warn(`⚠️ ${sheetName}: شركة غير معروفة (كود الشركة فارغ): ${compName}`);
    if (!catcode) warn(`⚠️ ${sheetName}: كود الفئة فارغ (العمود الرابع)`);

    // قراءة الصفوف مع رصد النواقص
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rows = [];
    let missSerial = 0, missCode = 0;
    for (let r = header.row + 1; r <= range.e.r; r++) {
      const code = normalizeCell(ws, r, header.codeCol);
      const serial = normalizeCell(ws, r, header.serialCol);
      if (!code && !serial) continue;            // صف فارغ تماماً = تجاهل
      if (code && !serial) missSerial++;          // كود بلا سيريال
      else if (serial && !code) missCode++;       // سيريال بلا كود
      rows.push([code, serial, ccode, catcode]); // الكود | السيريال | كود الشركة | كود الفئة
    }

    if (rows.length === 0) { warn(`⚠️ ${sheetName}: لا توجد بيانات في الورقة`); continue; }

    // تنبيه بوجود نواقص داخل هذه الفئة
    if (missSerial > 0 || missCode > 0) {
      const parts = [];
      if (missSerial > 0) parts.push(`${missSerial} كرت بدون سيريال`);
      if (missCode > 0) parts.push(`${missCode} كرت بدون كود`);
      warn(`⚠️ ${sheetName} (فئة ${catcode || "—"}): ${parts.join("، ")}`);
    }

    const company = safeName(compName);

    if (merge) {
      (mergedByCompany[company] = mergedByCompany[company] || []).push(...rows);
      log(`[+] ${sheetName}: ${rows.length} كرت  (فئة ${catcode || "—"})  ->  دمج ${company}`);
    } else {
      outputs.push({ path: `${safeName(base)}/${company}/${safeName(sheetName)}.csv`, rows });
      log(`[OK] ${sheetName}: ${rows.length} كرت  ->  ${company}/${safeName(sheetName)}.csv`);
    }

    totalRows += rows.length;
    summary[company] = (summary[company] || 0) + rows.length;
  }

  // ملف مدمج واحد لكل شركة
  if (merge) {
    for (const company of Object.keys(mergedByCompany)) {
      const compRows = mergedByCompany[company];
      outputs.push({ path: `${safeName(base)}/${company}_مدمج.csv`, rows: compRows });
      log(`[OK] دمج فئات ${company}: ${compRows.length} كرت -> ${company}_مدمج.csv`);
    }
  }

  // الملخص
  log("");
  log("──────── الملخص ────────");
  for (const comp of Object.keys(summary)) log(`   ${comp}: ${summary[comp]} كرت`);
  log(`إجمالي الملفات: ${outputs.length}   |   إجمالي الكروت: ${totalRows}`);

  // ملخص التنبيهات (النواقص)
  if (warnings.length > 0) {
    log(`⚠️ التنبيهات: ${warnings.length} — توجد فئة/ورقة فيها بيانات ناقصة (انظر الأسطر الصفراء أعلاه)`, "warn");
    showWarnBanner(warnings.length);
  } else {
    log("✓ لا توجد نواقص — كل الفئات مكتملة (كود + سيريال).", "ok");
  }

  // إثبات الخصوصية: عدد طلبات الشبكة أثناء المعالجة (يجب أن يكون 0)
  const netReqs = performance.getEntriesByType("resource").filter((e) => e.startTime >= t0);
  log(`🛡️ إثبات: طلبات الشبكة أثناء المعالجة = ${netReqs.length} — لم يُرفع ملفك إلى أي خادم.`,
      netReqs.length === 0 ? "ok" : "warn");

  if (outputs.length === 0) {
    log("لم يُستخرج أي ملف.");
    return;
  }

  // التنزيل: ملف واحد مباشرة، أو ZIP لعدة ملفات
  if (outputs.length === 1) {
    const f = outputs[0];
    const name = f.path.split("/").pop();
    downloadBlob(new Blob([rowsToCSV(f.rows)], { type: "text/csv;charset=utf-8" }), name);
    log(`⬇️ تم تنزيل: ${name}`);
  } else {
    const zip = new JSZip();
    for (const f of outputs) zip.file(f.path, rowsToCSV(f.rows));
    const blob = await zip.generateAsync({ type: "blob" });
    const zipName = `${safeName(base)}.zip`;
    downloadBlob(blob, zipName);
    log(`⬇️ تم تنزيل: ${zipName}`);
  }
}

// ----------------------------- ربط الأحداث -----------------------------
function setFile(file) {
  currentFile = file;
  if (file) {
    fileNameEl.textContent = "📎 " + file.name;
    runBtn.disabled = false;
  } else {
    fileNameEl.textContent = "";
    runBtn.disabled = true;
  }
}

dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => {
  if (e.target.files && e.target.files[0]) setFile(e.target.files[0]);
});

["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("drag"); })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("drag"); })
);
dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files && e.dataTransfer.files[0];
  if (f) setFile(f);
});

clearBtn.addEventListener("click", () => {
  setFile(null);
  fileInput.value = "";
  clearLog();
  hideWarnBanner();
});

runBtn.addEventListener("click", async () => {
  if (!currentFile) return;
  runBtn.disabled = true;
  clearLog();
  log("جاري الاستخراج…");
  try {
    await extract(currentFile);
  } catch (err) {
    log("خطأ: " + (err && err.message ? err.message : err));
    console.error(err);
  } finally {
    runBtn.disabled = false;
  }
});
