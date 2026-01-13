/**
 * BRIGHTEN ESTIMATOR — LIVE PROPOSAL + RAW→TAKEOFF TOTALS
 *
 * What's new:
 * - Takeoff sheet is a totals view generated from Takeoff Raw (keeps Unit overrides).
 * - Proposal is LIVE formulas (no rebuild/clear/reformat during sync).
 * - Build Proposal Template is ONE-TIME. Sync updates only data sheets.
 */

const CFG = {
  SHOP_ADDRESS: "512 S 70th Street Kansas City KS 66111",
  SPEED_MPH: 55,
  
  // Configuration constants
  MAX_PROPOSAL_ROWS: 2000,
  CACHE_DURATION_SECONDS: 43200, // 12 hours

  SHEETS: {
    PROJ: "Project Details",
    DB: "Accessory Item DB",
    TAKEOFF: "Takeoff",
    RAW: "Takeoff Raw",
    RAW_ALT: "TA Takeoff Raw",
    INSTALL: "Install Data",
    PROP: "Proposal"
  },

  // Project Details cells (same as before)
  PD: {
    PROJECT: "B4",
    ADDRESS: "B5",
    PHASED: "B9",
    PHASES: "D9",
    DELIVERY: "B10",
    STAIRS: "B11",
    FLIGHTS: "D11",
    DEMO: "B12",
    MILES: "B14",
    PLANS_DATED: "B16",
    ADDENDUMS: "B17",

    BASE_RATE: "G5",
    DIFF_MULT: "G6",
    EFFECTIVE_RATE: "G7",
    TRAVEL: "G9",

    ADDINS_TOTAL: "G12",

    MAT_MARKUP_PCT: "J13",
    TAX_PCT:        "J14",
    BOND_PCT:       "J15",
    TAX_AMOUNT:     "J23",
    BOND_AMOUNT:    "J24",
    GRAND_TOTAL:    "J26"
  },

  ADDINS: { BODY_START: 21, ROWS: 10 },

  SCOPES: [
    "Toilet Accessories",
    "Grab Bars",
    "Partitions",
    "Fire Equipment",
    "Lockers",
    "Visual Displays",
    "Cubicle Curtains/Track",
    "Signage",
    "Site Furnishings",
    "Wall Protection",
    "General"
  ]
};

/* =========================================================
   MENU
   ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Brighten Estimator")
    .addItem("🧱 Build / Refresh Project Details Form", "buildProjectDetailsForm")
    .addSeparator()
    .addItem("📥 Sync Takeoff Totals (Raw → Takeoff)", "syncTakeoffTotalsFromRaw")
    .addItem("🔄 Sync Install Data (Takeoff → Install Data)", "syncInstallFromTakeoff")
    .addSeparator()
    .addItem("🧾 Build Proposal Template (ONE TIME)", "buildProposalTemplateOnce")
    .addItem("✅ Full Sync (Raw→Takeoff→Install)", "fullSyncDataOnly")
    .addSeparator()
    .addItem("⚠️ Debug Connection", "testConnection")
    .addToUi();
}

function testConnection() {
  SpreadsheetApp.getUi().alert("✅ Script is Connected and Active!");
}

/**
 * Full Sync that DOES NOT touch Proposal formatting.
 * - Updates Takeoff totals from Raw
 * - Updates Install Data from Takeoff + DB
 */
function fullSyncDataOnly() {
  cleanupRawSheetName_();
  ensureTakeoffTemplate_();
  ensureInstallDataSkeleton_();
  syncTakeoffTotalsFromRaw();
  syncInstallFromTakeoff();
  SpreadsheetApp.getUi().alert("✅ Synced Takeoff totals + Install Data. Proposal stays live (no rebuild).");
}

/* =========================================================
   1) TAKEOFF TOTALS VIEW (Raw → Takeoff)
   ========================================================= */

/**
 * Reads Takeoff Raw and writes totals into Takeoff.
 * Preserves per-SKU overrides already entered in Takeoff:
 * - SCOPE (opt)
 * - DESCRIPTION (opt)
 * - UNIT MAT $ (opt)
 * - UNIT MINS (opt)
 */
function syncTakeoffTotalsFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cleanupRawSheetName_();
  ensureTakeoffTemplate_();

  const shTake = ss.getSheetByName(CFG.SHEETS.TAKEOFF);
  const shRaw = ss.getSheetByName(CFG.SHEETS.RAW);
  if (!shRaw) throw new Error(`Missing sheet: ${CFG.SHEETS.RAW}`);

  // Capture existing overrides from Takeoff before rewriting
  const overrideMap = readTakeoffOverrides_(shTake);

  // Aggregate totals from raw
  const totals = aggregateFromRaw_(ss); // [{sku, qty}]
  if (!totals.length) throw new Error("No quantities found in Takeoff Raw.");

  // Write header if needed (keeps your template)
  const header = shTake.getRange(1, 1, 1, 6).getValues()[0].map(x => String(x || "").trim());
  if (!header[0] || header[0].toUpperCase() !== "QTY") ensureTakeoffTemplate_();

  // Build new rows (keeping overrides)
  const rows = totals.map(t => {
    const skuKey = String(t.sku || "").trim();
    const ov = overrideMap[skuKey.toUpperCase()] || {};
    return [
      Number(t.qty || 0),
      skuKey,
      ov.scope || "",       // SCOPE (opt)
      ov.desc || "",        // DESCRIPTION (opt)
      ov.unitMat || "",     // UNIT MAT $ (opt)
      ov.unitMins || ""     // UNIT MINS (opt)
    ];
  });

  // Clear old body only (not header formatting)
  const lastRow = shTake.getLastRow();
  if (lastRow > 1) {
    shTake.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  // Paste new totals
  shTake.getRange(2, 1, rows.length, 6).setValues(rows);

  // Formats
  shTake.getRange(2, 1, rows.length, 1).setNumberFormat("0");
  shTake.getRange(2, 5, rows.length, 1).setNumberFormat("$#,##0.00");
  shTake.getRange(2, 6, rows.length, 1).setNumberFormat("0");

  // Make it obvious it's raw-driven totals
  shTake.getRange("H1").setValue("NOTE: Qty totals are synced from Takeoff Raw. Overrides remain editable here.")
    .setFontColor("#666").setFontSize(9);
}

/**
 * Read current override columns in Takeoff and return map keyed by SKU_KEY
 */
function readTakeoffOverrides_(shTake) {
  const lr = shTake.getLastRow();
  if (lr < 2) return {};
  const data = shTake.getRange(2, 1, lr - 1, 6).getValues();
  const map = {};
  data.forEach(r => {
    const sku = String(r[1] || "").trim();
    if (!sku) return;
    map[sku.toUpperCase()] = {
      scope: String(r[2] || "").trim(),
      desc: String(r[3] || "").trim(),
      unitMat: r[4],
      unitMins: r[5]
    };
  });
  return map;
}

/* =========================================================
   2) INSTALL DATA (Takeoff → Install Data)
   ========================================================= */

function syncInstallFromTakeoff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTakeoffTemplate_();
  ensureInstallDataSkeleton_();

  const shTake = ss.getSheetByName(CFG.SHEETS.TAKEOFF);
  const lr = shTake.getLastRow();
  if (lr < 2) throw new Error("Takeoff has no rows. Sync Takeoff Totals first.");

  const data = shTake.getRange(2, 1, lr - 1, 6).getValues();
  const items = data
    .map(r => ({
      qty: Number(r[0] || 0),
      sku: String(r[1] || "").trim(),
      scopeHint: String(r[2] || "").trim(),
      descOverride: String(r[3] || "").trim(),
      unitMatOverride: Number(r[4] || 0),
      unitMinsOverride: Number(r[5] || 0)
    }))
    .filter(x => x.qty > 0 && x.sku);

  if (!items.length) throw new Error("No valid lines in Takeoff (need qty + code).");

  buildInstallDataLive_(ss, items);
}

/* =========================================================
   3) PROPOSAL — LIVE TEMPLATE (ONE-TIME)
   ========================================================= */

/**
 * Builds Proposal sheet layout ONE time.
 * After this, do NOT rebuild. Proposal reads Install Data live.
 * Sync functions never clear/reformat the Proposal.
 */
function buildProposalTemplateOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreate_(ss, CFG.SHEETS.PROP);
  const already = String(sh.getRange("A1").getValue() || "").includes("BRIGHTEN INSTALLATION");

  // Only build if empty or user explicitly wants rebuild
  if (already && sh.getLastRow() > 10) {
    SpreadsheetApp.getUi().alert("Proposal template already exists. It is LIVE. Use Full Sync (Data Only).");
    return;
  }

  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, 300, 10).breakApart();
  sh.getRange(1, 1, 300, 10).clearFormat();

  // Columns (doc-like)
  sh.setColumnWidth(1, 60);   // A
  sh.setColumnWidth(2, 520);  // B
  sh.setColumnWidth(3, 20);   // C spacer
  sh.setColumnWidth(4, 140);  // D label
  sh.setColumnWidth(5, 160);  // E value
  sh.setColumnWidth(6, 20);   // F spacer

  // Header
  sh.getRange("A1:E2").merge()
    .setValue("BRIGHTEN INSTALLATION — PROPOSAL")
    .setBackground("#1F4E79").setFontColor("white")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 8);

  sh.getRange("A3:E3").merge()
    .setValue("512 S 70th St  |  Kansas City, KS 66111  |  install@brighten.com")
    .setHorizontalAlignment("center")
    .setFontColor("#333");

  // Project info block (live)
  sh.getRange("A5:E5").merge().setValue("PROJECT INFORMATION")
    .setBackground("#E8F0FE").setFontWeight("bold");
  sh.getRange("A6").setValue("Project:").setFontWeight("bold");
  sh.getRange("B6").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.PROJECT}`);

  sh.getRange("A7").setValue("Job Address:").setFontWeight("bold");
  sh.getRange("B7").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.ADDRESS}`).setWrap(true);
  sh.setRowHeight(7, 36);

  sh.getRange("D6").setValue("Date:").setFontWeight("bold");
  sh.getRange("E6").setFormula(`=TEXT(TODAY(),"mm/dd/yyyy")`);
  sh.getRange("D7").setValue("Plans Dated:").setFontWeight("bold");
  sh.getRange("E7").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.PLANS_DATED}`);
  sh.getRange("D8").setValue("Addendums:").setFontWeight("bold");
  sh.getRange("E8").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.ADDENDUMS}`);

  sh.getRange("A6:E8").setBorder(true,true,true,true,true,true);

  // Scope of work title
  sh.getRange("A10:E10").merge().setValue("SCOPE OF WORK")
    .setBackground("#1F4E79").setFontColor("white").setFontWeight("bold");

  // LIVE line-item table (no rebuild)
  // This gives you a normal table that grows/shrinks live.
  // Columns: Scope | Qty | Description
  sh.getRange("A11").setValue("QTY").setFontWeight("bold");
  sh.getRange("B11").setValue("DESCRIPTION").setFontWeight("bold");
  sh.getRange("D11").setValue("SCOPE").setFontWeight("bold");

  // Live formula starts row 12
  // Uses Install Data columns: A=QTY, C=DESC, D=SCOPE
  // Groups by Scope+Description so it behaves like totals and is stable for printing
  sh.getRange("A12").setFormula(
    `=IFERROR(QUERY({'${CFG.SHEETS.INSTALL}'!A3:A,'${CFG.SHEETS.INSTALL}'!C3:C,'${CFG.SHEETS.INSTALL}'!D3:D},
      "select sum(Col1), Col2, Col3 where Col1>0 group by Col2, Col3 label sum(Col1) 'QTY', Col2 'DESCRIPTION', Col3 'SCOPE' ",0),"")`
  );

  // Make the QUERY results land into A:C and we map to A,B,D visually:
  // QUERY returns 3 cols => A=QTY, B=DESC, C=SCOPE
  // We'll display scope also in D with a simple formula that references C
  sh.getRange("D12").setFormula(`=IF(C12="","",C12)`);
  sh.getRange(`D12:D${CFG.MAX_PROPOSAL_ROWS}`).setFormulaR1C1("=IF(RC[-1]=\"\",\"\",RC[-1])"); // copy down

  // Formatting for readability
  sh.getRange("A11:E11").setBorder(false,false,true,false,false,false);
  sh.getRange(`A12:A${CFG.MAX_PROPOSAL_ROWS}`).setHorizontalAlignment("center");
  sh.getRange(`B12:B${CFG.MAX_PROPOSAL_ROWS}`).setWrap(true);
  sh.getRange(`D12:D${CFG.MAX_PROPOSAL_ROWS}`).setFontColor("#555");

  // Work description block (live)
  sh.getRange("A30:E30").merge().setValue("WORK DESCRIPTION / INCLUDED CONDITIONS")
    .setBackground("#1F4E79").setFontColor("white").setFontWeight("bold");

  // Bullet list (live, pulls from Project Details)
  sh.getRange("A31:E37").merge().setFormula(`=WORK_BULLETS()`).setWrap(true);
  sh.setRowHeight(31, 110);

  // Pricing (live, no hours)
  sh.getRange("A39:E39").merge().setValue("PRICING SUMMARY")
    .setBackground("#1F4E79").setFontColor("white").setFontWeight("bold");

  sh.getRange("A40").setValue("Labor:").setFontWeight("bold");
  sh.getRange("B40").setFormula(`=IFERROR(SUM('${CFG.SHEETS.INSTALL}'!H3:H),0)`).setNumberFormat("$#,##0.00");

  sh.getRange("A41").setValue("Grand Total:").setFontWeight("bold");
  sh.getRange("B41").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.GRAND_TOTAL}`)
    .setNumberFormat("$#,##0.00").setFontWeight("bold");

  // Print settings
  try {
    const ps = sh.getPageSetup();
    ps.setOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
    ps.setFitToWidth(1);
    ps.setFitToHeight(0);
    ps.setTopMargin(0.5);
    ps.setBottomMargin(0.5);
    ps.setLeftMargin(0.5);
    ps.setRightMargin(0.5);
  } catch (e) {}

  SpreadsheetApp.getUi().alert("✅ Proposal Template built. It's LIVE now. Use Full Sync (Data Only) to update totals.");
}

/**
 * Custom function used by Proposal to produce bullet text live.
 * Returns a single text block with line breaks.
 */
function WORK_BULLETS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proj = ss.getSheetByName(CFG.SHEETS.PROJ);
  if (!proj) return "";

  const phased = String(proj.getRange(CFG.PD.PHASED).getValue() || "No");
  const phases = Number(proj.getRange(CFG.PD.PHASES).getValue() || 1);
  const stairs = String(proj.getRange(CFG.PD.STAIRS).getValue() || "No");
  const flights = Number(proj.getRange(CFG.PD.FLIGHTS).getValue() || 0);
  const demo = String(proj.getRange(CFG.PD.DEMO).getValue() || "No");
  const delivery = String(proj.getRange(CFG.PD.DELIVERY).getValue() || "Site");
  const miles = proj.getRange(CFG.PD.MILES).getDisplayValue();

  const out = [];
  out.push("• Furnish and install specified accessories per plans and site conditions.");
  out.push(`• Delivery assumed to: ${delivery}.`);
  if (miles) out.push(`• Travel calculated from Brighten shop to site: ${miles} miles (one-way).`);
  if (demo === "Yes") out.push("• Includes demo/retrofit conditions as required for installation.");
  if (stairs === "Yes") out.push(`• Includes material handling up stairs (${flights || 0} flights).`);
  if (phased === "Yes") out.push(`• Includes phased workflow (${Math.max(1, phases)} phase(s)).`);

  const addIns = getCheckedAddIns_(proj);
  if (addIns.length) {
    out.push("• Add-ins included in pricing:");
    addIns.forEach(a => out.push(`   - ${a}`));
  }
  return out.join("\n");
}

/* =========================================================
   PROJECT DETAILS FORM + INSTALL SKELETON + DB MATCHING
   (Same "working" pieces you already had; left minimal)
   ========================================================= */

function buildProjectDetailsForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreate_(ss, CFG.SHEETS.PROJ);
  
  // This is a placeholder function that can be expanded to build a complete Project Details form
  // For now, it ensures the sheet exists
  SpreadsheetApp.getUi().alert(
    "Project Details Form\n\n" +
    "This function can be customized to build your Project Details form layout.\n" +
    "Currently, the sheet is created and available for manual data entry."
  );
}

function ensureInstallDataSkeleton_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreate_(ss, CFG.SHEETS.INSTALL);
  if (sh.getLastRow() > 0) return;

  sh.clear();
  sh.setHiddenGridlines(true);

  sh.getRange("A1:K1").merge()
    .setBackground("#FFF2CC")
    .setFontWeight("bold")
    .setValue("UPDATED RATE FROM PROJECT DETAILS:");

  sh.getRange("K1").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.EFFECTIVE_RATE}`)
    .setNumberFormat("$0.00")
    .setFontWeight("bold");

  sh.getRange("A2:J2").setValues([[
    "QTY","SKU_KEY","DESCRIPTION","SCOPE","UNIT MINS","TOTAL MINS","EFF RATE","LABOR $","UNIT MAT $","MAT $"
  ]]).setBackground("#1F4E79").setFontColor("#fff").setFontWeight("bold");

  sh.setFrozenRows(2);
}

function ensureTakeoffTemplate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreate_(ss, CFG.SHEETS.TAKEOFF);

  // Always ensure header exists
  sh.getRange("A1:F1").setValues([[
    "QTY", "CODE / SKU", "SCOPE (opt)", "DESCRIPTION (opt)", "UNIT MAT $ (opt)", "UNIT MINS (opt)"
  ]]).setBackground("#1F4E79").setFontColor("#fff").setFontWeight("bold");

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 420);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);

  sh.getRange("E2:E500").setNumberFormat("$#,##0.00");
  sh.getRange("F2:F500").setNumberFormat("0");
}

function buildInstallDataLive_(ss, items) {
  const sh = getOrCreate_(ss, CFG.SHEETS.INSTALL);
  sh.clear();
  sh.setHiddenGridlines(true);

  sh.getRange("A1:K1").merge()
    .setBackground("#FFF2CC")
    .setFontWeight("bold")
    .setValue("UPDATED RATE FROM PROJECT DETAILS:");

  sh.getRange("K1").setFormula(`='${CFG.SHEETS.PROJ}'!${CFG.PD.EFFECTIVE_RATE}`)
    .setNumberFormat("$0.00")
    .setFontWeight("bold");

  sh.getRange("A2:J2").setValues([[
    "QTY","SKU_KEY","DESCRIPTION","SCOPE","UNIT MINS","TOTAL MINS","EFF RATE","LABOR $","UNIT MAT $","MAT $"
  ]]).setBackground("#1F4E79").setFontColor("#fff").setFontWeight("bold");

  const effRateCell = `'${CFG.SHEETS.PROJ}'!${CFG.PD.EFFECTIVE_RATE}`;
  const db = loadDb_(ss);

  const base = items.map(it => {
    const skuKey = normalizeTakeoffCode_(it.sku).skuKey;
    const qty = Number(it.qty || 0);

    const d = matchDb_(db, skuKey);
    const scope = (it.scopeHint || d.category || "General").trim();
    const desc  = (it.descOverride || d.desc || skuKey).trim();

    const mins = (Number(it.unitMinsOverride) > 0)
      ? Number(it.unitMinsOverride)
      : ((Number(d.mins) > 0) ? Number(d.mins) : defaultMins_(scope, skuKey, desc));

    const unitMat = (Number(it.unitMatOverride) > 0)
      ? Number(it.unitMatOverride)
      : Number(d.cost || 0);

    return [qty, skuKey, desc, scope, mins, "", "", "", unitMat, ""];
  });

  sh.getRange(3, 1, base.length, 10).setValues(base);

  for (let i = 0; i < base.length; i++) {
    const r = 3 + i;
    sh.getRange(r, 6).setFormula(`=A${r}*E${r}`);
    sh.getRange(r, 7).setFormula(`=${effRateCell}`);
    sh.getRange(r, 8).setFormula(`=(F${r}/60)*G${r}`);
    sh.getRange(r,10).setFormula(`=A${r}*I${r}`);
  }

  sh.getRange(3, 5, base.length, 2).setNumberFormat("0");
  sh.getRange(3, 7, base.length, 2).setNumberFormat("$#,##0.00");
  sh.getRange(3, 9, base.length, 2).setNumberFormat("$#,##0.00");

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 440);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 95);
  sh.setColumnWidth(7, 95);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 110);
  sh.setColumnWidth(10, 110);
}

/* =========================================================
   RAW PARSING (supports row-style OR matrix-style)
   ========================================================= */

function cleanupRawSheetName_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alt = ss.getSheetByName(CFG.SHEETS.RAW_ALT);
  const raw = ss.getSheetByName(CFG.SHEETS.RAW);
  if (!raw && alt) alt.setName(CFG.SHEETS.RAW);
  if (!ss.getSheetByName(CFG.SHEETS.RAW)) getOrCreate_(ss, CFG.SHEETS.RAW);
}

function aggregateFromRaw_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.RAW);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const header = data[0].map(x => String(x || "").toUpperCase());
  const qtyCol = header.findIndex(h => h.includes("QTY"));
  const codeCol = header.findIndex(h => h.includes("CODE") || h.includes("SKU") || h.includes("ITEM"));

  // Row-style raw
  if (qtyCol > -1 && codeCol > -1) {
    const totals = new Map();
    for (let i = 1; i < data.length; i++) {
      const qty = Number(data[i][qtyCol] || 0);
      const raw = String(data[i][codeCol] || "").trim();
      if (!raw || qty <= 0) continue;
      const skuKey = normalizeTakeoffCode_(raw).skuKey;
      totals.set(skuKey, (totals.get(skuKey) || 0) + qty);
    }
    return Array.from(totals, ([sku, qty]) => ({ sku, qty }))
      .sort((a,b)=>String(a.sku).localeCompare(String(b.sku)));
  }

  // Matrix-style raw (headers are codes, body are quantities)
  const headers = data[0];
  const codeCols = [];
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").trim();
    if (!h) continue;
    // skip obvious non-code headers
    const up = h.toUpperCase();
    if (up.includes("JOB") || up.includes("TOTAL") || up.includes("ROOM") || up.includes("LEVEL")) continue;
    const norm = normalizeTakeoffCode_(h);
    codeCols.push({ c, sku: norm.skuKey });
  }

  const totals = new Map();
  for (let r = 1; r < data.length; r++) {
    codeCols.forEach(col => {
      const q = Number(data[r][col.c] || 0);
      if (q > 0) totals.set(col.sku, (totals.get(col.sku) || 0) + q);
    });
  }

  return Array.from(totals, ([sku, qty]) => ({ sku, qty }))
    .sort((a,b)=>String(a.sku).localeCompare(String(b.sku)));
}

/* =========================================================
   NORMALIZATION: "GB 36" / "B1234" / "1234"
   ========================================================= */

function normalizeTakeoffCode_(raw) {
  const s = String(raw || "").trim();
  const up = s.toUpperCase().replace(/\s+/g, " ");

  // Manufacturer letter + 4 digits e.g. B1234
  const mfrDigits = up.match(/\b([A-Z])\s*[- ]?\s*([0-9]{4})\b/);

  // Grab bar length e.g. "GB 36" - preserve leading zeros
  const gbLen = up.match(/\bGB\b\s*([0-9]{2})\b/) || up.match(/\bGB([0-9]{2})\b/);

  if (up.includes("GB")) {
    if (mfrDigits) return { skuKey: `${mfrDigits[1]}${mfrDigits[2]}` };
    if (gbLen) return { skuKey: `GB${gbLen[1]}` }; // Keep as string to preserve leading zeros
  }
  if (mfrDigits) return { skuKey: `${mfrDigits[1]}${mfrDigits[2]}` };

  const digits = up.match(/\b([0-9]{4})\b/);
  if (digits) return { skuKey: digits[1] };

  return { skuKey: s };
}

/* =========================================================
   DB MATCH
   ========================================================= */

function loadDb_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.DB);
  if (!sh || sh.getLastRow() < 2) return {};
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  const map = {};
  values.forEach(r => {
    const sku = String(r[0] || "").trim();
    if (!sku) return;
    map[String(sku).toUpperCase()] = {
      sku,
      category: String(r[2] || "").trim(),
      desc: String(r[3] || "").trim(),
      cost: Number(r[5] || 0),
      mins: Number(r[6] || 0)
    };
  });
  return map;
}

function matchDb_(db, skuKey) {
  const key = String(skuKey || "").toUpperCase().trim();
  if (db[key]) return db[key];
  const digits = key.match(/\b([0-9]{4})\b/);
  if (digits && db[digits[1]]) return db[digits[1]];
  return {};
}

function defaultMins_(scope, sku, desc) {
  const s = String(scope || "").toUpperCase();
  const d = String(desc || "").toUpperCase();
  const k = String(sku || "").toUpperCase();
  if (s.includes("GRAB") || k.startsWith("GB") || d.includes("GRAB BAR")) return 20;
  if (d.includes("MIRROR")) return 30;
  if (d.includes("DISPENSER")) return 20;
  if (d.includes("DRYER")) return 45;
  if (d.includes("PARTITION")) return 60;
  return 15;
}

/* =========================================================
   ADD-INS helper (for Proposal bullets)
   ========================================================= */

function getCheckedAddIns_(projSheet) {
  const body = CFG.ADDINS.BODY_START;
  const rows = CFG.ADDINS.ROWS;
  if (projSheet.getLastRow() < body) return [];
  const range = projSheet.getRange(body, 1, rows, 6).getValues();
  const out = [];
  range.forEach(r => {
    const apply = r[0] === true;
    const name = String(r[1] || "").trim();
    const type = String(r[2] || "").trim();
    const val = r[3];
    const rule = String(r[4] || "").trim();
    if (!apply || !name) return;
    if (String(type).toUpperCase() === "MULT") out.push(`${name} (${Math.round(Number(val)*100)}% of labor) — ${rule||"Always"}`);
    else out.push(`${name} ($${Number(val).toFixed(0)}) — ${rule||"Always"}`);
  });
  return out;
}

/* =========================================================
   HELPERS
   ========================================================= */

function getOrCreate_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/* =========================================================
   CUSTOM FUNCTION: SHOP_MILES(address)
   ========================================================= */

function SHOP_MILES(destinationAddress) {
  const dest = String(destinationAddress || "").trim();
  if (!dest) return "";

  const cache = CacheService.getScriptCache();
  const key = "SHOP_MILES|" + dest.toUpperCase();
  const cached = cache.get(key);
  if (cached) return Number(cached);

  const dir = Maps.newDirectionFinder()
    .setOrigin(CFG.SHOP_ADDRESS)
    .setDestination(dest)
    .setMode(Maps.DirectionFinder.Mode.DRIVING);

  const res = dir.getDirections();
  const routes = (res && res.routes) ? res.routes : [];
  if (!routes.length) return "";

  const legs = routes[0].legs || [];
  if (!legs.length) return "";

  const meters = legs[0].distance && legs[0].distance.value ? Number(legs[0].distance.value) : 0;
  const miles = meters ? (meters / 1609.344) : "";
  if (miles !== "") cache.put(key, String(miles), CFG.CACHE_DURATION_SECONDS);

  return miles === "" ? "" : Math.round(miles * 10) / 10;
}