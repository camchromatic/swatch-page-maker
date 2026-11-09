// Swatch Page Maker: builds an .xlsx swatch sheet in the browser (no server).
// Groups are vertical strips (a code cell + a blank box per swatch),
// laid out "across" x "down" on each printed page.

// Column widths in Excel units, copied from the original "paint markers 240" sheet.
const WIDTH = { margin: 1.7109375, cell: 9.140625, gap: 1.5703125 };
const ROW_PT = 15;
const GAP_ROWS = 2;
const TITLE_PT = 24;
const DESC_LINE_PT = 13;
const SPACER_PT = 8;
const PAPER = { a4: { code: 9, w: 595, h: 842 }, letter: { code: 1, w: 612, h: 792 } };
const MARGIN_IN = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
const MIN_SCALE = 50; // layouts that would print smaller than this are not offered
const MAX_PAGES = 10; // no file is longer than this
// Excel, page and preview font. Arial matches Helvetica (the PDF's built-in font) letter for letter,
// and comes with every computer and with Google Sheets.
const FONT = "Arial";
// Swatch spacing: a gap (pt) between swatches in a group, and a wider gap
// column (Excel width units) between groups so groups still read as groups.
const SPACING = {
  none: { row: 0, gap: WIDTH.gap },
  small: { row: 4, gap: 3 },
  large: { row: 8, gap: 4.5 },
};

const $ = (id) => document.getElementById(id);

let chosenKey = null; // "perGroup-across-down" of the picked layout

function clampInt(v, min, max, fallback) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
}

const presetById = (id) => PRESETS.find((p) => p.id === id) || null;

// Whole lines ("All ..." sets) are always in number order: the "By color" order and
// the Order menu were removed. Only for lines whose codes (nearly)
// all hold a number; the few without one (blenders, "FV", "XB") go last. Copic
// style: 0000 before 000 before 00 before 0.
function codeCompare(a, b) {
  const ta = String(a).split(/(\d+)/), tb = String(b).split(/(\d+)/);
  for (let i = 0; i < Math.max(ta.length, tb.length); i++) {
    const x = ta[i] ?? "", y = tb[i] ?? "";
    if (x === y) continue;
    if (i % 2 && x && y) return (+x - +y) || (y.length - x.length);
    return x < y ? -1 : 1;
  }
  return 0;
}
function inOrder(p) {
  if (!p || p.collection || !p.name.startsWith("All ")) return p;
  const hasNum = (c) => /\d/.test(c);
  if (p.codes.filter(hasNum).length < p.codes.length * 0.9) return p; // names, not codes
  const order = p.codes.map((_, i) => i).sort((i, j) => (hasNum(p.codes[j]) - hasNum(p.codes[i])) ||
    (hasNum(p.codes[i]) ? codeCompare(p.codes[i], p.codes[j]) : i - j));
  return { ...p, codes: order.map((i) => p.codes[i]), legacy: p.legacy && order.map((i) => p.legacy[i]) };
}

// The set picked in the Set menu (whole lines in number order). A collection stands
// for its member sets; those are kept whole (keepAll) and headed by their short name
// ("Pitt Artist Pen Dual Marker - Wallet of 5" -> "Wallet of 5").
function chosenPresets() {
  const p = presetById($("preset").value);
  if (!p) return [];
  if (!p.collection) return [inOrder(p)];
  return p.collection.map(presetById).filter(Boolean)
    .map((m) => ({ ...m, keepAll: true, title: m.name.split(" - ").pop() }));
}
const hasCollection = () => !!presetById($("preset").value)?.collection;

// Each chosen set's codes, skipping colours an earlier set already has (legacy codes
// where a set has them). Repeats only count within one brand: Copic "Y08" and
// Ohuhu "Y08" are different colours. Sets from a collection stay complete.
// Sets left empty are dropped.
function mergedSegments(list, legacy) {
  const seen = new Set();
  const segments = [];
  for (const p of list) {
    const seg = [];
    for (const code of legacy && p.legacy ? p.legacy : p.codes) {
      const key = brandOf(p) + "|" + code;
      if (p.keepAll || !seen.has(key)) { seen.add(key); seg.push(code); }
    }
    seg.name = p.title || p.name; // shown in the header of the set's first group
    if (seg.length) segments.push(seg);
  }
  return segments;
}
const mergedCodes = (list, legacy) => mergedSegments(list, legacy).flat();

// Merged sets each start a new group: pad every set to a whole number of groups.
function padCodes(segments, perGroup) {
  const out = [];
  for (const seg of segments) {
    out.push(...seg);
    while (out.length % perGroup) out.push("");
  }
  return out;
}

// The codes as laid out for this layout (merged sets padded to group starts),
// plus each group's header text: the set's name on its first group.
// s.slotUsed marks which swatch slots hold a real swatch (the rest are spare blanks).
function withLayoutCodes(s, layout) {
  s.groupNames = [];
  if (!s.segments) {
    s.slotUsed = (i) => i < s.total;
    return s;
  }
  s.codes = padCodes(s.segments, layout.perGroup);
  const used = [];
  let group = 0;
  for (const seg of s.segments) {
    s.groupNames[group] = seg.name;
    for (let i = 0; i < seg.length; i++) used[group * layout.perGroup + i] = true;
    group += Math.ceil(seg.length / layout.perGroup);
  }
  s.slotUsed = (i) => !!used[i];
  return s;
}

// How many swatches to draw in a group starting at slot `first`: all of them, or,
// with spare blanks turned off, only up to the group's last real swatch.
function swatchesToDraw(s, first, perGroup) {
  if (!s.hideSpares || !s.slotUsed) return perGroup;
  let k = perGroup;
  while (k > 0 && !s.slotUsed(first + k - 1)) k--;
  return k;
}

// Screen colors for the chosen set (code -> hex), if its line has them.
// Colors are keyed by each set's main codes; legacy codes borrow them through the set's lists.
function previewColors(list, legacy) {
  if (!list.length || typeof PREVIEW_COLORS === "undefined") return null;
  const entry = PREVIEW_COLORS.find((e) => list.every((p) => e.lines.includes(p.group)));
  if (!entry) return null;
  if (!legacy || !list.some((p) => p.legacy)) return entry.colors;
  const map = {};
  for (const p of list) (p.legacy || p.codes).forEach((code, i) => { if (entry.colors[p.codes[i]]) map[code] = entry.colors[p.codes[i]]; });
  return map;
}

function readSettings() {
  const list = chosenPresets();
  const blank = $("no-codes").checked;
  const named = hasCollection(); // a collection names its groups, even when there is only one
  const segments = list.length > 1 || named
    ? mergedSegments(list, $("legacy-codes").checked).map((seg) => (blank ? Object.assign(seg.map(() => ""), { name: seg.name }) : seg))
    : null;
  const s = {
    segments: segments && (segments.length > 1 || named) ? segments : null,
    codes: list.length && !blank ? mergedCodes(list, $("legacy-codes").checked) : [],
    total: clampInt($("total").value, 1, 2000, 240),
    pages: clampInt($("pages").value, 1, MAX_PAGES, 1),
    paper: $("paper").value,
    orientation: $("orientation").value,
    align: ["left", "center", "right"].includes($("align").value) ? $("align").value : "center",
    headers: $("headers").checked,
    hideSpares: !$("spares").checked,
    spacing: SPACING[$("spacing").value] ? $("spacing").value : "none",
    sections: clampInt($("sections").value, 1, 5, 1),
    pageType: pageTypeOf($("page-type").value),
    stripColours: clampInt($("strip-colours").value, 3, 6, 5),
    palettes: clampInt($("palettes").value, 1, 500, 24),
    mixColours: clampInt($("mix-colours").value, 3, 12, 8),
    labelBold: $("label-bold").checked,
    labelItalic: $("label-italic").checked,
    labelSize: [8, 10, 11].includes(+$("label-size").value) ? +$("label-size").value : 10,
    labelColor: /^#[0-9a-f]{6}$/i.test($("label-color").value) ? $("label-color").value : "#808080",
    title: $("title").value.trim(),
    description: $("description").value.trim(),
    colorOf: blank ? null : previewColors(list, $("legacy-codes").checked),
  };
  s.labelUnits = labelUnits(s);
  s.autoPages = s.pageType === "grid" && $("auto-pages").checked;
  if (s.autoPages) s.pages = autoPageCount(s);
  return s;
}

// Auto pages: the fewest pages where the swatches still print at a comfortable
// size; if nothing reaches that, the fewest pages with any layout at all.
const AUTO_MIN_SCALE = 70;
function autoPageCount(s) {
  let firstFit = 0;
  for (let p = 1; p <= MAX_PAGES; p++) {
    const opts = layoutOptions({ ...s, pages: p });
    if (!opts.length) continue;
    if (!firstFit) firstFit = p;
    if (opts[0].scale >= AUTO_MIN_SCALE) return p;
  }
  return firstFit || 1;
}

// Excel stores widths in character units; one unit is about 7 px at 100%.
const widthPt = (w) => Math.round(w * 7) * 0.75;

// Sections: one box keeps the full cell width; a split box gets a narrower
// column per section, so each part stays big enough to colour in.
const SECTION_W = 5.5;
const sectionUnits = (s) => (s.sections > 1 ? SECTION_W : WIDTH.cell);

// Label column: the usual width, widened for sets labelled by color name
// (e.g. "Permanent Alizarin Crimson") so the text stays readable. Arial is
// about 5.5pt per character at 10pt; one width unit is about 5.25pt.
const LABEL_MAX = 20;
function labelUnits(s) {
  const all = s.segments ? s.segments.flat() : s.codes;
  const longest = all.reduce((m, c) => Math.max(m, String(c || "").length), 0);
  const need = (longest * 5.5 * (s.labelSize / 10)) / 5.25 + 1.5;
  return Math.min(LABEL_MAX, Math.max(WIDTH.cell, need));
}

function contentWidthPt(layout, s) {
  return widthPt(WIDTH.margin) + layout.across * (widthPt(s.labelUnits || WIDTH.cell) + s.sections * widthPt(sectionUnits(s)))
    + (layout.across - 1) * widthPt(SPACING[s.spacing].gap);
}

// Height of one row of groups: optional header, then swatches with gaps between.
function bandHeightPt(layout, s) {
  const sp = SPACING[s.spacing].row;
  return (s.headers ? ROW_PT + sp : 0) + layout.perGroup * ROW_PT + (layout.perGroup - 1) * sp;
}

// Lines the description wraps to, at about 5.5pt per character (Arial 10).
function descLines(text, contentW) {
  if (!text) return 0;
  const perLine = Math.max(10, Math.floor((contentW - widthPt(WIDTH.margin)) / 5.5));
  return text.split(/\r?\n/).reduce((n, line) => n + Math.max(1, Math.ceil(line.length / perLine)), 0);
}

// Rows at the top of every page: title, description, then a spacer.
function headRowsFor(contentW, s) {
  const rows = [];
  if (s.title) rows.push({ kind: "title", height: TITLE_PT });
  if (s.description) rows.push({ kind: "desc", height: descLines(s.description, contentW) * DESC_LINE_PT + 2 });
  if (rows.length) rows.push({ kind: "spacer", height: SPACER_PT });
  return rows;
}
const headRows = (layout, s) => headRowsFor(contentWidthPt(layout, s), s);

// Printable area of the paper in points.
function usableSize(s) {
  const paper = PAPER[s.paper];
  const [pw, ph] = s.orientation === "landscape" ? [paper.h, paper.w] : [paper.w, paper.h];
  return { w: pw - (MARGIN_IN.left + MARGIN_IN.right) * 72, h: ph - (MARGIN_IN.top + MARGIN_IN.bottom) * 72 };
}

// Left edge of the swatches (PDF/preview) for the chosen alignment. w is the
// drawn width; marginW is Excel's blank first column, kept on the left so the
// PDF lines up with the Excel print.
function alignX(s, pw, w, marginW = 0) {
  if (s.align === "left") return MARGIN_IN.left * 72 + marginW;
  if (s.align === "right") return pw - MARGIN_IN.right * 72 - w;
  return (pw - w) / 2;
}

// Excel page placement for the chosen alignment. Excel can only center or
// start at the left margin, so "right" widens the left margin by the spare
// width. contentW/H are the printed size at 100%; fixedScale is set for
// multi-page sheets.
function excelPlacement(s, contentW, contentH, fixedScale) {
  if (s.align === "center") return { horizontalCentered: true, margins: MARGIN_IN };
  if (s.align === "left") return { horizontalCentered: false, margins: MARGIN_IN };
  const { w, h } = usableSize(s);
  // Excel prints about 8% wider than our sums (see printScale): count that in,
  // so the swatches end at the right margin or just short of it, never past it.
  const printedW = contentW / 0.92;
  const k = fixedScale ? fixedScale / 100 : Math.min(1, w / printedW, h / contentH);
  const spareIn = Math.max(0, (w - printedW * k) / 72);
  return { horizontalCentered: false, margins: { ...MARGIN_IN, left: MARGIN_IN.left + spareIn } };
}

// Print scale (%) so one page of groups fits on the paper.
function printScale(layout, s) {
  const { w: usableW, h: usableH } = usableSize(s);
  const headH = headRows(layout, s).reduce((h, r) => h + r.height, 0);
  const contentH = headH + layout.down * bandHeightPt(layout, s) + (layout.down - 1) * GAP_ROWS * ROW_PT;
  // Printers lose some width to fonts and hard margins: tested in Excel, about 8%.
  const fit = Math.min((usableW / contentWidthPt(layout, s)) * 0.92, usableH / contentH) * 100;
  return Math.max(10, Math.min(100, Math.floor(fit) - 2)); // small safety margin
}

// All ways to split the swatches over the pages: groups of perGroup,
// across x down groups per page, with less than one group left blank.
function layoutOptions(s) {
  const perPage = Math.ceil(s.total / s.pages);
  const found = [];
  // Groups hold 4-30 swatches; fewer than 4 swatches make one group of exactly that many.
  for (let perGroup = Math.max(1, Math.min(4, perPage)); perGroup <= 30; perGroup++) {
    // Merged sets each start a new group, so count groups set by set.
    const needed = s.segments ? s.segments.reduce((t, seg) => t + Math.ceil(seg.length / perGroup), 0) : 0;
    for (let across = 1; across <= 10; across++) {
      const groups = s.segments ? Math.ceil(needed / s.pages) : Math.ceil(perPage / perGroup);
      if (groups % across !== 0) continue; // every row of groups is full
      const down = groups / across;
      if (down < 1 || down > 10) continue;
      const blanks = perGroup * across * down * s.pages - s.total;
      if (!s.segments && blanks >= perGroup) continue;
      if (s.segments && across * down * s.pages - needed >= across) continue; // at most one row of spare groups
      const layout = { perGroup, across, down };
      const scale = printScale(layout, s);
      if (scale < MIN_SCALE) continue;
      found.push({ ...layout, scale, blanks, key: `${perGroup}-${across}-${down}` });
    }
  }
  const exact = found.filter((o) => o.blanks === 0);
  const list = exact.length && !s.segments ? exact : found;
  // Grouped sets: layouts that keep every set in one group come first.
  const longest = s.segments ? Math.max(...s.segments.map((seg) => seg.length)) : Infinity;
  const whole = (o) => (o.perGroup >= longest ? 0 : 1);
  // Then best print size; on a tie, more (smaller) groups first.
  list.sort((a, b) => whole(a) - whole(b) || b.scale - a.scale || a.blanks - b.blanks || b.across * b.down - a.across * a.down);
  return list.slice(0, 8);
}

const thin = { style: "thin", color: { argb: "FF000000" } };
const medium = { style: "medium", color: { argb: "FF000000" } };

async function buildWorkbook(layout, s) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Swatches");

  const cols = [{ width: WIDTH.margin }];
  for (let g = 0; g < layout.across; g++) {
    cols.push({ width: s.labelUnits || WIDTH.cell });
    for (let j = 0; j < s.sections; j++) cols.push({ width: sectionUnits(s) });
    if (g < layout.across - 1) cols.push({ width: SPACING[s.spacing].gap });
  }
  ws.columns = cols;
  const sp = SPACING[s.spacing].row;
  const step = sp ? 2 : 1; // with spacing, a short spacer row follows each swatch row
  const lastCol = cols.length;

  // Title and description: Excel repeats these rows on every printed page.
  const head = headRows(layout, s);
  head.forEach((row, i) => {
    const r = i + 1;
    ws.getRow(r).height = row.height;
    if (row.kind === "spacer") return;
    ws.mergeCells(r, 2, r, lastCol);
    const cell = ws.getCell(r, 2);
    if (row.kind === "title") {
      cell.value = s.title;
      cell.font = { name: FONT, size: 16, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle", shrinkToFit: true };
    } else {
      cell.value = s.description;
      cell.font = { name: FONT, size: 10, color: { argb: "FF555555" } };
      cell.alignment = { horizontal: "center", vertical: "top", wrapText: true };
    }
  });

  let r = head.length + 1;
  for (let p = 0; p < s.pages; p++) {
    for (let b = 0; b < layout.down; b++) {
      for (let g = 0; g < layout.across; g++) {
        const codeCol = 2 + g * (s.sections + 2); // code, sections, gap
        const first = (((p * layout.down) + b) * layout.across + g) * layout.perGroup;
        const count = swatchesToDraw(s, first, layout.perGroup);
        if (!count) continue; // nothing but spare blanks: leave the space empty
        let top = r;
        if (s.headers) {
          ws.mergeCells(r, codeCol, r, codeCol + s.sections);
          const h = ws.getCell(r, codeCol);
          const name = (s.groupNames || [])[((p * layout.down) + b) * layout.across + g];
          if (name) h.value = name;
          h.font = { name: FONT, size: 10, bold: true };
          h.alignment = { horizontal: "center", vertical: "middle", shrinkToFit: true };
          h.border = { top: medium, left: medium, bottom: medium, right: medium };
          top += step;
        }
        for (let i = 0; i < count; i++) {
          const code = ws.getCell(top + i * step, codeCol);
          if (s.codes[first + i]) code.value = s.codes[first + i];
          code.font = {
            name: FONT, size: s.labelSize, bold: s.labelBold, italic: s.labelItalic,
            color: { argb: "FF" + s.labelColor.slice(1).toUpperCase() },
          };
          code.alignment = { horizontal: "center", shrinkToFit: true };
          code.border = { left: thin, top: thin, bottom: thin };
          // Touching swatches share one thick outline; spaced swatches get their own.
          // Sections inside a box are split by thin lines.
          for (let j = 1; j <= s.sections; j++) {
            ws.getCell(top + i * step, codeCol + j).border = {
              left: j === 1 ? medium : thin,
              right: j === s.sections ? medium : thin,
              top: sp || i === 0 ? medium : thin,
              bottom: sp || i === count - 1 ? medium : thin,
            };
          }
        }
      }
      const bandRows = (layout.perGroup + (s.headers ? 1 : 0)) * step - (step - 1);
      for (let i = 0; i < bandRows; i++) ws.getRow(r + i).height = sp && i % 2 ? sp : ROW_PT;
      r += bandRows;
      if (b < layout.down - 1) {
        for (let i = 0; i < GAP_ROWS; i++) ws.getRow(r + i).height = ROW_PT;
        r += GAP_ROWS;
      }
    }
    if (p < s.pages - 1) ws.getRow(r - 1).addPageBreak();
  }

  // One page: let Excel fit it (like the original sheet). Several pages: Excel ignores
  // page breaks when fitting, so use a fixed scale instead.
  const fitting = s.pages === 1
    ? { fitToPage: true, fitToWidth: 1, fitToHeight: 1 }
    : { scale: printScale(layout, s) };
  const contentH = head.reduce((h, r) => h + r.height, 0)
    + layout.down * bandHeightPt(layout, s) + (layout.down - 1) * GAP_ROWS * ROW_PT;
  ws.pageSetup = {
    paperSize: PAPER[s.paper].code,
    orientation: s.orientation,
    ...excelPlacement(s, contentWidthPt(layout, s), contentH, fitting.scale),
    ...(head.length ? { printTitlesRow: `1:${head.length}` } : {}),
    ...fitting,
  };

  return wb;
}

// ---- Drawing targets ------------------------------------------------------------
// buildPdf and buildModelPdf draw through a small set of jsPDF calls. The preview
// passes SvgDoc instead, which draws the same calls as SVG pages the size of the
// paper, so the preview is exactly what prints (white space included).

const pdfDoc = (s) => new window.jspdf.jsPDF({ unit: "pt", format: s.paper, orientation: s.orientation });

const NS = "http://www.w3.org/2000/svg";
const measureCtx = document.createElement("canvas").getContext("2d");

class SvgDoc {
  constructor(s) {
    const paper = PAPER[s.paper];
    [this.w, this.h] = s.orientation === "landscape" ? [paper.h, paper.w] : [paper.w, paper.h];
    this.pages = [];
    this.font = { style: "normal", size: 10 };
    this.textColor = "#000";
    this.lineWidth = 1;
    this.internal = { pageSize: { getWidth: () => this.w, getHeight: () => this.h } };
    this.addPage();
  }
  addPage() {
    const svg = document.createElementNS(NS, "svg");
    svg.setAttribute("viewBox", `0 0 ${this.w} ${this.h}`);
    svg.setAttribute("class", "paper");
    const bg = document.createElementNS(NS, "rect");
    Object.entries({ width: this.w, height: this.h, fill: "#fff" }).forEach(([a, v]) => bg.setAttribute(a, v));
    svg.append(bg);
    this.pages.push(svg);
  }
  get page() { return this.pages[this.pages.length - 1]; }
  setFont(_family, style = "normal") { this.font.style = style; }
  setFontSize(size) { this.font.size = size; }
  setDrawColor() {}
  setLineWidth(w) { this.lineWidth = w; }
  setTextColor(c) { this.textColor = typeof c === "number" ? `rgb(${c},${c},${c})` : c; }
  cssFont() {
    const st = this.font.style;
    return `${st.includes("italic") ? "italic " : ""}${st.includes("bold") ? "bold " : ""}${this.font.size}px Arial, Helvetica, sans-serif`;
  }
  getTextWidth(text) {
    measureCtx.font = this.cssFont();
    return measureCtx.measureText(text).width;
  }
  splitTextToSize(text, width) {
    const lines = [];
    for (const para of String(text).split(/\r?\n/)) {
      let line = "";
      for (const word of para.split(/\s+/)) {
        const next = line ? line + " " + word : word;
        if (line && this.getTextWidth(next) > width) { lines.push(line); line = word; } else line = next;
      }
      lines.push(line);
    }
    return lines;
  }
  fillBox(x, y, w, h, color) {
    const el = document.createElementNS(NS, "rect");
    Object.entries({ x, y, width: w, height: h, fill: color }).forEach(([a, v]) => el.setAttribute(a, v));
    this.page.append(el);
  }
  line(x1, y1, x2, y2) {
    const el = document.createElementNS(NS, "line");
    Object.entries({ x1, y1, x2, y2, stroke: "#000", "stroke-width": this.lineWidth, "stroke-linecap": "square" })
      .forEach(([a, v]) => el.setAttribute(a, v));
    this.page.append(el);
  }
  text(str, x, y, opts = {}) {
    const el = document.createElementNS(NS, "text");
    const st = this.font.style;
    Object.entries({
      x, y, fill: this.textColor, "font-size": this.font.size, "font-family": "Arial, Helvetica, sans-serif",
      "font-weight": st.includes("bold") ? "bold" : "normal", "font-style": st.includes("italic") ? "italic" : "normal",
      "text-anchor": opts.align === "center" ? "middle" : opts.align === "right" ? "end" : "start",
      "dominant-baseline": opts.baseline === "middle" ? "central" : opts.baseline === "top" ? "hanging" : "auto",
    }).forEach(([a, v]) => el.setAttribute(a, v));
    el.textContent = str;
    this.page.append(el);
  }
}

// PDF title size: 16pt, or smaller so a long title stays on one line.
function titleFontSize(doc, title, width) {
  doc.setFontSize(16);
  return Math.min(16, (16 * width) / doc.getTextWidth(title));
}

// The same sheet drawn straight to PDF (sizes in points, like the Excel file).
function buildPdf(layout, s, makeDoc = pdfDoc) {
  const doc = makeDoc(s);
  const pw = doc.internal.pageSize.getWidth();
  const ph = doc.internal.pageSize.getHeight();
  const cell = widthPt(s.labelUnits || WIDTH.cell);
  const gap = widthPt(SPACING[s.spacing].gap);
  const sp = SPACING[s.spacing].row;
  const secW = widthPt(sectionUnits(s));
  const boxW = s.sections * secW;
  const gridW = layout.across * (cell + boxW) + (layout.across - 1) * gap;

  // Wrap the description to the grid width at full size, then scale it with the rest.
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  const desc = s.description ? doc.splitTextToSize(s.description, gridW) : [];
  const headH = (s.title ? TITLE_PT : 0) + (desc.length ? desc.length * DESC_LINE_PT + 2 : 0)
    + (s.title || desc.length ? SPACER_PT : 0);
  const gridH = layout.down * bandHeightPt(layout, s) + (layout.down - 1) * GAP_ROWS * ROW_PT;
  const usableW = pw - (MARGIN_IN.left + MARGIN_IN.right) * 72;
  const usableH = ph - (MARGIN_IN.top + MARGIN_IN.bottom) * 72;
  const k = Math.min(1, usableW / gridW, usableH / (headH + gridH));
  const x0 = alignX(s, pw, gridW * k, widthPt(WIDTH.margin) * k);
  const midX = x0 + (gridW * k) / 2; // title and description center over the swatches

  const THIN = 0.5;
  const MEDIUM = 1.5;
  const line = (x1, y1, x2, y2, w) => {
    doc.setLineWidth(w * k);
    doc.line(x0 + x1 * k, y1, x0 + x2 * k, y2);
  };

  for (let p = 0; p < s.pages; p++) {
    if (p > 0) doc.addPage();
    let y = MARGIN_IN.top * 72;
    doc.setTextColor(0);
    if (s.title) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(titleFontSize(doc, s.title, gridW) * k); // one line, shrunk to fit like Excel
      doc.text(s.title, midX, y + (TITLE_PT * k) / 2, { align: "center", baseline: "middle" });
      y += TITLE_PT * k;
    }
    if (desc.length) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10 * k);
      doc.setTextColor(85);
      desc.forEach((text, i) => doc.text(text, midX, y + (2 + i * DESC_LINE_PT) * k, { align: "center", baseline: "top" }));
      y += (desc.length * DESC_LINE_PT + 2) * k;
    }
    if (s.title || desc.length) y += SPACER_PT * k;

    doc.setDrawColor(0);
    for (let b = 0; b < layout.down; b++) {
      for (let g = 0; g < layout.across; g++) {
        const cx = g * (cell + boxW + gap); // code cell, then box at cx + cell
        const bx = cx + cell;
        const ex = bx + boxW; // right edge of the box
        const first = (((p * layout.down) + b) * layout.across + g) * layout.perGroup;
        const count = swatchesToDraw(s, first, layout.perGroup);
        if (!count) continue; // nothing but spare blanks: leave the space empty
        let top = y;
        if (s.headers) {
          const h = ROW_PT * k;
          line(cx, top, ex, top, MEDIUM);
          line(cx, top + h, ex, top + h, MEDIUM);
          line(cx, top, cx, top + h, MEDIUM);
          line(ex, top, ex, top + h, MEDIUM);
          const name = (s.groupNames || [])[((p * layout.down) + b) * layout.across + g];
          if (name) {
            // Set name in the header, shrunk to fit the group's width.
            doc.setFont("helvetica", "bold");
            doc.setTextColor(0);
            const room = ex - cx - 4;
            doc.setFontSize(10);
            const size = Math.min(10, (10 * room) / doc.getTextWidth(name));
            doc.setFontSize(size * k);
            doc.text(name, x0 + ((cx + ex) / 2) * k, top + h / 2, { align: "center", baseline: "middle" });
          }
          top += h + sp * k;
        }
        const rowH = ROW_PT * k;
        const stepH = (ROW_PT + sp) * k;
        const last = count - 1;
        // Screen colors fill the boxes in the preview only (the PDF has no fillBox).
        if (doc.fillBox && s.colorOf) {
          for (let i = 0; i <= last; i++) {
            const color = s.colorOf[s.codes[first + i]];
            if (color) doc.fillBox(x0 + bx * k, top + i * stepH, boxW * k, rowH, color);
          }
        }
        // Thin lines first, so the thick outline sits on top.
        for (let i = 0; i <= last; i++) {
          const yT = top + i * stepH;
          const yB = yT + rowH;
          line(cx, yT, cx, yB, THIN);
          line(cx, yT, ex, yT, THIN);
          line(cx, yB, ex, yB, THIN);
          for (let j = 1; j < s.sections; j++) line(bx + j * secW, yT, bx + j * secW, yB, THIN);
        }
        // Touching swatches share one thick outline; spaced swatches get their own.
        for (let i = 0; i <= last; i++) {
          const yT = top + i * stepH;
          const yB = yT + rowH;
          line(bx, yT, bx, yB, MEDIUM);
          line(ex, yT, ex, yB, MEDIUM);
          if (sp || i === 0) line(bx, yT, ex, yT, MEDIUM);
          if (sp || i === last) line(bx, yB, ex, yB, MEDIUM);
        }
        const labelStyle = s.labelBold && s.labelItalic ? "bolditalic" : s.labelBold ? "bold" : s.labelItalic ? "italic" : "normal";
        doc.setFont("helvetica", labelStyle);
        doc.setTextColor(s.labelColor);
        for (let i = 0; i < count; i++) {
          const code = s.codes[first + i];
          if (!code) continue;
          let size = s.labelSize;
          doc.setFontSize(size);
          const w = doc.getTextWidth(code);
          if (w > cell - 4) size = (size * (cell - 4)) / w; // shrink to fit, like Excel
          doc.setFontSize(size * k);
          doc.text(code, x0 + (cx + cell / 2) * k, top + i * stepH + rowH / 2, { align: "center", baseline: "middle" });
        }
      }
      y += bandHeightPt(layout, s) * k;
      if (b < layout.down - 1) y += GAP_ROWS * ROW_PT * k;
    }
  }
  return doc;
}

// ---- Palette strips and mixing grids ------------------------------------------
// Both are described as a small sheet model (column widths in Excel units, row
// heights in pt, cells with borders and text) that the preview, Excel and PDF
// all draw from.

// Palette strips and mixing grids are always blank: no medium, brand or set.
// (Mixing grids are meant for alcohol or watercolor markers; the menu says so.)
function pageTypeOf(type) {
  return type === "strips" || type === "mixing" ? type : "grid";
}

const STRIP = { name: 16, box: 6.5, gap: 3, codePt: 14, rowPt: 34, gapPt: 10 };
const MIX = { maxBoxPt: 40 };
const T = "thin";
const M = "medium";

// Palette strips: a name cell, then n boxes with a thin row above them for
// writing each color's code, in two columns down the page. If two full-size
// strips don't fit side by side, the name cell and boxes get narrower (same height).
const toUnits = (pt) => Math.floor(pt / 0.75) / 7; // pt -> Excel width units (whole pixels)
function stripModel(s) {
  const { w, h } = usableSize(s);
  const n = s.stripColours;
  const across = 2;
  const gapW = widthPt(STRIP.gap);
  const room = (w * 0.92 - widthPt(WIDTH.margin) - gapW) / across; // width one strip may use
  const fullW = widthPt(STRIP.name) + n * widthPt(STRIP.box);
  const f = Math.min(1, room / fullW);
  const nameUnits = f < 1 ? toUnits(widthPt(STRIP.name) * f) : STRIP.name;
  const boxUnits = f < 1 ? toUnits(widthPt(STRIP.box) * f) : STRIP.box;
  const stripW = widthPt(nameUnits) + n * widthPt(boxUnits);
  const contentW = widthPt(WIDTH.margin) + across * stripW + (across - 1) * gapW;
  const head = headRowsFor(contentW, s);
  const headH = head.reduce((t, r) => t + r.height, 0);
  const stripH = STRIP.codePt + STRIP.rowPt;
  const down = Math.max(1, Math.floor((h * 0.97 - headH + STRIP.gapPt) / (stripH + STRIP.gapPt)));

  const cols = [WIDTH.margin];
  for (let a = 0; a < across; a++) {
    cols.push(nameUnits);
    for (let j = 0; j < n; j++) cols.push(boxUnits);
    if (a < across - 1) cols.push(STRIP.gap);
  }
  const rows = [];
  for (let d = 0; d < down; d++) {
    rows.push(STRIP.codePt, STRIP.rowPt);
    if (d < down - 1) rows.push(STRIP.gapPt);
  }
  const cells = [];
  for (let d = 0; d < down; d++) {
    for (let a = 0; a < across; a++) {
      const r = d * 3 + 1; // code row is r - 1
      const c0 = 1 + a * (n + 2);
      const strip = d * across + a; // palettes fill left to right, then down
      cells.push({ r, c: c0, strip, border: { top: T, left: T, bottom: T, right: M } }); // palette name
      for (let j = 0; j < n; j++) {
        // Code box above each color, then the color box itself.
        cells.push({ r: r - 1, c: c0 + 1 + j, strip, border: { top: T, bottom: M, left: T, right: T } });
        cells.push({ r, c: c0 + 1 + j, strip, border: { top: M, bottom: M, left: j === 0 ? M : T, right: j === n - 1 ? M : T } });
      }
    }
  }
  // As many pages as the chosen number of palettes needs (up to MAX_PAGES); the last
  // page holds the rest.
  const perPage = across * down;
  const palettes = Math.min(s.palettes, perPage * MAX_PAGES);
  const pages = Math.max(1, Math.ceil(palettes / perPage));
  const cellsOn = (p) => {
    const count = Math.min(perPage, palettes - p * perPage);
    return cells.filter((c) => c.strip < count);
  };
  return { cols, rows, cells, cellsOn, perPage, pages, palettes, head };
}

// Mixing grid: colours down the side and across the top; each inside box is
// where two colours meet. The diagonal (a colour with itself) gets a thick outline.
function mixModel(s) {
  const { w, h } = usableSize(s);
  const n = s.mixColours;
  const labels = Array(n).fill(""); // always blank: write your own colours in
  const head = headRowsFor(w, s);
  const headH = head.reduce((t, r) => t + r.height, 0);
  // Every box is the same square, label boxes included: n + 1 boxes across and down.
  const boxPt = Math.min(MIX.maxBoxPt, (w * 0.92 - widthPt(WIDTH.margin)) / (n + 1), (h * 0.97 - headH) / (n + 1));
  const boxUnits = Math.max(2, Math.floor(boxPt / 0.75) / 7);
  const side = widthPt(boxUnits);

  const cols = [WIDTH.margin];
  for (let j = 0; j <= n; j++) cols.push(boxUnits);
  const rows = [];
  for (let i = 0; i <= n; i++) rows.push(side); // square boxes
  const cells = [];
  for (let j = 0; j < n; j++) {
    cells.push({ r: 0, c: 2 + j, text: labels[j], label: true,
      border: { top: M, bottom: M, left: j === 0 ? M : T, right: j === n - 1 ? M : T } });
  }
  for (let i = 0; i < n; i++) {
    cells.push({ r: 1 + i, c: 1, text: labels[i], label: true,
      border: { left: M, right: M, top: i === 0 ? M : T, bottom: i === n - 1 ? M : T } });
    for (let j = 0; j < n; j++) {
      const diag = i === j;
      cells.push({ r: 1 + i, c: 2 + j, border: {
        top: diag || i === 0 ? M : T, bottom: diag || i === n - 1 ? M : T,
        left: diag || j === 0 ? M : T, right: diag || j === n - 1 ? M : T } });
    }
  }
  return { cols, rows, cells, cellsOn: () => cells, pages: 1, head };
}

const modelFor = (s) => (s.pageType === "strips" ? stripModel(s) : mixModel(s));
const modelWidthPt = (m) => m.cols.slice(1).reduce((t, u) => t + widthPt(u), 0); // without the margin column
const modelHeightPt = (m) => m.rows.reduce((t, h) => t + h, 0);

function modelScale(m, s) {
  const { w, h } = usableSize(s);
  const headH = m.head.reduce((t, r) => t + r.height, 0);
  const fit = Math.min((w / (modelWidthPt(m) + widthPt(WIDTH.margin))) * 0.92, h / (headH + modelHeightPt(m))) * 100;
  return Math.max(10, Math.min(100, Math.floor(fit) - 2));
}

const labelFont = (s) => ({
  name: FONT, size: s.labelSize, bold: s.labelBold, italic: s.labelItalic,
  color: { argb: "FF" + s.labelColor.slice(1).toUpperCase() },
});

async function buildModelWorkbook(m, s) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(s.pageType === "strips" ? "Palettes" : "Mixing grid");
  ws.columns = m.cols.map((width) => ({ width }));
  const lastCol = m.cols.length;
  m.head.forEach((row, i) => {
    const r = i + 1;
    ws.getRow(r).height = row.height;
    if (row.kind === "spacer") return;
    ws.mergeCells(r, 2, r, lastCol);
    const cell = ws.getCell(r, 2);
    cell.value = row.kind === "title" ? s.title : s.description;
    cell.font = row.kind === "title"
      ? { name: FONT, size: 16, bold: true }
      : { name: FONT, size: 10, color: { argb: "FF555555" } };
    cell.alignment = row.kind === "title"
      ? { horizontal: "center", vertical: "middle", shrinkToFit: true }
      : { horizontal: "center", vertical: "top", wrapText: true };
  });
  const line = { thin, medium };
  let r0 = m.head.length + 1;
  for (let p = 0; p < m.pages; p++) {
    m.rows.forEach((height, i) => { ws.getRow(r0 + i).height = height; });
    for (const c of m.cellsOn(p)) {
      const cell = ws.getCell(r0 + c.r, c.c + 1);
      const b = {};
      for (const side of ["top", "left", "bottom", "right"]) if (c.border[side]) b[side] = line[c.border[side]];
      cell.border = b;
      if (c.label) {
        if (c.text) cell.value = c.text;
        cell.font = labelFont(s);
        cell.alignment = { horizontal: "center", vertical: "middle", shrinkToFit: true };
      }
    }
    r0 += m.rows.length;
    if (p < m.pages - 1) ws.getRow(r0 - 1).addPageBreak();
  }
  const modelScaleFixed = m.pages === 1 ? 0 : modelScale(m, s);
  ws.pageSetup = {
    paperSize: PAPER[s.paper].code,
    orientation: s.orientation,
    ...excelPlacement(s, modelWidthPt(m) + widthPt(WIDTH.margin),
      m.head.reduce((t, r) => t + r.height, 0) + modelHeightPt(m), modelScaleFixed),
    ...(m.head.length ? { printTitlesRow: `1:${m.head.length}` } : {}),
    ...(m.pages === 1 ? { fitToPage: true, fitToWidth: 1, fitToHeight: 1 } : { scale: modelScaleFixed }),
  };
  return wb;
}

function buildModelPdf(m, s, makeDoc = pdfDoc) {
  const doc = makeDoc(s);
  const pw = doc.internal.pageSize.getWidth();
  const gridW = modelWidthPt(m);
  const gridH = modelHeightPt(m);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  const desc = s.description ? doc.splitTextToSize(s.description, gridW) : [];
  const headH = (s.title ? TITLE_PT : 0) + (desc.length ? desc.length * DESC_LINE_PT + 2 : 0)
    + (s.title || desc.length ? SPACER_PT : 0);
  const { w: usableW, h: usableH } = usableSize(s);
  const k = Math.min(1, usableW / gridW, usableH / (headH + gridH));
  const x0 = alignX(s, pw, gridW * k, widthPt(WIDTH.margin) * k);
  const midX = x0 + (gridW * k) / 2;
  // Column and row edges, in pt (column 0 is Excel's margin column and is skipped).
  const xs = [0, 0];
  for (let c = 1; c < m.cols.length; c++) xs.push(xs[c] + widthPt(m.cols[c]));
  const ys = [0];
  m.rows.forEach((h, i) => ys.push(ys[i] + h));
  const weight = { thin: 0.5, medium: 1.5 };
  const labelStyle = s.labelBold && s.labelItalic ? "bolditalic" : s.labelBold ? "bold" : s.labelItalic ? "italic" : "normal";

  for (let p = 0; p < m.pages; p++) {
    if (p > 0) doc.addPage();
    let y = MARGIN_IN.top * 72;
    doc.setTextColor(0);
    if (s.title) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(titleFontSize(doc, s.title, gridW) * k); // one line, shrunk to fit like Excel
      doc.text(s.title, midX, y + (TITLE_PT * k) / 2, { align: "center", baseline: "middle" });
      y += TITLE_PT * k;
    }
    if (desc.length) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10 * k);
      doc.setTextColor(85);
      desc.forEach((text, i) => doc.text(text, midX, y + (2 + i * DESC_LINE_PT) * k, { align: "center", baseline: "top" }));
      y += (desc.length * DESC_LINE_PT + 2) * k;
    }
    if (s.title || desc.length) y += SPACER_PT * k;
    const X = (x) => x0 + x * k;
    const Y = (v) => y + v * k;
    doc.setDrawColor(0);
    // Thin lines first, so the thick ones sit on top.
    for (const pass of ["thin", "medium"]) {
      doc.setLineWidth(weight[pass] * k);
      for (const c of m.cellsOn(p)) {
        const [l, r, t, b] = [xs[c.c], xs[c.c + 1], ys[c.r], ys[c.r + 1]];
        if (c.border.top === pass) doc.line(X(l), Y(t), X(r), Y(t));
        if (c.border.bottom === pass) doc.line(X(l), Y(b), X(r), Y(b));
        if (c.border.left === pass) doc.line(X(l), Y(t), X(l), Y(b));
        if (c.border.right === pass) doc.line(X(r), Y(t), X(r), Y(b));
      }
    }
    doc.setFont("helvetica", labelStyle);
    doc.setTextColor(s.labelColor);
    for (const c of m.cellsOn(p)) {
      if (!c.label || !c.text) continue;
      const cw = xs[c.c + 1] - xs[c.c];
      let size = s.labelSize;
      doc.setFontSize(size);
      const tw = doc.getTextWidth(c.text);
      if (tw > cw - 4) size = (size * (cw - 4)) / tw;
      doc.setFontSize(size * k);
      doc.text(c.text, X(xs[c.c] + cw / 2), Y((ys[c.r] + ys[c.r + 1]) / 2), { align: "center", baseline: "middle" });
    }
  }
  return doc;
}

// Preview: one picture per printed page, the full paper with its white space,
// drawn by the same code as the PDF download.
function showPages(doc) {
  const preview = $("preview");
  doc.pages.forEach((svg, i) => {
    const pageEl = document.createElement("div");
    pageEl.className = "page";
    const label = document.createElement("div");
    label.className = "page-label";
    label.textContent = `Page ${i + 1}`;
    pageEl.append(label, svg);
    preview.append(pageEl);
  });
}

function currentLayout(s) {
  const options = layoutOptions(s);
  return options.find((o) => o.key === chosenKey) || options[0] || null;
}

function fileName(s, ext = ".xlsx") {
  const base = s.title.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").trim();
  return (base || "swatch-sheet") + ext;
}

function downloadPdf() {
  const s = readSettings();
  if (s.pageType !== "grid") return buildModelPdf(modelFor(s), s).save(fileName(s, ".pdf"));
  const layout = currentLayout(s);
  if (!layout) return;
  buildPdf(layout, withLayoutCodes(s, layout)).save(fileName(s, ".pdf"));
}

async function download() {
  const s = readSettings();
  let wb;
  if (s.pageType !== "grid") {
    wb = await buildModelWorkbook(modelFor(s), s);
  } else {
    const layout = currentLayout(s);
    if (!layout) return;
    wb = await buildWorkbook(layout, withLayoutCodes(s, layout));
  }
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = fileName(s);
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function renderOptions(s) {
  const options = layoutOptions(s);
  const box = $("options");
  box.textContent = "";
  if (options.length === 0) {
    // Point to the nearest page count that fits: fewer or more pages.
    const fits = (pages) => layoutOptions({ ...s, pages }).length > 0;
    let nearest = null;
    for (let d = 1; d < MAX_PAGES && nearest === null; d++) {
      if (s.pages - d >= 1 && fits(s.pages - d)) nearest = s.pages - d;
      else if (s.pages + d <= MAX_PAGES && fits(s.pages + d)) nearest = s.pages + d;
    }
    const pagesText = (n) => `${n} page${n > 1 ? "s" : ""}`;
    box.innerHTML = nearest === null
      ? '<p class="note warn">No layout fits. Try landscape, or the other paper size.</p>'
      : nearest < s.pages
        ? `<p class="note warn">Too many pages for this many swatches. Try ${pagesText(nearest)}, or tick Auto.</p>`
        : `<p class="note warn">No layout fits on ${pagesText(s.pages)}. Try ${pagesText(nearest)}, or tick Auto.</p>`;
    return null;
  }
  if (!options.some((o) => o.key === chosenKey)) chosenKey = options[0].key;
  for (const o of options) {
    const label = document.createElement("label");
    label.className = "option";
    const input = document.createElement("input");
    input.type = "radio";
    input.name = "layout";
    input.checked = o.key === chosenKey;
    input.addEventListener("change", () => { chosenKey = o.key; render(); });
    const text = document.createElement("span");
    const groups = o.across * o.down * s.pages;
    text.textContent = `${groups} group${groups > 1 ? "s" : ""} of ${o.perGroup}: ${o.across} across × ${o.down} down`;
    const small = document.createElement("small");
    small.textContent = `Prints at about ${o.scale}%` + (o.blanks ? `, ${o.blanks} spare blank swatch${o.blanks > 1 ? "es" : ""}` : "");
    text.append(small);
    label.append(input, text);
    box.append(label);
  }
  return options.find((o) => o.key === chosenKey);
}

// Show only the settings that belong to the chosen page type.
function showPageTypeControls(s) {
  const grid = s.pageType === "grid";
  $("grid-only").hidden = !grid;
  $("layout-pick").hidden = !grid;
  $("total-row").hidden = !grid;
  $("strip-colours-row").hidden = s.pageType !== "strips";
  $("mix-colours-row").hidden = s.pageType !== "mixing";
  $("palettes-row").hidden = s.pageType !== "strips"; // strips: you choose how many palettes
  $("pages-row").hidden = !grid;
  $("pages").disabled = s.autoPages;
  if (s.autoPages) $("pages").value = s.pages;
  $("pages-note").textContent = "";
  // Blank pages (strips, mixing grid): no medium, brand or set to pick.
  const medium = $("medium");
  if (!grid) {
    for (const id of ["find-row", "medium-row", "brand-row", "set-row", "legacy-row", "no-codes-row"]) $(id).hidden = true;
  } else {
    $("find-row").hidden = false;
    $("medium-row").hidden = false;
    $("brand-row").hidden = !medium.value;
    $("set-row").hidden = !$("brand").value;
    showSetControls();
  }
  const note = $("page-type-note");
  note.hidden = grid;
  note.textContent = s.pageType === "strips"
    ? "Each strip has a box for the palette's name and a small box above each color for its code."
    : "Write your colors down the side and across the top. Each box is where two colors meet; the thick boxes are each color on its own.";
}

// One line above the preview: what the download will be.
function showSummary(s, pages, extra) {
  const paper = s.paper === "a4" ? "A4" : "US Letter";
  $("summary").textContent = [`${paper}, ${s.orientation}`, `${pages} page${pages > 1 ? "s" : ""}`, ...extra].join(" · ");
}

function render() {
  const s = readSettings();
  showPageTypeControls(s);
  $("color-note").hidden = !(s.pageType === "grid" && s.colorOf);
  const preview = $("preview");
  preview.textContent = "";
  if (s.pageType !== "grid") {
    const m = modelFor(s);
    if (s.pageType === "strips") {
      $("palettes-note").textContent = m.palettes < s.palettes
        ? `${m.perPage} per page. ${MAX_PAGES} pages (the most) hold ${m.palettes}, so only ${m.palettes} will print.`
        : `${m.perPage} per page, fits on ${m.pages} page${m.pages > 1 ? "s" : ""}.`;
    }
    showSummary(s, m.pages, [s.pageType === "strips" ? `${Math.min(s.palettes, m.palettes ?? s.palettes)} palettes` : `${s.mixColours} colors`]);
    showPages(buildModelPdf(m, s, (x) => new SvgDoc(x)));
    return;
  }
  const layout = renderOptions(s);
  if (!layout) {
    $("summary").textContent = "No layout fits yet: see the message under Pick a layout.";
    return;
  }
  withLayoutCodes(s, layout);
  if (s.autoPages) {
    $("pages-note").textContent = `Auto: fits on ${s.pages} page${s.pages > 1 ? "s" : ""} at about ${layout.scale}%.`;
  }
  showSummary(s, s.pages, [`${s.total} swatch${s.total > 1 ? "es" : ""}`, `prints at about ${layout.scale}%`]);
  showPages(buildPdf(layout, s, (x) => new SvgDoc(x)));
}

// Light, gray or dark page colours (dark by default: set in index.html before the
// page draws), remembered in this browser only.
function initLook() {
  const root = document.documentElement;
  const save = () => {
    try {
      localStorage.setItem("spm-look", JSON.stringify({ mode: root.dataset.mode || "light" }));
    } catch (e) {}
  };
  const mark = () => {
    const mode = root.dataset.mode || "light";
    document.querySelectorAll(".seg [data-mode]").forEach((b) => b.classList.toggle("on", b.dataset.mode === mode));
  };
  document.querySelectorAll(".seg [data-mode]").forEach((b) => b.addEventListener("click", () => {
    if (b.dataset.mode === "light") delete root.dataset.mode; else root.dataset.mode = b.dataset.mode;
    mark(); save();
  }));
  if (root.dataset.mode === "light") delete root.dataset.mode;
  mark();
}

// Picking a set fills in its size and a title; typing a new total drops the set.
// The brand is the first word of a preset's group ("Copic Sketch" -> "Copic").
const brandOf = (p) => p.brand || p.group.split(" ")[0]; // two-word brands carry their own

// Older preset files have no medium of their own; these lines set it.
const MEDIUM_BY_GROUP = [
  [/^(Guangna|Sharpie|Posca|Copic acrea)/, "Acrylic marker"],
  [/^(Copic|Ohuhu Honolulu|Ohuhu Oahu|Ohuhu Kaala)/, "Alcohol marker"],
];
const mediumOf = (p) => p.medium || (MEDIUM_BY_GROUP.find(([re]) => re.test(p.group)) || [])[1] || "Other";
const MEDIUM_ORDER = ["Alcohol marker", "Acrylic marker", "Marker", "Watercolor marker", "Felt-tip marker", "Colored pencil", "Watercolor pencil", "Pastel pencil", "Watercolor paint", "Watercolor stick", "Gel pen", "Ink pen", "Highlighter"];

// Only these mediums show in the menu for now (too many choices).
// The other preset files still load; add a medium back here to show it again.
const SHOWN_MEDIUMS = ["Alcohol marker", "Acrylic marker", "Marker", "Colored pencil"];

function initPresets() {
  const medium = $("medium");
  const brand = $("brand");
  const select = $("preset");
  const mediums = [...new Set(PRESETS.filter((p) => !p.hidden).map(mediumOf))].filter((m) => SHOWN_MEDIUMS.includes(m)).sort((a, b) =>
    ((MEDIUM_ORDER.indexOf(a) + 1) || 99) - ((MEDIUM_ORDER.indexOf(b) + 1) || 99) || a.localeCompare(b));
  for (const m of mediums) medium.add(new Option(m, m));
  const inMedium = () => PRESETS.filter((p) => mediumOf(p) === medium.value);

  // Fill the brand menu with the brands that make this medium.
  const fillBrands = () => {
    brand.textContent = "";
    brand.add(new Option("Pick a brand…", ""));
    for (const b of new Set(inMedium().filter((p) => !p.hidden).map(brandOf))) brand.add(new Option(b, b)); // a brand with every set hidden stays out
    $("brand-row").hidden = !medium.value;
  };

  // Fill the set menu for one brand, with its product lines as sub-headings.
  const fillSets = () => {
    select.textContent = "";
    select.add(new Option("Pick a set…", ""));
    const list = inMedium().filter((p) => brandOf(p) === brand.value && !p.hidden); // collection members hide
    const lines = new Set(list.map((p) => p.group));
    const groups = new Map();
    for (const p of list) {
      let parent = select;
      if (lines.size > 1) {
        if (!groups.has(p.group)) {
          const og = document.createElement("optgroup");
          og.label = p.group.slice(brand.value.length).trim();
          groups.set(p.group, og);
          select.append(og);
        }
        parent = groups.get(p.group);
      }
      const n = String(p.codes.length);
      const label = p.name.includes(n) ? p.name : `${p.name} (${n})`;
      parent.append(new Option(label, p.id));
    }
    $("set-row").hidden = !brand.value;
  };
  medium.addEventListener("change", () => {
    fillBrands();
    // Only one brand for this medium: pick it straight away.
    if (brand.options.length === 2) brand.selectedIndex = 1;
    fillSets();
    select.dispatchEvent(new Event("change"));
  });
  brand.addEventListener("change", () => {
    fillSets();
    select.dispatchEvent(new Event("change"));
  });
  fillBrands();
  fillSets();

  // "Find a set": every word typed must appear in the line, set or medium name.
  const findBox = $("find");
  const results = $("find-results");
  const findName = (p) => `${p.group} – ${p.name}`;
  const pick = (p) => {
    medium.value = mediumOf(p);
    fillBrands();
    brand.value = brandOf(p);
    fillSets();
    select.value = p.id;
    select.dispatchEvent(new Event("change"));
    findBox.value = "";
    results.hidden = true;
  };
  let found = [];
  findBox.addEventListener("input", () => {
    const words = findBox.value.toLowerCase().split(/\s+/).filter(Boolean);
    found = words.length
      ? PRESETS.filter((p) => !p.hidden && SHOWN_MEDIUMS.includes(mediumOf(p)))
        .filter((p) => words.every((w) => `${findName(p)} ${mediumOf(p)}`.toLowerCase().includes(w))).slice(0, 12)
      : [];
    results.textContent = "";
    for (const p of found) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = findName(p);
      const m = document.createElement("small");
      m.textContent = mediumOf(p);
      b.append(m);
      b.addEventListener("mousedown", (e) => { e.preventDefault(); pick(p); }); // before the box loses focus
      results.append(b);
    }
    if (words.length && !found.length) results.innerHTML = '<p class="note">No set matches. Try fewer words.</p>';
    results.hidden = !words.length;
  });
  findBox.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && found.length) { e.preventDefault(); pick(found[0]); }
    if (e.key === "Escape") results.hidden = true;
  });
  findBox.addEventListener("blur", () => { results.hidden = true; });
  // Preset notes stay in the data files as our own record; the page never shows them.
  // Picking, adding or removing a set updates the total. The title is never filled in:
  // it stays blank unless the user types one.
  // Collections switch headers on so each set's name shows; leaving one switches
  // them back off, unless the user ticked the box themselves.
  let autoHeaders = false;
  $("headers").addEventListener("input", () => { autoHeaders = false; });
  const applyChosen = () => {
    const list = chosenPresets();
    if (list.length) $("total").value = mergedCodes(list, $("legacy-codes").checked).length;
    if (hasCollection()) {
      if (!$("headers").checked) autoHeaders = true;
      $("headers").checked = true;
    } else if (autoHeaders) {
      $("headers").checked = false;
      autoHeaders = false;
    }
    // Some collections (Posca's color families) always show their headers.
    $("headers").disabled = !!presetById(select.value)?.forceHeaders;
    chosenKey = null; // a new set: start from the best layout again
    showSetControls();
    render();
  };
  select.addEventListener("change", () => {
    applyChosen();
  });
  $("legacy-codes").addEventListener("input", applyChosen); // legacy codes can change the repeat count
  $("total").addEventListener("input", () => {
    chosenKey = null; // a new total: start from the best layout again
    const list = chosenPresets();
    if (list.length && readSettings().total !== mergedCodes(list, $("legacy-codes").checked).length) {
      select.value = "";
      showSetControls();
    }
  });
}

// Shows the code checkboxes that belong to the picked set.
function showSetControls() {
  const list = chosenPresets();
  $("no-codes-row").hidden = !list.length;
  $("legacy-row").hidden = !list.some((p) => p.legacy); // only sets that have old codes (Ohuhu)
}

// The "?" button opens the tips box; ✕, Esc or a click outside closes it.
function initHelp() {
  const box = $("help");
  $("help-btn").addEventListener("click", () => box.showModal());
  $("help-close").addEventListener("click", () => box.close());
  box.addEventListener("click", (e) => {
    const r = box.getBoundingClientRect();
    if (e.target === box && (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom)) box.close();
  });
}

// "Reset": every setting back to how the page first opens (the theme stays).
function resetAll() {
  for (const el of document.querySelectorAll("main input, main select, main textarea")) {
    if (el.type === "checkbox") el.checked = el.defaultChecked;
    else if (el.tagName === "SELECT") el.selectedIndex = Math.max(0, [...el.options].findIndex((o) => o.defaultSelected));
    else el.value = el.defaultValue;
  }
  $("headers").disabled = false;
  $("find-results").hidden = true;
  $("more").open = false;
  chosenKey = null;
  $("medium").dispatchEvent(new Event("change")); // empties Brand and Set, then redraws
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function init() {
  initLook();
  initHelp();
  initPresets();
  for (const id of ["page-type", "strip-colours", "palettes", "mix-colours", "total", "pages", "auto-pages", "paper", "orientation", "align", "headers", "spares", "spacing", "sections", "label-bold", "label-italic", "label-size", "label-color", "legacy-codes", "no-codes", "title", "description"]) {
    $(id).addEventListener("input", render);
  }
  // A number outside the box's limits is used as the nearest limit: show that number in the box too.
  for (const [id, min, max] of [["total", 1, 2000], ["pages", 1, MAX_PAGES], ["palettes", 1, 500]]) {
    $(id).addEventListener("change", () => {
      const v = clampInt($(id).value, min, max, NaN);
      if (Number.isFinite(v) && String(v) !== $(id).value) { $(id).value = v; render(); }
    });
  }
  $("download").addEventListener("click", download);
  $("download-pdf").addEventListener("click", downloadPdf);
  // Reset only works once Confirm is ticked; resetting unticks it again.
  $("reset-ok").addEventListener("input", () => { $("reset").disabled = !$("reset-ok").checked; });
  $("reset").addEventListener("click", () => {
    if (!$("reset-ok").checked) return;
    resetAll();
    $("reset").disabled = true;
  });
  render();
}

init();
