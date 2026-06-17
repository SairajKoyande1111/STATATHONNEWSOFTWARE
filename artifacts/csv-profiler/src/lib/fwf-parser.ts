import * as XLSX from "xlsx";

// ── Types ────────────────────────────────────────────────────────────────────

export interface FieldDef {
  srlNo: number;
  fullName: string;
  varName: string;   // short code name → CSV column header
  start: number;     // 1-indexed byte start
  end: number;       // 1-indexed byte end
  length: number;
}

export interface ParseLayoutResult {
  fields: FieldDef[];
  sheetName: string;
  warnings: string[];
}

export interface ExcelFileInfo {
  sheetNames: string[];
  buf: ArrayBuffer;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function nh(s: unknown): string {
  return String(s ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

const HEADER_ALIASES: Record<string, string> = {
  // variable name
  fieldname: "varName",
  fieldnames: "varName",
  varname: "varName",
  variable: "varName",
  columnname: "varName",
  colname: "varName",
  // full / display name
  fullname: "fullName",
  name: "fullName",
  item: "fullName",
  description: "fullName",
  label: "fullName",
  // byte start
  bytepositionstart: "start",
  bytestart: "start",
  startbyte: "start",
  byteposstart: "start",
  start: "start",
  from: "start",
  // byte end
  bytepositionend: "end",
  byteend: "end",
  endbyte: "end",
  byteposend: "end",
  end: "end",
  to: "end",
  // length
  fieldlength: "length",
  length: "length",
  len: "length",
  size: "length",
  width: "length",
};

function detectColumns(headers: string[]): Record<string, number> {
  const map: Record<string, number> = {};
  headers.forEach((h, i) => {
    const key = HEADER_ALIASES[nh(h)];
    if (key && !(key in map)) map[key] = i;
  });
  return map;
}

function rowToFieldDef(
  cells: unknown[],
  colMap: Record<string, number>,
  idx: number
): FieldDef | null {
  const getNum = (key: string): number => {
    const v = colMap[key] !== undefined ? Number(cells[colMap[key]]) : NaN;
    return isNaN(v) ? 0 : v;
  };
  const getStr = (key: string): string => {
    const v = colMap[key] !== undefined ? String(cells[colMap[key]] ?? "") : "";
    return v.trim();
  };

  let start = getNum("start");
  let end = getNum("end");
  const length = getNum("length");

  if (!start && !end && !length) return null;

  // Derive missing values
  if (start && !end && length) end = start + length - 1;
  if (!start || !end) return null;
  const len = length || end - start + 1;

  const varName =
    getStr("varName") ||
    getStr("fullName")
      .toLowerCase()
      .replace(/[^a-z0-9_]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "") ||
    `field_${idx + 1}`;

  const fullName = getStr("fullName") || varName;

  return {
    srlNo: idx + 1,
    fullName,
    varName,
    start,
    end,
    length: len,
  };
}

// ── Get Excel sheet names ─────────────────────────────────────────────────────

export function getExcelSheetNames(buf: ArrayBuffer): string[] {
  const wb = XLSX.read(buf, { type: "array" });
  return wb.SheetNames;
}

// ── Get row count for a specific sheet ───────────────────────────────────────

export function getSheetRowCount(buf: ArrayBuffer, sheetName: string): number {
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[sheetName];
  if (!ws) return 0;
  const aoa: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  return aoa.length;
}

// ── Parse layout from Excel ArrayBuffer (auto-detect sheet) ──────────────────

export function parseLayoutFromExcel(
  buf: ArrayBuffer,
  options?: { sheetName?: string; startRow?: number; endRow?: number }
): ParseLayoutResult {
  const wb = XLSX.read(buf, { type: "array" });
  const warnings: string[] = [];

  let targetSheet: string;

  if (options?.sheetName && wb.SheetNames.includes(options.sheetName)) {
    targetSheet = options.sheetName;
  } else {
    // Auto-detect: find a sheet that has start/end columns
    targetSheet = wb.SheetNames[0];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const aoa: unknown[][] = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        defval: "",
      });
      const firstNonEmpty = aoa.find((row) => row.some((c) => String(c).trim()));
      if (!firstNonEmpty) continue;
      const colMap = detectColumns(firstNonEmpty.map(String));
      if (colMap.start && colMap.end) {
        targetSheet = name;
        break;
      }
    }
  }

  const ws = wb.Sheets[targetSheet];
  let aoa: unknown[][] = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
  });

  // Extract merge info for merged-header detection
  const merges: MergeRange[] = (ws["!merges"] as MergeRange[]) ?? [];

  // Apply row range if specified (1-indexed, inclusive)
  if (options?.startRow !== undefined || options?.endRow !== undefined) {
    const s = Math.max(0, (options.startRow ?? 1) - 1);
    const e = options.endRow !== undefined ? options.endRow : aoa.length;
    // Adjust merge row indices when slicing rows
    const adjustedMerges = merges
      .filter((m) => m.s.r >= s && m.s.r < e)
      .map((m) => ({ s: { r: m.s.r - s, c: m.s.c }, e: { r: m.e.r - s, c: m.e.c } }));
    aoa = aoa.slice(s, e);
    warnings.push(`Scanning rows ${options.startRow ?? 1}–${Math.min(e, aoa.length + s)} of sheet "${targetSheet}".`);
    return parseFromAOA(aoa, targetSheet, warnings, adjustedMerges);
  }

  return parseFromAOA(aoa, targetSheet, warnings, merges);
}

// ── Parse layout from CSV text ────────────────────────────────────────────────

export function parseLayoutFromCSV(text: string): ParseLayoutResult {
  const wb = XLSX.read(text, { type: "string" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa: unknown[][] = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
  });
  return parseFromAOA(aoa, "CSV", []);
}

// ── Shared AOA→FieldDef logic ─────────────────────────────────────────────────

type MergeRange = { s: { r: number; c: number }; e: { r: number; c: number } };

/**
 * Expand merged "Byte Position" headers in a header row.
 * When a cell normalises to "byteposition" (or similar) and the worksheet has a
 * merge that spans exactly 2 columns at that cell, we inject "Byte Position
 * (Start)" at the left column and "Byte Position (End)" at the right column so
 * that detectColumns() can find them.
 */
function expandMergedBytePositionHeaders(
  headerRow: string[],
  rowIdx: number,
  merges: MergeRange[]
): string[] {
  const expanded = [...headerRow];
  for (let c = 0; c < headerRow.length; c++) {
    const norm = nh(headerRow[c]);
    // Detect "Byte Position" (merged) — normalises to "byteposition"
    if (norm === "byteposition" || norm === "bytepos" || norm === "bytepositions") {
      // Find a merge covering (rowIdx, c)
      const merge = merges.find(
        (m) => m.s.r === rowIdx && m.s.c === c && m.e.c > c
      );
      if (merge) {
        expanded[c] = "Byte Position (Start)";
        expanded[merge.e.c] = "Byte Position (End)";
      } else {
        // No merge info — assume next blank column is "end"
        if (c + 1 < headerRow.length && headerRow[c + 1].trim() === "") {
          expanded[c] = "Byte Position (Start)";
          expanded[c + 1] = "Byte Position (End)";
        }
      }
    }
  }
  return expanded;
}

function parseFromAOA(
  aoa: unknown[][],
  sheetName: string,
  warnings: string[],
  merges: MergeRange[] = []
): ParseLayoutResult {
  // Find header row: first row where start/end columns can be detected
  let headerRowIdx = -1;
  let colMap: Record<string, number> = {};

  for (let r = 0; r < Math.min(aoa.length, 15); r++) {
    const raw = aoa[r].map(String);
    // Expand merged "Byte Position" headers before detection
    const row = expandMergedBytePositionHeaders(raw, r, merges);
    const cm = detectColumns(row);
    if (cm.start !== undefined && cm.end !== undefined) {
      headerRowIdx = r;
      colMap = cm;
      break;
    }
    // Also accept length + start (end derivable)
    if (cm.start !== undefined && cm.length !== undefined) {
      headerRowIdx = r;
      colMap = cm;
      break;
    }
  }

  if (headerRowIdx === -1) {
    warnings.push(
      "Could not detect header row. Expected columns like: Field_Name, Byte Position (Start), Byte Position (End)."
    );
    return { fields: [], sheetName, warnings };
  }

  const fields: FieldDef[] = [];
  let srl = 0;
  for (let r = headerRowIdx + 1; r < aoa.length; r++) {
    const row = aoa[r];
    const fd = rowToFieldDef(row, colMap, srl);
    if (fd) {
      fields.push(fd);
      srl++;
    }
  }

  if (fields.length === 0) {
    warnings.push("Layout parsed but no valid field rows found.");
  }

  return { fields, sheetName, warnings };
}

// ── Convert FWF text → row objects ───────────────────────────────────────────

export function fwfToRows(text: string, fields: FieldDef[]): Record<string, string>[] {
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  return lines.map((line) => {
    const row: Record<string, string> = {};
    for (const f of fields) {
      row[f.varName] = line.padEnd(f.end).substring(f.start - 1, f.end).trim();
    }
    return row;
  });
}

// ── Convert row objects → CSV Blob ────────────────────────────────────────────

export function rowsToCSVBlob(
  rows: Record<string, string>[],
  columns: string[]
): Blob {
  const header = columns.map(csvCell).join(",");
  const dataLines = rows.map((row) =>
    columns.map((c) => csvCell(row[c] ?? "")).join(",")
  );
  return new Blob([[header, ...dataLines].join("\n")], {
    type: "text/csv;charset=utf-8;",
  });
}

// ── Convert FWF text → CSV Blob ───────────────────────────────────────────────

export interface ConvertOptions {
  onProgress?: (pct: number) => void;
}

const CHUNK = 50_000;

export async function convertFWFToCSV(
  text: string,
  fields: FieldDef[],
  options: ConvertOptions = {}
): Promise<Blob> {
  const lines = text.split(/\r?\n/);

  // Build CSV header
  const header = fields.map((f) => csvCell(f.varName)).join(",");
  const chunks: string[] = [header + "\n"];

  const dataLines = lines.filter((l) => l.length > 0);
  const total = dataLines.length;

  for (let i = 0; i < total; i += CHUNK) {
    const batch = dataLines.slice(i, i + CHUNK);
    const rows = batch.map((line) => {
      const cells = fields.map((f) => {
        // Byte positions are 1-indexed; pad line with spaces if too short
        const s = f.start - 1;
        const e = f.end;
        const raw = line.padEnd(e).substring(s, e);
        return csvCell(raw.trim());
      });
      return cells.join(",");
    });
    chunks.push(rows.join("\n") + "\n");

    if (options.onProgress) {
      options.onProgress(Math.min(99, Math.round(((i + batch.length) / total) * 100)));
    }

    // Yield to the event loop every chunk
    await new Promise((r) => setTimeout(r, 0));
  }

  return new Blob(chunks, { type: "text/csv;charset=utf-8;" });
}

function csvCell(value: string): string {
  if (value.includes(",") || value.includes('"') || value.includes("\n")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

// ── Parse layout file (auto-detect format) ───────────────────────────────────

export async function parseLayoutFile(
  file: File,
  options?: { sheetName?: string; startRow?: number; endRow?: number }
): Promise<ParseLayoutResult> {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv")) {
    const text = await file.text();
    return parseLayoutFromCSV(text);
  }
  // Excel
  const buf = await file.arrayBuffer();
  return parseLayoutFromExcel(buf, options);
}

// ── Read Excel file info (sheet names) for UI ─────────────────────────────────

export async function readExcelFileInfo(file: File): Promise<ExcelFileInfo> {
  const buf = await file.arrayBuffer();
  const sheetNames = getExcelSheetNames(buf);
  return { sheetNames, buf };
}
