export type ColumnType = "numeric" | "text" | "boolean" | "date" | "mixed";

export interface ValueFrequency {
  value: string;
  count: number;
  percent: number;
}

export interface ColumnLayout {
  srlNo: number;
  name: string;
  // Questionnaire reference — auto-detected or left blank
  qSec: string;
  qItem: string;
  qCol: string;
  // Field width (max string length of any value in this column)
  length: number;
  // Cumulative byte positions in a fixed-width representation
  byteStart: number;
  byteEnd: number;
  // Auto-inferred remarks
  remarks: string;
  // Extra stats for the detail panel
  type: ColumnType;
  totalCount: number;
  nonNullCount: number;
  nullCount: number;
  fillRate: number;
  uniqueCount: number;
  topValues: ValueFrequency[];
  sampleValues: string[];
  min?: number;
  max?: number;
  mean?: number;
  median?: number;
}

export interface DataProfile {
  fileName: string;
  totalRows: number;
  totalColumns: number;
  fileSize?: number;
  totalRecordLength: number; // sum of all field widths
  columns: ColumnLayout[];
  previewRows: Record<string, string>[];
}

function detectType(sample: string[]): ColumnType {
  const nonEmpty = sample.filter((v) => v !== "");
  if (nonEmpty.length === 0) return "text";

  const numericCount = nonEmpty.filter((v) => !isNaN(Number(v)) && v.trim() !== "").length;
  if (numericCount / nonEmpty.length > 0.85) return "numeric";

  const boolValues = new Set(["true", "false", "yes", "no", "1", "0", "y", "n"]);
  const boolCount = nonEmpty.filter((v) => boolValues.has(v.toLowerCase())).length;
  if (boolCount / nonEmpty.length > 0.9) return "boolean";

  const dateCount = nonEmpty.filter((v) => {
    const d = new Date(v);
    return !isNaN(d.getTime()) && v.length > 5;
  }).length;
  if (dateCount / nonEmpty.length > 0.8) return "date";

  return "text";
}

function inferRemarks(
  values: string[],
  nonNullValues: string[],
  topValues: ValueFrequency[],
  type: ColumnType,
  nullCount: number,
  uniqueCount: number
): string {
  // If all non-null values are identical → 'VALUE' Generated
  if (uniqueCount === 1 && nonNullValues.length > 0) {
    const val = topValues[0]?.value ?? "";
    return `'${val}' Generated`;
  }

  // If column has blanks/nulls and few unique values → blank-generated note
  if (nullCount > 0 && uniqueCount <= 10) {
    return "If not selected blank generated";
  }

  // If it looks like a multiplier / weight column (large numbers, consistent length)
  if (type === "numeric" && nonNullValues.length > 0) {
    const nums = nonNullValues.map(Number).filter((n) => !isNaN(n));
    const allLarge = nums.every((n) => n > 1000);
    const maxLen = nonNullValues.reduce((a, b) => (b.length > a ? b.length : a), 0);
    if (allLarge && maxLen >= 5) {
      return "Final weight/multiplier";
    }
  }

  return "";
}

function computeTopValues(
  nonNullValues: string[],
  limit = 10
): ValueFrequency[] {
  const freqMap = new Map<string, number>();
  for (const v of nonNullValues) {
    freqMap.set(v, (freqMap.get(v) ?? 0) + 1);
  }
  return Array.from(freqMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([value, count]) => ({
      value,
      count,
      percent: nonNullValues.length > 0 ? (count / nonNullValues.length) * 100 : 0,
    }));
}

function medianOf(sorted: number[]): number {
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0
    ? sorted[mid]
    : (sorted[mid - 1] + sorted[mid]) / 2;
}

export function profileData(
  data: Record<string, string>[],
  headers: string[],
  fileName: string,
  fileSize?: number
): DataProfile {
  let bytePos = 1;
  const columns: ColumnLayout[] = [];

  for (let i = 0; i < headers.length; i++) {
    const name = headers[i];
    const rawValues = data.map((row) => row[name] ?? "");
    const nonNullValues = rawValues.filter((v) => v !== "");
    const totalCount = rawValues.length;
    const nullCount = totalCount - nonNullValues.length;
    const fillRate = totalCount > 0 ? (nonNullValues.length / totalCount) * 100 : 0;

    // Field width = max string length of actual data values (NOT the column name)
    // This mirrors how fixed-width layout files define field width from the data, not the label.
    let fieldWidth = 1; // minimum 1 byte
    for (const v of rawValues) {
      if (v.length > fieldWidth) fieldWidth = v.length;
    }

    const type = detectType(nonNullValues.slice(0, 200));
    const topValues = computeTopValues(nonNullValues);
    const uniqueCount = topValues.length < 10
      ? topValues.length
      : (() => {
          const s = new Set(nonNullValues);
          return s.size;
        })();

    const sampleValues = Array.from(new Set(nonNullValues.slice(0, 5)));

    const remarks = inferRemarks(rawValues, nonNullValues, topValues, type, nullCount, uniqueCount);

    const col: ColumnLayout = {
      srlNo: i + 1,
      name,
      qSec: "",
      qItem: "",
      qCol: "",
      length: fieldWidth,
      byteStart: bytePos,
      byteEnd: bytePos + fieldWidth - 1,
      remarks,
      type,
      totalCount,
      nonNullCount: nonNullValues.length,
      nullCount,
      fillRate,
      uniqueCount,
      topValues,
      sampleValues,
    };

    // Numeric stats
    if (type === "numeric") {
      const nums = nonNullValues.map(Number).filter((n) => !isNaN(n));
      if (nums.length > 0) {
        const sorted = [...nums].sort((a, b) => a - b);
        col.min = sorted[0];
        col.max = sorted[sorted.length - 1];
        col.mean = nums.reduce((a, b) => a + b, 0) / nums.length;
        col.median = medianOf(sorted);
      }
    }

    columns.push(col);
    bytePos += fieldWidth;
  }

  return {
    fileName,
    totalRows: data.length,
    totalColumns: headers.length,
    fileSize,
    totalRecordLength: bytePos - 1,
    columns,
    previewRows: data.slice(0, 100),
  };
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function formatNumber(n: number, decimals = 2): string {
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(1) + "B";
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(1) + "M";
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + "K";
  return n.toFixed(decimals);
}
