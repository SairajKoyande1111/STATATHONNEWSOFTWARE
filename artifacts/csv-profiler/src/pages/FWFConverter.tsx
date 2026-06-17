import { useState, useRef, useCallback } from "react";
import {
  Upload, FileText, FileSpreadsheet, CheckCircle2, AlertTriangle,
  X, ArrowRight, Download, Eye, Layers, ChevronRight, RotateCcw,
  ShieldCheck, Key, Lock, Shuffle,
} from "lucide-react";
import {
  parseLayoutFile, readExcelFileInfo, getSheetRowCount,
  convertFWFToCSV, fwfToRows, rowsToCSVBlob,
  type FieldDef, type ParseLayoutResult, type ExcelFileInfo,
} from "@/lib/fwf-parser";
import { anonymizeRows, type AnonymizeOptions } from "@/lib/anonymize";

type Step = "layout" | "data" | "converted" | "anon-done";
type LayoutSubStep = "upload" | "sheet-select" | "done";

export default function FWFConverter() {
  const [step, setStep] = useState<Step>("layout");

  // Layout state
  const [layoutSubStep, setLayoutSubStep] = useState<LayoutSubStep>("upload");
  const [layoutResult, setLayoutResult] = useState<ParseLayoutResult | null>(null);
  const [layoutFileName, setLayoutFileName] = useState<string>("");
  const [layoutError, setLayoutError] = useState<string>("");

  // Sheet selection state
  const [excelInfo, setExcelInfo] = useState<ExcelFileInfo | null>(null);
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [rowFrom, setRowFrom] = useState<string>("");
  const [rowTo, setRowTo] = useState<string>("");
  const [sheetRowCount, setSheetRowCount] = useState<number>(0);
  const [applyingSheet, setApplyingSheet] = useState(false);

  // Data file state
  const [dataFileName, setDataFileName] = useState<string>("");
  const [dataText, setDataText] = useState<string>("");
  const [dataLineCount, setDataLineCount] = useState<number>(0);
  const [dataError, setDataError] = useState<string>("");

  // Conversion state
  const [progress, setProgress] = useState<number>(0);
  const [converting, setConverting] = useState<boolean>(false);
  const [parsedRows, setParsedRows] = useState<Record<string, string>[] | null>(null);
  const [outputBaseName, setOutputBaseName] = useState<string>("");
  const [preview, setPreview] = useState<string[][]>([]);

  // Anonymize state
  const [anonCols, setAnonCols] = useState<Set<string>>(new Set());
  const [anonKeyMode, setAnonKeyMode] = useState<"random" | "pbkdf2">("random");
  const [anonSeed, setAnonSeed] = useState<number>(42);
  const [anonPassphrase, setAnonPassphrase] = useState<string>("");
  const [anonPbkdf2Iter, setAnonPbkdf2Iter] = useState<number>(100000);
  const [anonDeterministic, setAnonDeterministic] = useState<boolean>(true);
  const [anonRunning, setAnonRunning] = useState<boolean>(false);
  const [anonProgress, setAnonProgress] = useState<number>(0);
  const [anonKeyHex, setAnonKeyHex] = useState<string | null>(null);
  const [anonBlob, setAnonBlob] = useState<Blob | null>(null);
  const [anonError, setAnonError] = useState<string>("");
  const [keyCopied, setKeyCopied] = useState(false);

  const layoutInputRef = useRef<HTMLInputElement>(null);
  const dataInputRef = useRef<HTMLInputElement>(null);

  // ── Layout file upload ────────────────────────────────────────────────────

  const handleLayoutFile = useCallback(async (file: File) => {
    setLayoutError("");
    setLayoutResult(null);
    setLayoutFileName(file.name);
    const name = file.name.toLowerCase();

    if (name.endsWith(".csv") || name.endsWith(".tsv")) {
      try {
        const result = await parseLayoutFile(file);
        if (result.fields.length === 0) {
          setLayoutError(result.warnings.join(" ") || "No fields found.");
          return;
        }
        setLayoutResult(result);
        setLayoutSubStep("done");
        setStep("data");
      } catch (e) {
        setLayoutError(`Failed to parse layout file: ${(e as Error).message}`);
      }
      return;
    }

    try {
      const info = await readExcelFileInfo(file);
      if (info.sheetNames.length > 1) {
        setExcelInfo(info);
        setPendingFile(file);
        setSelectedSheet(info.sheetNames[0]);
        setSheetRowCount(getSheetRowCount(info.buf, info.sheetNames[0]));
        setRowFrom(""); setRowTo("");
        setLayoutSubStep("sheet-select");
      } else {
        const result = await parseLayoutFile(file);
        if (result.fields.length === 0) {
          setLayoutError(result.warnings.join(" ") || "No fields found.");
          return;
        }
        setLayoutResult(result);
        setLayoutSubStep("done");
        setStep("data");
      }
    } catch (e) {
      setLayoutError(`Failed to read layout file: ${(e as Error).message}`);
    }
  }, []);

  const handleSheetChange = useCallback((sheet: string) => {
    setSelectedSheet(sheet);
    if (excelInfo) setSheetRowCount(getSheetRowCount(excelInfo.buf, sheet));
    setRowFrom(""); setRowTo("");
  }, [excelInfo]);

  const handleConfirmSheetSelection = useCallback(async () => {
    if (!pendingFile || !selectedSheet) return;
    setApplyingSheet(true); setLayoutError("");
    try {
      const startRow = rowFrom ? parseInt(rowFrom, 10) : undefined;
      const endRow = rowTo ? parseInt(rowTo, 10) : undefined;
      const result = await parseLayoutFile(pendingFile, { sheetName: selectedSheet, startRow, endRow });
      if (result.fields.length === 0) {
        setLayoutError(result.warnings.filter(w => !w.startsWith("Scanning")).join(" ") || "No fields found.");
        setApplyingSheet(false); return;
      }
      setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
    } catch (e) {
      setLayoutError(`Failed to parse layout: ${(e as Error).message}`);
    } finally { setApplyingSheet(false); }
  }, [pendingFile, selectedSheet, rowFrom, rowTo]);

  const handleAutoDetect = useCallback(async () => {
    if (!pendingFile) return;
    setApplyingSheet(true); setLayoutError("");
    try {
      const result = await parseLayoutFile(pendingFile);
      if (result.fields.length === 0) {
        setLayoutError(result.warnings.join(" ") || "No fields found.");
        setApplyingSheet(false); return;
      }
      setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
    } catch (e) {
      setLayoutError(`Failed to parse layout file: ${(e as Error).message}`);
    } finally { setApplyingSheet(false); }
  }, [pendingFile]);

  // ── Data file upload ──────────────────────────────────────────────────────

  const handleDataFile = useCallback(async (file: File) => {
    setDataError("");
    setDataFileName(file.name);
    setDataText(""); setDataLineCount(0); setParsedRows(null); setPreview([]);

    const text = await file.text();
    const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
    setDataText(text);
    setDataLineCount(lines.length);

    if (layoutResult) {
      const prev = lines.slice(0, 5).map((line) =>
        layoutResult.fields.map((f) => line.padEnd(f.end).substring(f.start - 1, f.end).trim())
      );
      setPreview(prev);
    }

    setOutputBaseName(file.name.replace(/\.[^.]+$/, ""));
    setStep("data");
  }, [layoutResult]);

  // ── Convert ───────────────────────────────────────────────────────────────

  const handleConvert = useCallback(async () => {
    if (!layoutResult || !dataText) return;
    setConverting(true); setProgress(0); setParsedRows(null);
    setAnonBlob(null); setAnonKeyHex(null); setAnonError("");
    try {
      // Convert FWF → in-memory rows (for anonymization)
      const rows = fwfToRows(dataText, layoutResult.fields);
      setParsedRows(rows);

      // Also build default anon columns (all fields)
      setAnonCols(new Set(layoutResult.fields.map(f => f.varName)));

      setProgress(100);
      setStep("converted");
    } catch (e) {
      setDataError(`Conversion failed: ${(e as Error).message}`);
    } finally {
      setConverting(false);
    }
  }, [layoutResult, dataText]);

  // ── Anonymize ─────────────────────────────────────────────────────────────

  const handleAnonymize = useCallback(async () => {
    if (!parsedRows || !layoutResult) return;
    if (anonCols.size === 0) { setAnonError("Select at least one column to anonymize."); return; }
    setAnonRunning(true); setAnonProgress(0); setAnonError(""); setAnonBlob(null); setAnonKeyHex(null);

    try {
      const opts: AnonymizeOptions = {
        keyMode: anonKeyMode,
        seed: anonSeed,
        passphrase: anonPassphrase,
        pbkdf2Iterations: anonPbkdf2Iter,
        deterministic: anonDeterministic,
      };
      const { rows: anonRows, keyHex } = await anonymizeRows(
        parsedRows, [...anonCols], opts, setAnonProgress
      );
      const columns = layoutResult.fields.map(f => f.varName);
      const blob = rowsToCSVBlob(anonRows, columns);
      setAnonBlob(blob);
      setAnonKeyHex(keyHex);
      setStep("anon-done");
    } catch (e) {
      setAnonError(`Anonymization failed: ${(e as Error).message}`);
    } finally {
      setAnonRunning(false);
    }
  }, [parsedRows, layoutResult, anonCols, anonKeyMode, anonSeed, anonPassphrase, anonPbkdf2Iter, anonDeterministic]);

  const handleDownload = (blob: Blob, name: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name; a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadOriginal = useCallback(() => {
    if (!parsedRows || !layoutResult) return;
    const columns = layoutResult.fields.map(f => f.varName);
    const blob = rowsToCSVBlob(parsedRows, columns);
    handleDownload(blob, `${outputBaseName}.csv`);
  }, [parsedRows, layoutResult, outputBaseName]);

  const handleCopyKey = () => {
    if (!anonKeyHex) return;
    navigator.clipboard.writeText(anonKeyHex).then(() => {
      setKeyCopied(true);
      setTimeout(() => setKeyCopied(false), 2000);
    });
  };

  const handleReset = () => {
    setStep("layout"); setLayoutSubStep("upload");
    setLayoutResult(null); setLayoutFileName(""); setLayoutError("");
    setExcelInfo(null); setPendingFile(null); setSelectedSheet("");
    setRowFrom(""); setRowTo(""); setSheetRowCount(0);
    setDataFileName(""); setDataText(""); setDataLineCount(0);
    setDataError(""); setProgress(0);
    setParsedRows(null); setPreview([]);
    setAnonCols(new Set()); setAnonBlob(null); setAnonKeyHex(null);
    setAnonError(""); setAnonProgress(0);
  };

  const fields: FieldDef[] = layoutResult?.fields ?? [];
  const allColNames = fields.map(f => f.varName);
  const isConverted = step === "converted" || step === "anon-done";

  return (
    <div className="space-y-6">
      {/* Step indicator */}
      <div className="flex items-center gap-2 text-sm flex-wrap">
        <StepBadge n={1} label="Upload layout" active={step === "layout"} done={step !== "layout"} />
        <ArrowRight className="w-4 h-4 text-muted-foreground" />
        <StepBadge n={2} label="Upload data file" active={step === "data"} done={isConverted} />
        <ArrowRight className="w-4 h-4 text-muted-foreground" />
        <StepBadge n={3} label="Convert" active={step === "data" && !!parsedRows === false} done={isConverted} />
        <ArrowRight className="w-4 h-4 text-muted-foreground" />
        <StepBadge n={4} label="Anonymize & download" active={step === "converted"} done={step === "anon-done"} />
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[1fr_1fr] min-w-0">
        {/* ── Step 1: Layout file ─────────────────────────────────────────── */}
        <div className="bg-card border border-border rounded-xl p-5 space-y-4 min-w-0 overflow-hidden">
          <div className="flex items-center justify-between">
            <div>
              <h2 className="text-sm font-semibold text-foreground">Step 1 — Layout file</h2>
              <p className="text-xs text-muted-foreground mt-0.5">
                Excel (.xlsx) or CSV with columns: Field_Name, Byte Position (Start), Byte Position (End)
              </p>
            </div>
            {(layoutResult || layoutSubStep === "sheet-select") && (
              <button onClick={handleReset} className="text-xs text-muted-foreground hover:text-foreground" title="Start over">
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {layoutSubStep === "upload" && (
            <>
              <DropZone accept=".xlsx,.xls,.csv" icon={<FileSpreadsheet className="w-8 h-8 text-primary" />}
                label="Drop layout file here" sublabel="Excel or CSV"
                inputRef={layoutInputRef} onFile={handleLayoutFile} />
              {layoutError && <ErrorBox message={layoutError} />}
            </>
          )}

          {layoutSubStep === "sheet-select" && excelInfo && (
            <div className="space-y-4">
              <div className="flex items-center gap-2 text-xs text-blue-700 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2">
                <FileSpreadsheet className="w-4 h-4 flex-shrink-0" />
                <span className="font-medium truncate">{layoutFileName}</span>
                <span className="ml-auto text-blue-500 whitespace-nowrap flex-shrink-0">{excelInfo.sheetNames.length} sheets found</span>
              </div>

              <div className="space-y-2">
                <label className="text-xs font-semibold text-foreground flex items-center gap-1.5">
                  <Layers className="w-3.5 h-3.5" /> Select sheet
                </label>
                <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                  {excelInfo.sheetNames.map((name) => (
                    <label key={name} className={`flex items-center gap-2.5 px-3 py-2 rounded-lg border cursor-pointer transition-colors text-xs ${
                      selectedSheet === name ? "border-primary bg-primary/5 text-foreground" : "border-border hover:border-primary/40 hover:bg-accent/20 text-muted-foreground"
                    }`}>
                      <input type="radio" name="sheet" value={name} checked={selectedSheet === name}
                        onChange={() => handleSheetChange(name)} className="accent-primary" />
                      <span className="font-medium">{name}</span>
                    </label>
                  ))}
                </div>
              </div>

              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-semibold text-foreground">
                    Row range {sheetRowCount > 0 && <span className="ml-1 font-normal text-muted-foreground">(sheet has {sheetRowCount} rows)</span>}
                  </label>
                  {(rowFrom || rowTo) && (
                    <button onClick={() => { setRowFrom(""); setRowTo(""); }}
                      className="text-[11px] text-muted-foreground hover:text-foreground flex items-center gap-1">
                      <RotateCcw className="w-3 h-3" /> All rows
                    </button>
                  )}
                </div>
                <div className="flex items-center gap-2">
                  <div className="flex-1 space-y-1">
                    <p className="text-[11px] text-muted-foreground">From row</p>
                    <input type="number" min={1} placeholder="1" value={rowFrom}
                      onChange={(e) => setRowFrom(e.target.value)}
                      className="w-full px-2 py-1.5 text-xs rounded-lg border border-border bg-background focus:outline-none focus:ring-1 focus:ring-primary" />
                  </div>
                  <ChevronRight className="w-4 h-4 text-muted-foreground mt-4 flex-shrink-0" />
                  <div className="flex-1 space-y-1">
                    <p className="text-[11px] text-muted-foreground">To row</p>
                    <input type="number" min={1} placeholder={sheetRowCount ? String(sheetRowCount) : "last"} value={rowTo}
                      onChange={(e) => setRowTo(e.target.value)}
                      className="w-full px-2 py-1.5 text-xs rounded-lg border border-border bg-background focus:outline-none focus:ring-1 focus:ring-primary" />
                  </div>
                </div>
                <p className="text-[11px] text-muted-foreground">Leave blank to scan all rows.</p>
              </div>

              {layoutError && <ErrorBox message={layoutError} />}

              <div className="flex flex-col gap-2 pt-1">
                <button onClick={handleConfirmSheetSelection} disabled={applyingSheet || !selectedSheet}
                  className="w-full flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-primary text-primary-foreground text-sm font-semibold hover:opacity-90 disabled:opacity-60 transition-opacity">
                  {applyingSheet ? <><div className="w-4 h-4 border-2 border-primary-foreground/30 border-t-primary-foreground rounded-full animate-spin" />Parsing…</> : (
                    <><ArrowRight className="w-4 h-4" />Use selected sheet{rowFrom || rowTo ? ` (rows ${rowFrom || "1"}–${rowTo || "end"})` : ""}</>
                  )}
                </button>
                <button onClick={handleAutoDetect} disabled={applyingSheet}
                  className="w-full flex items-center justify-center gap-2 px-4 py-2 rounded-xl border border-border text-xs text-muted-foreground hover:text-foreground hover:border-primary/40 disabled:opacity-60 transition-colors">
                  <Upload className="w-3.5 h-3.5" /> Auto-detect layout sheet (skip selection)
                </button>
              </div>
            </div>
          )}

          {layoutSubStep === "done" && layoutResult && (
            <div className="space-y-3">
              <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                <CheckCircle2 className="w-4 h-4 flex-shrink-0" />
                <span><strong>{layoutFileName}</strong> — {fields.length} fields detected{layoutResult.sheetName ? ` (sheet: ${layoutResult.sheetName})` : ""}</span>
              </div>
              {layoutResult.warnings.length > 0 && (
                <div className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">{layoutResult.warnings.join(" ")}</div>
              )}
              <div className="overflow-auto max-h-72 rounded-lg border border-border">
                <table className="w-full text-xs border-collapse">
                  <thead className="bg-muted sticky top-0">
                    <tr>
                      <th className="px-2 py-1.5 text-left border-r border-border/50 text-muted-foreground font-medium">#</th>
                      <th className="px-2 py-1.5 text-left border-r border-border/50 text-muted-foreground font-medium">Variable</th>
                      <th className="px-2 py-1.5 text-left border-r border-border/50 text-muted-foreground font-medium">Full Name</th>
                      <th className="px-2 py-1.5 text-center border-r border-border/50 text-muted-foreground font-medium">Start</th>
                      <th className="px-2 py-1.5 text-center border-r border-border/50 text-muted-foreground font-medium">End</th>
                      <th className="px-2 py-1.5 text-center text-muted-foreground font-medium">Len</th>
                    </tr>
                  </thead>
                  <tbody>
                    {fields.map((f) => (
                      <tr key={f.srlNo} className="border-t border-border/40 hover:bg-muted/30">
                        <td className="px-2 py-1 text-muted-foreground font-mono border-r border-border/30">{f.srlNo}</td>
                        <td className="px-2 py-1 font-medium text-foreground border-r border-border/30 whitespace-nowrap">{f.varName}</td>
                        <td className="px-2 py-1 text-muted-foreground border-r border-border/30">{f.fullName}</td>
                        <td className="px-2 py-1 text-center font-mono border-r border-border/30">{f.start}</td>
                        <td className="px-2 py-1 text-center font-mono border-r border-border/30">{f.end}</td>
                        <td className="px-2 py-1 text-center font-mono">{f.length}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* ── Step 2: Data file & Convert ──────────────────────────────────── */}
        <div className="bg-card border border-border rounded-xl p-5 space-y-4 min-w-0 overflow-hidden">
          <div>
            <h2 className="text-sm font-semibold text-foreground">Step 2 — Fixed-width data file</h2>
            <p className="text-xs text-muted-foreground mt-0.5">The .TXT file containing the actual records</p>
          </div>

          {!layoutResult ? (
            <div className="flex flex-col items-center justify-center h-40 text-center text-sm text-muted-foreground border-2 border-dashed border-border/40 rounded-xl">
              Complete Step 1 first
            </div>
          ) : !dataFileName ? (
            <DropZone accept=".txt,.dat,.fwf,.data" icon={<FileText className="w-8 h-8 text-primary" />}
              label="Drop fixed-width data file here" sublabel=".TXT, .DAT or any fixed-width file"
              inputRef={dataInputRef} onFile={handleDataFile} />
          ) : (
            <div className="space-y-4">
              <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                <CheckCircle2 className="w-4 h-4 flex-shrink-0" />
                <span><strong>{dataFileName}</strong> — {dataLineCount.toLocaleString()} records</span>
                {!isConverted && (
                  <button onClick={() => { setDataFileName(""); setDataText(""); setDataLineCount(0); setParsedRows(null); setPreview([]); setStep("data"); }}
                    className="ml-auto text-muted-foreground hover:text-foreground"><X className="w-3.5 h-3.5" /></button>
                )}
              </div>

              {preview.length > 0 && (
                <div className="space-y-1">
                  <div className="flex items-center gap-1 text-xs text-muted-foreground">
                    <Eye className="w-3.5 h-3.5" /> Preview (first {preview.length} rows)
                  </div>
                  <div className="overflow-auto max-h-48 rounded-lg border border-border text-[11px]">
                    <table className="w-full border-collapse">
                      <thead className="bg-muted sticky top-0">
                        <tr>{fields.map((f) => (
                          <th key={f.srlNo} className="px-2 py-1 text-left font-medium text-muted-foreground border-r border-border/50 whitespace-nowrap">{f.varName}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {preview.map((row, ri) => (
                          <tr key={ri} className="border-t border-border/40">
                            {row.map((cell, ci) => (
                              <td key={ci} className="px-2 py-1 font-mono border-r border-border/30 whitespace-nowrap">
                                {cell || <span className="text-muted-foreground/40 italic">—</span>}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {dataError && <ErrorBox message={dataError} />}

              {!isConverted ? (
                <button onClick={handleConvert} disabled={converting}
                  className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-primary text-primary-foreground text-sm font-semibold hover:opacity-90 disabled:opacity-60 transition-opacity">
                  {converting ? (
                    <><div className="w-4 h-4 border-2 border-primary-foreground/30 border-t-primary-foreground rounded-full animate-spin" />Converting… {progress}%</>
                  ) : (
                    <><ArrowRight className="w-4 h-4" />Convert {dataLineCount.toLocaleString()} records — proceed to anonymize</>
                  )}
                </button>
              ) : (
                <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                  <CheckCircle2 className="w-4 h-4 flex-shrink-0" />
                  {dataLineCount.toLocaleString()} records ready · {fields.length} columns · Proceed to Step 4 below
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* ── Step 3 (4): Anonymize ─────────────────────────────────────────────── */}
      {isConverted && parsedRows && layoutResult && (
        <div className="bg-card border border-border rounded-xl p-5 space-y-5 min-w-0 overflow-hidden">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-5 h-5 text-primary" />
            <div>
              <h2 className="text-sm font-semibold text-foreground">Step 4 — Anonymize columns</h2>
              <p className="text-xs text-muted-foreground mt-0.5">
                Select columns to encrypt using AES-256-GCM format-preserving encryption, then download the anonymized CSV.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 gap-5 lg:grid-cols-[1fr_320px]">
            {/* ── Column selector ── */}
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-foreground">
                  Columns to anonymize <span className="font-normal text-muted-foreground">({anonCols.size} / {allColNames.length} selected)</span>
                </span>
                <div className="flex gap-2">
                  <button onClick={() => setAnonCols(new Set(allColNames))}
                    className="text-[11px] px-2 py-1 rounded border border-border hover:border-primary/40 text-muted-foreground hover:text-foreground transition-colors">
                    Select all
                  </button>
                  <button onClick={() => setAnonCols(new Set())}
                    className="text-[11px] px-2 py-1 rounded border border-border hover:border-primary/40 text-muted-foreground hover:text-foreground transition-colors">
                    Clear
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-1.5 max-h-60 overflow-y-auto pr-1 border border-border/50 rounded-lg p-2 bg-muted/20">
                {allColNames.map((col) => (
                  <label key={col} className={`flex items-center gap-1.5 px-2 py-1.5 rounded-md border cursor-pointer text-[11px] transition-colors ${
                    anonCols.has(col) ? "border-primary/40 bg-primary/5 text-foreground" : "border-transparent hover:border-border text-muted-foreground hover:text-foreground"
                  }`}>
                    <input type="checkbox" checked={anonCols.has(col)}
                      onChange={(e) => {
                        const next = new Set(anonCols);
                        if (e.target.checked) next.add(col); else next.delete(col);
                        setAnonCols(next);
                      }}
                      className="accent-primary w-3 h-3 flex-shrink-0" />
                    <span className="truncate font-mono">{col}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* ── Encryption settings ── */}
            <div className="space-y-4">
              <div className="space-y-2">
                <p className="text-xs font-semibold text-foreground flex items-center gap-1.5"><Key className="w-3.5 h-3.5" /> Key derivation</p>
                <div className="space-y-1.5">
                  {(["random", "pbkdf2"] as const).map((mode) => (
                    <label key={mode} className={`flex items-center gap-2 px-3 py-2 rounded-lg border cursor-pointer text-xs transition-colors ${
                      anonKeyMode === mode ? "border-primary bg-primary/5 text-foreground" : "border-border hover:border-primary/40 text-muted-foreground"
                    }`}>
                      <input type="radio" name="keymode" value={mode} checked={anonKeyMode === mode}
                        onChange={() => setAnonKeyMode(mode)} className="accent-primary" />
                      <span>{mode === "random" ? "Random Key (seed-based)" : "PBKDF2 (passphrase)"}</span>
                    </label>
                  ))}
                </div>

                {anonKeyMode === "random" && (
                  <div className="space-y-1 pt-1">
                    <p className="text-[11px] text-muted-foreground">Key seed</p>
                    <input type="number" value={anonSeed} onChange={(e) => setAnonSeed(Number(e.target.value))}
                      className="w-full px-2 py-1.5 text-xs font-mono rounded-lg border border-border bg-background focus:outline-none focus:ring-1 focus:ring-primary" />
                    <p className="text-[11px] text-muted-foreground">Same seed = reproducible key</p>
                  </div>
                )}

                {anonKeyMode === "pbkdf2" && (
                  <div className="space-y-2 pt-1">
                    <div className="space-y-1">
                      <p className="text-[11px] text-muted-foreground">Passphrase</p>
                      <input type="password" value={anonPassphrase} onChange={(e) => setAnonPassphrase(e.target.value)}
                        placeholder="Enter encryption passphrase…"
                        className="w-full px-2 py-1.5 text-xs rounded-lg border border-border bg-background focus:outline-none focus:ring-1 focus:ring-primary" />
                    </div>
                    <div className="space-y-1">
                      <p className="text-[11px] text-muted-foreground">Iterations: {anonPbkdf2Iter.toLocaleString()}</p>
                      <input type="range" min={10000} max={500000} step={10000} value={anonPbkdf2Iter}
                        onChange={(e) => setAnonPbkdf2Iter(Number(e.target.value))}
                        className="w-full accent-primary" />
                    </div>
                  </div>
                )}
              </div>

              <label className={`flex items-center gap-2.5 px-3 py-2 rounded-lg border cursor-pointer text-xs transition-colors ${
                anonDeterministic ? "border-primary bg-primary/5 text-foreground" : "border-border hover:border-primary/40 text-muted-foreground"
              }`}>
                <input type="checkbox" checked={anonDeterministic} onChange={(e) => setAnonDeterministic(e.target.checked)}
                  className="accent-primary w-3.5 h-3.5 flex-shrink-0" />
                <div>
                  <p className="font-medium">Deterministic mode</p>
                  <p className="text-[10px] text-muted-foreground mt-0.5">Same value → same encrypted output</p>
                </div>
              </label>

              {/* Cipher suite info */}
              <div className="bg-muted/40 border border-border/50 rounded-lg p-3 space-y-1 text-[11px]">
                <p className="font-semibold text-foreground uppercase tracking-wide text-[10px] text-muted-foreground mb-1.5">Cipher suite</p>
                {[
                  ["Algorithm", "AES-256-GCM"],
                  ["Key size", "256-bit (32 bytes)"],
                  ["IV / Nonce", "96-bit"],
                  ["Auth tag", "128-bit (GHASH)"],
                  ["Standard", "NIST FIPS 197 + SP 800-38D"],
                  ["FPE", "Digits→Digits, Letters→Letters"],
                ].map(([k, v]) => (
                  <div key={k} className="flex gap-1 text-muted-foreground">
                    <span className="font-medium text-foreground w-20 shrink-0">{k}</span>
                    <span>{v}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {anonError && <ErrorBox message={anonError} />}

          {/* Progress bar */}
          {anonRunning && (
            <div className="space-y-1.5">
              <div className="flex justify-between text-xs text-muted-foreground">
                <span className="flex items-center gap-1.5"><Shuffle className="w-3.5 h-3.5 animate-spin" /> Encrypting {anonCols.size} column{anonCols.size !== 1 ? "s" : ""}…</span>
                <span>{anonProgress}%</span>
              </div>
              <div className="h-2 rounded-full bg-muted overflow-hidden">
                <div className="h-full bg-primary transition-all duration-300 rounded-full" style={{ width: `${anonProgress}%` }} />
              </div>
            </div>
          )}

          {/* Action buttons */}
          {step !== "anon-done" && (
            <div className="flex flex-col sm:flex-row gap-2">
              <button onClick={handleAnonymize} disabled={anonRunning || anonCols.size === 0}
                className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-primary text-primary-foreground text-sm font-semibold hover:opacity-90 disabled:opacity-60 transition-opacity">
                {anonRunning ? <><div className="w-4 h-4 border-2 border-primary-foreground/30 border-t-primary-foreground rounded-full animate-spin" />Encrypting…</>
                  : <><Lock className="w-4 h-4" />Apply AES-256-GCM anonymization</>}
              </button>
              <button onClick={handleDownloadOriginal} disabled={anonRunning}
                className="flex items-center justify-center gap-2 px-4 py-3 rounded-xl border border-border text-xs text-muted-foreground hover:text-foreground hover:border-primary/40 disabled:opacity-60 transition-colors">
                <Download className="w-3.5 h-3.5" /> Skip — download original CSV
              </button>
            </div>
          )}

          {/* Results after anonymization */}
          {step === "anon-done" && anonBlob && anonKeyHex && (
            <div className="space-y-4">
              <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                <CheckCircle2 className="w-4 h-4 flex-shrink-0" />
                Anonymization complete — {anonCols.size} column{anonCols.size !== 1 ? "s" : ""} encrypted across {parsedRows.length.toLocaleString()} records
              </div>

              {/* Key material box */}
              <div className="border border-amber-300 bg-amber-50 rounded-xl p-4 space-y-2">
                <div className="flex items-center gap-2 text-xs font-semibold text-amber-800">
                  <Key className="w-4 h-4" />
                  Symmetric Key — save this to decrypt later
                </div>
                <div className="font-mono text-[11px] bg-white border border-amber-200 rounded-lg px-3 py-2 break-all select-all cursor-text text-foreground">
                  {anonKeyHex}
                </div>
                <div className="flex items-center justify-between">
                  <p className="text-[11px] text-amber-700">
                    AES-256 symmetric key · {anonKeyMode === "random" ? `seed = ${anonSeed}` : `PBKDF2 (${anonPbkdf2Iter.toLocaleString()} iter)`}
                  </p>
                  <button onClick={handleCopyKey}
                    className="text-[11px] px-2 py-1 rounded border border-amber-300 text-amber-800 hover:bg-amber-100 transition-colors">
                    {keyCopied ? "Copied!" : "Copy key"}
                  </button>
                </div>
                <p className="text-[11px] text-amber-700/80">
                  ⚠ AES-256-GCM is symmetric — the same key encrypts and decrypts. Store in a secure vault (HSM, AWS KMS, etc.).
                </p>
              </div>

              <div className="flex flex-col sm:flex-row gap-2">
                <button onClick={() => handleDownload(anonBlob!, `${outputBaseName}_anonymized.csv`)}
                  className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-emerald-600 text-white text-sm font-semibold hover:bg-emerald-700 transition-colors">
                  <Download className="w-4 h-4" /> Download anonymized CSV
                </button>
                <button onClick={handleDownloadOriginal}
                  className="flex items-center justify-center gap-2 px-4 py-3 rounded-xl border border-border text-xs text-muted-foreground hover:text-foreground hover:border-primary/40 transition-colors">
                  <Download className="w-3.5 h-3.5" /> Download original CSV
                </button>
              </div>

              <button onClick={() => { setAnonBlob(null); setAnonKeyHex(null); setAnonProgress(0); setStep("converted"); }}
                className="w-full text-xs text-muted-foreground hover:text-foreground text-center">
                Change column selection or settings
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ── Sub-components ────────────────────────────────────────────────────────────

function StepBadge({ n, label, active, done }: { n: number; label: string; active: boolean; done: boolean }) {
  return (
    <div className={`flex items-center gap-1.5 text-xs font-medium ${active ? "text-primary" : done ? "text-emerald-600" : "text-muted-foreground"}`}>
      <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold flex-shrink-0 ${
        active ? "bg-primary text-primary-foreground" : done ? "bg-emerald-100 text-emerald-600" : "bg-muted text-muted-foreground"
      }`}>
        {done ? <CheckCircle2 className="w-3 h-3" /> : n}
      </span>
      {label}
    </div>
  );
}

function DropZone({ accept, icon, label, sublabel, inputRef, onFile }: {
  accept: string; icon: React.ReactNode; label: string; sublabel: string;
  inputRef: React.RefObject<HTMLInputElement | null>; onFile: (f: File) => void;
}) {
  const [dragging, setDragging] = useState(false);
  return (
    <div className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all ${
      dragging ? "border-primary bg-primary/5" : "border-border hover:border-primary/50 hover:bg-accent/20"
    }`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
      onClick={() => inputRef.current?.click()}>
      <input ref={inputRef} type="file" accept={accept} className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />
      <div className="flex flex-col items-center gap-3">
        <div className="w-14 h-14 rounded-xl bg-primary/10 flex items-center justify-center">{icon}</div>
        <div>
          <p className="text-sm font-medium text-foreground">{label}</p>
          <p className="text-xs text-muted-foreground mt-1">{sublabel}</p>
        </div>
        <span className="text-xs px-3 py-1.5 rounded-lg border border-border bg-background text-muted-foreground">Browse</span>
      </div>
    </div>
  );
}

function ErrorBox({ message }: { message: string }) {
  return (
    <div className="flex items-start gap-2 text-xs text-destructive bg-destructive/10 border border-destructive/20 rounded-lg px-3 py-2">
      <AlertTriangle className="w-4 h-4 flex-shrink-0 mt-0.5" />
      {message}
    </div>
  );
}
