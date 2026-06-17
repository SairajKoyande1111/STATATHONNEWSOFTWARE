import { useState, useRef, useCallback } from "react";
import {
  Upload, FileText, FileSpreadsheet, CheckCircle2, AlertTriangle,
  X, ArrowRight, Download, Eye, Layers, RotateCcw,
  ShieldCheck, Key, Lock, Shuffle, LockOpen,
} from "lucide-react";
import {
  parseLayoutFile, readExcelFileInfo, getSheetRowCount, convertFWFToCSV,
  type FieldDef, type ParseLayoutResult, type ExcelFileInfo,
} from "@/lib/fwf-parser";
import {
  encryptFWFToBlob, decryptCSVToBlob, readCSVHeaders,
  type AnonymizeOptions,
} from "@/lib/anonymize";

type Step = "layout" | "data" | "converted" | "anon-done";
type LayoutSubStep = "upload" | "sheet-select" | "done";
type AnonMode = "encrypt" | "decrypt";

export default function FWFConverter() {
  const [step, setStep] = useState<Step>("layout");

  const [layoutSubStep, setLayoutSubStep] = useState<LayoutSubStep>("upload");
  const [layoutResult, setLayoutResult] = useState<ParseLayoutResult | null>(null);
  const [layoutFileName, setLayoutFileName] = useState("");
  const [layoutError, setLayoutError] = useState("");

  const [excelInfo, setExcelInfo] = useState<ExcelFileInfo | null>(null);
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [rowFrom, setRowFrom] = useState("");
  const [rowTo, setRowTo] = useState("");
  const [sheetRowCount, setSheetRowCount] = useState(0);
  const [applyingSheet, setApplyingSheet] = useState(false);

  const [dataFileName, setDataFileName] = useState("");
  const [dataText, setDataText] = useState("");
  const [dataLineCount, setDataLineCount] = useState(0);
  const [dataError, setDataError] = useState("");
  const [preview, setPreview] = useState<string[][]>([]);
  const [outputBaseName, setOutputBaseName] = useState("");

  const [converting, setConverting] = useState(false);

  const [anonMode, setAnonMode] = useState<AnonMode>("encrypt");
  const [anonKeyMode, setAnonKeyMode] = useState<"random" | "pbkdf2" | "hex">("random");
  const [anonSeed, setAnonSeed] = useState(42);
  const [anonPassphrase, setAnonPassphrase] = useState("");
  const [anonPbkdf2Iter, setAnonPbkdf2Iter] = useState(100000);
  const [anonDeterministic, setAnonDeterministic] = useState(true);
  const [anonKeyHexInput, setAnonKeyHexInput] = useState("");

  const [encCols, setEncCols] = useState<Set<string>>(new Set());
  const [encRunning, setEncRunning] = useState(false);
  const [encProgress, setEncProgress] = useState(0);
  const [encResultKey, setEncResultKey] = useState<string | null>(null);
  const [encResultBlob, setEncResultBlob] = useState<Blob | null>(null);
  const [encError, setEncError] = useState("");
  const [keyCopied, setKeyCopied] = useState(false);

  const [origDownloading, setOrigDownloading] = useState(false);
  const [origProgress, setOrigProgress] = useState(0);

  const [decryptFileName, setDecryptFileName] = useState("");
  const [decryptCsvText, setDecryptCsvText] = useState<string | null>(null);
  const [decryptHeaders, setDecryptHeaders] = useState<string[]>([]);
  const [decryptCols, setDecryptCols] = useState<Set<string>>(new Set());
  const [decryptRunning, setDecryptRunning] = useState(false);
  const [decryptProgress, setDecryptProgress] = useState(0);
  const [decryptBlob, setDecryptBlob] = useState<Blob | null>(null);
  const [decryptError, setDecryptError] = useState("");

  const layoutInputRef = useRef<HTMLInputElement>(null);
  const dataInputRef = useRef<HTMLInputElement>(null);
  const decryptInputRef = useRef<HTMLInputElement>(null);

  const fields: FieldDef[] = layoutResult?.fields ?? [];
  const allColNames = fields.map((f) => f.varName);
  const isConverted = step === "converted" || step === "anon-done";

  const triggerDownload = (blob: Blob, name: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  };

  const buildOpts = (): AnonymizeOptions => ({
    keyMode: anonKeyMode,
    seed: anonSeed,
    passphrase: anonPassphrase,
    pbkdf2Iterations: anonPbkdf2Iter,
    deterministic: anonDeterministic,
    keyHex: anonKeyHexInput,
  });

  const keyModeLabel =
    anonKeyMode === "random" ? `seed = ${anonSeed}`
    : anonKeyMode === "pbkdf2" ? `PBKDF2 (${anonPbkdf2Iter.toLocaleString()} iter)`
    : "raw hex key";

  const handleLayoutFile = useCallback(async (file: File) => {
    setLayoutError(""); setLayoutResult(null); setLayoutFileName(file.name);
    const name = file.name.toLowerCase();
    if (name.endsWith(".csv") || name.endsWith(".tsv")) {
      try {
        const result = await parseLayoutFile(file);
        if (!result.fields.length) { setLayoutError(result.warnings.join(" ") || "No fields found."); return; }
        setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
      } catch (e) { setLayoutError(`Parse error: ${(e as Error).message}`); }
      return;
    }
    try {
      const info = await readExcelFileInfo(file);
      if (info.sheetNames.length > 1) {
        setExcelInfo(info); setPendingFile(file); setSelectedSheet(info.sheetNames[0]);
        setSheetRowCount(getSheetRowCount(info.buf, info.sheetNames[0]));
        setRowFrom(""); setRowTo(""); setLayoutSubStep("sheet-select");
      } else {
        const result = await parseLayoutFile(file);
        if (!result.fields.length) { setLayoutError(result.warnings.join(" ") || "No fields found."); return; }
        setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
      }
    } catch (e) { setLayoutError(`Read error: ${(e as Error).message}`); }
  }, []);

  const handleSheetChange = useCallback((sheet: string) => {
    setSelectedSheet(sheet);
    if (excelInfo) setSheetRowCount(getSheetRowCount(excelInfo.buf, sheet));
    setRowFrom(""); setRowTo("");
  }, [excelInfo]);

  const handleConfirmSheet = useCallback(async () => {
    if (!pendingFile || !selectedSheet) return;
    setApplyingSheet(true); setLayoutError("");
    try {
      const result = await parseLayoutFile(pendingFile, {
        sheetName: selectedSheet,
        startRow: rowFrom ? parseInt(rowFrom, 10) : undefined,
        endRow: rowTo ? parseInt(rowTo, 10) : undefined,
      });
      if (!result.fields.length) { setLayoutError(result.warnings.join(" ") || "No fields found."); return; }
      setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
    } catch (e) { setLayoutError(`Parse error: ${(e as Error).message}`); }
    finally { setApplyingSheet(false); }
  }, [pendingFile, selectedSheet, rowFrom, rowTo]);

  const handleAutoDetect = useCallback(async () => {
    if (!pendingFile) return;
    setApplyingSheet(true); setLayoutError("");
    try {
      const result = await parseLayoutFile(pendingFile);
      if (!result.fields.length) { setLayoutError(result.warnings.join(" ") || "No fields found."); return; }
      setLayoutResult(result); setLayoutSubStep("done"); setStep("data");
    } catch (e) { setLayoutError(`Parse error: ${(e as Error).message}`); }
    finally { setApplyingSheet(false); }
  }, [pendingFile]);

  const handleDataFile = useCallback(async (file: File) => {
    setDataError(""); setDataFileName(file.name);
    setDataText(""); setDataLineCount(0); setPreview([]);
    setEncResultBlob(null); setEncResultKey(null); setEncError("");

    const text = await file.text();
    const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
    setDataText(text);
    setDataLineCount(lines.length);

    if (layoutResult) {
      setPreview(lines.slice(0, 5).map((line) =>
        layoutResult.fields.map((f) => line.padEnd(f.end).substring(f.start - 1, f.end).trim())
      ));
    }
    setOutputBaseName(file.name.replace(/\.[^.]+$/, ""));
    setStep("data");
  }, [layoutResult]);

  const handleConvert = useCallback(async () => {
    if (!layoutResult || !dataText) return;
    setConverting(true);
    await new Promise((r) => setTimeout(r, 30));
    setEncCols(new Set(layoutResult.fields.map((f) => f.varName)));
    setStep("converted");
    setConverting(false);
  }, [layoutResult, dataText]);

  const handleEncrypt = useCallback(async () => {
    if (!layoutResult || !dataText) return;
    if (encCols.size === 0) { setEncError("Select at least one column to encrypt."); return; }
    setEncRunning(true); setEncProgress(0); setEncError("");
    setEncResultBlob(null); setEncResultKey(null);
    try {
      const { blob, keyHex } = await encryptFWFToBlob(dataText, layoutResult.fields, encCols, buildOpts(), setEncProgress);
      setEncResultBlob(blob);
      setEncResultKey(keyHex);
      setStep("anon-done");
    } catch (e) { setEncError(`Encryption failed: ${(e as Error).message}`); }
    finally { setEncRunning(false); }
  }, [layoutResult, dataText, encCols, anonKeyMode, anonSeed, anonPassphrase, anonPbkdf2Iter, anonDeterministic, anonKeyHexInput]);

  const handleDownloadOriginal = useCallback(async () => {
    if (!layoutResult || !dataText) return;
    setOrigDownloading(true); setOrigProgress(0);
    try {
      const blob = await convertFWFToCSV(dataText, layoutResult.fields, { onProgress: setOrigProgress });
      triggerDownload(blob, `${outputBaseName}.csv`);
    } finally { setOrigDownloading(false); setOrigProgress(0); }
  }, [layoutResult, dataText, outputBaseName]);

  const handleCopyKey = (k: string) => {
    navigator.clipboard.writeText(k).then(() => {
      setKeyCopied(true); setTimeout(() => setKeyCopied(false), 2000);
    });
  };

  const handleDownloadKey = (k: string) => {
    const txt = [
      "AES-256-GCM Symmetric Key",
      "=".repeat(40),
      "",
      `Key (256-bit hex): ${k}`,
      "",
      `Key derivation: ${keyModeLabel}`,
      `Deterministic mode: ${anonDeterministic ? "ON" : "OFF"}`,
      `Generated: ${new Date().toISOString()}`,
      "",
      "IMPORTANT — Store in a secure vault (HSM, AWS KMS, etc.).",
      "This key is required to decrypt the anonymized CSV.",
      "AES-256-GCM is symmetric — the same key encrypts and decrypts.",
    ].join("\n");
    triggerDownload(new Blob([txt], { type: "text/plain" }), `aes256_key_${outputBaseName || "export"}.txt`);
  };

  const handleDecryptFile = useCallback(async (file: File) => {
    setDecryptError(""); setDecryptBlob(null); setDecryptFileName(file.name);
    setDecryptCsvText(null); setDecryptHeaders([]);
    const text = await file.text();
    const headers = readCSVHeaders(text);
    if (!headers.length) { setDecryptError("Could not read CSV headers — check the file."); return; }
    setDecryptCsvText(text);
    setDecryptHeaders(headers);
    setDecryptCols(new Set(headers));
  }, []);

  const handleDecrypt = useCallback(async () => {
    if (!decryptCsvText) { setDecryptError("Upload an encrypted CSV first."); return; }
    if (decryptCols.size === 0) { setDecryptError("Select at least one column to decrypt."); return; }
    if (anonKeyMode === "hex" && anonKeyHexInput.trim().length !== 64) {
      setDecryptError("Raw hex key must be exactly 64 hex characters."); return;
    }
    if (anonKeyMode === "pbkdf2" && !anonPassphrase.trim()) {
      setDecryptError("Enter the passphrase used during encryption."); return;
    }
    setDecryptRunning(true); setDecryptProgress(0); setDecryptError(""); setDecryptBlob(null);
    try {
      const blob = await decryptCSVToBlob(decryptCsvText, decryptCols, buildOpts(), setDecryptProgress);
      setDecryptBlob(blob);
    } catch (e) { setDecryptError(`Decryption failed: ${(e as Error).message}`); }
    finally { setDecryptRunning(false); }
  }, [decryptCsvText, decryptCols, anonKeyMode, anonSeed, anonPassphrase, anonPbkdf2Iter, anonDeterministic, anonKeyHexInput]);

  const handleReset = () => {
    setStep("layout"); setLayoutSubStep("upload");
    setLayoutResult(null); setLayoutFileName(""); setLayoutError("");
    setExcelInfo(null); setPendingFile(null); setSelectedSheet("");
    setRowFrom(""); setRowTo(""); setSheetRowCount(0);
    setDataFileName(""); setDataText(""); setDataLineCount(0);
    setDataError(""); setPreview([]);
    setEncCols(new Set()); setEncResultBlob(null); setEncResultKey(null); setEncError("");
    setDecryptFileName(""); setDecryptCsvText(null); setDecryptHeaders([]);
    setDecryptCols(new Set()); setDecryptBlob(null); setDecryptError("");
  };

  return (
    <div className="space-y-8">
      {/* Step indicator */}
      <div className="flex items-center gap-3 flex-wrap">
        {(["Upload layout", "Upload data file", "Convert", "Anonymize & download"] as const).map((label, idx) => {
          const n = idx + 1;
          const done =
            n === 1 ? step !== "layout" :
            n === 2 ? isConverted :
            n === 3 ? isConverted :
            step === "anon-done";
          const active =
            n === 1 ? step === "layout" :
            n === 2 ? step === "data" :
            n === 3 ? step === "data" && !isConverted :
            step === "converted";
          return (
            <span key={n} className="flex items-center gap-3">
              {idx > 0 && <ArrowRight className="w-4 h-4 text-gray-300 flex-shrink-0" />}
              <StepBadge n={n} label={label} active={active} done={done} />
            </span>
          );
        })}
      </div>

      <div className="grid grid-cols-1 gap-8 lg:grid-cols-[1fr_1fr] min-w-0">
        {/* ── Step 1: Layout ─────────────────────────────────────────────── */}
        <div className="border border-gray-200 rounded-2xl p-6 space-y-5 min-w-0 overflow-hidden">
          <div className="flex items-start justify-between">
            <div>
              <h2 className="text-lg font-semibold text-black">Step 1 — Layout file</h2>
              <p className="text-sm text-gray-500 mt-1">Excel (.xlsx) or CSV with Field_Name, Start, End columns</p>
            </div>
            {(layoutResult || layoutSubStep === "sheet-select") && (
              <button onClick={handleReset} className="text-gray-400 hover:text-black mt-1"><X className="w-5 h-5" /></button>
            )}
          </div>

          {layoutSubStep === "upload" && (
            <>
              <DropZone accept=".xlsx,.xls,.csv" icon={<FileSpreadsheet className="w-9 h-9 text-blue-600" />}
                label="Drop layout file here" sublabel="Excel or CSV"
                inputRef={layoutInputRef} onFile={handleLayoutFile} />
              {layoutError && <ErrorBox message={layoutError} />}
            </>
          )}

          {layoutSubStep === "sheet-select" && excelInfo && (
            <div className="space-y-5">
              <InfoBadge icon={<FileSpreadsheet className="w-4 h-4" />} text={`${layoutFileName} — ${excelInfo.sheetNames.length} sheets`} />
              <div className="space-y-2">
                <label className="text-sm font-semibold text-black flex items-center gap-2"><Layers className="w-4 h-4" />Select sheet</label>
                <div className="space-y-2 max-h-44 overflow-y-auto pr-1">
                  {excelInfo.sheetNames.map((name) => (
                    <label key={name} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl border cursor-pointer text-sm transition-colors ${selectedSheet === name ? "border-blue-500 bg-blue-50 text-black" : "border-gray-200 hover:border-blue-300 text-gray-500"}`}>
                      <input type="radio" name="sheet" value={name} checked={selectedSheet === name} onChange={() => handleSheetChange(name)} className="accent-blue-600" />
                      <span className="font-medium">{name}</span>
                    </label>
                  ))}
                </div>
              </div>
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <label className="text-sm font-semibold text-black">Row range {sheetRowCount > 0 && <span className="font-normal text-gray-500">({sheetRowCount} rows)</span>}</label>
                  {(rowFrom || rowTo) && <button onClick={() => { setRowFrom(""); setRowTo(""); }} className="text-sm text-gray-400 hover:text-black flex items-center gap-1"><RotateCcw className="w-3.5 h-3.5" />All rows</button>}
                </div>
                <div className="flex items-center gap-3">
                  {[{ label: "From", val: rowFrom, set: setRowFrom, ph: "1" }, { label: "To", val: rowTo, set: setRowTo, ph: sheetRowCount ? String(sheetRowCount) : "last" }].map(({ label, val, set, ph }, i) => (
                    <div key={i} className="flex-1 space-y-1">
                      <p className="text-sm text-gray-500">{label} row</p>
                      <input type="number" min={1} placeholder={ph} value={val} onChange={(e) => set(e.target.value)}
                        className="w-full px-3 py-2 text-sm rounded-xl border border-gray-200 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-black" />
                    </div>
                  ))}
                </div>
              </div>
              {layoutError && <ErrorBox message={layoutError} />}
              <div className="flex flex-col gap-2">
                <button onClick={handleConfirmSheet} disabled={applyingSheet || !selectedSheet}
                  className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-black text-white text-base font-semibold hover:bg-gray-800 disabled:opacity-50 transition-colors">
                  {applyingSheet ? <><Spin />Parsing…</> : <><ArrowRight className="w-4 h-4" />Use selected sheet</>}
                </button>
                <button onClick={handleAutoDetect} disabled={applyingSheet}
                  className="w-full flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl border border-gray-200 text-sm text-gray-500 hover:text-black hover:border-gray-400 disabled:opacity-50 transition-colors">
                  <Upload className="w-4 h-4" />Auto-detect layout sheet
                </button>
              </div>
            </div>
          )}

          {layoutSubStep === "done" && layoutResult && (
            <div className="space-y-4">
              <SuccessBadge text={`${layoutFileName} — ${fields.length} fields${layoutResult.sheetName ? ` (${layoutResult.sheetName})` : ""}`} />
              {layoutResult.warnings.length > 0 && <WarnBox message={layoutResult.warnings.join(" ")} />}
              <div className="overflow-auto max-h-72 rounded-xl border border-gray-200">
                <table className="w-full text-sm border-collapse">
                  <thead className="bg-gray-50 sticky top-0">
                    <tr>{["#", "Variable", "Full Name", "Start", "End", "Len"].map((h) => (
                      <th key={h} className="px-3 py-2.5 text-left border-r last:border-r-0 border-gray-200 text-gray-500 font-semibold">{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {fields.map((f) => (
                      <tr key={f.srlNo} className="border-t border-gray-100 hover:bg-gray-50">
                        <td className="px-3 py-2 text-gray-400 font-mono border-r border-gray-100">{f.srlNo}</td>
                        <td className="px-3 py-2 font-semibold text-black border-r border-gray-100 whitespace-nowrap">{f.varName}</td>
                        <td className="px-3 py-2 text-gray-600 border-r border-gray-100">{f.fullName}</td>
                        <td className="px-3 py-2 text-center font-mono text-black border-r border-gray-100">{f.start}</td>
                        <td className="px-3 py-2 text-center font-mono text-black border-r border-gray-100">{f.end}</td>
                        <td className="px-3 py-2 text-center font-mono text-black">{f.length}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* ── Step 2 + 3: Data & Convert ─────────────────────────────────── */}
        <div className="border border-gray-200 rounded-2xl p-6 space-y-5 min-w-0 overflow-hidden">
          <div>
            <h2 className="text-lg font-semibold text-black">Step 2 — Fixed-width data file</h2>
            <p className="text-sm text-gray-500 mt-1">The .TXT file containing the actual records</p>
          </div>

          {!layoutResult ? (
            <div className="flex items-center justify-center h-44 text-base text-gray-400 border-2 border-dashed border-gray-200 rounded-2xl">
              Complete Step 1 first
            </div>
          ) : !dataFileName ? (
            <DropZone accept=".txt,.dat,.fwf,.data" icon={<FileText className="w-9 h-9 text-blue-600" />}
              label="Drop fixed-width data file here" sublabel=".TXT, .DAT or any fixed-width file"
              inputRef={dataInputRef} onFile={handleDataFile} />
          ) : (
            <div className="space-y-5">
              <div className="flex items-center gap-2">
                <SuccessBadge text={`${dataFileName} — ${dataLineCount.toLocaleString()} records`} />
                {!isConverted && (
                  <button onClick={() => { setDataFileName(""); setDataText(""); setDataLineCount(0); setPreview([]); setStep("data"); }}
                    className="ml-auto text-gray-400 hover:text-black flex-shrink-0"><X className="w-4 h-4" /></button>
                )}
              </div>

              {preview.length > 0 && (
                <div className="space-y-2">
                  <p className="text-sm text-gray-500 flex items-center gap-1.5"><Eye className="w-4 h-4" />Preview (first {preview.length} rows)</p>
                  <div className="overflow-auto max-h-48 rounded-xl border border-gray-200 text-xs">
                    <table className="w-full border-collapse">
                      <thead className="bg-gray-50 sticky top-0">
                        <tr>{fields.map((f) => <th key={f.srlNo} className="px-3 py-2 text-left font-semibold text-gray-500 border-r border-gray-200 whitespace-nowrap">{f.varName}</th>)}</tr>
                      </thead>
                      <tbody>
                        {preview.map((row, ri) => (
                          <tr key={ri} className="border-t border-gray-100">
                            {row.map((cell, ci) => (
                              <td key={ci} className="px-3 py-1.5 font-mono border-r border-gray-100 whitespace-nowrap text-black">
                                {cell || <span className="text-gray-300 italic">—</span>}
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
                  className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-black text-white text-base font-semibold hover:bg-gray-800 disabled:opacity-50 transition-colors">
                  {converting ? <><Spin />Preparing…</> : <><ArrowRight className="w-4 h-4" />Convert {dataLineCount.toLocaleString()} records → proceed to anonymize</>}
                </button>
              ) : (
                <SuccessBadge text={`${dataLineCount.toLocaleString()} records ready · ${fields.length} columns · see Step 4 below ↓`} />
              )}
            </div>
          )}
        </div>
      </div>

      {/* ── Step 4: Anonymize ─────────────────────────────────────────────── */}
      {isConverted && layoutResult && (
        <div className="border border-gray-200 rounded-2xl p-6 space-y-6 min-w-0 overflow-hidden">
          <div className="flex flex-col sm:flex-row sm:items-center gap-4">
            <div className="flex items-center gap-3 flex-1 min-w-0">
              <ShieldCheck className="w-6 h-6 text-blue-600 flex-shrink-0" />
              <div className="min-w-0">
                <h2 className="text-lg font-semibold text-black">Step 4 — AES-256-GCM Encrypt / Decrypt</h2>
                <p className="text-sm text-gray-500 mt-0.5">Format-preserving: digits→digits, letters→letters</p>
              </div>
            </div>
            <div className="flex items-center rounded-xl border border-gray-200 overflow-hidden text-sm font-semibold flex-shrink-0">
              {(["encrypt", "decrypt"] as const).map((m) => (
                <button key={m} onClick={() => { setAnonMode(m); setEncError(""); setDecryptError(""); }}
                  className={`flex items-center gap-2 px-4 py-2.5 transition-colors ${m !== "encrypt" ? "border-l border-gray-200" : ""} ${anonMode === m ? "bg-black text-white" : "hover:bg-gray-50 text-gray-500"}`}>
                  {m === "encrypt" ? <Lock className="w-4 h-4" /> : <LockOpen className="w-4 h-4" />}
                  {m === "encrypt" ? "Encrypt" : "Decrypt"}
                </button>
              ))}
            </div>
          </div>

          <KeySettings
            keyMode={anonKeyMode} setKeyMode={setAnonKeyMode}
            seed={anonSeed} setSeed={setAnonSeed}
            passphrase={anonPassphrase} setPassphrase={setAnonPassphrase}
            pbkdf2Iter={anonPbkdf2Iter} setPbkdf2Iter={setAnonPbkdf2Iter}
            deterministic={anonDeterministic} setDeterministic={setAnonDeterministic}
            keyHexInput={anonKeyHexInput} setKeyHexInput={setAnonKeyHexInput}
          />

          {/* ENCRYPT */}
          {anonMode === "encrypt" && (
            <div className="space-y-5">
              <ColSelector allCols={allColNames} selected={encCols} onChange={setEncCols} label="Columns to encrypt" />

              {encError && <ErrorBox message={encError} />}
              {encRunning && <ProgressBar pct={encProgress} label={`Encrypting ${encCols.size} column${encCols.size !== 1 ? "s" : ""} across ${dataLineCount.toLocaleString()} records…`} icon={<Shuffle className="w-4 h-4 animate-spin" />} />}
              {origDownloading && <ProgressBar pct={origProgress} label="Building original CSV…" icon={<Download className="w-4 h-4 animate-pulse" />} />}

              {step !== "anon-done" && (
                <div className="flex flex-col sm:flex-row gap-3">
                  <button onClick={handleEncrypt} disabled={encRunning || origDownloading || encCols.size === 0}
                    className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-black text-white text-base font-semibold hover:bg-gray-800 disabled:opacity-50 transition-colors">
                    {encRunning ? <><Spin />Encrypting…</> : <><Lock className="w-4 h-4" />Apply AES-256-GCM encryption</>}
                  </button>
                  <button onClick={handleDownloadOriginal} disabled={encRunning || origDownloading}
                    className="flex items-center justify-center gap-2 px-4 py-3 rounded-xl border border-gray-200 text-sm text-gray-500 hover:text-black hover:border-gray-400 disabled:opacity-50 transition-colors">
                    {origDownloading ? <><Spin />Building…</> : <><Download className="w-4 h-4" />Skip — download original</>}
                  </button>
                </div>
              )}

              {step === "anon-done" && encResultBlob && encResultKey && (
                <div className="space-y-5">
                  <SuccessBadge text={`Encryption complete — ${encCols.size} column${encCols.size !== 1 ? "s" : ""} encrypted across ${dataLineCount.toLocaleString()} records`} />

                  <div className="border-l-4 border-amber-400 bg-amber-50 rounded-r-xl p-5 space-y-3">
                    <p className="text-sm font-semibold text-amber-800 flex items-center gap-2"><Key className="w-4 h-4" />Symmetric Key — save this to decrypt later</p>
                    <div className="font-mono text-xs bg-white rounded-lg px-4 py-3 break-all select-all cursor-text leading-relaxed text-black border border-amber-200">
                      {encResultKey}
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-sm text-amber-700 flex-1 min-w-0">AES-256 · {keyModeLabel} · det. {anonDeterministic ? "ON" : "OFF"}</span>
                      <button onClick={() => handleCopyKey(encResultKey)} className="text-sm px-3 py-1.5 rounded-lg border border-amber-300 text-amber-800 hover:bg-amber-100 transition-colors whitespace-nowrap font-medium">
                        {keyCopied ? "✓ Copied!" : "Copy key"}
                      </button>
                      <button onClick={() => handleDownloadKey(encResultKey)} className="flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-lg border border-amber-300 text-amber-800 hover:bg-amber-100 transition-colors whitespace-nowrap font-medium">
                        <Download className="w-3.5 h-3.5" />Download key (.txt)
                      </button>
                    </div>
                    <p className="text-sm text-amber-700">⚠ Same key decrypts. Store in a secure vault — never log or share.</p>
                  </div>

                  <div className="flex flex-col sm:flex-row gap-3">
                    <button onClick={() => triggerDownload(encResultBlob!, `${outputBaseName}_anonymized.csv`)}
                      className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-emerald-600 text-white text-base font-semibold hover:bg-emerald-700 transition-colors">
                      <Download className="w-4 h-4" />Download anonymized CSV
                    </button>
                    <button onClick={handleDownloadOriginal} disabled={origDownloading}
                      className="flex items-center justify-center gap-2 px-4 py-3 rounded-xl border border-gray-200 text-sm text-gray-500 hover:text-black hover:border-gray-400 disabled:opacity-50 transition-colors">
                      {origDownloading ? <><Spin />Building…</> : <><Download className="w-4 h-4" />Download original CSV</>}
                    </button>
                  </div>
                  {origDownloading && <ProgressBar pct={origProgress} label="Building original CSV…" icon={<Download className="w-4 h-4 animate-pulse" />} />}

                  <button onClick={() => { setEncResultBlob(null); setEncResultKey(null); setEncProgress(0); setStep("converted"); }}
                    className="w-full text-sm text-gray-400 hover:text-black text-center transition-colors">
                    ← Change column selection or key settings
                  </button>
                </div>
              )}
            </div>
          )}

          {/* DECRYPT */}
          {anonMode === "decrypt" && (
            <div className="space-y-5">
              <div className="space-y-2">
                <p className="text-base font-semibold text-black">Upload anonymized CSV to decrypt</p>
                <p className="text-sm text-gray-500">Must have been encrypted by this tool with matching key settings.</p>

                {!decryptCsvText ? (
                  <DropZone accept=".csv" icon={<LockOpen className="w-9 h-9 text-blue-600" />}
                    label="Drop anonymized CSV here" sublabel=".CSV encrypted by this tool"
                    inputRef={decryptInputRef} onFile={handleDecryptFile} />
                ) : (
                  <div className="space-y-4">
                    <div className="flex items-center gap-2">
                      <SuccessBadge text={`${decryptFileName} — ${decryptHeaders.length} columns detected`} />
                      <button onClick={() => { setDecryptFileName(""); setDecryptCsvText(null); setDecryptHeaders([]); setDecryptCols(new Set()); setDecryptBlob(null); }}
                        className="ml-auto text-gray-400 hover:text-black flex-shrink-0"><X className="w-4 h-4" /></button>
                    </div>
                    <ColSelector allCols={decryptHeaders} selected={decryptCols} onChange={setDecryptCols} label="Columns to decrypt" />
                  </div>
                )}
              </div>

              {decryptError && <ErrorBox message={decryptError} />}
              {decryptRunning && <ProgressBar pct={decryptProgress} label={`Decrypting ${decryptCols.size} column${decryptCols.size !== 1 ? "s" : ""}…`} icon={<Shuffle className="w-4 h-4 animate-spin" />} />}

              {!decryptBlob ? (
                <button onClick={handleDecrypt} disabled={decryptRunning || !decryptCsvText || decryptCols.size === 0}
                  className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-black text-white text-base font-semibold hover:bg-gray-800 disabled:opacity-50 transition-colors">
                  {decryptRunning ? <><Spin />Decrypting…</> : <><LockOpen className="w-4 h-4" />Apply AES-256-GCM decryption</>}
                </button>
              ) : (
                <div className="space-y-4">
                  <SuccessBadge text="Decryption complete — original values restored" />
                  <div className="flex flex-col sm:flex-row gap-3">
                    <button onClick={() => triggerDownload(decryptBlob!, `${decryptFileName.replace(/\.csv$/i, "")}_decrypted.csv`)}
                      className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-emerald-600 text-white text-base font-semibold hover:bg-emerald-700 transition-colors">
                      <Download className="w-4 h-4" />Download decrypted CSV
                    </button>
                    <button onClick={() => { setDecryptBlob(null); setDecryptProgress(0); }}
                      className="flex items-center justify-center gap-2 px-4 py-3 rounded-xl border border-gray-200 text-sm text-gray-500 hover:text-black hover:border-gray-400 transition-colors">
                      ← Change settings
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ── Key Settings ──────────────────────────────────────────────────────────────

function KeySettings({ keyMode, setKeyMode, seed, setSeed, passphrase, setPassphrase, pbkdf2Iter, setPbkdf2Iter, deterministic, setDeterministic, keyHexInput, setKeyHexInput }: {
  keyMode: "random" | "pbkdf2" | "hex"; setKeyMode: (m: "random" | "pbkdf2" | "hex") => void;
  seed: number; setSeed: (n: number) => void;
  passphrase: string; setPassphrase: (s: string) => void;
  pbkdf2Iter: number; setPbkdf2Iter: (n: number) => void;
  deterministic: boolean; setDeterministic: (b: boolean) => void;
  keyHexInput: string; setKeyHexInput: (s: string) => void;
}) {
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 pt-4 border-t border-gray-100">
      <div className="space-y-3">
        <p className="text-sm font-semibold text-black flex items-center gap-2"><Key className="w-4 h-4" />Key derivation</p>
        <div className="space-y-2">
          {(["random", "pbkdf2", "hex"] as const).map((m) => (
            <label key={m} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl border cursor-pointer text-sm transition-colors ${keyMode === m ? "border-blue-500 bg-blue-50 text-black" : "border-gray-200 hover:border-blue-300 text-gray-500"}`}>
              <input type="radio" name="keymode" checked={keyMode === m} onChange={() => setKeyMode(m)} className="accent-blue-600" />
              {m === "random" ? "Random (seed)" : m === "pbkdf2" ? "PBKDF2 passphrase" : "Paste hex key"}
            </label>
          ))}
        </div>
      </div>

      <div className="space-y-3">
        <p className="text-sm font-semibold text-black">{keyMode === "random" ? "Key seed" : keyMode === "pbkdf2" ? "Passphrase" : "256-bit hex key"}</p>
        {keyMode === "random" && (
          <>
            <input type="number" value={seed} onChange={(e) => setSeed(Number(e.target.value))}
              className="w-full px-3 py-2.5 text-sm font-mono rounded-xl border border-gray-200 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-black" />
            <p className="text-sm text-gray-500">Same seed → same key (reproducible)</p>
          </>
        )}
        {keyMode === "pbkdf2" && (
          <div className="space-y-3">
            <input type="password" value={passphrase} onChange={(e) => setPassphrase(e.target.value)} placeholder="Enter passphrase…"
              className="w-full px-3 py-2.5 text-sm rounded-xl border border-gray-200 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-black" />
            <div>
              <p className="text-sm text-gray-500 mb-1">Iterations: {pbkdf2Iter.toLocaleString()}</p>
              <input type="range" min={10000} max={500000} step={10000} value={pbkdf2Iter} onChange={(e) => setPbkdf2Iter(Number(e.target.value))} className="w-full accent-blue-600" />
            </div>
          </div>
        )}
        {keyMode === "hex" && (
          <div className="space-y-2">
            <textarea value={keyHexInput} onChange={(e) => setKeyHexInput(e.target.value)} placeholder="Paste 64-char hex key…" rows={2}
              className={`w-full px-3 py-2.5 text-xs font-mono rounded-xl border bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none text-black ${keyHexInput && keyHexInput.trim().length !== 64 ? "border-red-400" : "border-gray-200"}`} />
            <p className={`text-sm ${keyHexInput.trim().length === 64 ? "text-emerald-600" : "text-gray-500"}`}>
              {keyHexInput.trim().length === 64 ? "✓ Valid 256-bit key" : `${keyHexInput.trim().length}/64 hex chars`}
            </p>
          </div>
        )}
      </div>

      <div className="space-y-3">
        <label className={`flex items-start gap-3 px-4 py-3 rounded-xl border cursor-pointer text-sm transition-colors ${deterministic ? "border-blue-500 bg-blue-50 text-black" : "border-gray-200 hover:border-blue-300 text-gray-500"}`}>
          <input type="checkbox" checked={deterministic} onChange={(e) => setDeterministic(e.target.checked)} className="accent-blue-600 w-4 h-4 mt-0.5 flex-shrink-0" />
          <div>
            <p className="font-semibold">Deterministic mode</p>
            <p className="text-xs mt-1 opacity-70">Same value → same output. Required for consistent round-trip.</p>
          </div>
        </label>
        <div className="space-y-1 text-sm text-gray-500">
          {[["Cipher", "AES-256-GCM"], ["Key", "256-bit"], ["IV", "96-bit"], ["Tag", "128-bit GHASH"], ["Std", "NIST FIPS 197"]].map(([k, v]) => (
            <div key={k} className="flex gap-3"><span className="font-semibold text-black w-10 shrink-0">{k}</span><span>{v}</span></div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Column selector ───────────────────────────────────────────────────────────

function ColSelector({ allCols, selected, onChange, label }: { allCols: string[]; selected: Set<string>; onChange: (s: Set<string>) => void; label: string }) {
  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold text-black">{label} <span className="font-normal text-gray-500">({selected.size}/{allCols.length})</span></span>
        <div className="flex gap-2">
          <button onClick={() => onChange(new Set(allCols))} className="text-sm px-3 py-1 rounded-lg border border-gray-200 hover:border-gray-400 text-gray-500 hover:text-black transition-colors">Select all</button>
          <button onClick={() => onChange(new Set())} className="text-sm px-3 py-1 rounded-lg border border-gray-200 hover:border-gray-400 text-gray-500 hover:text-black transition-colors">Clear</button>
        </div>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-5 gap-2 max-h-52 overflow-y-auto pr-1 pt-1">
        {allCols.map((col) => (
          <label key={col} className={`flex items-center gap-2 px-3 py-2 rounded-xl border cursor-pointer text-sm transition-colors ${selected.has(col) ? "border-blue-500 bg-blue-50 text-black" : "border-gray-200 hover:border-blue-300 text-gray-500 hover:text-black"}`}>
            <input type="checkbox" checked={selected.has(col)} onChange={(e) => { const n = new Set(selected); if (e.target.checked) n.add(col); else n.delete(col); onChange(n); }} className="accent-blue-600 w-3.5 h-3.5 flex-shrink-0" />
            <span className="truncate font-mono text-xs">{col}</span>
          </label>
        ))}
      </div>
    </div>
  );
}

// ── Small reusable components ─────────────────────────────────────────────────

function ProgressBar({ pct, label, icon }: { pct: number; label: string; icon?: React.ReactNode }) {
  return (
    <div className="space-y-2">
      <div className="flex justify-between text-sm text-gray-500">
        <span className="flex items-center gap-2 min-w-0 truncate">{icon}{label}</span>
        <span className="flex-shrink-0 ml-2 font-semibold text-black">{pct}%</span>
      </div>
      <div className="h-2 rounded-full bg-gray-100 overflow-hidden">
        <div className="h-full bg-black rounded-full transition-all duration-200" style={{ width: `${pct}%` }} />
      </div>
    </div>
  );
}

function StepBadge({ n, label, active, done }: { n: number; label: string; active: boolean; done: boolean }) {
  return (
    <div className={`flex items-center gap-2 text-sm font-semibold ${active ? "text-black" : done ? "text-emerald-600" : "text-gray-400"}`}>
      <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 ${active ? "bg-black text-white" : done ? "bg-emerald-100 text-emerald-600" : "bg-gray-100 text-gray-400"}`}>
        {done ? <CheckCircle2 className="w-3.5 h-3.5" /> : n}
      </span>
      {label}
    </div>
  );
}

function DropZone({ accept, icon, label, sublabel, inputRef, onFile }: { accept: string; icon: React.ReactNode; label: string; sublabel: string; inputRef: React.RefObject<HTMLInputElement | null>; onFile: (f: File) => void }) {
  const [dragging, setDragging] = useState(false);
  return (
    <div className={`border-2 border-dashed rounded-2xl p-10 text-center cursor-pointer transition-all ${dragging ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-blue-400 hover:bg-gray-50"}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
      onClick={() => inputRef.current?.click()}>
      <input ref={inputRef} type="file" accept={accept} className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />
      <div className="flex flex-col items-center gap-4">
        <div className="w-16 h-16 rounded-2xl bg-blue-50 flex items-center justify-center">{icon}</div>
        <div>
          <p className="text-base font-semibold text-black">{label}</p>
          <p className="text-sm text-gray-500 mt-1">{sublabel}</p>
        </div>
        <span className="text-sm px-4 py-2 rounded-xl border border-gray-200 bg-white text-gray-500 font-medium">Browse</span>
      </div>
    </div>
  );
}

function SuccessBadge({ text }: { text: string }) {
  return (
    <div className="flex items-center gap-2.5 text-sm text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-xl px-4 py-2.5 font-medium">
      <CheckCircle2 className="w-4 h-4 flex-shrink-0" /><span>{text}</span>
    </div>
  );
}

function InfoBadge({ icon, text }: { icon: React.ReactNode; text: string }) {
  return (
    <div className="flex items-center gap-2.5 text-sm text-blue-700 bg-blue-50 border border-blue-200 rounded-xl px-4 py-2.5 font-medium">
      {icon}<span>{text}</span>
    </div>
  );
}

function WarnBox({ message }: { message: string }) {
  return (
    <div className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-xl px-4 py-2.5 font-medium">{message}</div>
  );
}

function ErrorBox({ message }: { message: string }) {
  return (
    <div className="flex items-start gap-2.5 text-sm text-red-700 bg-red-50 border border-red-200 rounded-xl px-4 py-2.5 font-medium">
      <AlertTriangle className="w-4 h-4 flex-shrink-0 mt-0.5" />{message}
    </div>
  );
}

function Spin() {
  return <div className="w-4 h-4 border-2 border-current/30 border-t-current rounded-full animate-spin flex-shrink-0" />;
}
