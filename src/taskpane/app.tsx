// src/taskpane/App.tsx
import React, { useEffect, useState } from "react";
import {
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
} from "@fluentui/react";
import * as XLSX from "xlsx";
import {
  compareSheetRows, ChangeItem, ChangeType, COLORS as ENGINE_COLORS,
} from "./compareEngine";
import { requestTrialToken, validateTrialToken } from "./firebaseClient";

type Change = ChangeItem;

const COLORS = ENGINE_COLORS;

export default function App() {
  const [file2, setFile2] = useState<File | null>(null);
  const [changes, setChanges] = useState<Change[]>([]);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [trialUses, setTrialUses] = useState<number>(0);
  const [loading, setLoading] = useState(false);
  const [summary, setSummary] = useState("");
  const [error, setError] = useState("");

  const TRIAL_TOKEN_KEY = "excelDiffTrialToken";

  useEffect(() => {
    (async () => {
      const token = localStorage.getItem(TRIAL_TOKEN_KEY);
      if (token) {
        try {
          const data = await validateTrialToken(token, false); // just validate, don't consume
          if (data.valid) setTrialUses(data.usesLeft);
          else await obtainTrialToken();
        } catch {
          // network issues -> try to obtain trial (best-effort)
          await obtainTrialToken();
        }
      } else {
        await obtainTrialToken();
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function obtainTrialToken() {
    try {
      const data = await requestTrialToken("");
      if (data && data.token) {
        localStorage.setItem(TRIAL_TOKEN_KEY, data.token);
        setTrialUses(data.usesLeft ?? 0);
      } else {
        setError("Could not obtain trial token from server.");
      }
    } catch (err) {
      console.error(err);
      setError("Failed to request trial token.");
    }
  }

  async function validateTrialAndConsume(): Promise<boolean> {
    const token = localStorage.getItem(TRIAL_TOKEN_KEY);
    if (!token) {
      await obtainTrialToken();
      return false;
    }
    try {
      const data = await validateTrialToken(token, true); // consume one use
      if (data.valid) {
        setTrialUses(data.usesLeft);
        if ((data.usesLeft ?? 0) <= 0) setError("Free trial exhausted; please upgrade.");
        return (data.usesLeft ?? 0) > 0;
      } else {
        // token invalid -> get a new one
        localStorage.removeItem(TRIAL_TOKEN_KEY);
        await obtainTrialToken();
        return false;
      }
    } catch (err) {
      console.error("validateTrial error", err);
      setError("Unable to validate trial (network).");
      return false;
    }
  }

  const handleFile2 = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setFile2(f);
  };

  const parseFileToSheets = async (file: File) => {
    const name = file.name.toLowerCase();
    if (name.endsWith(".csv")) {
      const txt = await file.text();
      return XLSX.read(txt, { type: "string", raw: true });
    } else {
      const ab = await file.arrayBuffer();
      return XLSX.read(ab, { type: "array", raw: true });
    }
  };

  const compareWithOpenWorkbook = async () => {
    if (!file2) {
      setError("Select a modified file (File 2) first.");
      return;
    }

    // validate trial and decrement one use BEFORE expensive work
    setLoading(true);
    setError("");
    setChanges([]);
    setSummary("");

    const trialOk = await validateTrialAndConsume();
    if (!trialOk) {
      setLoading(false);
      return;
    }

    try {
      const wb2 = await parseFileToSheets(file2);
      const allChanges: Change[] = [];
      const sheetFormatRequests: Record<string, { addresses: string[]; color: string }[]> = {};

      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const openSheets = workbook.worksheets;
        openSheets.load("items/name");
        await context.sync();
        const openSheetNames = openSheets.items.map((s) => s.name);

        // Sheet-level detection: added / removed
        const uploadedNames = wb2.SheetNames.slice();
        const sheetsAdded = uploadedNames.filter((n) => !openSheetNames.includes(n));
        const sheetsRemoved = openSheetNames.filter((n) => !uploadedNames.includes(n));
        for (const s of sheetsAdded) {
          allChanges.push({
            sheet: s,
            address: "",
            row: -1,
            col: -1,
            oldVal: "",
            newVal: `Sheet "${s}" added`,
            type: "added",
          });
        }
        for (const s of sheetsRemoved) {
          allChanges.push({
            sheet: s,
            address: "",
            row: -1,
            col: -1,
            oldVal: `Sheet "${s}" removed`,
            newVal: "",
            type: "deleted",
          });
        }

        // Only deeply diff common sheets
        const commonSheets = uploadedNames.filter((n) => openSheetNames.includes(n));

        for (const sheetName of commonSheets) {
          const ws2 = wb2.Sheets[sheetName];
          if (!ws2) continue;

          // read used range once for open sheet
          const openSheet = workbook.worksheets.getItem(sheetName);
          const used = openSheet.getUsedRangeOrNullObject();
          used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
          await context.sync();

          // convert open sheet to values/formulas arrays (may be empty)
          const baseRowIndex = used.isNullObject ? 0 : (used.rowIndex || 0);
          const baseColIndex = used.isNullObject ? 0 : (used.columnIndex || 0);
          const rowCount = used.isNullObject ? 0 : (used.rowCount || 0);
          const colCount = used.isNullObject ? 0 : (used.columnCount || 0);
          const values1 = (used.isNullObject ? [] : (used.values as any[][])) ?? [];
          const formulas1 = (used.isNullObject ? [] : (used.formulas as any[][])) ?? [];

          // uploaded sheet -> rows2
          const rows2 = ws2["!ref"] ? (XLSX.utils.sheet_to_json(ws2, { header: 1, defval: "" }) as any[][]) : [];

          // call compareEngine to get changes for this sheet (pure computation)
          const sheetChanges = compareSheetRows({
            sheetName,
            rowsOpen: { baseRowIndex, baseColIndex, values: values1, formulas: formulas1 },
            rowsUploaded: rows2,
          });

          // accumulate and prepare formatting requests
          for (const ch of sheetChanges) {
            allChanges.push(ch);
            // sheet-format grouping
            if (!sheetFormatRequests[ch.sheet]) sheetFormatRequests[ch.sheet] = [];
            sheetFormatRequests[ch.sheet].push({ addresses: [ch.address], color: COLORS[ch.type] });
          }
        } // end sheets loop

        // Apply highlights grouped by color and chunked
        for (const [sheetName, reqs] of Object.entries(sheetFormatRequests)) {
          const sheet = workbook.worksheets.getItem(sheetName);
          // group addresses by color
          const byColor: Record<string, string[]> = {};
          for (const r of reqs) {
            const color = r.color;
            if (!byColor[color]) byColor[color] = [];
            byColor[color].push(...r.addresses);
          }
          const CHUNK = 300;
          for (const color of Object.keys(byColor)) {
            const addrs = byColor[color];
            for (let i = 0; i < addrs.length; i += CHUNK) {
              const chunk = addrs.slice(i, i + CHUNK);
              const range = sheet.getRange(chunk.join(","));
              range.format.fill.color = color;
            }
          }
        }

        await context.sync();
      }); // end Excel.run

      // Build UI summary
      const added = allChanges.filter((c) => c.type === "added").length;
      const deleted = allChanges.filter((c) => c.type === "deleted").length;
      const modified = allChanges.filter((c) => c.type === "modified").length;
      setChanges(allChanges);
      setCurrentIndex(allChanges.length > 0 ? 0 : -1);
      setSummary(`${allChanges.length} changes — ${added} added • ${deleted} deleted • ${modified} modified`);
    } catch (err: any) {
      console.error(err);
      setError("Comparison failed: " + (err?.message ?? "Unknown error"));
    } finally {
      setLoading(false);
    }
  };

  const goToChange = async (idx: number) => {
    if (idx < 0 || idx >= changes.length) return;
    setCurrentIndex(idx);
    const ch = changes[idx];
    await Excel.run(async (context) => {
      if (!ch.address) return;
      const sheet = context.workbook.worksheets.getItem(ch.sheet);
      const range = sheet.getRange(ch.address);
      range.select();
      sheet.activate();
      await context.sync();
    });
  };

  const clearHighlights = async () => {
    setLoading(true);
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        for (const s of sheets.items) {
          const used = s.getUsedRangeOrNullObject(true);
          used.load("address");
          await context.sync();
          if (!used.isNullObject) used.format.fill.clear();
        }
        await context.sync();
      });
      setChanges([]);
      setCurrentIndex(-1);
      setSummary("");
      setError("");
    } catch (e: any) {
      console.error(e);
      setError("Failed to clear highlights: " + (e?.message ?? ""));
    } finally {
      setLoading(false);
    }
  };

  return (
    <Stack tokens={{ padding: 20, childrenGap: 12 }} style={{ width: "100%", maxWidth: 640 }}>
      <Text variant="xxLarge">Excel Visual Diff</Text>

      <MessageBar messageBarType={MessageBarType.warning}>
        Open the original file in Excel (File 1), then upload the modified file (File 2) below.
      </MessageBar>

      {trialUses <= 0 ? (
        <MessageBar messageBarType={MessageBarType.error}>
          Free trial expired • <a href="https://yourdomain.com/pro" target="_blank" rel="noreferrer">Go Pro → $12/mo</a>
        </MessageBar>
      ) : (
        <MessageBar messageBarType={MessageBarType.info}>
          {trialUses} free comparisons remaining
        </MessageBar>
      )}

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError("")}>
          {error}
        </MessageBar>
      )}

      <Text variant="medium">Upload modified file (File 2)</Text>
      <input type="file" accept=".xlsx,.xlsm,.csv" onChange={handleFile2} />
      <Text variant="small">{file2?.name || "No file selected"}</Text>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton onClick={compareWithOpenWorkbook} disabled={!file2 || loading || trialUses <= 0}>
          {loading ? "Analyzing..." : "Highlight Changes"}
        </PrimaryButton>
        <DefaultButton onClick={clearHighlights} disabled={loading || changes.length === 0}>
          Clear Highlights
        </DefaultButton>
      </Stack>

      {loading && <ProgressIndicator label="Comparing files..." />}

      {summary && <MessageBar messageBarType={MessageBarType.success}>{summary}</MessageBar>}

      {changes.length > 0 && (
        <>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <DefaultButton onClick={() => goToChange(currentIndex - 1)} disabled={currentIndex <= 0}>
              ← Previous
            </DefaultButton>
            <Text>{currentIndex + 1} / {changes.length}</Text>
            <DefaultButton onClick={() => goToChange(currentIndex + 1)} disabled={currentIndex >= changes.length - 1}>
              Next →
            </DefaultButton>
          </Stack>

          <Text variant="small">Tip: select a difference to jump to it inside Excel.</Text>
        </>
      )}
    </Stack>
  );
}
