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
  Link,
  Icon,
} from "@fluentui/react";
import {
  Road20Regular,
  CheckmarkCircle20Filled,
  ErrorCircle20Filled,
  Warning20Filled,
} from "@fluentui/react-icons";
import * as XLSX from "xlsx";
import {
  compareSheetRows,
  ChangeItem,
  ChangeType,
  COLORS as ENGINE_COLORS,
} from "./compareEngine";
import { requestTrialToken, validateTrialToken } from "./firebaseClient";

type Change = ChangeItem & {
  navigable?: boolean;
};

const COLORS = ENGINE_COLORS;
const TRIAL_TOKEN_KEY = "excelDiffTrialToken";

export default function App() {
  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);
  const [changes, setChanges] = useState<Change[]>([]);
  const changeMap = React.useMemo(() => {
  const map = new Map<string, Change>();
  changes.forEach(c => map.set(c.address, c));
  return map;
}, [changes]);

const changeIndexMap = React.useMemo(() => {
  const map = new Map<string, number>();
  changes.forEach((c, i) => map.set(c.address, i));
  return map;
}, [changes]);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [selectedChange, setSelectedChange] = useState<Change | null>(null);
  const [trialUses, setTrialUses] = useState<number | null>(null);
  const [loading, setLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState("Preparing...");
  const [summary, setSummary] = useState("");
  const [error, setError] = useState("");

  // Inicjalizacja trial token (bez zmian)
  useEffect(() => {
    (async () => {
      const token = localStorage.getItem(TRIAL_TOKEN_KEY);
      if (token) {
        try {
          const data = await validateTrialToken(token, false);
          if (data.valid) setTrialUses(data.usesLeft);
          else await obtainTrialToken();
        } catch {
          await obtainTrialToken();
        }
      } else {
        await obtainTrialToken();
      }
    })();
  }, []);

  async function obtainTrialToken() {
    try {
      const data = await requestTrialToken("");
      if (data?.token) {
        localStorage.setItem(TRIAL_TOKEN_KEY, data.token);
        setTrialUses(data.usesLeft ?? 0);
      } else {
        setError("Could not connect to licensing server.");
      }
    } catch (err) {
      console.error(err);
      setError("Failed to obtain trial. Check internet connection.");
    }
  }

  async function validateTrialAndConsume(): Promise<boolean> {
    const token = localStorage.getItem(TRIAL_TOKEN_KEY);
    if (!token) {
      await obtainTrialToken();
      return false;
    }

    try {
      const data = await validateTrialToken(token, true);
      if (data.valid) {
        setTrialUses(data.usesLeft);
        if (data.usesLeft <= 0) {
          setError("Free trial exhausted — upgrade to continue.");
        }
        return data.usesLeft > 0;
      } else {
        localStorage.removeItem(TRIAL_TOKEN_KEY);
        await obtainTrialToken();
        return false;
      }
    } catch (err) {
      console.error(err);
      setError("License check failed (offline?). Try again later.");
      return false;
    }
  }

  const handleFile1 = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setFile1(f);
  };

  const handleFile2 = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setFile2(f);
  };

  const parseFileToSheets = async (file: File): Promise<XLSX.WorkBook> => {
    const name = file.name.toLowerCase();
    if (name.endsWith(".csv")) {
      const txt = await file.text();
      return XLSX.read(txt, { type: "string", raw: true });
    } else {
      const ab = await file.arrayBuffer();
      return XLSX.read(ab, { type: "array", raw: true });
    }
  };

  const compareFiles = async () => {
    if (!file1 || !file2) {
      setError("Please select both File 1 (original) and File 2 (modified).");
      return;
    }

    setLoading(true);
    setLoadingMessage("Validating license...");
    setError("");
    setChanges([]);
    setSummary("");
    setSelectedChange(null);
    setCurrentIndex(-1);

    // Odkomentuj w produkcji
    // const trialOk = await validateTrialAndConsume();
    const trialOk = true;

    if (!trialOk) {
      setLoading(false);
      return;
    }

    try {
      setLoadingMessage("Parsing File 1 (original)...");
      const wb1 = await parseFileToSheets(file1);

      setLoadingMessage("Parsing File 2 (modified)...");
      const wb2 = await parseFileToSheets(file2);

      // Prosty check rozmiaru plików
      const totalCells1 = wb1.SheetNames.reduce((sum, name) => {
        const ws = wb1.Sheets[name];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        return sum + (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
      }, 0);

      const totalCells2 = wb2.SheetNames.reduce((sum, name) => {
        const ws = wb2.Sheets[name];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        return sum + (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
      }, 0);

      if (totalCells1 + totalCells2 > 50000) {
        setError("Files are large – comparison may be slow. Limited to first 50k cells per file.");
      }

      setLoadingMessage("Comparing sheets...");
      const allChanges: Change[] = [];

      const sheetNames1 = wb1.SheetNames;
      const sheetNames2 = wb2.SheetNames;
      const commonSheets = sheetNames1.filter(name => sheetNames2.includes(name));

      for (const sheetName of commonSheets) {
        const ws1 = wb1.Sheets[sheetName];
        const ws2 = wb2.Sheets[sheetName];

        const rows1 = XLSX.utils.sheet_to_json(ws1, { header: 1, raw: false, defval: "" });
        const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, raw: false, defval: "" });

        const changesInSheet = compareSheetRows({
          sheetName,
          rowsOpen: {
            baseRowIndex: 0,
            baseColIndex: 0,
            values: rows1 as any[][],
            formulas: [], // brak formuł przy uploadzie – do poprawy później
          },
          rowsUploaded: rows2 as any[][],
        });

        allChanges.push(...changesInSheet.map(ch => ({ ...ch, navigable: false })));
      }

      setChanges(allChanges);
      setSummary(`Found ${allChanges.length} changes across ${commonSheets.length} sheets.`);

      // Jeśli chcesz highlight w otwartym workbooku – dodaj logikę później
      // na razie tylko lista zmian

    } catch (err: any) {
      setError("Comparison failed: " + (err.message || "Unknown error"));
      console.error(err);
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  const clearHighlights = async () => {
    setLoading(true);
    setLoadingMessage("Clearing highlights...");
    try {
      await Excel.run(async (context: Excel.RequestContext) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

const usedRanges: Excel.Range[] = [];

for (const s of sheets.items) {
  const used = s.getUsedRangeOrNullObject(true);
  usedRanges.push(used);
  used.load("isNullObject");
}

await context.sync();

for (const used of usedRanges) {
  if (!used.isNullObject) {
    used.format.fill.clear();
  }
}

await context.sync();

      setChanges([]);
      setCurrentIndex(-1);
      setSelectedChange(null);
      setSummary("");
      setError("");
    } catch (e: any) {
      setError("Failed to clear highlights: " + (e?.message || ""));
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  const goToChange = (index: number) => {
    if (index < 0 || index >= changes.length) return;
    setCurrentIndex(index);
    const ch = changes[index];
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(ch.sheet);
      await context.sync();
      if (sheet.isNullObject) return;

      const range = sheet.getRange(ch.address);
      range.select();
      await context.sync();
    }).catch((err) => console.error("Go to change failed", err));
  };

  // Podłącz selection changed (po załadowaniu zmian)
 useEffect(() => {
  let handler: OfficeExtension.EventHandlerResult<any> | null = null;

    Excel.run(async (context) => {
    const workbook = context.workbook;

  if (changes.length === 0) return;

    handler = workbook.onSelectionChanged.add((event) => {
      const address = event.address.split('!')[1] || event.address;

      const change = changeMap.get(address);
      if (change) {
        setSelectedChange(change);
        setCurrentIndex(changeIndexMap.get(address) ?? -1);
      } else {
        setSelectedChange(null);
      }
    });

    await context.sync();
  }).catch(err => console.error("Selection listener failed", err));

  return () => {
    if (handler) {
      handler.remove();
    }
  };
}, [changes]);

  return (
    <Stack tokens={{ padding: 20, childrenGap: 16 }} style={{ width: "100%", maxWidth: 640 }}>
      <Text variant="xxLarge" style={{ fontWeight: 600 }}>
        Excel Visual Diff
      </Text>

      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        <Warning20Filled style={{ marginRight: 8 }} />
        Upload two files to compare: <strong>original (File 1)</strong> and <strong>modified (File 2)</strong>.
        <br />
        (Comparing with open workbook coming soon)
      </MessageBar>

      {trialUses === null ? (
        <MessageBar>Checking license...</MessageBar>
      ) : trialUses <= 0 ? (
        <MessageBar messageBarType={MessageBarType.error}>
          <ErrorCircle20Filled style={{ marginRight: 8 }} />
          Free trial expired •{" "}
          <Link href="https://yourdomain.com/pro" target="_blank">
            Go Pro → $11.99/mo (Basic)
          </Link>
        </MessageBar>
      ) : (
        <MessageBar messageBarType={MessageBarType.info}>
          <CheckmarkCircle20Filled style={{ marginRight: 8 }} />
          {trialUses} free comparison{trialUses === 1 ? "" : "s"} remaining
        </MessageBar>
      )}

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError("")}>
          {error}
        </MessageBar>
      )}

      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="medium">Upload original file (File 1)</Text>
        <input
          type="file"
          accept=".xlsx,.xlsm,.xls,.csv"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) setFile1(f);
          }}
          style={{ fontSize: 14 }}
        />
        {file1 && <Text variant="small">{file1.name}</Text>}
      </Stack>

      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="medium">Upload modified file (File 2)</Text>
        <input
          type="file"
          accept=".xlsx,.xlsm,.xls,.csv"
          onChange={handleFile2}
          style={{ fontSize: 14 }}
        />
        {file2 && <Text variant="small">{file2.name}</Text>}
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 12 }}>
        <PrimaryButton
          onClick={compareFiles}
          disabled={!file1 || !file2 || loading || trialUses === 0 || trialUses === null}
        >
          {loading ? loadingMessage : (
            <>
              <Road20Regular style={{ marginRight: 8 }} />
              Compare & Highlight Changes
            </>
          )}
        </PrimaryButton>

        <DefaultButton onClick={clearHighlights} disabled={loading || changes.length === 0}>
          Clear Highlights
        </DefaultButton>
      </Stack>

      {loading && <ProgressIndicator description={loadingMessage} />}

      {summary && (
        <MessageBar messageBarType={MessageBarType.success}>
          <Icon iconName="Completed" style={{ marginRight: 8 }} />
          {summary}
        </MessageBar>
      )}

      {selectedChange && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline>
          <strong>Selected change:</strong> {selectedChange.sheet} {selectedChange.address}
          <br />
          Type: <span style={{ color: COLORS[selectedChange.type], fontWeight: 600 }}>
            {selectedChange.type.toUpperCase()}
          </span>
          <br />
          Old: {String(selectedChange.oldVal || '—')}
          <br />
          New: {String(selectedChange.newVal || '—')}
        </MessageBar>
      )}

      {changes.length > 0 && (
        <Stack tokens={{ childrenGap: 12 }}>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
            Found {changes.length} changes
          </Text>

          <Stack
            tokens={{ childrenGap: 8 }}
            style={{
              maxHeight: 300,
              overflowY: 'auto',
              border: '1px solid #ddd',
              padding: 8,
              borderRadius: 4,
              background: '#f9f9f9',
            }}
          >
            {changes.map((ch, idx) => (
              <div
                key={`${ch.sheet}-${ch.address}`}
                onClick={() => goToChange(idx)}
                style={{
                  cursor: 'pointer',
                  padding: '8px 12px',
                  background: idx === currentIndex ? '#e6f7ff' : 'white',
                  borderRadius: 4,
                  border: '1px solid #eee',
                  transition: 'background 0.2s',
                }}
              >
                <Text variant="smallPlus" style={{ fontWeight: 600 }}>
                  {ch.sheet} {ch.address}
                </Text>
                <br />
                <Text variant="small" style={{ color: COLORS[ch.type] }}>
                  {ch.type.toUpperCase()}: {String(ch.oldVal || '').slice(0, 30)} → {String(ch.newVal || '').slice(0, 30)}
                </Text>
              </div>
            ))}
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} horizontalAlign="center">
            <DefaultButton
              onClick={() => goToChange(currentIndex - 1)}
              disabled={currentIndex <= 0}
            >
              ← Previous
            </DefaultButton>
            <Text>
              {currentIndex + 1} / {changes.length}
            </Text>
            <DefaultButton
              onClick={() => goToChange(currentIndex + 1)}
              disabled={currentIndex >= changes.length - 1}
            >
              Next →
            </DefaultButton>
          </Stack>

          <Text variant="small" style={{ color: "#666", textAlign: "center" }}>
            Click any change to jump to cell • Use ← → arrows too
          </Text>
        </Stack>
      )}
    </Stack>
  );
}
