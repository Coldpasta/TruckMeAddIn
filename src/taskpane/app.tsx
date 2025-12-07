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
    navigable?: boolean; // true if user can jump to it
  };

  const COLORS = ENGINE_COLORS;

  const TRIAL_TOKEN_KEY = "excelDiffTrialToken";

  export default function App() {
    const [file2, setFile2] = useState<File | null>(null);
    const [changes, setChanges] = useState<Change[]>([]);
    const [currentIndex, setCurrentIndex] = useState(-1);
    const [trialUses, setTrialUses] = useState<number | null>(null);
    const [loading, setLoading] = useState(false);
    const [loadingMessage, setLoadingMessage] = useState("Preparing...");
    const [summary, setSummary] = useState("");
    const [error, setError] = useState("");

    // Initialize trial token
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

    const compareWithOpenWorkbook = async () => {
      if (!file2) {
        setError("Please select a modified file first.");
        return;
      }

      setLoading(true);
      setLoadingMessage("Validating license...");
      setError("");
      setChanges([]);
      setSummary("");


//       const trialOk = await validateTrialAndConsume();
      const trialOk = true;
//       if (!trialOk) {
//         setLoading(false);
//         return;
//       }

      try {
        setLoadingMessage("Parsing uploaded file...");
        const wb2 = await parseFileToSheets(file2);
        const allChanges: Change[] = [];

        setLoadingMessage("Reading open workbook...");
        await Excel.run(async (context: Excel.RequestContext) => {
          const workbook = context.workbook;
          const openSheets = workbook.worksheets;
          openSheets.load("items/name");
          await context.sync();

          const openSheetNames = openSheets.items.map((s: Excel.Worksheet) => s.name);
          const uploadedNames = wb2.SheetNames as string[];

          // Detect added / removed sheets
          const sheetsAdded = uploadedNames.filter((n: string) => !openSheetNames.includes(n));
          const sheetsRemoved = openSheetNames.filter((n: string) => !uploadedNames.includes(n));

          for (const s of sheetsAdded) {
            allChanges.push({
              sheet: s,
              address: "A1",
              row: 0,
              col: 0,
              oldVal: "",
              newVal: `Sheet "${s}" added`,
              type: "added",
              navigable: true,
            });
          }
          for (const s of sheetsRemoved) {
            allChanges.push({
              sheet: s,
              address: "A1",
              row: 0,
              col: 0,
              oldVal: `Sheet "${s}" removed`,
              newVal: "",
              type: "deleted",
              navigable: true,
            });
          }

          const commonSheets = uploadedNames.filter((n) => openSheetNames.includes(n));

          setLoadingMessage(`Comparing ${commonSheets.length} sheets...`);
          const sheetFormatRequests: Record<string, Change[]> = {};

          for (const sheetName of commonSheets) {
            const ws2 = wb2.Sheets[sheetName];
            if (!ws2) continue;

            const openSheet = workbook.worksheets.getItem(sheetName);
            const used = openSheet.getUsedRangeOrNullObject();
            used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
            await context.sync();

            const baseRowIndex = used.isNullObject ? 0 : used.rowIndex || 0;
            const baseColIndex = used.isNullObject ? 0 : used.columnIndex || 0;
            const values1 = used.isNullObject ? [] : (used.values as any[][]);
            const formulas1 = used.isNullObject ? [] : (used.formulas as any[][]);

            const rows2 = ws2["!ref"]
              ? (XLSX.utils.sheet_to_json(ws2, { header: 1, defval: "" }) as any[][])
              : [];

            const sheetChanges: Change[] = compareSheetRows({
              sheetName,
              rowsOpen: { baseRowIndex, baseColIndex, values: values1, formulas: formulas1 },
              rowsUploaded: rows2,
            });

            for (const ch of sheetChanges) {
              allChanges.push({ ...ch, navigable: true });
              if (!sheetFormatRequests[ch.sheet]) sheetFormatRequests[ch.sheet] = [];
              sheetFormatRequests[ch.sheet].push(ch);
            }
          }

          // Apply all highlights
          setLoadingMessage("Applying highlights...");
          for (const [sheetName, changesInSheet] of Object.entries(sheetFormatRequests)) {
            const sheet = workbook.worksheets.getItem(sheetName);

            const byColor: Record<string, string[]> = {};
            for (const ch of changesInSheet) {
              const color = COLORS[ch.type];
              if (!byColor[color]) byColor[color] = [];
              byColor[color].push(ch.address);
            }

            const CHUNK_SIZE = 300;
            for (const [color, addresses] of Object.entries(byColor)) {
              for (let i = 0; i < addresses.length; i += CHUNK_SIZE) {
                const chunk = addresses.slice(i, i + CHUNK_SIZE);
                const range = sheet.getRange(chunk.join(","));
                range.format.fill.color = color;
              }
            }
          }

          await context.sync();
        });

        // Final summary
        const added = allChanges.filter((c) => c.type === "added").length;
        const deleted = allChanges.filter((c) => c.type === "deleted").length;
        const modified = allChanges.filter((c) => c.type === "modified").length;

        setChanges(allChanges);
        setCurrentIndex(allChanges.length > 0 ? 0 : -1);
        setSummary(
          `${allChanges.length} change${allChanges.length === 1 ? "" : "s"} — ` +
            `${added} added • ${deleted} deleted • ${modified} modified`
        );
      } catch (err: any) {
        console.error("Comparison failed:", err);
        setError("Comparison failed: " + (err?.message || "Unknown error"));
      } finally {
        setLoading(false);
        setLoadingMessage("");
      }
    };

    const goToChange = async (idx: number) => {
      if (idx < 0 || idx >= changes.length) return;
      setCurrentIndex(idx);
      const ch = changes[idx];

      await Excel.run(async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getItem(ch.sheet);
        sheet.activate();

        if (ch.navigable && ch.address && ch.row >= 0) {
          const range = sheet.getRange(ch.address);
          range.select();
        }
        await context.sync();
      });
    };

    const clearHighlights = async () => {
      setLoading(true);
      setLoadingMessage("Clearing highlights...");
      try {
        await Excel.run(async (context: Excel.RequestContext) => {
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
        setError("Failed to clear highlights: " + (e?.message || ""));
      } finally {
        setLoading(false);
        setLoadingMessage("");
      }
    };

    return (
      <Stack tokens={{ padding: 20, childrenGap: 16 }} style={{ width: "100%", maxWidth: 640 }}>
        <Text variant="xxLarge" style={{ fontWeight: 600 }}>
          Excel Visual Diff
        </Text>

        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <Warning20Filled style={{ marginRight: 8 }} />
          Open the <strong>original file</strong> in Excel → then upload the <strong>modified version</strong> below.
        </MessageBar>

        {trialUses === null ? (
          <MessageBar>Checking license...</MessageBar>
        ) : trialUses <= 0 ? (
          <MessageBar messageBarType={MessageBarType.error}>
            <ErrorCircle20Filled style={{ marginRight: 8 }} />
            Free trial expired •{" "}
            <Link href="https://yourdomain.com/pro" target="_blank">
              Go Pro → $12/mo
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
            onClick={compareWithOpenWorkbook}
            disabled={!file2 || loading || trialUses === 0 || trialUses === null}
          >
            {loading ? (
              loadingMessage
            ) : (
              <>
                <Road20Regular style={{ marginRight: 8 }} />
                Highlight Changes
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

        {changes.length > 0 && (
          <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
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

            <Text variant="small" style={{ color: "#666" }}>
              Tip: Use ← → arrows or click a change to jump directly in Excel.
            </Text>
          </Stack>
        )}
      </Stack>
    );
  }