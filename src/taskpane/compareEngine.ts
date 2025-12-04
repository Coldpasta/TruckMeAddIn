// src/taskpane/compareEngine.ts
export type ChangeType = "added" | "deleted" | "modified";
export interface ChangeItem {
  sheet: string;
  address: string;
  row: number;
  col: number;
  oldVal: any;
  newVal: any;
  type: ChangeType;
}

export const COLORS: Record<ChangeType, string> = {
  added: "#d4edda",
  deleted: "#f8d7da",
  modified: "#fff3cd",
};

function normalizeCell(v: any, settings?: { trim?: boolean; ignoreCase?: boolean }) {
  let s = v === null || v === undefined ? "" : String(v);
  if (settings?.trim) s = s.trim();
  if (settings?.ignoreCase) s = s.toLowerCase();
  return s;
}

export function rowSignature(row: any[], settings?: { trim?: boolean; ignoreCase?: boolean }) {
  return row.map((c) => normalizeCell(c, settings)).join("\u0001");
}

export function lcsIndices(a: string[], b: string[]) {
  const n = a.length, m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? 1 + dp[i + 1][j + 1] : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const pairs: Array<{ i: number; j: number }> = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      pairs.push({ i, j });
      i++; j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      i++;
    } else {
      j++;
    }
  }
  return pairs;
}

interface OpenSheetRows {
  baseRowIndex: number;
  baseColIndex: number;
  values: any[][];
  formulas: any[][];
}

interface CompareSheetRowsArgs {
  sheetName: string;
  rowsOpen: OpenSheetRows;
  rowsUploaded: any[][];
}

/**
 * Produces ChangeItem[] for a single sheet.
 * rowsOpen: values/formulas arrays from usedRange (may be empty)
 * rowsUploaded: SheetJS sheet_to_json(..., header:1)
 */
export function compareSheetRows(args: CompareSheetRowsArgs): ChangeItem[] {
  const { sheetName, rowsOpen, rowsUploaded } = args;
  const { baseRowIndex, baseColIndex, values, formulas } = rowsOpen;
  const rows2 = rowsUploaded || [];

  const rowCount = values?.length ?? 0;
  const colCount = values && values[0] ? values[0].length : 0;

  // build signatures
  const sigs1: string[] = [];
  for (let r = 0; r < rowCount; r++) {
    const rowArr: any[] = [];
    for (let c = 0; c < colCount; c++) {
      const formula = formulas?.[r]?.[c];
      const val = (typeof formula === "string" && formula.startsWith("=")) ? formula : values?.[r]?.[c];
      rowArr.push(val ?? "");
    }
    sigs1.push(rowSignature(rowArr));
  }

  const sigs2: string[] = [];
  for (let r = 0; r < rows2.length; r++) sigs2.push(rowSignature(rows2[r] ?? []));

  const matches = lcsIndices(sigs1, sigs2);
  const matchedMap1 = new Map(matches.map((p) => [p.i, p.j]));
  const matchedMap2 = new Map(matches.map((p) => [p.j, p.i]));

  const changes: ChangeItem[] = [];
  let p1 = 0, p2 = 0;
  const maxRows = Math.max(rowCount, rows2.length);

  while (p1 < maxRows || p2 < maxRows) {
    const m1 = matchedMap1.get(p1);
    const m2 = matchedMap2.get(p2);

    if (m1 !== undefined && m1 === p2) {
      // matched row: compare per-cell
      const colsToCheck = Math.max(colCount, (rows2[p2] || []).length);
      for (let C = 0; C < colsToCheck; C++) {
        const rIndex = p1;
        const addr = encodeCell(rIndex + baseRowIndex, C + baseColIndex);
        const formula1 = formulas?.[rIndex]?.[C];
        const value1 = values?.[rIndex]?.[C];
        const val1 = (typeof formula1 === "string" && formula1.startsWith("=")) ? formula1 : (value1 ?? "");
        const val2 = (rows2[p2] && C < rows2[p2].length) ? (rows2[p2][C] ?? "") : "";
        if (val1 !== val2) {
          let type: ChangeType = "modified";
          if ((val1 === "" || val1 === null) && val2 !== "") type = "added";
          if (val1 !== "" && (val2 === "" || val2 === null)) type = "deleted";
          changes.push({
            sheet: sheetName,
            address: addr,
            row: rIndex + baseRowIndex,
            col: C + baseColIndex,
            oldVal: val1,
            newVal: val2,
            type,
          });
        }
      }
      p1++; p2++;
    } else {
      // inserted row in uploaded (added)
      if ((m1 === undefined) && p2 < rows2.length && !matchedMap2.has(p2)) {
        const rowArr2 = rows2[p2] || [];
        for (let C = 0; C < rowArr2.length; C++) {
          const addr = encodeCell(p1 + baseRowIndex, C + baseColIndex);
          const newVal = rowArr2[C] ?? "";
          if (newVal !== "") {
            changes.push({
              sheet: sheetName,
              address: addr,
              row: p1 + baseRowIndex,
              col: C + baseColIndex,
              oldVal: "",
              newVal,
              type: "added",
            });
          }
        }
        p2++;
        // do not increment p1 (insertion shifts)
      } else if ((m2 === undefined) && p1 < rowCount && !matchedMap1.has(p1)) {
        // row deleted in uploaded (present in open workbook)
        for (let C = 0; C < colCount; C++) {
          const addr = encodeCell(p1 + baseRowIndex, C + baseColIndex);
          const formula1 = formulas?.[p1]?.[C];
          const value1 = values?.[p1]?.[C];
          const val1 = (typeof formula1 === "string" && formula1.startsWith("=")) ? formula1 : (value1 ?? "");
          if (val1 !== "") {
            changes.push({
              sheet: sheetName,
              address: addr,
              row: p1 + baseRowIndex,
              col: C + baseColIndex,
              oldVal: val1,
              newVal: "",
              type: "deleted",
            });
          }
        }
        p1++;
      } else {
        // fallback advance
        if (p1 < rowCount) p1++;
        if (p2 < rows2.length) p2++;
      }
    }
  }

  return changes;
}

// Helper to convert zero-based row/col to A1 address
function encodeCell(zeroRow: number, zeroCol: number) {
  // column to letters
  let col = zeroCol;
  let s = "";
  while (col >= 0) {
    s = String.fromCharCode((col % 26) + 65) + s;
    col = Math.floor(col / 26) - 1;
  }
  return `${s}${zeroRow + 1}`;
}
