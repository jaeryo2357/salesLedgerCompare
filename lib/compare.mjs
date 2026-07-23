import XLSX from "xlsx-js-style";

const FIELDS = [
  { a: "A", b: "C", type: "date" },
  { a: "D", b: "G", type: "text" },
  { a: "E", b: "E", type: "text" },
  { a: "F", b: "AA", fallbackB: "Y", type: "text" },
  { a: "G", b: "AC", type: "number" },
  { a: "H", b: "AD", type: "number" },
  { a: "K", b: "P", fallbackB: "AE", type: "number" },
  { a: "L", b: "Q", fallbackB: "AF", type: "number" },
];
const COLORS = { missing: "FFF2CC", different: "F4CCCC" };
const text = (cell) => (cell?.w ?? cell?.v ?? "").toString().trim();
const usable = (cell) => text(cell) !== "";
const digits = (cell) => text(cell).replace(/[^0-9A-Za-z]/g, "").toUpperCase();
const numeric = (cell) => {
  const value = Number(text(cell).replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(value) ? value : null;
};
const dateKey = (cell) => {
  const match = text(cell).match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);
  return match ? `${match[1]}${match[2].padStart(2, "0")}${match[3].padStart(2, "0")}` : digits(cell);
};
const addr = (column, row) => `${column}${row}`;
const bCell = (sheet, row, field) => sheet[addr(field.b, row)] ?? (field.fallbackB ? sheet[addr(field.fallbackB, row)] : undefined);

function rowsWithData(sheet, firstRow, lastRow, columns) {
  const rows = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    if (columns.some((column) => usable(sheet[addr(column, row)]))) rows.push(row);
  }
  return rows;
}
function exactKeyA(sheet, row) {
  return [digits(sheet[addr("E", row)]), dateKey(sheet[addr("A", row)]), numeric(sheet[addr("K", row)]) ?? "", numeric(sheet[addr("L", row)]) ?? ""].join("|");
}
function exactKeyB(sheet, row) {
  return [digits(sheet[addr("E", row)]), dateKey(sheet[addr("C", row)]), numeric(sheet[addr("P", row)]) ?? numeric(sheet[addr("AE", row)]) ?? "", numeric(sheet[addr("Q", row)]) ?? numeric(sheet[addr("AF", row)]) ?? ""].join("|");
}
function partyDateA(sheet, row) { return [digits(sheet[addr("E", row)]), dateKey(sheet[addr("A", row)])].join("|"); }
function partyDateB(sheet, row) { return [digits(sheet[addr("E", row)]), dateKey(sheet[addr("C", row)])].join("|"); }
function sameValue(aCell, bCellValue, type) {
  if (type === "number") {
    const a = numeric(aCell); const b = numeric(bCellValue);
    return a !== null && b !== null && Math.abs(a - b) < 0.000001;
  }
  if (type === "date") return dateKey(aCell) === dateKey(bCellValue);
  return text(aCell).replace(/\s+/g, " ").toUpperCase() === text(bCellValue).replace(/\s+/g, " ").toUpperCase();
}
function highlight(cell, rgb) {
  cell.s = { ...(cell.s ?? {}), fill: { patternType: "solid", fgColor: { rgb } } };
}
function add(map, key, row) { map.set(key, [...(map.get(key) ?? []), row]); }

/** B의 구성과 값을 유지하면서 비교 결과 배경색만 적용한다. */
export function comparePurchaseLedgers(aBuffer, bBuffer) {
  const aWorkbook = XLSX.read(aBuffer, { type: "buffer", cellStyles: true, cellDates: true });
  const bWorkbook = XLSX.read(bBuffer, { type: "buffer", cellStyles: true, cellDates: true });
  const aSheet = aWorkbook.Sheets[aWorkbook.SheetNames[0]];
  const bSheet = bWorkbook.Sheets[bWorkbook.SheetNames[0]];
  const aEndRow = XLSX.utils.decode_range(aSheet["!ref"] ?? "A1:A1").e.r + 1;
  const bEndRow = XLSX.utils.decode_range(bSheet["!ref"] ?? "A1:A1").e.r + 1;
  const aRows = rowsWithData(aSheet, 3, aEndRow, ["A", "D", "E", "F", "K", "L"]);
  const bRows = rowsWithData(bSheet, 7, bEndRow, ["C", "E", "G", "P", "Q", "AA", "AE"]);
  const exactRows = new Map();
  const partyDateRows = new Map();
  for (const row of aRows) {
    add(exactRows, exactKeyA(aSheet, row), row);
    add(partyDateRows, partyDateA(aSheet, row), row);
  }

  const usedARows = new Set();
  let yellowCells = 0;
  let redCells = 0;
  for (const bRow of bRows) {
    const exact = (exactRows.get(exactKeyB(bSheet, bRow)) ?? []).filter((row) => !usedARows.has(row));
    const fallback = (partyDateRows.get(partyDateB(bSheet, bRow)) ?? []).filter((row) => !usedARows.has(row));
    const aRow = exact.length === 1 ? exact[0] : fallback.length === 1 ? fallback[0] : undefined;
    if (!aRow) {
      for (const field of FIELDS) {
        const target = bCell(bSheet, bRow, field);
        if (usable(target)) { highlight(target, COLORS.missing); yellowCells += 1; }
      }
      continue;
    }
    usedARows.add(aRow);
    for (const field of FIELDS) {
      const target = bCell(bSheet, bRow, field);
      if (!usable(target)) continue;
      const source = aSheet[addr(field.a, aRow)];
      if (!usable(source)) { highlight(target, COLORS.missing); yellowCells += 1; }
      else if (!sameValue(source, target, field.type)) { highlight(target, COLORS.different); redCells += 1; }
    }
  }

  return {
    output: XLSX.write(bWorkbook, { type: "buffer", bookType: "xlsx", cellStyles: true }),
    stats: { yellowCells, redCells, aRows: aRows.length, bRows: bRows.length },
  };
}
