/* global XLSX */
// 노랑은 값 불일치, 빨강은 B에 대응하는 A 행/값이 없는 경우에만 사용합니다.
const COLORS = { missing: "F4CCCC", different: "FFF2CC" };
const A_HEADERS = {
  business: ["사업자(주민)번호", "사업자등록번호", "사업자번호"],
  supply: ["매입 공급가액", "매입공급가액"],
  tax: ["매입 부가세", "매입부가세", "매입세액"],
  taxType: ["과세유형", "과세 구분", "과세구분"],
};
const B_HEADERS = {
  business: ["공급자사업자등록번호", "사업자등록번호", "사업자번호"],
  supply: ["공급가액", "품목공급가액"],
  tax: ["세액", "품목세액", "부가세"],
};

const text = (cell) => (cell?.w ?? cell?.v ?? "").toString().trim();
const usable = (cell) => text(cell) !== "";
const digits = (cell) => text(cell).replace(/[^0-9A-Za-z]/g, "").toUpperCase();
const numeric = (cell) => {
  if (typeof cell?.v === "number" && Number.isFinite(cell.v)) return cell.v;
  const raw = text(cell);
  if (!raw) return null;
  const value = Number(raw.replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(value) ? value : null;
};
const headerKey = (value) => value.toString().replace(/[\s(){}\[\]·ㆍ._\-/:]/g, "").toLowerCase();
const address = (column, row) => `${column}${row}`;
const cellAt = (sheet, column, row) => sheet[address(column, row)];
function headerTextAt(sheet, columnIndex, row) {
  const column = XLSX.utils.encode_col(columnIndex);
  const direct = text(cellAt(sheet, column, row));
  if (direct) return direct;
  const rowIndex = row - 1;
  const merge = (sheet["!merges"] ?? []).find((range) => range.s.r <= rowIndex && range.e.r >= rowIndex && range.s.c <= columnIndex && range.e.c >= columnIndex);
  return merge ? text(sheet[XLSX.utils.encode_cell(merge.s)]) : "";
}

function findHeaderRow(sheet, definitions, requiredKeys = Object.keys(definitions)) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const aliases = Object.fromEntries(Object.entries(definitions).map(([key, values]) => [key, values.map(headerKey)]));
  const maxRow = Math.min(range.e.r + 1, 40);
  for (let row = 1; row <= maxRow; row += 1) {
    const columns = {};
    for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
      const column = XLSX.utils.encode_col(columnIndex);
      const current = headerKey(headerTextAt(sheet, columnIndex, row));
      const previous = row > 1 ? headerKey(headerTextAt(sheet, columnIndex, row - 1)) : "";
      // 일부 매입매출장은 "매 입"과 "공급가액"처럼 헤더를 두 줄로 나눕니다.
      // 현재 줄, 이전 줄, 두 줄을 합친 값을 모두 헤더 후보로 봅니다.
      const values = [current, previous, `${previous}${current}`];
      for (const [key, names] of Object.entries(aliases)) {
        if (!columns[key] && values.some((value) => names.includes(value))) columns[key] = column;
      }
    }
    if (requiredKeys.every((key) => columns[key])) return { row, columns };
  }
  return null;
}

function findWorksheet(book, definitions, fileName, requiredKeys) {
  for (const name of book.SheetNames) {
    const sheet = book.Sheets[name];
    const header = findHeaderRow(sheet, definitions, requiredKeys);
    if (header) return { sheet, name, ...header };
  }
  throw new Error(`${fileName}에서 필요한 헤더를 찾지 못했습니다. 지원 헤더명을 확인해 주세요.`);
}

function rowsWithData(sheet, firstRow, lastRow, columns) {
  const rows = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    if (columns.some((column) => usable(cellAt(sheet, column, row)))) rows.push(row);
  }
  return rows;
}

const sameNumber = (a, b) => {
  const left = numeric(a); const right = numeric(b);
  return left !== null && right !== null && Math.abs(left - right) < 0.000001;
};
const taxMode = (cell) => {
  const value = headerKey(text(cell));
  if (value.includes("불공")) return "disallowed";
  if (value.includes("면세")) return "exempt";
  return "normal";
};

function color(cell, rgb) {
  cell.s = { ...(cell.s ?? {}), fill: { patternType: "solid", fgColor: { rgb } } };
}

function add(map, key, row) {
  map.set(key, [...(map.get(key) ?? []), row]);
}

function isExactMatch(aSheet, aColumns, aRow, bSheet, bColumns, bRow) {
  if (!sameNumber(cellAt(aSheet, aColumns.supply, aRow), cellAt(bSheet, bColumns.supply, bRow))) return false;
  if (taxMode(cellAt(aSheet, aColumns.taxType, aRow)) !== "disallowed" || !bColumns.tax) return true;
  return sameNumber(cellAt(aSheet, aColumns.tax, aRow), cellAt(bSheet, bColumns.tax, bRow));
}

function differenceCost(aSheet, aColumns, aRow, bSheet, bColumns, bRow) {
  const aSupply = numeric(cellAt(aSheet, aColumns.supply, aRow));
  const bSupply = numeric(cellAt(bSheet, bColumns.supply, bRow));
  let cost = aSupply === null || bSupply === null ? Number.MAX_SAFE_INTEGER : Math.abs(aSupply - bSupply);
  if (taxMode(cellAt(aSheet, aColumns.taxType, aRow)) === "disallowed" && bColumns.tax) {
    const aTax = numeric(cellAt(aSheet, aColumns.tax, aRow));
    const bTax = numeric(cellAt(bSheet, bColumns.tax, bRow));
    cost += aTax === null || bTax === null ? Number.MAX_SAFE_INTEGER : Math.abs(aTax - bTax);
  }
  return cost;
}

function compareByBusinessGroups(aBytes, bBytes) {
  const aBook = XLSX.read(aBytes, { type: "array", cellStyles: true, cellDates: true });
  const bBook = XLSX.read(bBytes, { type: "array", cellStyles: true, cellDates: true });
  const a = findWorksheet(aBook, A_HEADERS, "A 매입매출장");
  const b = findWorksheet(bBook, B_HEADERS, "B 전자세금계산서", ["business", "supply"]);
  const aLastRow = XLSX.utils.decode_range(a.sheet["!ref"] || "A1:A1").e.r + 1;
  const bLastRow = XLSX.utils.decode_range(b.sheet["!ref"] || "A1:A1").e.r + 1;
  const aRows = rowsWithData(a.sheet, a.row + 1, aLastRow, [a.columns.business, a.columns.supply]);
  const bRows = rowsWithData(b.sheet, b.row + 1, bLastRow, [b.columns.business, b.columns.supply]);
  const aGroups = new Map(); const bGroups = new Map();
  let blankA = 0; let blankB = 0;
  for (const row of aRows) {
    const businessNumber = digits(cellAt(a.sheet, a.columns.business, row));
    if (businessNumber) add(aGroups, businessNumber, row); else blankA += 1;
  }
  for (const row of bRows) {
    const businessNumber = digits(cellAt(b.sheet, b.columns.business, row));
    if (businessNumber) add(bGroups, businessNumber, row); else blankB += 1;
  }

  let yellow = 0; let red = 0; let countMismatch = 0; let aOnlyRows = blankA; let bOnlyRows = blankB;
  const paintMissingBRow = (row) => {
    for (const key of ["business", "supply", "tax"].filter((key) => b.columns[key])) {
      const target = cellAt(b.sheet, b.columns[key], row);
      if (usable(target)) { color(target, COLORS.missing); red += 1; }
    }
  };
  for (const row of bRows) if (!digits(cellAt(b.sheet, b.columns.business, row))) paintMissingBRow(row);

  const businessNumbers = new Set([...aGroups.keys(), ...bGroups.keys()]);
  for (const businessNumber of businessNumbers) {
    const aBusinessRows = aGroups.get(businessNumber) ?? [];
    const bBusinessRows = bGroups.get(businessNumber) ?? [];
    if (aBusinessRows.length !== bBusinessRows.length) countMismatch += 1;
    if (aBusinessRows.length > bBusinessRows.length) aOnlyRows += aBusinessRows.length - bBusinessRows.length;
    if (bBusinessRows.length > aBusinessRows.length) bOnlyRows += bBusinessRows.length - aBusinessRows.length;

    // 사업자번호 그룹 안에서는 공급가액(불공은 공급가액+세액)으로 대응 행을 찾습니다.
    // 두 파일의 행 순서는 비교에 사용하지 않습니다.
    const unusedA = new Set(aBusinessRows);
    const pairs = new Map();
    for (const bRow of bBusinessRows) {
      const exactA = [...unusedA].find((aRow) => isExactMatch(a.sheet, a.columns, aRow, b.sheet, b.columns, bRow));
      if (exactA !== undefined) { pairs.set(bRow, exactA); unusedA.delete(exactA); }
    }
    for (const bRow of bBusinessRows) {
      if (pairs.has(bRow)) continue;
      const nearestA = [...unusedA].sort((left, right) => differenceCost(a.sheet, a.columns, left, b.sheet, b.columns, bRow) - differenceCost(a.sheet, a.columns, right, b.sheet, b.columns, bRow))[0];
      if (nearestA === undefined) { paintMissingBRow(bRow); continue; }
      pairs.set(bRow, nearestA); unusedA.delete(nearestA);
    }
    for (const [bRow, aRow] of pairs) {
      const aSupply = cellAt(a.sheet, a.columns.supply, aRow);
      const bSupply = cellAt(b.sheet, b.columns.supply, bRow);
      if (!usable(aSupply)) { color(bSupply, COLORS.missing); red += 1; }
      else if (!sameNumber(aSupply, bSupply)) { color(bSupply, COLORS.different); yellow += 1; }

      if (taxMode(cellAt(a.sheet, a.columns.taxType, aRow)) === "disallowed" && b.columns.tax) {
        const aTax = cellAt(a.sheet, a.columns.tax, aRow);
        const bTax = cellAt(b.sheet, b.columns.tax, bRow);
        if (!usable(aTax)) { color(bTax, COLORS.missing); red += 1; }
        else if (!sameNumber(aTax, bTax)) { color(bTax, COLORS.different); yellow += 1; }
      }
    }
  }
  return {
    data: XLSX.write(bBook, { type: "array", bookType: "xlsx", cellStyles: true }),
    yellow, red, countMismatch, aOnlyRows, bOnlyRows, aSheetName: a.name, bSheetName: b.name,
  };
}

const form = document.querySelector("#compare-form");
const status = document.querySelector("#status");
const button = form.querySelector("button");
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  button.disabled = true;
  status.textContent = "이 브라우저 안에서 엑셀 헤더와 사업자번호를 확인하는 중입니다…";
  try {
    const aFile = form.elements.aFile.files[0]; const bFile = form.elements.bFile.files[0];
    const result = compareByBusinessGroups(new Uint8Array(await aFile.arrayBuffer()), new Uint8Array(await bFile.arrayBuffer()));
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([result.data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    link.download = "B_색상표시_비교결과.xlsx";
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 0);
    status.textContent = `완료: 노랑 ${result.yellow}셀 · 빨강 ${result.red}셀 · 사업자별 행 수 차이 ${result.countMismatch}개 · A에만 ${result.aOnlyRows}행 · B에만 ${result.bOnlyRows}행`;
  } catch (error) {
    status.textContent = `처리 오류: ${error.message}`;
  } finally {
    button.disabled = false;
  }
});
