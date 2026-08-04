/* global XLSX */
// 노랑은 값 불일치, 빨강은 B에 대응하는 A 행/값이 없는 경우에만 사용합니다.
const COLORS = { missing: "F4CCCC", different: "FFF2CC" };
const A_HEADERS = {
  date: ["일자", "거래일자", "작성일자"],
  business: ["사업자(주민)번호", "사업자등록번호", "사업자번호"],
  supply: ["매입 공급가액", "매입공급가액"],
  tax: ["매입 부가세", "매입부가세", "매입세액"],
  taxType: ["과세유형", "과세 구분", "과세구분"],
};
const B_HEADERS = {
  date: ["작성일자", "발급일자", "일자", "거래일자"],
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
const dateKey = (cell) => {
  if (cell?.v instanceof Date && !Number.isNaN(cell.v.valueOf())) {
    return `${cell.v.getFullYear()}${String(cell.v.getMonth() + 1).padStart(2, "0")}${String(cell.v.getDate()).padStart(2, "0")}`;
  }
  if (typeof cell?.v === "number" && XLSX.SSF?.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (parsed) return `${parsed.y}${String(parsed.m).padStart(2, "0")}${String(parsed.d).padStart(2, "0")}`;
  }
  const value = text(cell);
  const match = value.match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);
  return match ? `${match[1]}${match[2].padStart(2, "0")}${match[3].padStart(2, "0")}` : "";
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

function invoiceKey(sheet, columns, row) {
  const businessNumber = digits(cellAt(sheet, columns.business, row));
  const date = dateKey(cellAt(sheet, columns.date, row));
  return businessNumber && date ? `${businessNumber}|${date}` : "";
}

function isDocumentTypeCompatible(aSheet, aColumns, aRow, bColumns) {
  const aMode = taxMode(cellAt(aSheet, aColumns.taxType, aRow));
  // B에 세액 열이 없으면 전자계산서(면세), 있으면 전자세금계산서로 처리합니다.
  return bColumns.tax ? aMode !== "exempt" : aMode === "exempt";
}

function isExactMatch(aSheet, aColumns, aRow, bSheet, bColumns, bRow) {
  if (!sameNumber(cellAt(aSheet, aColumns.supply, aRow), cellAt(bSheet, bColumns.supply, bRow))) return false;
  if (taxMode(cellAt(aSheet, aColumns.taxType, aRow)) !== "disallowed" || !bColumns.tax) return true;
  return sameNumber(cellAt(aSheet, aColumns.tax, aRow), cellAt(bSheet, bColumns.tax, bRow));
}

function compareByBusinessGroups(aBytes, bBytes) {
  const aBook = XLSX.read(aBytes, { type: "array", cellStyles: true, cellDates: true });
  const bBook = XLSX.read(bBytes, { type: "array", cellStyles: true, cellDates: true });
  const a = findWorksheet(aBook, A_HEADERS, "A 매입매출장");
  const b = findWorksheet(bBook, B_HEADERS, "B 전자세금계산서", ["date", "business", "supply"]);
  const aLastRow = XLSX.utils.decode_range(a.sheet["!ref"] || "A1:A1").e.r + 1;
  const bLastRow = XLSX.utils.decode_range(b.sheet["!ref"] || "A1:A1").e.r + 1;
  const aRows = rowsWithData(a.sheet, a.row + 1, aLastRow, [a.columns.date, a.columns.business, a.columns.supply]);
  const bRows = rowsWithData(b.sheet, b.row + 1, bLastRow, [b.columns.date, b.columns.business, b.columns.supply]);
  const aGroups = new Map(); const bGroups = new Map();
  let blankA = 0; let blankB = 0;
  for (const row of aRows) {
    const key = invoiceKey(a.sheet, a.columns, row);
    if (key) add(aGroups, key, row); else blankA += 1;
  }
  for (const row of bRows) {
    const key = invoiceKey(b.sheet, b.columns, row);
    if (key) add(bGroups, key, row); else blankB += 1;
  }

  let yellow = 0; let red = 0; let countMismatch = 0; let aOnlyRows = blankA; let bOnlyRows = blankB;
  let matchedRows = 0; let supplyDifferences = 0; let taxDifferences = 0; let unmatchedBRows = 0; let documentTypeMismatches = 0; let ambiguousBRows = 0;
  const paintMissingBRow = (row) => {
    unmatchedBRows += 1;
    // 대응 불가 계산서는 해당 행의 공급가액 셀 하나만 빨강으로 표시합니다.
    // 여러 셀을 모두 칠해 한 건이 여러 오류처럼 보이는 일을 막습니다.
    const target = cellAt(b.sheet, b.columns.supply, row) ?? cellAt(b.sheet, b.columns.business, row);
    if (usable(target)) { color(target, COLORS.missing); red += 1; }
  };
  for (const row of bRows) if (!invoiceKey(b.sheet, b.columns, row)) paintMissingBRow(row);

  const invoiceKeys = new Set([...aGroups.keys(), ...bGroups.keys()]);
  for (const key of invoiceKeys) {
    const aBusinessRows = aGroups.get(key) ?? [];
    const bBusinessRows = bGroups.get(key) ?? [];
    if (aBusinessRows.length !== bBusinessRows.length) countMismatch += 1;
    if (aBusinessRows.length > bBusinessRows.length) aOnlyRows += aBusinessRows.length - bBusinessRows.length;
    if (bBusinessRows.length > aBusinessRows.length) bOnlyRows += bBusinessRows.length - aBusinessRows.length;

    // 사업자번호·날짜·문서유형이 같은 그룹 안에서만 비교합니다.
    // 같은 그룹에 여러 행이 남아 대응 관계가 모호하면 노랑 대신 빨강으로 표시합니다.
    const unusedA = new Set(aBusinessRows);
    const pairs = new Map();
    for (const bRow of bBusinessRows) {
      const exactA = [...unusedA].find((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns) && isExactMatch(a.sheet, a.columns, aRow, b.sheet, b.columns, bRow));
      if (exactA !== undefined) { pairs.set(bRow, exactA); unusedA.delete(exactA); }
    }
    const remainingBRows = bBusinessRows.filter((bRow) => !pairs.has(bRow));
    const compatibleA = [...unusedA].filter((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns));
    if (remainingBRows.length === 1 && compatibleA.length === 1) {
      pairs.set(remainingBRows[0], compatibleA[0]);
      unusedA.delete(compatibleA[0]);
    } else {
      for (const bRow of remainingBRows) {
        if (aBusinessRows.length > 0 && !aBusinessRows.some((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns))) documentTypeMismatches += 1;
        else if (aBusinessRows.length > 0) ambiguousBRows += 1;
        paintMissingBRow(bRow);
      }
    }
    for (const [bRow, aRow] of pairs) {
      const aSupply = cellAt(a.sheet, a.columns.supply, aRow);
      const bSupply = cellAt(b.sheet, b.columns.supply, bRow);
      if (!usable(aSupply)) { color(bSupply, COLORS.missing); red += 1; }
      else if (!sameNumber(aSupply, bSupply)) { color(bSupply, COLORS.different); yellow += 1; supplyDifferences += 1; }

      if (taxMode(cellAt(a.sheet, a.columns.taxType, aRow)) === "disallowed" && b.columns.tax) {
        const aTax = cellAt(a.sheet, a.columns.tax, aRow);
        const bTax = cellAt(b.sheet, b.columns.tax, bRow);
        if (!usable(aTax)) { color(bTax, COLORS.missing); red += 1; }
        else if (!sameNumber(aTax, bTax)) { color(bTax, COLORS.different); yellow += 1; taxDifferences += 1; }
      }
    }
    matchedRows += pairs.size;
  }
  return {
    data: XLSX.write(bBook, { type: "array", bookType: "xlsx", cellStyles: true }),
    yellow, red, countMismatch, aOnlyRows, bOnlyRows, matchedRows, supplyDifferences, taxDifferences, unmatchedBRows, documentTypeMismatches, ambiguousBRows,
    aRowCount: aRows.length, bRowCount: bRows.length, bHasTax: Boolean(b.columns.tax), aSheetName: a.name, bSheetName: b.name,
  };
}

const form = document.querySelector("#compare-form");
const status = document.querySelector("#status");
const button = form.querySelector("button");
const summary = document.querySelector("#summary");
const summaryTitle = document.querySelector("#summary-title");
const summaryDescription = document.querySelector("#summary-description");
const summaryCards = document.querySelector("#summary-cards");
const summaryNotes = document.querySelector("#summary-notes");
function summaryCard(label, value, tone = "") {
  const card = document.createElement("div"); card.className = `summary-card ${tone}`;
  const labelElement = document.createElement("span"); labelElement.textContent = label;
  const valueElement = document.createElement("strong"); valueElement.textContent = value;
  card.append(labelElement, valueElement); return card;
}
function summaryNote(textValue, tone = "") {
  const note = document.createElement("li"); note.className = tone; note.textContent = textValue; return note;
}
function renderSummary(result) {
  const hasIssues = result.yellow > 0 || result.red > 0;
  summary.hidden = false;
  summaryTitle.textContent = hasIssues ? "비교 결과를 확인해 주세요" : "모든 비교가 정상 완료됐어요";
  summaryDescription.textContent = `${result.aSheetName} ↔ ${result.bSheetName} · B 문서 유형: ${result.bHasTax ? "전자세금계산서" : "전자계산서(면세)"}`;
  summaryCards.replaceChildren(
    summaryCard("읽은 행", `A ${result.aRowCount} · B ${result.bRowCount}`),
    summaryCard("A와 대응된 행", `${result.matchedRows}행`, result.matchedRows === result.bRowCount ? "success" : ""),
    summaryCard("공급가액 차이", `${result.supplyDifferences}건`, result.supplyDifferences ? "warning" : ""),
    summaryCard("세액 차이", result.bHasTax ? `${result.taxDifferences}건` : "해당 없음", result.taxDifferences ? "warning" : ""),
    summaryCard("대응 불가 B행", `${result.unmatchedBRows}행`, result.unmatchedBRows ? "danger" : ""),
    summaryCard("A 단독 행", `${result.aOnlyRows}행`, result.aOnlyRows ? "danger" : ""),
  );
  const notes = [];
  if (result.supplyDifferences) notes.push(summaryNote(`공급가액이 다른 행이 ${result.supplyDifferences}건 있어 B의 공급가액 셀을 노랑으로 표시했습니다.`, "warning"));
  if (result.taxDifferences) notes.push(summaryNote(`불공 행 중 세액이 다른 행이 ${result.taxDifferences}건 있어 B의 세액 셀을 노랑으로 표시했습니다.`, "warning"));
  if (result.documentTypeMismatches) notes.push(summaryNote(`같은 사업자번호·날짜지만 면세/세금계산서 문서 유형이 맞지 않는 B 행이 ${result.documentTypeMismatches}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.ambiguousBRows) notes.push(summaryNote(`같은 사업자번호·날짜에 여러 행이 남아 대응 관계가 모호한 B 행이 ${result.ambiguousBRows}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.unmatchedBRows > result.documentTypeMismatches + result.ambiguousBRows) notes.push(summaryNote(`A에서 대응 행을 찾지 못한 B 행이 ${result.unmatchedBRows - result.documentTypeMismatches - result.ambiguousBRows}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.countMismatch) notes.push(summaryNote(`사업자번호·날짜별 행 수가 다른 그룹이 ${result.countMismatch}개입니다.`, "danger"));
  if (result.aOnlyRows) notes.push(summaryNote(`A에만 존재하는 행은 ${result.aOnlyRows}건입니다. B 결과 파일에는 대응할 셀이 없어 상태 정보로만 표시합니다.`, "danger"));
  if (!notes.length) notes.push(summaryNote("사업자번호, 문서 유형, 공급가액 기준으로 모든 B 행이 정상 대응됐습니다."));
  summaryNotes.replaceChildren(...notes);
}
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  button.disabled = true;
  summary.hidden = true;
  status.textContent = "이 브라우저 안에서 엑셀 헤더와 사업자번호를 확인하는 중입니다…";
  try {
    const aFile = form.elements.aFile.files[0]; const bFile = form.elements.bFile.files[0];
    const result = compareByBusinessGroups(new Uint8Array(await aFile.arrayBuffer()), new Uint8Array(await bFile.arrayBuffer()));
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([result.data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    link.download = "B_색상표시_비교결과.xlsx";
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 0);
    status.textContent = "비교가 완료되어 결과 엑셀을 다운로드했습니다.";
    renderSummary(result);
  } catch (error) {
    status.textContent = `처리 오류: ${error.message}`;
  } finally {
    button.disabled = false;
  }
});
