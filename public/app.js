/* global XLSX */
// 노랑은 값 불일치, 빨강은 B에 대응하는 A 행 없음, 파랑은 A에 대응하는 B 행 없음입니다.
const COLORS = { missing: "F4CCCC", different: "FFF2CC", aOnly: "CFE2F3" };
const CLOSE_SUPPLY_DIFFERENCE = 1000;
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
const supplyDifference = (a, b) => {
  const left = numeric(a); const right = numeric(b);
  return left === null || right === null ? null : Math.abs(left - right);
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

function invoiceKey(sheet, columns, row) { return digits(cellAt(sheet, columns.business, row)); }

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

const valueLabel = (cell) => text(cell) || "입력 없음";
const amountLabel = (cell) => {
  const value = numeric(cell);
  return value === null ? valueLabel(cell) : new Intl.NumberFormat("ko-KR", { maximumFractionDigits: 6 }).format(value);
};
const aDocumentLabel = (sheet, columns, row) => {
  const mode = taxMode(cellAt(sheet, columns.taxType, row));
  if (mode === "disallowed") return "불공 (공급가액·세액 비교)";
  if (mode === "exempt") return "면세";
  return "과세/기타 (공급가액 비교)";
};
const bDocumentLabel = (columns) => columns.tax ? "전자세금계산서" : "전자계산서 (면세)";

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
    const key = invoiceKey(a.sheet, a.columns, row);
    if (key) add(aGroups, key, row); else blankA += 1;
  }
  for (const row of bRows) {
    const key = invoiceKey(b.sheet, b.columns, row);
    if (key) add(bGroups, key, row); else blankB += 1;
  }

  let yellow = 0; let red = 0; let countMismatch = 0; let aOnlyRows = blankA; let bOnlyRows = blankB;
  let matchedRows = 0; let supplyDifferences = 0; let taxDifferences = 0; let nearSupplyMatches = 0; let unmatchedBRows = 0; let documentTypeMismatches = 0; let ambiguousBRows = 0; let differentInvoiceBRows = 0; let blueRows = 0;
  const aRowsToPaint = new Set();
  const details = [];
  const bDetail = (row, values) => details.push({
    source: "B", sourceRow: row,
    businessNumber: valueLabel(cellAt(b.sheet, b.columns.business, row)),
    bDocument: bDocumentLabel(b.columns),
    bSupply: amountLabel(cellAt(b.sheet, b.columns.supply, row)),
    bTax: b.columns.tax ? amountLabel(cellAt(b.sheet, b.columns.tax, row)) : "해당 없음",
    ...values,
  });
  const aDetail = (row, values) => details.push({
    source: "A", sourceRow: row,
    businessNumber: valueLabel(cellAt(a.sheet, a.columns.business, row)),
    aDocument: aDocumentLabel(a.sheet, a.columns, row),
    aSupply: amountLabel(cellAt(a.sheet, a.columns.supply, row)),
    aTax: amountLabel(cellAt(a.sheet, a.columns.tax, row)),
    ...values,
  });
  const paintMissingBRow = (row) => {
    unmatchedBRows += 1;
    // 대응 불가 계산서는 해당 행의 공급가액 셀 하나만 빨강으로 표시합니다.
    // 여러 셀을 모두 칠해 한 건이 여러 오류처럼 보이는 일을 막습니다.
    const target = cellAt(b.sheet, b.columns.supply, row) ?? cellAt(b.sheet, b.columns.business, row);
    if (usable(target)) { color(target, COLORS.missing); red += 1; }
  };
  for (const row of bRows) {
    if (!invoiceKey(b.sheet, b.columns, row)) {
      paintMissingBRow(row);
      bDetail(row, {
        tone: "danger", status: "대응 불가", isIssue: true,
        reason: "사업자번호가 비어 있어 A 파일에서 비교할 후보를 찾을 수 없습니다.",
      });
    }
  }
  const invoiceKeys = new Set([...aGroups.keys(), ...bGroups.keys()]);
  for (const key of invoiceKeys) {
    const aBusinessRows = aGroups.get(key) ?? [];
    const bBusinessRows = bGroups.get(key) ?? [];
    if (aBusinessRows.length !== bBusinessRows.length) countMismatch += 1;
    if (aBusinessRows.length > bBusinessRows.length) aOnlyRows += aBusinessRows.length - bBusinessRows.length;
    if (bBusinessRows.length > aBusinessRows.length) bOnlyRows += bBusinessRows.length - aBusinessRows.length;

    // 사업자번호와 문서유형이 같은 그룹 안에서만 비교합니다. 날짜는 비교하지 않습니다.
    // 공급가액이 1,000원 미만으로 차이나는 경우에만 같은 계산서로 추정합니다.
    const unusedA = new Set(aBusinessRows);
    const pairs = new Map();
    for (const bRow of bBusinessRows) {
      const exactA = [...unusedA].find((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns) && isExactMatch(a.sheet, a.columns, aRow, b.sheet, b.columns, bRow));
      if (exactA !== undefined) { pairs.set(bRow, { aRow: exactA, matchMethod: "exact" }); unusedA.delete(exactA); }
    }
    let remainingBRows = bBusinessRows.filter((bRow) => !pairs.has(bRow));
    let compatibleA = [...unusedA].filter((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns));

    // 여러 행이 남아도 가장 가까운 금액 후보가 A/B 양쪽에서 하나씩으로 확정되면 대응합니다.
    // 최소 차액 후보가 동률이면 임의로 연결하지 않고 빨강으로 남깁니다.
    while (remainingBRows.length && compatibleA.length) {
      const nearestCandidates = [];
      for (const bRow of remainingBRows) {
        const candidates = compatibleA
          .map((aRow) => ({ aRow, difference: supplyDifference(cellAt(a.sheet, a.columns.supply, aRow), cellAt(b.sheet, b.columns.supply, bRow)) }))
          .filter((candidate) => candidate.difference !== null && candidate.difference < CLOSE_SUPPLY_DIFFERENCE);
        if (!candidates.length) continue;
        const minimum = Math.min(...candidates.map((candidate) => candidate.difference));
        const nearest = candidates.filter((candidate) => Math.abs(candidate.difference - minimum) < 0.000001);
        if (nearest.length === 1) nearestCandidates.push({ bRow, ...nearest[0] });
      }
      const candidateCountByA = new Map();
      for (const candidate of nearestCandidates) candidateCountByA.set(candidate.aRow, (candidateCountByA.get(candidate.aRow) ?? 0) + 1);
      const confirmed = nearestCandidates.filter((candidate) => candidateCountByA.get(candidate.aRow) === 1);
      if (!confirmed.length) break;
      for (const candidate of confirmed) {
        pairs.set(candidate.bRow, { aRow: candidate.aRow, matchMethod: "near", difference: candidate.difference });
        unusedA.delete(candidate.aRow);
        if (candidate.difference > 0.000001) nearSupplyMatches += 1;
      }
      remainingBRows = bBusinessRows.filter((bRow) => !pairs.has(bRow));
      compatibleA = [...unusedA].filter((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns));
    }
    const hasNearCandidate = remainingBRows.some((bRow) => compatibleA.some((aRow) => {
      const difference = supplyDifference(cellAt(a.sheet, a.columns.supply, aRow), cellAt(b.sheet, b.columns.supply, bRow));
      return difference !== null && difference < CLOSE_SUPPLY_DIFFERENCE;
    }));
    if (remainingBRows.length) {
      for (const bRow of remainingBRows) {
        const noCompatibleDocument = aBusinessRows.length > 0 && !aBusinessRows.some((aRow) => isDocumentTypeCompatible(a.sheet, a.columns, aRow, b.columns));
        const differences = compatibleA
          .map((aRow) => supplyDifference(cellAt(a.sheet, a.columns.supply, aRow), cellAt(b.sheet, b.columns.supply, bRow)))
          .filter((difference) => difference !== null);
        const smallestDifference = differences.length ? Math.min(...differences) : null;
        const hasCloseCandidate = smallestDifference !== null && smallestDifference < CLOSE_SUPPLY_DIFFERENCE;
        if (noCompatibleDocument) documentTypeMismatches += 1;
        else if (hasCloseCandidate) ambiguousBRows += 1;
        else if (smallestDifference !== null) differentInvoiceBRows += 1;
        paintMissingBRow(bRow);
        const candidateRows = aBusinessRows.map((row) => `${row}행`).join(", ") || "없음";
        const reason = noCompatibleDocument
          ? `같은 사업자번호의 A 후보(${candidateRows})는 있지만 문서 유형이 다릅니다. B는 ${bDocumentLabel(b.columns)}이고 A 후보와는 비교하지 않았습니다.`
          : hasCloseCandidate
            ? `같은 사업자번호에 공급가액 차이 1,000원 미만인 A 후보가 여러 개 남았습니다. 어느 행과 대응하는지 확정할 수 없어 임의로 노랑 처리하지 않고 빨강으로 표시했습니다.`
            : smallestDifference !== null
              ? `가장 가까운 A 후보와의 공급가액 차이가 ${amountLabel({ v: smallestDifference })}원으로 1,000원 이상입니다. 같은 계산서가 아닌 것으로 판단해 빨강으로 표시했습니다.`
            : "같은 사업자번호의 A 행을 찾지 못했습니다.";
        bDetail(bRow, { tone: "danger", status: noCompatibleDocument ? "문서 유형 불일치" : hasCloseCandidate ? "대응 관계 모호" : smallestDifference !== null ? "다른 계산서 추정" : "대응 불가", isIssue: true, reason, candidateRows });
      }
    }
    for (const [bRow, pair] of pairs) {
      const { aRow } = pair;
      const aSupply = cellAt(a.sheet, a.columns.supply, aRow);
      const bSupply = cellAt(b.sheet, b.columns.supply, bRow);
      const reasons = [];
      let hasMissingValue = false;
      if (!usable(aSupply)) { color(bSupply, COLORS.missing); red += 1; hasMissingValue = true; reasons.push("A의 공급가액이 비어 있습니다."); }
      else if (!sameNumber(aSupply, bSupply)) { color(bSupply, COLORS.different); yellow += 1; supplyDifferences += 1; reasons.push(`공급가액이 다릅니다 (A ${amountLabel(aSupply)} / B ${amountLabel(bSupply)}).`); }
      if (pair.matchMethod === "near" && pair.difference > 0.000001) reasons.unshift(`공급가액 차이가 ${amountLabel({ v: pair.difference })}원으로 1,000원 미만이어서 같은 계산서로 추정했습니다.`);

      if (taxMode(cellAt(a.sheet, a.columns.taxType, aRow)) === "disallowed" && b.columns.tax) {
        const aTax = cellAt(a.sheet, a.columns.tax, aRow);
        const bTax = cellAt(b.sheet, b.columns.tax, bRow);
        if (!usable(aTax)) { color(bTax, COLORS.missing); red += 1; hasMissingValue = true; reasons.push("불공 행인데 A의 세액이 비어 있습니다."); }
        else if (!sameNumber(aTax, bTax)) { color(bTax, COLORS.different); yellow += 1; taxDifferences += 1; reasons.push(`불공 행의 세액이 다릅니다 (A ${amountLabel(aTax)} / B ${amountLabel(bTax)}).`); }
      }
      if (hasMissingValue || reasons.length) {
        bDetail(bRow, {
          tone: hasMissingValue ? "danger" : "warning",
          status: hasMissingValue ? "입력값 확인 필요" : "값 차이",
          isIssue: true,
          aRow,
          aDocument: aDocumentLabel(a.sheet, a.columns, aRow),
          aSupply: amountLabel(aSupply),
          aTax: amountLabel(cellAt(a.sheet, a.columns.tax, aRow)),
          reason: reasons.join(" "),
        });
      }
    }
    // 1,000원 이상 차이만 남은 경우는 서로 다른 계산서로 보므로 A 행도 파랑으로 표시합니다.
    // 1,000원 미만 후보가 여러 개라 모호한 경우에는 파랑으로 단정하지 않습니다.
    if (remainingBRows.length === 0 || compatibleA.length === 0 || !hasNearCandidate) {
      for (const aRow of unusedA) aRowsToPaint.add(aRow);
    }
    matchedRows += pairs.size;
  }
  for (const row of aRowsToPaint) {
    const target = cellAt(a.sheet, a.columns.supply, row);
    if (usable(target)) {
      color(target, COLORS.aOnly); blueRows += 1;
      aDetail(row, {
        tone: "info", status: "B 대응 없음", isIssue: true,
        reason: "B 파일에서 같은 사업자번호·문서 유형으로 확정할 수 있는 대응 행이 없어 A 결과 파일의 공급가액 셀을 파랑으로 표시했습니다.",
      });
    }
  }
  details.sort((left, right) => (left.source === right.source ? left.sourceRow - right.sourceRow : left.source === "B" ? -1 : 1));
  return {
    bData: XLSX.write(bBook, { type: "array", bookType: "xlsx", cellStyles: true }),
    aData: XLSX.write(aBook, { type: "array", bookType: "xlsx", cellStyles: true }),
    yellow, red, blueRows, countMismatch, aOnlyRows, bOnlyRows, matchedRows, supplyDifferences, taxDifferences, nearSupplyMatches, unmatchedBRows, documentTypeMismatches, ambiguousBRows, differentInvoiceBRows,
    aRowCount: aRows.length, bRowCount: bRows.length, bHasTax: Boolean(b.columns.tax), aSheetName: a.name, bSheetName: b.name, details,
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
const detailSection = document.querySelector("#detail-section");
const detailDescription = document.querySelector("#detail-description");
const detailList = document.querySelector("#detail-list");
function summaryCard(label, value, tone = "") {
  const card = document.createElement("div"); card.className = `summary-card ${tone}`;
  const labelElement = document.createElement("span"); labelElement.textContent = label;
  const valueElement = document.createElement("strong"); valueElement.textContent = value;
  card.append(labelElement, valueElement); return card;
}
function summaryNote(textValue, tone = "") {
  const note = document.createElement("li"); note.className = tone; note.textContent = textValue; return note;
}
function downloadWorkbook(data, filename) {
  const link = document.createElement("a");
  const url = URL.createObjectURL(new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}
function detailValue(label, value) {
  const item = document.createElement("span");
  const name = document.createElement("strong");
  name.textContent = `${label}: `;
  item.append(name, value);
  return item;
}
function detailItem(detail) {
  const item = document.createElement("article");
  item.className = `detail-item ${detail.tone}`;
  const top = document.createElement("div"); top.className = "detail-top";
  const status = document.createElement("span"); status.className = "detail-status"; status.textContent = detail.status;
  const row = document.createElement("span"); row.className = "detail-row"; row.textContent = `${detail.source} 파일 ${detail.sourceRow}행`;
  top.append(status, row);
  const key = document.createElement("p"); key.className = "detail-key"; key.textContent = `사업자번호 ${detail.businessNumber}`;
  const values = document.createElement("div"); values.className = "detail-values";
  if (detail.source === "B") {
    values.append(
      detailValue("B 문서", detail.bDocument),
      detailValue("A 대응", detail.aRow ? `${detail.aRow}행${detail.aDocument ? ` · ${detail.aDocument}` : ""}` : detail.candidateRows ? `${detail.candidateRows} 후보` : "없음"),
      detailValue("공급가액", detail.aSupply ? `A ${detail.aSupply} / B ${detail.bSupply}` : `B ${detail.bSupply}`),
    );
    if (detail.bTax !== "해당 없음") values.append(detailValue("세액", detail.aTax ? `A ${detail.aTax} / B ${detail.bTax}` : `B ${detail.bTax}`));
  } else {
    values.append(detailValue("A 문서", detail.aDocument), detailValue("공급가액", detail.aSupply), detailValue("세액", detail.aTax));
  }
  const reason = document.createElement("p"); reason.className = "detail-reason"; reason.textContent = detail.reason;
  item.append(top, key, values, reason);
  return item;
}
function renderDetails(result) {
  detailSection.hidden = result.details.length === 0;
  detailDescription.textContent = `색을 표시한 ${result.details.length}행의 판정 근거입니다. B 행을 먼저, B에 대응하지 않는 A 행을 뒤에 표시합니다.`;
  detailList.replaceChildren(...result.details.map(detailItem));
}
function renderSummary(result) {
  const hasIssues = result.details.length > 0;
  summary.hidden = false;
  summaryTitle.textContent = hasIssues ? "비교 결과를 확인해 주세요" : "모든 비교가 정상 완료됐어요";
  summaryDescription.textContent = `${result.aSheetName} ↔ ${result.bSheetName} · B 문서 유형: ${result.bHasTax ? "전자세금계산서" : "전자계산서(면세)"}`;
  summaryCards.replaceChildren(
    summaryCard("읽은 행", `A ${result.aRowCount} · B ${result.bRowCount}`),
    summaryCard("A와 대응된 행", `${result.matchedRows}행`, result.matchedRows === result.bRowCount ? "success" : ""),
    summaryCard("공급가액 차이", `${result.supplyDifferences}건`, result.supplyDifferences ? "warning" : ""),
    summaryCard("1천원 미만 추정", `${result.nearSupplyMatches}행`, result.nearSupplyMatches ? "warning" : ""),
    summaryCard("세액 차이", result.bHasTax ? `${result.taxDifferences}건` : "해당 없음", result.taxDifferences ? "warning" : ""),
    summaryCard("대응 불가 B행", `${result.unmatchedBRows}행`, result.unmatchedBRows ? "danger" : ""),
    summaryCard("파랑 A행", `${result.blueRows}행`, result.blueRows ? "danger" : ""),
  );
  const notes = [];
  if (result.supplyDifferences) notes.push(summaryNote(`공급가액이 다른 행이 ${result.supplyDifferences}건 있어 B의 공급가액 셀을 노랑으로 표시했습니다.`, "warning"));
  if (result.nearSupplyMatches) notes.push(summaryNote(`공급가액 차이가 1,000원 미만인 ${result.nearSupplyMatches}행은 같은 계산서로 추정해 노랑으로 표시했습니다.`, "warning"));
  if (result.taxDifferences) notes.push(summaryNote(`불공 행 중 세액이 다른 행이 ${result.taxDifferences}건 있어 B의 세액 셀을 노랑으로 표시했습니다.`, "warning"));
  if (result.documentTypeMismatches) notes.push(summaryNote(`같은 사업자번호지만 면세/세금계산서 문서 유형이 맞지 않는 B 행이 ${result.documentTypeMismatches}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.ambiguousBRows) notes.push(summaryNote(`같은 사업자번호에 여러 행이 남아 대응 관계가 모호한 B 행이 ${result.ambiguousBRows}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.differentInvoiceBRows) notes.push(summaryNote(`가장 가까운 공급가액과도 1,000원 이상 차이 나는 B 행이 ${result.differentInvoiceBRows}건 있어 다른 계산서로 보고 빨강으로 표시했습니다.`, "danger"));
  if (result.unmatchedBRows > result.documentTypeMismatches + result.ambiguousBRows + result.differentInvoiceBRows) notes.push(summaryNote(`A에서 대응 행을 찾지 못한 B 행이 ${result.unmatchedBRows - result.documentTypeMismatches - result.ambiguousBRows - result.differentInvoiceBRows}건 있어 빨강으로 표시했습니다.`, "danger"));
  if (result.countMismatch) notes.push(summaryNote(`사업자번호별 행 수가 다른 그룹이 ${result.countMismatch}개입니다.`, "danger"));
  if (result.blueRows) notes.push(summaryNote(`B에 대응 행이 없는 A 계산서 ${result.blueRows}건을 A 결과 파일의 공급가액 셀에 파랑으로 표시했습니다.`, "danger"));
  if (!notes.length) notes.push(summaryNote("사업자번호, 문서 유형, 공급가액 기준으로 모든 B 행이 정상 대응됐습니다."));
  summaryNotes.replaceChildren(...notes);
  renderDetails(result);
}
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  button.disabled = true;
  summary.hidden = true;
  status.textContent = "이 브라우저 안에서 엑셀 헤더와 사업자번호를 확인하는 중입니다…";
  try {
    const aFile = form.elements.aFile.files[0]; const bFile = form.elements.bFile.files[0];
    const result = compareByBusinessGroups(new Uint8Array(await aFile.arrayBuffer()), new Uint8Array(await bFile.arrayBuffer()));
    downloadWorkbook(result.bData, "B_색상표시_비교결과.xlsx");
    downloadWorkbook(result.aData, "A_파랑표시_비교결과.xlsx");
    status.textContent = "비교가 완료되어 B·A 결과 엑셀 파일을 각각 다운로드했습니다.";
    renderSummary(result);
  } catch (error) {
    status.textContent = `처리 오류: ${error.message}`;
  } finally {
    button.disabled = false;
  }
});
