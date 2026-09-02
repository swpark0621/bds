import fs from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const snapshotDate = new Date('2026-09-02T12:00:00+09:00');
const outputDir = resolve(root, 'outputs', '신흥1_신흥3_실거래_호가_20260902');
const outputPath = resolve(outputDir, '성남_신흥1_신흥3_실거래_호가_20260902.xlsx');
const renderDir = resolve(outputDir, 'renders');

const sources = {
  molit: 'https://rt.molit.go.kr/pt/qut/qut.do',
  naverMap: 'https://fin.land.naver.com/map?center=3zoOEu-2AH2TC&realEstateTypes=F01&tradeTypes=A1-B1-B2-B3&zoom=17.49880630642608&showOnlySelectedRegion=true',
  shin1Project: 'https://jaegebal.co.kr/develops/4027',
  shin1Asks: 'https://jaegebal.co.kr/develops/4027/asks',
  shin3Project: 'https://jaegebal.co.kr/develops/4072',
  shin3Asks: 'https://jaegebal.co.kr/develops/4072/asks',
};

const zones = [
  { key: '신흥1구역', id: 4027, projectUrl: sources.shin1Project, asksUrl: sources.shin1Asks },
  { key: '신흥3구역', id: 4072, projectUrl: sources.shin3Project, asksUrl: sources.shin3Asks },
];

const naverHistory = [
  ['신흥1구역', new Date('2025-05-28'), '2528723018', '단독/다가구', 68000, 7800, 70, null, 72.1, null, '신흥1구역 명시 과거 확인 매물. 월세 포함으로 순투입금 비교 제외.', 'https://fin.land.naver.com/articles/2528723018'],
  ['신흥3구역', new Date('2025-05-12'), '2525197013', '다세대', 120000, 25000, 0, null, 53.25, 44300, '신흥3구역 명시 과거 확인 매물.', 'https://fin.land.naver.com/articles/2525197013'],
  ['신흥3구역', new Date('2025-06-10'), '2531179199', '다가구', 62000, 24000, 0, 38000, 62.5, 23100, '신흥3구역 명시 과거 확인 매물.', 'https://fin.land.naver.com/articles/2531179199'],
  ['신흥3구역', new Date('2025-06-12'), '2531686054', '단독/다가구', 65000, 24000, 0, 41000, 59.8, null, '신흥3구역 명시 과거 확인 매물.', 'https://fin.land.naver.com/articles/2531686054'],
  ['신흥3구역', new Date('2025-06-12'), '2531686562', '단독', 120000, 24000, 150, 96000, 133.6, 50600, '신흥3구역 명시 과거 확인 매물. 월세 포함이라 초투입금 평당 비교 제외.', 'https://fin.land.naver.com/articles/2531686562'],
];

function colName(index) {
  let value = index + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function cellRange(startRow, startCol, rowCount, colCount) {
  const start = `${colName(startCol)}${startRow}`;
  const end = `${colName(startCol + colCount - 1)}${startRow + rowCount - 1}`;
  return `${start}:${end}`;
}

function median(values) {
  const filtered = values.filter((value) => Number.isFinite(value)).sort((a, b) => a - b);
  if (!filtered.length) return null;
  const mid = Math.floor(filtered.length / 2);
  return filtered.length % 2 ? filtered[mid] : (filtered[mid - 1] + filtered[mid]) / 2;
}

function formatDate(value) {
  return value.toISOString().slice(0, 10);
}

function emptyToNull(value) {
  return value === undefined || value === null || value === -1 || value === '' ? null : value;
}

function displayPropertyType(value) {
  const normalized = String(value ?? '').trim().toLowerCase();
  const labels = {
    dandok: '단독',
    dasedae: '다세대',
    yeonlib: '연립',
    sangga: '상가',
  };
  return labels[normalized] ?? value ?? '공개값 미표기';
}

function findJsonArray(text, start) {
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (inString) {
      if (escaped) escaped = false;
      else if (char === '\\') escaped = true;
      else if (char === '"') inString = false;
      continue;
    }
    if (char === '"') inString = true;
    else if (char === '[') depth += 1;
    else if (char === ']') {
      depth -= 1;
      if (depth === 0) return text.slice(start, index + 1);
    }
  }
  throw new Error('공개 호가 배열의 끝을 찾지 못했습니다.');
}

async function fetchTransactions(zone) {
  const response = await fetch(`https://jaegebal.co.kr/api/develops/${zone.id}/property-transactions`, {
    method: 'POST',
    headers: {
      'X-JDC-App-Token': '0418b9cadfb234a22035c05e982492cf02b62fa91bc89e217f8f93c7e6f0e10f',
      'Content-Type': 'application/json',
      Origin: 'https://jaegebal.co.kr',
      Referer: zone.projectUrl,
    },
    body: JSON.stringify({
      limit: 100,
      offset: 0,
      filters: {
        types: [],
        made_at: { start: '2025-09-01', end: '2026-08-31' },
        building_age: { start: null, end: null },
        junyoung_area: { start: null, end: null },
        land_area: { start: null, end: null },
        price: { start: null, end: null },
        official_price: { start: null, end: null },
        junyoung_type: null,
        is_apt: null,
      },
    }),
  });
  if (!response.ok) throw new Error(`${zone.key} 실거래 API 응답 오류: ${response.status}`);
  const payload = await response.json();
  return (payload.data ?? payload.transactions ?? []).slice(0, 20).map((record) => ({ ...record, zone: zone.key }));
}

async function fetchAsks(zone) {
  const response = await fetch(zone.asksUrl, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (!response.ok) throw new Error(`${zone.key} 공개 호가 페이지 응답 오류: ${response.status}`);
  const html = await response.text();
  const marker = '\\"asks\\":{\\"asks\\":';
  const markerIndex = html.indexOf(marker);
  if (markerIndex < 0) throw new Error(`${zone.key} 공개 호가 데이터 마커를 찾지 못했습니다.`);
  const decoded = html.slice(markerIndex + marker.length).replaceAll('\\"', '"');
  const arrayStart = decoded.indexOf('[');
  if (arrayStart < 0) throw new Error(`${zone.key} 공개 호가 배열 시작점을 찾지 못했습니다.`);
  const asks = JSON.parse(findJsonArray(decoded, arrayStart));
  return asks.map((ask) => ({ ...ask, zone: zone.key }));
}

function prepareTransactions(rows) {
  return rows
    .sort((a, b) => String(b.made_at).localeCompare(String(a.made_at)))
    .map((record) => ({
      zone: record.zone,
      madeAt: new Date(`${record.made_at}T12:00:00Z`),
      type: displayPropertyType(record.type_value ?? record.type),
      address: record.address ?? '',
      building: record.building ?? '',
      floor: emptyToNull(record.floor),
      constructionYear: emptyToNull(record.construction_year),
      housingType: record.housing_type ?? '',
      price: emptyToNull(record.price),
      landArea: emptyToNull(record.land_area),
      sourceLandPyeongPrice: emptyToNull(record.land_price),
      officialPrice: emptyToNull(record.official_price),
      officialPriceYear: emptyToNull(record.official_price_year),
      canceledAt: record.canceled_at ? new Date(`${record.canceled_at}T12:00:00Z`) : null,
      transactionId: record.id ?? record._id ?? '',
    }));
}

function prepareAsks(rows) {
  return rows
    .sort((a, b) => (b.price ?? 0) - (a.price ?? 0))
    .map((ask) => ({
      zone: ask.zone,
      askId: ask.id ?? ask._id ?? '',
      type: displayPropertyType(ask.property?.type_value ?? ask.property?.type ?? ask.type_value ?? ask.type),
      price: emptyToNull(ask.price),
      address: ask.property?.address ?? ask.address ?? '',
      building: ask.property?.building ?? ask.building ?? '',
      floor: emptyToNull(ask.property?.floor ?? ask.floor),
      floorArea: emptyToNull(ask.property?.floor_area ?? ask.floor_area),
      constructionYear: emptyToNull(ask.property?.construction_year ?? ask.construction_year),
      createdAt: ask.created_at ? new Date(ask.created_at) : null,
      landArea: emptyToNull(ask.property?.land_area ?? ask.land_area),
      officialPrice: emptyToNull(ask.property?.official_price ?? ask.official_price),
    }));
}

function title(sheet, text, endColumn) {
  const range = sheet.getRange(`A1:${endColumn}1`);
  range.merge();
  range.values = [[text]];
  range.format = {
    fill: '#17324D',
    font: { bold: true, color: '#FFFFFF', size: 16 },
    horizontalAlignment: 'left',
    verticalAlignment: 'center',
  };
  range.format.rowHeight = 30;
}

function subtitle(sheet, text, endColumn) {
  const range = sheet.getRange(`A2:${endColumn}2`);
  range.merge();
  range.values = [[text]];
  range.format = {
    fill: '#E8F0F5',
    font: { color: '#34495E', italic: true, size: 10 },
    wrapText: true,
    verticalAlignment: 'center',
  };
  range.format.rowHeight = 34;
}

function section(sheet, row, text, endColumn) {
  const range = sheet.getRange(`A${row}:${endColumn}${row}`);
  range.merge();
  range.values = [[text]];
  range.format = {
    fill: '#2F6B5F',
    font: { bold: true, color: '#FFFFFF' },
    verticalAlignment: 'center',
  };
  range.format.rowHeight = 22;
}

function header(sheet, row, labels) {
  const range = sheet.getRange(cellRange(row, 0, 1, labels.length));
  range.values = [labels];
  range.format = {
    fill: '#DCE7EE',
    font: { bold: true, color: '#183047' },
    horizontalAlignment: 'center',
    verticalAlignment: 'center',
    wrapText: true,
    borders: { preset: 'all', style: 'thin', color: '#C9D5DE' },
  };
  range.format.rowHeight = 30;
}

function bodyBorders(sheet, rangeAddress) {
  sheet.getRange(rangeAddress).format.borders = { preset: 'inside', style: 'thin', color: '#E2E8ED' };
}

function setWidths(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.getRange(`${colName(index)}:${colName(index)}`).format.columnWidth = width;
  });
}

function decorateTable(sheet, headerRow, lastRow, lastCol) {
  const range = sheet.getRange(`A${headerRow}:${lastCol}${lastRow}`);
  range.format.borders = { preset: 'outside', style: 'thin', color: '#C9D5DE' };
  bodyBorders(sheet, `A${headerRow + 1}:${lastCol}${lastRow}`);
  sheet.freezePanes.freezeRows(headerRow);
  sheet.getRange(`A${headerRow}:${lastCol}${lastRow}`).format.wrapText = true;
}

function transactionLink(zone) {
  return zone === '신흥1구역' ? sources.shin1Project : sources.shin3Project;
}

const [shin1Transactions, shin3Transactions, shin1Asks, shin3Asks] = await Promise.all([
  fetchTransactions(zones[0]),
  fetchTransactions(zones[1]),
  fetchAsks(zones[0]),
  fetchAsks(zones[1]),
]);

const preparedTransactions = prepareTransactions([...shin1Transactions, ...shin3Transactions]);
const preparedAsks = prepareAsks([...shin1Asks, ...shin3Asks]);
const transactions = zones.flatMap((zone) => preparedTransactions.filter((row) => row.zone === zone.key));
const asks = zones.flatMap((zone) => preparedAsks.filter((row) => row.zone === zone.key));
const transactionRowsByZone = new Map(zones.map((zone) => [zone.key, transactions.filter((row) => row.zone === zone.key)]));
const askRowsByZone = new Map(zones.map((zone) => [zone.key, asks.filter((row) => row.zone === zone.key)]));

const workbook = Workbook.create();
const summary = workbook.worksheets.add('요약');
const transactionsSheet = workbook.worksheets.add('실거래_최신20');
const asksSummarySheet = workbook.worksheets.add('현호가_보조');
const asksRawSheet = workbook.worksheets.add('현호가_원자료');
const naverSheet = workbook.worksheets.add('네이버_확인이력');
const sourceSheet = workbook.worksheets.add('출처_방법');

[summary, transactionsSheet, asksSummarySheet, asksRawSheet, naverSheet, sourceSheet].forEach((sheet) => { sheet.showGridLines = false; });

// Summary
title(summary, '성남 신흥1·신흥3 재개발 | 실거래·호가 비교', 'J');
subtitle(summary, '기준일 2026-09-02. 실거래는 공개 재개발 플랫폼의 최근 신고 표본(구역별 20건)을 정리한 것이며, 계약 전 국토교통부 실거래가 공개시스템 원문으로 재확인해야 합니다.', 'J');
section(summary, 4, '1. 구역별 비교', 'J');
header(summary, 5, ['구역', '법정 단계', '표본 기간', '실거래 표본', '거래가 중앙값\n(억원)', '대지평당가 중앙값\n(만원/평)', '대지평당가 평균\n(만원/평)', '현 호가 수\n(보조)', '현 호가 최저\n(억원)', '판독']);
const summaryRows = zones.map((zone, index) => {
  const start = 6 + (index === 0 ? 0 : transactionRowsByZone.get(zones[0].key).length);
  const end = start + transactionRowsByZone.get(zone.key).length - 1;
  const askStart = 6 + (index === 0 ? 0 : askRowsByZone.get(zones[0].key).length);
  const askEnd = askStart + askRowsByZone.get(zone.key).length - 1;
  return [
    zone.key,
    zone.key === '신흥1구역' ? '사업시행계획인가' : '주민대표회의 구성\n(사업시행인가 전)',
    `${formatDate(transactionRowsByZone.get(zone.key).at(-1).madeAt)} ~ ${formatDate(transactionRowsByZone.get(zone.key)[0].madeAt)}`,
    `=COUNT('실거래_최신20'!$I$${start}:$I$${end})`,
    `=MEDIAN('실거래_최신20'!$J$${start}:$J$${end})`,
    `=MEDIAN('실거래_최신20'!$M$${start}:$M$${end})`,
    `=AVERAGE('실거래_최신20'!$M$${start}:$M$${end})`,
    `=COUNT('현호가_원자료'!$D$${askStart}:$D$${askEnd})`,
    `=MIN('현호가_원자료'!$E$${askStart}:$E$${askEnd})`,
    zone.key === '신흥1구역' ? '단계 우세. 다만 관리처분·이주 변수 남음.' : '시공자 선정과 법정 인가 단계는 별개.'
  ];
});
summary.getRange('A6:J7').values = summaryRows.map((row) => row.map((value, index) => ([3, 4, 5, 6, 7, 8].includes(index) ? null : value)));
summary.getRange('D6:I7').formulas = summaryRows.map((row) => row.slice(3, 9));
summary.getRange('A5:J7').format.wrapText = true;
decorateTable(summary, 5, 7, 'J');
summary.getRange('D6:D7').format.numberFormat = '#,##0';
summary.getRange('E6:I7').format.numberFormat = '#,##0.0';

section(summary, 10, '2. 프리미엄 해석', 'J');
summary.getRange('A11:C11').values = [['항목', '엑셀 표시', '산식']];
summary.getRange('D11:E11').merge();
summary.getRange('D11').values = [['투자 판단에서의 위치']];
summary.getRange('F11:J11').merge();
summary.getRange('F11').values = [['주의']];
const premiumRows = [
  ['대지평당 거래가', '실거래_최신20', '실거래가 ÷ (대지면적㎡ ÷ 3.305785)', '같은 주택 유형·대지지분 비교의 출발점', '단독/다가구와 다세대는 대지권 구조가 달라 분리 비교'],
  ['공시가격 차액', '실거래_최신20', '실거래가 - 공시가격', '가격 수준 참고', '재개발 프리미엄이 아님'],
  ['실제 재개발 프리미엄', '확인 불가', '매매가 - 종전자산 감정평가액 또는 권리가액', '추가분담금·예상 분양가 산정의 출발점', '개별 물건의 감정평가/권리가액 없이는 산출 금지'],
  ['현 호가', '현호가_보조', '공개 호가 가격', '거래 성사 가능 범위 참고', '네이버 현행 개별 매물의 구역별 대지지분을 일괄 추출할 수 없어 대지평당 호가 비교는 제외'],
];
premiumRows.forEach((row, index) => {
  const currentRow = 12 + index;
  summary.getRange(`A${currentRow}:C${currentRow}`).values = [row.slice(0, 3)];
  summary.getRange(`D${currentRow}:E${currentRow}`).merge();
  summary.getRange(`D${currentRow}`).values = [[row[3]]];
  summary.getRange(`F${currentRow}:J${currentRow}`).merge();
  summary.getRange(`F${currentRow}`).values = [[row[4]]];
});
summary.getRange('A11:J15').format.borders = { preset: 'all', style: 'thin', color: '#D4DDE4' };
summary.getRange('A11:J11').format = { fill: '#DCE7EE', font: { bold: true, color: '#183047' }, wrapText: true };
summary.getRange('A12:J15').format.wrapText = true;
summary.getRange('C12:C15').format.font = { color: '#008000' };
summary.getRange('A11:J15').format.rowHeight = 42;

section(summary, 18, '3. 사용 전 확인', 'J');
summary.getRange('A19:C23').values = [
  ['체크', '상태', '설명'],
  ['정비구역 편입', '필수', '후보 지번을 정비계획 도면·토지이음으로 대조'],
  ['분양자격', '필수', '소유권 취득일·공유지분·다물권·권리산정기준일 확인'],
  ['실제 재개발 P', '필수', '종전자산 감정평가·권리가액·추정 분담금 자료가 있는 동일 물건으로 산정'],
  ['실거래 확정', '필수', '계약 전 국토교통부 실거래가 공개시스템에서 계약일·취소 여부·주소 재확인'],
];
for (let row = 19; row <= 23; row += 1) summary.getRange(`C${row}:J${row}`).merge();
summary.getRange('A19:J23').format.borders = { preset: 'all', style: 'thin', color: '#D4DDE4' };
summary.getRange('A19:J19').format = { fill: '#DCE7EE', font: { bold: true, color: '#183047' } };
summary.getRange('B20:B23').format = { fill: '#FFF4CC', font: { bold: true, color: '#7A4F00' }, horizontalAlignment: 'center' };
summary.getRange('C19:J23').format.wrapText = true;
summary.getRange('A20:J23').format.rowHeight = 34;
setWidths(summary, [16, 22, 30, 14, 16, 17, 17, 14, 14, 28]);
summary.freezePanes.freezeRows(5);

// Transactions
title(transactionsSheet, '실거래 신고 표본 | 구역별 최신 20건', 'T');
subtitle(transactionsSheet, '단위: 가격은 만원/억원, 면적은 ㎡/평. 원자료는 공개 재개발 플랫폼의 구역별 2025-09-01~2026-08-31 필터 결과 중 최신 20건이며, 원계약 사실은 국토부 시스템에서 재확인 필요.', 'T');
header(transactionsSheet, 5, ['구역', '계약일', '유형', '주소', '동/건물', '층', '준공', '주택형태', '거래가\n(만원)', '거래가\n(억원)', '대지면적\n(㎡)', '대지면적\n(평)', '대지평당가\n(만원/평)', '원자료 대지평당가\n(만원/평)', '공시가격\n(만원)', '공시가격\n(억원)', '공시가격 차액\n(억원)', '매매/공시\n(배)', '취소일', '출처']);
const transactionValueRows = transactions.map((row) => [
  row.zone, row.madeAt, row.type, row.address, row.building, row.floor, row.constructionYear, row.housingType,
  row.price, null, row.landArea, null, null, row.sourceLandPyeongPrice, row.officialPrice, null, null, null, row.canceledAt, transactionLink(row.zone),
]);
const transactionStartRow = 6;
const transactionEndRow = transactionStartRow + transactionValueRows.length - 1;
transactionsSheet.getRange(`A${transactionStartRow}:T${transactionEndRow}`).values = transactionValueRows;
transactionsSheet.getRange(`J${transactionStartRow}:J${transactionEndRow}`).formulas = transactions.map((_, index) => [`=I${transactionStartRow + index}/10000`]);
transactionsSheet.getRange(`L${transactionStartRow}:M${transactionEndRow}`).formulas = transactions.map((_, index) => [
  `=IFERROR(K${transactionStartRow + index}/3.305785,"")`,
  `=IFERROR(I${transactionStartRow + index}/L${transactionStartRow + index},"")`,
]);
transactionsSheet.getRange(`P${transactionStartRow}:R${transactionEndRow}`).formulas = transactions.map((_, index) => [
  `=IFERROR(O${transactionStartRow + index}/10000,"")`,
  `=IFERROR((I${transactionStartRow + index}-O${transactionStartRow + index})/10000,"")`,
  `=IFERROR(I${transactionStartRow + index}/O${transactionStartRow + index},"")`,
]);
decorateTable(transactionsSheet, 5, transactionEndRow, 'T');
transactionsSheet.getRange(`B${transactionStartRow}:B${transactionEndRow}`).format.numberFormat = 'yyyy-mm-dd';
transactionsSheet.getRange(`S${transactionStartRow}:S${transactionEndRow}`).format.numberFormat = 'yyyy-mm-dd';
transactionsSheet.getRange(`I${transactionStartRow}:I${transactionEndRow}`).format.numberFormat = '#,##0';
transactionsSheet.getRange(`J${transactionStartRow}:J${transactionEndRow}`).format.numberFormat = '#,##0.00';
transactionsSheet.getRange(`K${transactionStartRow}:K${transactionEndRow}`).format.numberFormat = '#,##0.0';
transactionsSheet.getRange(`L${transactionStartRow}:N${transactionEndRow}`).format.numberFormat = '#,##0.0';
transactionsSheet.getRange(`O${transactionStartRow}:O${transactionEndRow}`).format.numberFormat = '#,##0';
transactionsSheet.getRange(`P${transactionStartRow}:Q${transactionEndRow}`).format.numberFormat = '#,##0.00';
transactionsSheet.getRange(`R${transactionStartRow}:R${transactionEndRow}`).format.numberFormat = '0.00x';
transactionsSheet.getRange(`T${transactionStartRow}:T${transactionEndRow}`).format.font = { color: '#008000' };
setWidths(transactionsSheet, [13, 13, 14, 16, 16, 8, 9, 13, 13, 11, 12, 12, 14, 16, 14, 12, 15, 11, 13, 35]);

// Current asking-price helper summary
title(asksSummarySheet, '현 호가 보조 스냅샷 | 재개발닷컴 공개 페이지', 'J');
subtitle(asksSummarySheet, '2026-09-02 조회. 공개 호가 원자료에는 대지면적·공시가격이 비공개(-1)인 행이 많아 대지평당 호가와 실제 재개발 프리미엄은 계산하지 않았습니다. 네이버부동산 현행 지도는 참고용으로만 확인했습니다.', 'J');
header(asksSummarySheet, 5, ['구역', '공개 호가 수', '최저 호가\n(억원)', '중앙 호가\n(억원)', '최고 호가\n(억원)', '대지면적 공개', '대지평당 호가', '실거래와의 비교', '출처', '해석']);
const askSummaryRows = zones.map((zone, index) => {
  const start = 6 + (index === 0 ? 0 : askRowsByZone.get(zones[0].key).length);
  const end = start + askRowsByZone.get(zone.key).length - 1;
  return [
    zone.key,
    `=COUNT('현호가_원자료'!$D$${start}:$D$${end})`,
    `=MIN('현호가_원자료'!$E$${start}:$E$${end})`,
    `=MEDIAN('현호가_원자료'!$E$${start}:$E$${end})`,
    `=MAX('현호가_원자료'!$E$${start}:$E$${end})`,
    '대부분 비공개',
    '산출 제외',
    '유형·임차조건 통일 후 개별 비교',
    zone.asksUrl,
    '호가는 거래 성사 가격이 아님',
  ];
});
asksSummarySheet.getRange('A6:J7').values = askSummaryRows.map((row) => row.map((value, index) => ([1, 2, 3, 4].includes(index) ? null : value)));
asksSummarySheet.getRange('B6:E7').formulas = askSummaryRows.map((row) => row.slice(1, 5));
decorateTable(asksSummarySheet, 5, 7, 'J');
asksSummarySheet.getRange('B6:B7').format.numberFormat = '#,##0';
asksSummarySheet.getRange('C6:E7').format.numberFormat = '#,##0.00';
asksSummarySheet.getRange('I6:I7').format.font = { color: '#008000' };
section(asksSummarySheet, 10, '네이버부동산 확인 범위와 한계', 'J');
asksSummarySheet.getRange('A11:C14').values = [
  ['항목', '확인 결과', '워크북 반영'],
  ['현행 지도 필터', '성남시 수정구 신흥동 · 재개발(F01) · 매매/전세/월세/단기임대', '출처_방법에 지도 URL 기록'],
  ['구역별 실시간 개별 매물', '동적 지도에서 구역 경계·대지지분을 일괄 식별할 수 없음', '현재 네이버 호가를 신흥1·3별 대지평당가로 산출하지 않음'],
  ['개별 확인 이력', '신흥1 1건, 신흥3 4건의 구역명 명시 매물을 2025년 확인', '네이버_확인이력에 과거 확인일과 원문 URL 기록'],
];
for (let row = 11; row <= 14; row += 1) asksSummarySheet.getRange(`C${row}:J${row}`).merge();
asksSummarySheet.getRange('A11:J14').format.borders = { preset: 'all', style: 'thin', color: '#D4DDE4' };
asksSummarySheet.getRange('A11:J11').format = { fill: '#DCE7EE', font: { bold: true, color: '#183047' } };
asksSummarySheet.getRange('A11:J14').format.wrapText = true;
asksSummarySheet.getRange('A11:J14').format.rowHeight = 40;
setWidths(asksSummarySheet, [15, 14, 15, 15, 15, 16, 16, 28, 36, 26]);
asksSummarySheet.freezePanes.freezeRows(5);

// Current asking-price raw data
title(asksRawSheet, '현 호가 공개 원자료 | 보조 지표', 'L');
subtitle(asksRawSheet, '공개 웹 페이지에 표시된 호가를 스냅샷으로 정리했습니다. 원자료의 대지면적·공시가격이 비공개(-1)이면 빈칸으로 처리했습니다. 호가=계약가가 아니며, 네이버부동산 자료가 아닙니다.', 'L');
header(asksRawSheet, 5, ['구역', '매물 ID', '유형', '호가\n(만원)', '호가\n(억원)', '주소', '동/건물', '층', '연면적\n(㎡)', '준공', '대지면적\n(㎡)', '공시가격\n(만원)']);
const askStartRow = 6;
const askEndRow = askStartRow + asks.length - 1;
asksRawSheet.getRange(`A${askStartRow}:L${askEndRow}`).values = asks.map((row) => [
  row.zone, row.askId, row.type, row.price, null, row.address, row.building, row.floor, row.floorArea, row.constructionYear, row.landArea, row.officialPrice,
]);
asksRawSheet.getRange(`E${askStartRow}:E${askEndRow}`).formulas = asks.map((_, index) => [`=D${askStartRow + index}/10000`]);
decorateTable(asksRawSheet, 5, askEndRow, 'L');
asksRawSheet.getRange(`D${askStartRow}:D${askEndRow}`).format.numberFormat = '#,##0';
asksRawSheet.getRange(`E${askStartRow}:E${askEndRow}`).format.numberFormat = '#,##0.00';
asksRawSheet.getRange(`I${askStartRow}:K${askEndRow}`).format.numberFormat = '#,##0.0';
asksRawSheet.getRange(`L${askStartRow}:L${askEndRow}`).format.numberFormat = '#,##0';
setWidths(asksRawSheet, [15, 15, 14, 13, 11, 16, 16, 8, 12, 9, 12, 14]);

// Naver historical listing confirmation
title(naverSheet, '네이버부동산 | 구역명 명시 매물 확인이력', 'O');
subtitle(naverSheet, '2025년 개별 매물 상세페이지에서 구역명이 명시된 사례만 정리했습니다. 현재 매물 여부·가격 유지는 보장되지 않으며, "공시가격 차액"은 재개발 프리미엄이 아닙니다.', 'O');
header(naverSheet, 5, ['구역', '확인일', '매물번호', '유형', '호가\n(만원)', '보증금\n(만원)', '월세\n(만원)', '광고 표기\n초투입금(만원)', '대지면적\n(㎡)', '대지면적\n(평)', '호가 평당\n(만원/평)', '공시가격\n(만원)', '공시가격 차액\n(억원)', '상태', '원문 URL']);
const naverStartRow = 6;
const naverEndRow = naverStartRow + naverHistory.length - 1;
naverSheet.getRange(`A${naverStartRow}:O${naverEndRow}`).values = naverHistory.map((row) => [
  row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], null, null, row[9], null, '과거 확인 / 현행 여부 미확인', row[11],
]);
naverSheet.getRange(`J${naverStartRow}:K${naverEndRow}`).formulas = naverHistory.map((_, index) => [
  `=IFERROR(I${naverStartRow + index}/3.305785,"")`,
  `=IFERROR(E${naverStartRow + index}/J${naverStartRow + index},"")`,
]);
naverSheet.getRange(`M${naverStartRow}:M${naverEndRow}`).formulas = naverHistory.map((_, index) => [`=IFERROR((E${naverStartRow + index}-L${naverStartRow + index})/10000,"")`]);
decorateTable(naverSheet, 5, naverEndRow, 'O');
naverSheet.getRange(`B${naverStartRow}:B${naverEndRow}`).format.numberFormat = 'yyyy-mm-dd';
naverSheet.getRange(`E${naverStartRow}:I${naverEndRow}`).format.numberFormat = '#,##0';
naverSheet.getRange(`J${naverStartRow}:K${naverEndRow}`).format.numberFormat = '#,##0.0';
naverSheet.getRange(`L${naverStartRow}:L${naverEndRow}`).format.numberFormat = '#,##0';
naverSheet.getRange(`M${naverStartRow}:M${naverEndRow}`).format.numberFormat = '#,##0.00';
naverSheet.getRange(`O${naverStartRow}:O${naverEndRow}`).format.font = { color: '#008000' };
section(naverSheet, 13, '매물별 메모', 'O');
naverSheet.getRange('A14:C14').values = [['구역', '매물번호', '메모']];
naverSheet.getRange('A15:C19').values = naverHistory.map((row) => [row[0], row[2], row[10]]);
for (let row = 14; row <= 19; row += 1) naverSheet.getRange(`C${row}:O${row}`).merge();
naverSheet.getRange('A14:O19').format.borders = { preset: 'all', style: 'thin', color: '#D4DDE4' };
naverSheet.getRange('A14:O14').format = { fill: '#DCE7EE', font: { bold: true, color: '#183047' } };
naverSheet.getRange('A14:O19').format.wrapText = true;
naverSheet.getRange('A15:O19').format.rowHeight = 34;
setWidths(naverSheet, [14, 13, 14, 14, 13, 13, 12, 17, 13, 13, 15, 14, 15, 26, 38]);

// Sources and method
title(sourceSheet, '출처·방법·한계', 'F');
subtitle(sourceSheet, '실거래와 호가를 같은 개념으로 취급하지 않고, 출처별 기준일과 한계를 함께 기록합니다.', 'F');
header(sourceSheet, 5, ['구분', '내용', '기준일', 'URL', '사용 범위', '한계']);
const sourceRows = [
  ['실거래 원자료', '재개발닷컴 구역별 property-transactions 공개 응답, 각 최신 20건', snapshotDate, sources.shin1Project, '거래가·대지면적·공시가격·계약일 비교', '플랫폼 재가공 데이터이므로 계약 전 국토부 원문 확인 필요'],
  ['실거래 원자료', '재개발닷컴 구역별 property-transactions 공개 응답, 각 최신 20건', snapshotDate, sources.shin3Project, '거래가·대지면적·공시가격·계약일 비교', '플랫폼 재가공 데이터이므로 계약 전 국토부 원문 확인 필요'],
  ['법정 실거래 확인', '국토교통부 실거래가 공개시스템', snapshotDate, sources.molit, '계약일·취소·주소의 최종 대조', '원자료 조회 조건에 따라 확인 필요'],
  ['현 호가 보조', '재개발닷컴 공개 호가 페이지', snapshotDate, sources.shin1Asks, '공개 호가 개수와 가격 분포', '호가이며 대지면적·공시가격이 비공개인 행 다수'],
  ['현 호가 보조', '재개발닷컴 공개 호가 페이지', snapshotDate, sources.shin3Asks, '공개 호가 개수와 가격 분포', '호가이며 대지면적·공시가격이 비공개인 행 다수'],
  ['네이버 현행 화면', '성남 수정구 신흥동 / 재개발(F01) 지도 필터', snapshotDate, sources.naverMap, '현행 지도 범위 확인', '동적 지도에서 신흥1·3별 대지지분·가격을 일괄 추출할 수 없음'],
  ['네이버 확인이력', '구역명 명시 개별 매물 5건', new Date('2025-06-12'), '네이버_확인이력 시트 URL 참조', '과거 호가·대지면적 사례', '현재 매물·현재 가격이 아님'],
];
sourceSheet.getRange(`A6:F${5 + sourceRows.length}`).values = sourceRows;
decorateTable(sourceSheet, 5, 5 + sourceRows.length, 'F');
sourceSheet.getRange(`C6:C${5 + sourceRows.length}`).format.numberFormat = 'yyyy-mm-dd';
sourceSheet.getRange(`D6:D${5 + sourceRows.length}`).format.font = { color: '#008000' };
sourceSheet.getRange(`A6:F${5 + sourceRows.length}`).format.wrapText = true;
section(sourceSheet, 15, '산식과 해석 원칙', 'F');
sourceSheet.getRange('A16:C21').values = [
  ['지표', '산식', '판단'],
  ['대지평당 거래가', '거래가(만원) ÷ [대지면적(㎡) ÷ 3.305785]', '같은 유형·같은 권리구조에서만 비교'],
  ['공시가격 차액', '거래가 - 공시가격', '세금 기준가격과의 차이일 뿐, 재개발 P가 아님'],
  ['재개발 프리미엄', '매매가 - 종전자산 감정평가액 또는 권리가액', '개별 평가자료 없이는 공란으로 둠'],
  ['호가-실거래 차이', '동일 물건·동일 조건의 호가와 체결가 비교', '현재는 물건별 대지지분 미공개로 구역 평균 산출 제외'],
  ['유형 분리', '단독/다가구와 다세대/연립 분리', '다세대는 대지지분이 작아 평당 단가가 높게 보일 수 있음'],
];
for (let row = 16; row <= 21; row += 1) sourceSheet.getRange(`C${row}:F${row}`).merge();
sourceSheet.getRange('A16:F21').format.borders = { preset: 'all', style: 'thin', color: '#D4DDE4' };
sourceSheet.getRange('A16:F16').format = { fill: '#DCE7EE', font: { bold: true, color: '#183047' } };
sourceSheet.getRange('A16:F21').format.wrapText = true;
sourceSheet.getRange('A17:F21').format.rowHeight = 42;
setWidths(sourceSheet, [17, 36, 18, 32, 30, 30]);
sourceSheet.freezePanes.freezeRows(5);

// Formula and data checks, documented on the summary tab.
const actualTransactionStats = zones.map((zone) => {
  const rows = transactionRowsByZone.get(zone.key);
  const pyeongPrices = rows.map((row) => row.price && row.landArea ? row.price / (row.landArea / 3.305785) : null);
  return {
    zone: zone.key,
    count: rows.length,
    minDate: rows.at(-1)?.madeAt?.toISOString().slice(0, 10),
    maxDate: rows[0]?.madeAt?.toISOString().slice(0, 10),
    priceMedian: median(rows.map((row) => row.price)) / 10000,
    pyeongMedian: median(pyeongPrices),
    pyeongAverage: pyeongPrices.filter(Number.isFinite).reduce((sum, value) => sum + value, 0) / pyeongPrices.filter(Number.isFinite).length,
  };
});
const actualAskStats = zones.map((zone) => {
  const rows = askRowsByZone.get(zone.key);
  const prices = rows.map((row) => row.price).filter(Number.isFinite);
  return { zone: zone.key, count: rows.length, min: Math.min(...prices) / 10000, median: median(prices) / 10000, max: Math.max(...prices) / 10000 };
});

await fs.mkdir(renderDir, { recursive: true });
const inspectSummary = await workbook.inspect({ kind: 'table,formula', sheetId: '요약', range: 'A1:J23', maxChars: 6000, tableMaxRows: 12, tableMaxCols: 10 });
const formulaErrors = await workbook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'formula error scan' });
if (formulaErrors.recordCount > 1) throw new Error(`수식 오류 표식이 ${formulaErrors.recordCount - 1}건 발견됐습니다.`);
for (const sheetName of ['요약', '실거래_최신20', '현호가_보조', '현호가_원자료', '네이버_확인이력', '출처_방법']) {
  const image = await workbook.render({ sheetName, autoCrop: 'all', scale: 1, format: 'png' });
  await fs.writeFile(resolve(renderDir, `${sheetName}.png`), new Uint8Array(await image.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await fs.mkdir(outputDir, { recursive: true });
await xlsx.save(outputPath);
await fs.writeFile(resolve(outputDir, 'build_summary.json'), JSON.stringify({
  snapshotDate: '2026-09-02',
  transactions: actualTransactionStats,
  asks: actualAskStats,
  inspectSummary,
  formulaErrors,
}, null, 2));
console.log(JSON.stringify({ outputPath, actualTransactionStats, actualAskStats, formulaErrors }, null, 2));
