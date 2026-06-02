const XLSX = require('./node_modules/xlsx');
const path = require('path');
const fs = require('fs');

// ================================================================
// [1] 제품코드 접두어 설정 (날짜 부분 - 공통)
// ================================================================
const PREFIX = 'LG260602';  // LG + YYMMDD

// ================================================================
// [2] 카테고리별 파일 및 제품코드 끝번호 시작값 설정
//     - startNum: 이 번호부터 시작해서 새 모델마다 자동 증가
//     - numWidth: 끝번호 자릿수 (예: 3 -> 001, 002, ...)
//     - null: 기존 제품코드 유지
// ================================================================
const CATEGORIES = {
  jeongsugi:   { path: 'D:/COWAY_OVERVIEW/LG전자/6월정책/안대잖아.xlsx',                          startNum: null, numWidth: 3 },
  TV:          { path: 'C:/Users/Designer/Downloads/TV_테스트_LG신전산.xlsx',                     startNum: null, numWidth: 3 },
  gongcheong:  { path: 'C:/Users/Designer/Downloads/공기청정기_테스트_LG신전산.xlsx',              startNum: null, numWidth: 3 },
  naengjango:  { path: 'C:/Users/Designer/Downloads/냉장고_테스트_LG신전산.xlsx',                 startNum: null, numWidth: 3 },
  aircon:      { path: 'C:/Users/Designer/Downloads/에어컨_테스트_LG신전산.xlsx',                 startNum: null, numWidth: 3 },
  cooking:     { path: 'C:/Users/Designer/Downloads/쿠킹_테스트_LG신전산.xlsx',                   startNum: null, numWidth: 3 },
  all:         { path: 'C:/Users/Designer/Downloads/LG전자_최종본_LG신전산.xlsx',                 startNum: null, numWidth: 3 },
};
// ================================================================

const bindMap = { 1: '일반', 27: '신규결합', 28: '기존결합' };

function visitLabel(v) {
  if (!v || v === '') return '';
  const s = String(v).trim();
  if (s === '관리없음' || s === '자가관리') return '(자가관리)';
  const m = s.match(/(\d+개월)/);
  if (m) return '(' + m[1] + '방문)';
  return '(' + s + ')';
}

// ── 모델 목록 출력 모드 (--list <카테고리키>) ──────────────────
function listModels(catKey) {
  const cat = CATEGORIES[catKey];
  if (!cat) { console.log('없는 카테고리: ' + catKey); return; }
  if (!fs.existsSync(cat.path)) { console.log('파일 없음: ' + cat.path); return; }

  const wb = XLSX.readFile(cat.path, { cellText: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const seen = [];
  const modelSet = new Set();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[1]) continue;
    const model = row[2] ? String(row[2]).trim() : '';
    if (model && !modelSet.has(model)) {
      modelSet.add(model);
      seen.push(model);
    }
  }

  const start = cat.startNum || 1;
  const w = cat.numWidth || 3;
  const mapObj = {};
  seen.forEach((m, idx) => {
    mapObj[m] = String(start + idx).padStart(w, '0');
  });

  const outFile = 'D:/COWAY_OVERVIEW/model_codes_' + catKey + '.json';
  fs.writeFileSync(outFile, JSON.stringify({ prefix: PREFIX, numWidth: w, models: mapObj }, null, 2), 'utf8');
  console.log('[' + catKey + '] 모델 수: ' + seen.length + '개');
  console.log('저장됨: ' + outFile);
  console.log('→ 파일 열어서 번호 수정 후 다시 node update_all.js 실행하세요\n');
  console.log('샘플:');
  Object.entries(mapObj).slice(0, 5).forEach(([k, v]) => console.log('  ' + PREFIX + v + '  ←  ' + k));
}

// ── 파일 업데이트 ─────────────────────────────────────────────
function updateFile(catKey) {
  const cat = CATEGORIES[catKey];
  if (!cat) return;
  const filePath = cat.path;
  if (!fs.existsSync(filePath)) { console.log('없음: ' + filePath); return; }

  // 모델 코드 매핑 파일 로드 (있으면 사용, 없으면 startNum으로 자동)
  const mapFile = 'D:/COWAY_OVERVIEW/model_codes_' + catKey + '.json';
  let modelCodeMap = null;
  let mapPrefix = PREFIX;
  let mapWidth = cat.numWidth || 3;

  if (fs.existsSync(mapFile)) {
    const json = JSON.parse(fs.readFileSync(mapFile, 'utf8'));
    modelCodeMap = json.models;
    mapPrefix = json.prefix || PREFIX;
    mapWidth = json.numWidth || mapWidth;
    console.log('[' + catKey + '] model_codes_' + catKey + '.json 사용');
  } else if (cat.startNum !== null) {
    // 자동 생성 모드
    console.log('[' + catKey + '] startNum=' + cat.startNum + ' 자동 생성');
  }

  const wb = XLSX.readFile(filePath, { cellText: false });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // 자동 생성 모드용 맵
  const autoMap = {};
  let autoSeq = cat.startNum || 1;

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[1]) continue;

    const model = row[2] ? String(row[2]).trim() : '';

    // 제품코드 할당
    if (modelCodeMap && model && modelCodeMap[model]) {
      // JSON 파일 기반
      row[0] = mapPrefix + modelCodeMap[model];
    } else if (!modelCodeMap && cat.startNum !== null && model) {
      // 자동 증가
      if (!autoMap[model]) {
        autoMap[model] = String(autoSeq).padStart(mapWidth, '0');
        autoSeq++;
      }
      row[0] = PREFIX + autoMap[model];
    }
    // null이면 기존 코드 유지

    // 품목명 업데이트
    const s = String(row[1]).trim();
    let base = s;
    if (s.match(/\d년의무/)) {
      const m = s.match(/^(.+?)\s+\d년의무/);
      if (m) base = m[1].trim();
    }

    const color = row[3] ? String(row[3]).trim() : '';
    const codeNum = Math.round(Number(row[6]));
    const term = row[7] ? String(row[7]).trim() : '';
    const visit = visitLabel(row[10]);
    const combine = bindMap[codeNum] || String(codeNum);

    if (color && base.endsWith('[' + color + ']')) {
      base = base.slice(0, base.lastIndexOf('[' + color + ']')).trim();
    }

    const colorPart = color ? ' [' + color + ']' : '';
    row[1] = base + colorPart + ' ' + term + ' ' + visit + ' (' + combine + ')';
    count++;
  }

  XLSX.writeFile(wb, filePath);
  const fname = path.basename(filePath);
  console.log('완료 ' + fname + ': ' + count + '행');
  for (let i = 1; i <= 3 && i < data.length; i++) {
    if (data[i] && data[i][0]) process.stdout.write('  [' + data[i][0] + '] ');
    if (data[i] && data[i][1]) console.log(data[i][1]);
  }
}

// ── 실행 ────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args[0] === '--list') {
  // node update_all.js --list jeongsugi
  const key = args[1];
  if (key && CATEGORIES[key]) {
    listModels(key);
  } else {
    console.log('사용법: node update_all.js --list <카테고리>');
    console.log('카테고리: ' + Object.keys(CATEGORIES).join(', '));
  }
} else {
  // 전체 업데이트
  for (const key of Object.keys(CATEGORIES)) {
    updateFile(key);
  }
  console.log('\n전체 완료');
}
