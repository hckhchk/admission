import os
import sys
import json
import pandas as pd

MEDICAL_KEYWORDS = ['의예', '의학', '의과', '치의', '치과', '한의', '약학', '수의예']
SPECIAL_KEYWORDS = [
    '기초생활', '기회균형', '기회균등', '고른기회', '사회통합', '차상위',
    '저소득', '이웃사랑', '국가보훈', '농어촌', '다문화', '특수교육',
    '사회적배려', '사회기여', '연세한마음', '교육기회균형', '사회공헌',
]


def classify_admission(전형명칭, 전형유형):
    n = str(전형명칭)
    t = str(전형유형)
    if any(kw in n for kw in SPECIAL_KEYWORDS):
        return '특별전형'
    if t == '논술' or '논술' in n:
        return '논술'
    if t == '종합' or '종합' in n:
        return '학생부종합'
    return '기타'


def prepare_data(file_path):
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Excel 파일 로드 오류: {e}")
        raise

    def find_col(df_cols, keywords, exclude=None):
        exclude = exclude or set()
        available = [c for c in df_cols if c not in exclude]
        for kw in keywords:
            for col in available:
                if str(col) == kw:
                    return col
        for kw in keywords:
            for col in available:
                if kw in str(col):
                    return col
        return None

    candidates = {
        '대입연도':      ['대입연도', '연도', '입시연도'],
        '1차결과':       ['1차결과', '1차', '1차 합격', '1차_결과'],
        '결과':          ['최종결과', '최종 합격', '최종결과', '결과'],
        '대학교명':      ['대학교명', '대학', '학교'],
        '국영수과 등평': ['국영수과 등평', '국영수과등평'],
        '전교과 등평':   ['전교과 등평', '전교과등평'],
        '학년':          ['학년'],
        '모집단위':      ['모집 단위(학과)', '모집단위', '학과', '모집'],
        '전형명칭':      ['전형명칭', '전형명'],
        '전형유형':      ['전형\n유형', '전형유형', '유형'],
        '최초합격':      ['최초합격'],
    }

    col_map = {}
    used_cols = set()
    for key, kwlist in candidates.items():
        found = find_col(df.columns, kwlist, exclude=used_cols)
        if found:
            col_map[key] = found
            used_cols.add(found)

    missing = [c for c in ['대입연도', '1차결과', '결과', '대학교명', '국영수과 등평'] if c not in col_map]
    if missing:
        print("필수 컬럼 매핑 실패:", missing)
        raise KeyError(f"Missing required columns: {missing}")

    df = df.rename(columns={v: k for k, v in col_map.items()})
    df['국영수과 등평'] = pd.to_numeric(df['국영수과 등평'], errors='coerce')
    if '전교과 등평' in df.columns:
        df['전교과 등평'] = pd.to_numeric(df['전교과 등평'], errors='coerce')
    else:
        df['전교과 등평'] = float('nan')
    df = df.dropna(subset=['국영수과 등평', '대학교명', '대입연도'])
    df['대입연도'] = df['대입연도'].astype(int)

    def categorize_status(row):
        first = str(row['1차결과']).strip()
        final = str(row['결과']).strip()
        if final == '합격':
            return '최종합격'
        if first == '합격':
            return '1차합격_최종탈락'
        return '1차탈락'

    df['상태구분'] = df.apply(categorize_status, axis=1)

    medical_pat = '|'.join(MEDICAL_KEYWORDS)
    df['is_medical'] = df['모집단위'].str.contains(medical_pat, na=False) if '모집단위' in df.columns else False

    df['admission_type'] = df.apply(
        lambda r: classify_admission(r.get('전형명칭', ''), r.get('전형유형', '')), axis=1
    )

    # 추가합격: 최종합격이지만 최초합격이 아닌 경우
    if '최초합격' in df.columns:
        df['is_additional'] = (df['상태구분'] == '최종합격') & \
                              (df['최초합격'].astype(str).str.strip() != '합격')
    else:
        df['is_additional'] = False

    cols = ['대입연도', '학년', '대학교명', '국영수과 등평', '전교과 등평', '모집단위', '상태구분', 'is_additional', 'is_medical', 'admission_type']
    records = df[cols].to_dict('records')
    years = sorted(df['대입연도'].unique().tolist())
    return records, years


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<meta name="theme-color" content="#1a2a4a">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="입시 분석">
<link rel="manifest" href="manifest.json">
<title>입시 결과 분석</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; margin: 0; background: #f0f2f5; color: #222; }
  .page { max-width: 1400px; margin: 0 auto; padding: 24px 20px; }
  h2 { margin: 0 0 20px; font-size: 20px; color: #1a2a4a; }

  .controls { background: white; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px;
              box-shadow: 0 1px 4px rgba(0,0,0,0.08); display: flex; flex-wrap: wrap; gap: 20px; align-items: flex-start; }
  .ctrl-group { display: flex; flex-direction: column; gap: 8px; }
  .ctrl-label { font-size: 12px; font-weight: 700; color: #666; text-transform: uppercase; letter-spacing: .5px; }
  .btn-row { display: flex; gap: 6px; flex-wrap: wrap; }

  .year-btn, .filter-btn, .cat-btn, .view-btn {
    padding: 6px 14px; border: 2px solid #d0d7e3; border-radius: 20px;
    background: white; cursor: pointer; font-size: 13px; font-weight: 600;
    color: #555; transition: all .15s;
  }
  .year-btn.on   { background: #0c4da2; border-color: #0c4da2; color: white; }
  .filter-btn.on { background: #2d7dd2; border-color: #2d7dd2; color: white; }
  .cat-btn.on    { background: #c0392b; border-color: #c0392b; color: white; }
  .view-btn.on   { background: #16a085; border-color: #16a085; color: white; }

  .grade-input {
    width: 100px; padding: 6px 12px; border: 2px solid #d0d7e3; border-radius: 20px;
    font-size: 13px; font-weight: 600; color: #333; outline: none; transition: border .15s;
  }
  .grade-input:focus { border-color: #e74c3c; }

  .chart-wrap { background: white; border-radius: 10px; padding: 16px;
                box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
  .chart-scroll-wrap { width: 100%; overflow: hidden; }
  #chart { width: 100%; }
  .empty-msg { text-align: center; padding: 60px; color: #aaa; font-size: 15px; }

  /* 드릴다운 모달 */
  #modal-overlay {
    display: none; position: fixed; inset: 0;
    background: rgba(0,0,0,0.45); z-index: 1000;
    align-items: center; justify-content: center;
  }
  #modal-overlay.open { display: flex; }
  #modal-box {
    background: white; border-radius: 14px; padding: 24px;
    width: min(820px, 95vw); box-shadow: 0 8px 32px rgba(0,0,0,0.2);
    position: relative;
  }
  #modal-title { font-size: 17px; font-weight: 700; color: #1a2a4a; margin-bottom: 4px; }
  #modal-sub   { font-size: 12px; color: #aaa; margin-bottom: 10px; }
  #modal-close {
    position: absolute; top: 16px; right: 18px;
    background: none; border: none; font-size: 20px; cursor: pointer; color: #888;
  }
  .modal-tabs { display: flex; gap: 6px; margin-bottom: 14px; }
  .modal-tab {
    padding: 5px 16px; border: 2px solid #d0d7e3; border-radius: 20px;
    background: white; cursor: pointer; font-size: 13px; font-weight: 600; color: #555;
  }
  .modal-tab.on { background: #16a085; border-color: #16a085; color: white; }
  #modal-chart { width: 100%; }

  #grade-summary { margin-top: 14px; display: flex; flex-wrap: wrap; gap: 12px; }
  .summary-group { background: white; border-radius: 10px; padding: 12px 16px;
                   box-shadow: 0 1px 4px rgba(0,0,0,0.08); flex: 1; min-width: 200px; }
  .summary-title { font-size: 12px; font-weight: 700; margin-bottom: 8px; }
  .summary-chips { display: flex; flex-wrap: wrap; gap: 5px; }
  .chip { padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
  .chip-stable { background: #d5f5e3; color: #1e8449; }
  .chip-proper { background: #d6eaf8; color: #1a5276; }
  .chip-reach  { background: #fdebd0; color: #784212; }
  .hint { font-size: 11px; color: #aaa; margin-top: 4px; }

  /* 토글 스위치 */
  .toggle-group { display: flex; flex-direction: column; gap: 6px; }
  .toggle-item { display: flex; align-items: center; gap: 8px; }
  .toggle-label { font-size: 13px; font-weight: 600; color: #555; cursor: pointer; user-select: none; }
  .toggle-switch { position: relative; display: inline-block; width: 38px; height: 20px; flex-shrink: 0; }
  .toggle-switch input { opacity: 0; width: 0; height: 0; }
  .toggle-slider {
    position: absolute; cursor: pointer; inset: 0;
    background: #ccc; border-radius: 20px; transition: .2s;
  }
  .toggle-slider:before {
    content: ''; position: absolute;
    width: 14px; height: 14px; left: 3px; bottom: 3px;
    background: white; border-radius: 50%; transition: .2s;
  }
  .toggle-switch input:checked + .toggle-slider { background: #e67e22; }
  .toggle-switch input:checked + .toggle-slider:before { transform: translateX(18px); }
  .toggle-switch.blue input:checked + .toggle-slider { background: #2d7dd2; }

  /* 모바일 반응형 */
  @media (max-width: 680px) {
    body { background: #eef0f4; }
    .page { padding: 0 0 24px; }
    h2 { font-size: 15px; margin: 0; padding: 14px 14px 10px; background: #1a2a4a; color: white; }

    /* 컨트롤 패널: 카드들을 세로로 쌓음 */
    .controls {
      flex-direction: column; gap: 0; padding: 0;
      border-radius: 0; box-shadow: none; background: transparent; margin-bottom: 8px;
    }
    /* 각 필터 그룹을 가로 스크롤 카드로 */
    .ctrl-group {
      background: white; margin: 4px 8px 0;
      border-radius: 10px; padding: 10px 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
      flex-direction: column; gap: 6px;
    }
    .ctrl-group[style*="margin-left:auto"] { display: none; }
    .ctrl-label { font-size: 11px; }
    /* 버튼 행은 가로 스크롤 */
    .btn-row {
      display: flex; flex-wrap: nowrap; overflow-x: auto;
      gap: 6px; padding-bottom: 2px;
      scrollbar-width: none; -ms-overflow-style: none;
    }
    .btn-row::-webkit-scrollbar { display: none; }
    .year-btn, .filter-btn, .cat-btn, .view-btn {
      padding: 7px 14px; font-size: 13px; white-space: nowrap; flex-shrink: 0;
      min-height: 36px;
    }
    .grade-input { width: 100%; max-width: 160px; min-height: 36px; font-size: 14px; }

    /* 차트 영역: 가로 스크롤 */
    .chart-wrap { border-radius: 0; margin: 0 4px; padding: 8px 4px; }
    .chart-scroll-wrap {
      overflow-x: auto; overflow-y: hidden;
      -webkit-overflow-scrolling: touch;
    }
    #chart { min-width: 800px; width: max-content; }

    /* 토글 그룹 */
    .toggle-group { flex-direction: row; flex-wrap: wrap; gap: 16px; }
    .toggle-item { gap: 8px; }
    .toggle-label { font-size: 13px; }

    /* 모달: 바텀 시트 */
    #modal-overlay { align-items: flex-end; }
    #modal-box {
      width: 100%; max-width: 100%;
      border-radius: 18px 18px 0 0; padding: 20px 16px 32px;
      max-height: 88vh; overflow-y: auto;
      animation: slideUp .25s cubic-bezier(0.32,0.72,0,1);
    }
    @keyframes slideUp {
      from { transform: translateY(100%); }
      to   { transform: translateY(0); }
    }
    /* 바텀시트 드래그 핸들 */
    #modal-box::before {
      content: ''; display: block; width: 36px; height: 4px;
      background: #d0d7e3; border-radius: 2px; margin: 0 auto 16px;
    }
    #modal-title { font-size: 16px; }
    #modal-sub { font-size: 12px; }
    #modal-chart { min-height: 280px; }
    .modal-tab { padding: 6px 14px; font-size: 13px; min-height: 34px; }

    /* 합격 요약 */
    #grade-summary { gap: 8px; padding: 0 8px; }
    .summary-group { min-width: 100%; padding: 12px 14px; }
    .chip { font-size: 12px; padding: 4px 10px; }
  }
</style>
</head>
<body>
<div class="page">
  <h2>대학교별 내신 등급 입시 결과 분석</h2>

  <div class="controls">
    <div class="ctrl-group">
      <div class="ctrl-label">학년</div>
      <div class="btn-row" id="grade-year-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">입시 연도</div>
      <div class="btn-row" id="year-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">전형 유형</div>
      <div class="btn-row" id="filter-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">계열</div>
      <div class="btn-row" id="cat-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">등급 기준</div>
      <div class="btn-row" id="grade-row"></div>
    </div>
    <div class="ctrl-group" id="mygrade-group">
      <div class="ctrl-label">내 등급</div>
      <input type="number" id="my-grade" class="grade-input" min="1" max="9" step="0.01" placeholder="예: 2.30">
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">표시 옵션</div>
      <div class="toggle-group">
        <div class="toggle-item">
          <label class="toggle-switch blue">
            <input type="checkbox" id="toggle-nopass">
            <span class="toggle-slider"></span>
          </label>
          <span class="toggle-label" onclick="document.getElementById('toggle-nopass').click()">합격 없는 대학 표시</span>
        </div>
        <div class="toggle-item">
          <label class="toggle-switch">
            <input type="checkbox" id="toggle-additional">
            <span class="toggle-slider"></span>
          </label>
          <span class="toggle-label" onclick="document.getElementById('toggle-additional').click()">추가합격 구분 표시</span>
        </div>
      </div>
    </div>
    <div class="ctrl-group" style="margin-left:auto; align-self:flex-end;">
      <div style="font-size:11px; color:#bbb; line-height:1.7; text-align:right;">
        x축: <b>대학명 (최종합격/총지원)</b><br>
        대학 클릭 → 연도별 상세 보기
      </div>
    </div>
  </div>

  <div class="chart-wrap">
    <div class="chart-scroll-wrap">
      <div id="chart"></div>
    </div>
  </div>
  <div id="grade-summary"></div>
</div>

<!-- 드릴다운 모달 -->
<div id="modal-overlay">
  <div id="modal-box">
    <button id="modal-close" title="닫기">✕</button>
    <div id="modal-title"></div>
    <div id="modal-sub"></div>
    <div class="modal-tabs">
      <button class="modal-tab on" onclick="switchModalTab('year', this)">연도별</button>
      <button class="modal-tab"    onclick="switchModalTab('dept', this)">학과별</button>
    </div>
    <div id="modal-chart"></div>
  </div>
</div>

<script>
const RAW = __DATA__;
const YEARS = __YEARS__;

const COLORS = {
  '최종합격':        '#0c4da2',
  '1차합격_최종탈락': '#5a9fd4',
  '1차탈락':         '#cccccc',
};

// ── state ──────────────────────────────────────────────
const state = {
  years: new Set([2026].filter(y => YEARS.includes(y)).concat(YEARS.includes(2026) ? [] : [YEARS[YEARS.length-1]])),
  gradeYears: new Set([2, 3]),
  filter: '학생부종합',
  cat: '일반계열',
  grade: '국영수과 등평',
  myGrade: null,
  showNoPass: false,
  showAdditional: false,
};

const ADD_COLOR    = '#e67e22'; // 추가합격 (주황)
const FIRST_COLOR  = '#0c4da2'; // 최초합격 (기존 최종합격 색)

// ── UI setup ───────────────────────────────────────────
function makeBtn(text, cls, active, onClick) {
  const b = document.createElement('button');
  b.textContent = text;
  b.className = cls + (active ? ' on' : '');
  b.addEventListener('click', onClick);
  return b;
}

// 학년 버튼 (multi-select)
const gradeYearRow = document.getElementById('grade-year-row');
[2, 3].forEach(g => {
  const b = makeBtn(g + '학년', 'year-btn', true, () => {
    if (state.gradeYears.has(g) && state.gradeYears.size === 1) return;
    state.gradeYears.has(g) ? state.gradeYears.delete(g) : state.gradeYears.add(g);
    b.classList.toggle('on', state.gradeYears.has(g));
    render();
  });
  gradeYearRow.appendChild(b);
});

// Year buttons (multi-select)
const yearRow = document.getElementById('year-row');
YEARS.forEach(y => {
  const b = makeBtn(y + '년', 'year-btn', state.years.has(y), () => {
    if (state.years.has(y) && state.years.size === 1) return;
    state.years.has(y) ? state.years.delete(y) : state.years.add(y);
    b.classList.toggle('on', state.years.has(y));
    render();
  });
  yearRow.appendChild(b);
});

// Filter buttons
const FILTERS = ['전체', '학생부종합', '논술', '특별전형'];
const filterRow = document.getElementById('filter-row');
FILTERS.forEach(f => {
  const b = makeBtn(f, 'filter-btn', f === '학생부종합', () => {
    state.filter = f;
    filterRow.querySelectorAll('.filter-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  filterRow.appendChild(b);
});

// Category buttons
const CATS = ['일반계열', '메디컬계열'];
const catRow = document.getElementById('cat-row');
CATS.forEach(c => {
  const b = makeBtn(c, 'cat-btn', c === state.cat, () => {
    state.cat = c;
    catRow.querySelectorAll('.cat-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  catRow.appendChild(b);
});

// Grade type buttons
const GRADES = ['국영수과 등평', '전교과 등평'];
const gradeRow = document.getElementById('grade-row');
GRADES.forEach(g => {
  const b = makeBtn(g, 'filter-btn', g === state.grade, () => {
    state.grade = g;
    gradeRow.querySelectorAll('.filter-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  gradeRow.appendChild(b);
});

// My grade input
document.getElementById('my-grade').addEventListener('input', e => {
  const v = parseFloat(e.target.value);
  state.myGrade = (!isNaN(v) && v >= 1 && v <= 9) ? v : null;
  render();
});

// 표시 옵션 토글
document.getElementById('toggle-nopass').addEventListener('change', e => {
  state.showNoPass = e.target.checked;
  render();
});
document.getElementById('toggle-additional').addEventListener('change', e => {
  state.showAdditional = e.target.checked;
  render();
});

// ── helpers ────────────────────────────────────────────
function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m-1] + s[m]) / 2;
}

function quantile(arr, q) {
  const s = [...arr].sort((a, b) => a - b);
  const pos = (s.length - 1) * q;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return s[lo] + (s[hi] - s[lo]) * (pos - lo);
}

function getOrder(byUniv, gradeKey) {
  const sortKey = {};
  for (const [u, rows] of Object.entries(byUniv)) {
    const finals = rows.filter(r => r.상태구분 === '최종합격').map(r => r[gradeKey]).filter(v => v != null);
    const firsts = rows.filter(r => r.상태구분 === '1차합격_최종탈락').map(r => r[gradeKey]).filter(v => v != null);
    if (finals.length) sortKey[u] = median(finals);
    else if (firsts.length) sortKey[u] = median(firsts);
  }
  return Object.entries(sortKey).sort((a, b) => a[1] - b[1]).map(x => x[0]);
}

// ── 분포도 ─────────────────────────────────────────────
function renderBox() {
  const gradeKey = state.grade;
  const isMedical = state.cat === '메디컬계열';

  const rows = RAW.filter(r =>
    state.years.has(r['대입연도']) &&
    state.gradeYears.has(r['학년']) &&
    r['is_medical'] === isMedical &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null
  );

  const byUniv = {};
  for (const r of rows) (byUniv[r['대학교명']] ??= []).push(r);
  for (const u of Object.keys(byUniv)) {
    if (!state.showNoPass && !byUniv[u].some(r => r['상태구분'] !== '1차탈락')) delete byUniv[u];
  }

  const orderedUnivs = getOrder(byUniv, gradeKey).filter(u => byUniv[u]);
  // 합격 없는 대학(정렬키 없음)은 뒤에 추가
  const noPassUnivs = Object.keys(byUniv).filter(u => !orderedUnivs.includes(u));
  const univs = [...orderedUnivs, ...noPassUnivs];
  if (!univs.length) {
    document.getElementById('chart').innerHTML = '<div class="empty-msg">해당 조건의 데이터가 없습니다.</div>';
    document.getElementById('grade-summary').innerHTML = '';
    return;
  }

  const nFinal = {}, nTotal = {};
  for (const u of univs) {
    nFinal[u] = byUniv[u].filter(r => r['상태구분'] === '최종합격').length;
    nTotal[u] = byUniv[u].length;
  }

  const traces = [];

  // 1차탈락
  const rejX = [], rejY = [], rejD = [];
  for (const u of univs)
    for (const r of byUniv[u])
      if (r['상태구분'] === '1차탈락') { rejX.push(u); rejY.push(r[gradeKey]); rejD.push(r['모집단위'] || ''); }
  if (rejX.length) traces.push({
    type: 'box', x: rejX, y: rejY, name: '1차탈락 (참고)',
    customdata: rejD,
    marker: { color: COLORS['1차탈락'], opacity: 0.45, size: 4 },
    line: { color: COLORS['1차탈락'] }, fillcolor: 'rgba(204,204,204,0.15)',
    boxpoints: 'all', jitter: 0.4, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<extra>1차탈락</extra>',
  });

  // 1차합격_최종탈락
  {
    const sx = [], sy = [], sd = [];
    for (const u of univs)
      for (const r of byUniv[u])
        if (r['상태구분'] === '1차합격_최종탈락') { sx.push(u); sy.push(r[gradeKey]); sd.push(r['모집단위'] || ''); }
    if (sx.length) traces.push({
      type: 'box', x: sx, y: sy, name: '1차합격_최종탈락',
      customdata: sd,
      marker: { color: COLORS['1차합격_최종탈락'], size: 5 }, line: { color: COLORS['1차합격_최종탈락'] },
      boxpoints: 'all', jitter: 0.3, pointpos: 0,
      hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<extra>1차합격_최종탈락</extra>',
    });
  }

  // 최종합격: 항상 전체 통계 box 유지, 추가합격 ON 시 ghost box로 점 색 구분
  {
    const sx = [], sy = [], sd = [];
    const f1x = [], f1y = [], f1d = [];
    const f2x = [], f2y = [], f2d = [];
    for (const u of univs)
      for (const r of byUniv[u])
        if (r['상태구분'] === '최종합격') {
          const isAdd = !!r['is_additional'];
          const dept = r['모집단위'] || '';
          sx.push(u); sy.push(r[gradeKey]); sd.push([dept, isAdd ? '추가합격' : '최초합격']);
          if (isAdd) { f2x.push(u); f2y.push(r[gradeKey]); f2d.push(dept); }
          else       { f1x.push(u); f1y.push(r[gradeKey]); f1d.push(dept); }
        }
    if (sx.length) {
      if (!state.showAdditional) {
        // 기본: 단일 trace, 점 포함. customdata 2D → 호버에서 학과+구분 모두 표시
        traces.push({
          type: 'box', x: sx, y: sy, name: '최종합격',
          customdata: sd,
          marker: { color: COLORS['최종합격'], size: 5 }, line: { color: COLORS['최종합격'] },
          boxpoints: 'all', jitter: 0.3, pointpos: 0,
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}<br>구분: %{customdata[1]}<extra>최종합격</extra>',
        });
      } else {
        // 추가합격 구분: 통계 box (점 숨김) + ghost box 2개로 점만 표시
        // ghost box에 width:0 → 그룹 폭 계산에 영향 없음
        traces.push({
          type: 'box', x: sx, y: sy, name: '최종합격',
          line: { color: COLORS['최종합격'] },
          fillcolor: 'rgba(12,77,162,0.12)',
          boxpoints: false,
          offsetgroup: 'final',
        });
        if (f1x.length) traces.push({
          type: 'box', x: f1x, y: f1y, name: '└ 최초합격',
          customdata: f1d,
          boxpoints: 'all', jitter: 0.35, pointpos: 0,
          fillcolor: 'rgba(0,0,0,0)', whiskerwidth: 0, width: 0,
          line: { color: 'rgba(0,0,0,0)', width: 0 },
          marker: { color: FIRST_COLOR, size: 6, opacity: 0.85 },
          offsetgroup: 'final',
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<br>구분: 최초합격<extra></extra>',
        });
        if (f2x.length) traces.push({
          type: 'box', x: f2x, y: f2y, name: '└ 추가합격',
          customdata: f2d,
          boxpoints: 'all', jitter: 0.35, pointpos: 0,
          fillcolor: 'rgba(0,0,0,0)', whiskerwidth: 0, width: 0,
          line: { color: 'rgba(0,0,0,0)', width: 0 },
          marker: { color: ADD_COLOR, size: 6, opacity: 0.9 },
          offsetgroup: 'final',
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<br>구분: 추가합격<extra></extra>',
        });
      }
    }
  }

  // 내 등급 수평선
  const shapes = [], annotations = [];
  if (state.myGrade) {
    shapes.push({
      type: 'line', x0: 0, x1: 1, xref: 'paper',
      y0: state.myGrade, y1: state.myGrade, yref: 'y',
      line: { color: '#e74c3c', width: 2, dash: 'dash' },
    });
    annotations.push({
      x: 1.01, xref: 'paper', xanchor: 'left',
      y: state.myGrade, yref: 'y',
      text: `내 등급 ${state.myGrade.toFixed(2)}`,
      showarrow: false, font: { color: '#e74c3c', size: 12, weight: 700 },
      bgcolor: 'rgba(255,255,255,0.85)',
    });
  }

  const yearLabel = [...state.years].sort().join(', ');
  const typeLabel = state.filter === '전체' ? '' : ` · ${state.filter}`;
  const gradeLabel = gradeKey === '국영수과 등평' ? '국영수과' : '전교과';

  // 모바일: 가로 스크롤 + Plotly 터치 인터랙션 비활성화
  const isMobile = window.innerWidth <= 680;
  const chartW = isMobile ? Math.max(860, univs.length * 52 + 220) : undefined;
  const chartH = isMobile ? 460 : 620;

  Plotly.react('chart', traces, {
    title: { text: `${yearLabel}년 ${state.cat}${typeLabel} · ${gradeLabel} 기준`, font: { size: isMobile ? 13 : 16 } },
    xaxis: {
      title: { text: '대학명 (최종합격수 / 총지원수)', font: { size: 11 } },
      categoryorder: 'array', categoryarray: univs,
      tickangle: -38, tickvals: univs, ticktext: univs.map(u => `${u}(${nFinal[u]}/${nTotal[u]})`),
      fixedrange: isMobile,
    },
    yaxis: { title: gradeKey, range: [9.2, 0.8], fixedrange: isMobile },
    boxmode: 'group',
    legend: { orientation: 'h', y: 1.08, x: 1, xanchor: 'right' },
    height: chartH,
    width: chartW,
    margin: { b: 170, t: 60, r: isMobile ? 10 : 80, l: isMobile ? 40 : 60 },
    plot_bgcolor: '#fafbff',
    shapes, annotations,
  }, { responsive: !isMobile, scrollZoom: false, displayModeBar: !isMobile });

  // 대학 클릭 → 연도별 드릴다운
  document.getElementById('chart').on('plotly_click', data => {
    const univ = data.points[0].x;
    openDrilldown(univ);
  });

  // 합격권 요약
  updateSummary(byUniv, univs, gradeKey);
}

function updateSummary(byUniv, univs, gradeKey) {
  const el = document.getElementById('grade-summary');
  if (!state.myGrade) { el.innerHTML = ''; return; }
  const g = state.myGrade;
  const stable = [], proper = [], reach = [];

  for (const u of univs) {
    const grades = byUniv[u].filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]).filter(v => v != null);
    if (!grades.length) continue;
    const q1 = quantile(grades, 0.25);
    const med = median(grades);
    const q3 = quantile(grades, 0.75);
    const n = grades.length;
    const entry = { u, n };
    if (g < q1)        stable.push(entry);
    else if (g <= med) proper.push(entry);
    else if (g <= q3)  reach.push(entry);
  }

  const LOW_N = 5;

  function chip(entry, cls) {
    const lowN = entry.n < LOW_N;
    const style = lowN ? 'opacity:0.5; border: 1px dashed #aaa;' : '';
    const nLabel = lowN ? ` (n=${entry.n}⚠)` : `(n=${entry.n})`;
    return `<span class="chip ${cls}" style="${style}" title="${lowN ? '합격 사례 수가 적어 참고용으로만 활용하세요' : ''}">${entry.u} ${nLabel}</span>`;
  }

  function group(title, color, cls, items, hint) {
    if (!items.length) return '';
    return `<div class="summary-group">
      <div class="summary-title" style="color:${color}">${title} (${items.length}개)</div>
      <div class="summary-chips">${items.map(e => chip(e, cls)).join('')}</div>
      <div class="hint">${hint}</div>
    </div>`;
  }

  const disclaimer = `<div style="font-size:11px;color:#aaa;margin-top:10px;line-height:1.6;">
    ※ 위 분류는 과거 합격자 등급 분포 내 위치를 나타낸 참고 자료입니다. 실제 합격 여부와 다를 수 있으며,
    전형별 특성·수능최저·면접 등 다양한 요소가 결과에 영향을 줍니다.
    ⚠ 표시는 합격 사례 ${LOW_N}건 미만으로 통계적 신뢰도가 낮습니다.
  </div>`;

  el.innerHTML =
    group('📊 합격자 상위 25% 이내', '#1e8449', 'chip-stable', stable, `입력한 등급(${g.toFixed(2)})이 최종합격자 하위 25% 등급보다 우수한 구간`) +
    group('📊 합격자 중앙값 이내',   '#1a5276', 'chip-proper', proper, `입력한 등급(${g.toFixed(2)})이 최종합격자 중앙값 이내 구간`) +
    group('📊 합격자 중앙~상위 25% 구간', '#784212', 'chip-reach', reach, `입력한 등급(${g.toFixed(2)})이 최종합격자 중앙값~상위 25% 등급 사이 구간`) +
    disclaimer;
}

// ── 드릴다운 모달 ───────────────────────────────────────
let drillUniv = null;
let drillTab = 'year';

function openDrilldown(univ) {
  drillUniv = univ;
  drillTab = 'year';
  // 탭 버튼 초기화
  document.querySelectorAll('.modal-tab').forEach((b, i) => b.classList.toggle('on', i === 0));
  const gradeLabel = state.grade === '국영수과 등평' ? '국영수과' : '전교과';
  const typeLabel  = state.filter === '전체' ? '전체 전형' : state.filter;
  document.getElementById('modal-title').textContent = univ;
  document.getElementById('modal-sub').textContent   = `${typeLabel} · ${gradeLabel} 기준`;
  document.getElementById('modal-overlay').classList.add('open');
  renderModalYear();
}

function switchModalTab(tab, btn) {
  drillTab = tab;
  document.querySelectorAll('.modal-tab').forEach(b => b.classList.remove('on'));
  btn.classList.add('on');
  tab === 'year' ? renderModalYear() : renderModalDept();
}

function modalShapes() {
  const shapes = [], annotations = [];
  if (state.myGrade) {
    shapes.push({ type: 'line', x0: 0, x1: 1, xref: 'paper',
      y0: state.myGrade, y1: state.myGrade, yref: 'y',
      line: { color: '#e74c3c', width: 2, dash: 'dash' } });
    annotations.push({ x: 1.01, xref: 'paper', xanchor: 'left',
      y: state.myGrade, yref: 'y',
      text: `내 등급 ${state.myGrade.toFixed(2)}`,
      showarrow: false, font: { color: '#e74c3c', size: 11, weight: 700 },
      bgcolor: 'rgba(255,255,255,0.85)' });
  }
  return { shapes, annotations };
}

function buildBoxTraces(rows, xKey, order) {
  const traces = [];

  // 1차탈락
  const r1x = [], r1y = [], r1d = [];
  for (const r of rows)
    if (r['상태구분'] === '1차탈락') { r1x.push(r[xKey]); r1y.push(r[state.grade]); r1d.push(r['모집단위'] || ''); }
  if (r1x.length) traces.push({
    type: 'box', x: r1x, y: r1y, name: '1차탈락 (참고)',
    customdata: r1d,
    marker: { color: COLORS['1차탈락'], opacity: 0.4, size: 5 }, line: { color: COLORS['1차탈락'] },
    fillcolor: 'rgba(204,204,204,0.1)',
    boxpoints: 'all', jitter: 0.35, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<extra>1차탈락</extra>',
  });

  // 1차합격_최종탈락
  const r2x = [], r2y = [], r2d = [];
  for (const r of rows)
    if (r['상태구분'] === '1차합격_최종탈락') { r2x.push(r[xKey]); r2y.push(r[state.grade]); r2d.push(r['모집단위'] || ''); }
  if (r2x.length) traces.push({
    type: 'box', x: r2x, y: r2y, name: '1차합격_최종탈락',
    customdata: r2d,
    marker: { color: COLORS['1차합격_최종탈락'], size: 5 }, line: { color: COLORS['1차합격_최종탈락'] },
    boxpoints: 'all', jitter: 0.35, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<extra>1차합격_최종탈락</extra>',
  });

  // 최종합격: 항상 전체 통계 box 유지, 추가합격 ON 시 ghost box로 점 색 구분
  {
    const sx = [], sy = [], sd = [];
    const f1x = [], f1y = [], f1d = [];
    const f2x = [], f2y = [], f2d = [];
    for (const r of rows)
      if (r['상태구분'] === '최종합격') {
        const isAdd = !!r['is_additional'];
        const dept = r['모집단위'] || '';
        sx.push(r[xKey]); sy.push(r[state.grade]); sd.push([dept, isAdd ? '추가합격' : '최초합격']);
        if (isAdd) { f2x.push(r[xKey]); f2y.push(r[state.grade]); f2d.push(dept); }
        else       { f1x.push(r[xKey]); f1y.push(r[state.grade]); f1d.push(dept); }
      }
    if (sx.length) {
      if (!state.showAdditional) {
        traces.push({
          type: 'box', x: sx, y: sy, name: '최종합격',
          customdata: sd,
          marker: { color: COLORS['최종합격'], size: 5 }, line: { color: COLORS['최종합격'] },
          boxpoints: 'all', jitter: 0.35, pointpos: 0,
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}<br>구분: %{customdata[1]}<extra>최종합격</extra>',
        });
      } else {
        traces.push({
          type: 'box', x: sx, y: sy, name: '최종합격',
          line: { color: COLORS['최종합격'] },
          fillcolor: 'rgba(12,77,162,0.12)',
          boxpoints: false, offsetgroup: 'final',
        });
        if (f1x.length) traces.push({
          type: 'box', x: f1x, y: f1y, name: '└ 최초합격',
          customdata: f1d,
          boxpoints: 'all', jitter: 0.35, pointpos: 0,
          fillcolor: 'rgba(0,0,0,0)', whiskerwidth: 0, width: 0,
          line: { color: 'rgba(0,0,0,0)', width: 0 },
          marker: { color: FIRST_COLOR, size: 6, opacity: 0.85 },
          offsetgroup: 'final',
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<br>구분: 최초합격<extra></extra>',
        });
        if (f2x.length) traces.push({
          type: 'box', x: f2x, y: f2y, name: '└ 추가합격',
          customdata: f2d,
          boxpoints: 'all', jitter: 0.35, pointpos: 0,
          fillcolor: 'rgba(0,0,0,0)', whiskerwidth: 0, width: 0,
          line: { color: 'rgba(0,0,0,0)', width: 0 },
          marker: { color: ADD_COLOR, size: 6, opacity: 0.9 },
          offsetgroup: 'final',
          hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata}<br>구분: 추가합격<extra></extra>',
        });
      }
    }
  }
  return traces;
}

// 연도별 뷰
function renderModalYear() {
  const gradeKey = state.grade;
  const allYears = [...YEARS].sort();
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
    state.gradeYears.has(r['학년']) &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null
  ).map(r => ({ ...r, _yearLabel: String(r['대입연도']) + '년' }));

  if (!rows.length) return;
  const traces = buildBoxTraces(rows, '_yearLabel', allYears.map(y => y + '년'));
  const { shapes, annotations } = modalShapes();

  Plotly.react('modal-chart', traces, {
    xaxis: { title: '입시 연도', categoryorder: 'array',
             categoryarray: allYears.map(y => y + '년') },
    yaxis: { title: gradeKey, range: [9.2, 0.8] },
    boxmode: 'group',
    legend: { orientation: 'h', y: 1.1, x: 1, xanchor: 'right' },
    height: 420, margin: { b: 60, t: 16, r: 90 },
    plot_bgcolor: '#fafbff', shapes, annotations,
  }, { responsive: true });
}

// 학과별 뷰
function renderModalDept() {
  const gradeKey = state.grade;
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
    state.years.has(r['대입연도']) &&
    state.gradeYears.has(r['학년']) &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null &&
    r['모집단위']
  );

  if (!rows.length) {
    Plotly.react('modal-chart', [], { height: 200 }, { responsive: true });
    return;
  }

  // 학과 정렬: 최종합격 중앙값 기준
  const byDept = {};
  for (const r of rows) (byDept[r['모집단위']] ??= []).push(r);
  const deptOrder = Object.entries(byDept)
    .map(([d, rs]) => {
      const g = rs.filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]);
      const g2 = rs.filter(r => r['상태구분'] === '1차합격_최종탈락').map(r => r[gradeKey]);
      return { d, med: g.length ? median(g) : (g2.length ? median(g2) : 99) };
    })
    .sort((a, b) => a.med - b.med)
    .map(x => x.d);

  const traces = buildBoxTraces(rows, '모집단위', deptOrder);
  const { shapes, annotations } = modalShapes();
  const yearLabel = [...state.years].sort().join(', ');

  Plotly.react('modal-chart', traces, {
    title: { text: `${yearLabel}년 선택`, font: { size: 12, color: '#aaa' } },
    xaxis: { title: '모집단위(학과)', categoryorder: 'array',
             categoryarray: deptOrder, tickangle: -35 },
    yaxis: { title: gradeKey, range: [9.2, 0.8] },
    boxmode: 'group',
    legend: { orientation: 'h', y: 1.1, x: 1, xanchor: 'right' },
    height: 440, margin: { b: 130, t: 30, r: 90 },
    plot_bgcolor: '#fafbff', shapes, annotations,
  }, { responsive: true });
}

// 모달 닫기
document.getElementById('modal-close').addEventListener('click', () => {
  document.getElementById('modal-overlay').classList.remove('open');
});
document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay'))
    document.getElementById('modal-overlay').classList.remove('open');
});

// ── dispatcher ─────────────────────────────────────────
function render() { renderBox(); }

render();
</script>
</body>
</html>"""


def prepare_and_visualize(file_path):
    records, years = prepare_data(file_path)
    html = HTML_TEMPLATE \
        .replace('__DATA__', json.dumps(records, ensure_ascii=False)) \
        .replace('__YEARS__', json.dumps(years))

    output_path = os.path.splitext(file_path)[0] + '_admission_analysis.html'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'저장: {output_path}')

    import webbrowser
    webbrowser.open(output_path)
    return output_path


if __name__ == '__main__':
    excel_name = 'University_Admission_Results_2022-2026.xlsx'
    default_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), excel_name)
    excel_path = sys.argv[1] if len(sys.argv) > 1 else default_path

    if not os.path.exists(excel_path):
        print(f"파일 없음: {excel_path}")
        sys.exit(1)

    prepare_and_visualize(excel_path)
