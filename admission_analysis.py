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
        '모집단위':      ['모집 단위(학과)', '모집단위', '학과', '모집'],
        '전형명칭':      ['전형명칭', '전형명'],
        '전형유형':      ['전형\n유형', '전형유형', '유형'],
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

    cols = ['대입연도', '대학교명', '국영수과 등평', '전교과 등평', '모집단위', '상태구분', 'is_medical', 'admission_type']
    records = df[cols].to_dict('records')
    years = sorted(df['대입연도'].unique().tolist())
    return records, years


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
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
</style>
</head>
<body>
<div class="page">
  <h2>대학교별 내신 등급 입시 결과 분석</h2>

  <div class="controls">
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
    <div class="ctrl-group" style="margin-left:auto; align-self:flex-end;">
      <div style="font-size:11px; color:#bbb; line-height:1.7; text-align:right;">
        x축: <b>대학명(지원학생수)</b><br>
        대학 클릭 → 연도별 상세 보기
      </div>
    </div>
  </div>

  <div class="chart-wrap">
    <div id="chart"></div>
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
  filter: '학생부종합',
  cat: '일반계열',
  grade: '국영수과 등평',
  myGrade: null,
};

// ── UI setup ───────────────────────────────────────────
function makeBtn(text, cls, active, onClick) {
  const b = document.createElement('button');
  b.textContent = text;
  b.className = cls + (active ? ' on' : '');
  b.addEventListener('click', onClick);
  return b;
}

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
    r['is_medical'] === isMedical &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null
  );

  const byUniv = {};
  for (const r of rows) (byUniv[r['대학교명']] ??= []).push(r);
  for (const u of Object.keys(byUniv)) {
    if (!byUniv[u].some(r => r['상태구분'] !== '1차탈락')) delete byUniv[u];
  }

  const univs = getOrder(byUniv, gradeKey).filter(u => byUniv[u]);
  if (!univs.length) {
    document.getElementById('chart').innerHTML = '<div class="empty-msg">해당 조건의 데이터가 없습니다.</div>';
    document.getElementById('grade-summary').innerHTML = '';
    return;
  }

  const nCount = {};
  for (const u of univs) nCount[u] = byUniv[u].filter(r => r['상태구분'] !== '1차탈락').length;

  const traces = [];

  // 1차탈락
  const rejX = [], rejY = [];
  for (const u of univs)
    for (const r of byUniv[u])
      if (r['상태구분'] === '1차탈락') { rejX.push(u); rejY.push(r[gradeKey]); }
  if (rejX.length) traces.push({
    type: 'box', x: rejX, y: rejY, name: '1차탈락 (참고)',
    marker: { color: COLORS['1차탈락'], opacity: 0.45, size: 4 },
    line: { color: COLORS['1차탈락'] }, fillcolor: 'rgba(204,204,204,0.15)',
    boxpoints: 'all', jitter: 0.4, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<extra>1차탈락</extra>',
  });

  for (const [status, color] of [['최종합격', COLORS['최종합격']], ['1차합격_최종탈락', COLORS['1차합격_최종탈락']]]) {
    const sx = [], sy = [];
    for (const u of univs)
      for (const r of byUniv[u])
        if (r['상태구분'] === status) { sx.push(u); sy.push(r[gradeKey]); }
    if (!sx.length) continue;
    traces.push({
      type: 'box', x: sx, y: sy, name: status,
      marker: { color, size: 5 }, line: { color },
      boxpoints: 'all', jitter: 0.3, pointpos: 0,
      hovertemplate: '%{x}<br>등평: %{y:.2f}<extra>' + status + '</extra>',
    });
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

  Plotly.react('chart', traces, {
    title: { text: `${yearLabel}년 ${state.cat}${typeLabel} · ${gradeLabel} 기준`, font: { size: 16 } },
    xaxis: {
      title: '대학교 (클릭 시 연도별 상세)',
      categoryorder: 'array', categoryarray: univs,
      tickangle: -38, tickvals: univs, ticktext: univs.map(u => `${u}(${nCount[u]})`),
    },
    yaxis: { title: gradeKey, range: [9.2, 0.8], fixedrange: false },
    boxmode: 'group',
    legend: { orientation: 'h', y: 1.08, x: 1, xanchor: 'right' },
    height: 620, margin: { b: 170, t: 60, r: 80 },
    plot_bgcolor: '#fafbff',
    shapes, annotations,
  }, { responsive: true });

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
  for (const [status, color] of [
    ['1차탈락', COLORS['1차탈락']],
    ['1차합격_최종탈락', COLORS['1차합격_최종탈락']],
    ['최종합격', COLORS['최종합격']],
  ]) {
    const sx = [], sy = [];
    for (const r of rows) {
      if (r['상태구분'] === status) { sx.push(r[xKey]); sy.push(r[state.grade]); }
    }
    if (!sx.length) continue;
    const isRej = status === '1차탈락';
    traces.push({
      type: 'box', x: sx, y: sy, name: status,
      marker: { color, opacity: isRej ? 0.4 : 1, size: 5 },
      line: { color },
      fillcolor: isRej ? 'rgba(204,204,204,0.1)' : undefined,
      boxpoints: 'all', jitter: 0.35, pointpos: 0,
      hovertemplate: '%{x}<br>등평: %{y:.2f}<extra>' + status + '</extra>',
    });
  }
  return traces;
}

// 연도별 뷰
function renderModalYear() {
  const gradeKey = state.grade;
  const allYears = [...YEARS].sort();
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
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
