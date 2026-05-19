<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>팀장 REINVENT과제 실행/점검</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    --lg-red: #A50034;
    --lg-red-dark: #7A0026;
    --sidebar-bg: #2C3540;
    --sidebar-bg-2: #232A33;
    --sidebar-border: #3A4452;
    --sidebar-text: #E5E7EB;
    --sidebar-text-mute: #9CA3AF;
    --sidebar-label: #7A8597;

    --status-green:  #16A34A;
    --status-green-bg:  #E8F5EC;
    --status-orange: #EA8A1F;
    --status-orange-bg: #FDF1E1;
    --status-red:    #DC2626;
    --status-red-bg: #FCE7E7;

    --bg: #F4F5F7;
    --card: #FFFFFF;
    --text: #1F2937;
    --text-soft: #4B5563;
    --text-mute: #9CA3AF;
    --border: #E5E7EB;
    --border-soft: #F0F1F4;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  html, body {
    font-family: 'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, "Malgun Gothic", "맑은 고딕", sans-serif;
    color: var(--text);
    background: var(--bg);
    font-feature-settings: 'tnum' 1;
    -webkit-font-smoothing: antialiased;
    font-size: 13.5px;
    height: 100%;
  }

  .app {
    display: grid;
    grid-template-columns: 240px 1fr;
    min-height: 100vh;
  }

  /* =================== SIDEBAR =================== */
  aside.sidebar {
    background: var(--sidebar-bg);
    color: var(--sidebar-text);
    display: flex;
    flex-direction: column;
    border-right: 1px solid var(--sidebar-border);
  }
  .brand {
    display: flex; align-items: center; gap: 10px;
    padding: 18px 18px 16px;
    border-bottom: 1px solid var(--sidebar-border);
  }
  .brand .logo {
    width: 30px; height: 30px;
    background: var(--lg-red);
    color: #fff; font-weight: 800;
    border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-size: 14px; letter-spacing: -0.5px;
  }
  .brand .title {
    font-size: 13.5px; font-weight: 700; color: #fff;
    letter-spacing: -0.2px;
    line-height: 1.25;
  }

  .section {
    padding: 16px 18px 6px;
    border-bottom: 1px solid var(--sidebar-border);
  }
  .section:last-of-type { border-bottom: none; flex: 1; }

  .section-head {
    font-size: 10.5px; font-weight: 700;
    color: var(--sidebar-label);
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin-bottom: 12px;
    display: flex; align-items: center; gap: 6px;
  }

  .upload-btn {
    display: flex; align-items: center; gap: 8px;
    width: 100%;
    background: var(--sidebar-bg-2);
    border: 1px solid var(--sidebar-border);
    color: var(--sidebar-text);
    padding: 10px 12px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 12.5px;
    font-family: inherit;
    transition: background 0.15s;
    margin-bottom: 10px;
  }
  .upload-btn:hover { background: #353F4D; }
  .upload-btn .ico { color: #FBBF24; }
  .upload-btn input[type=file] { display: none; }

  .upload-status {
    font-size: 11px;
    color: var(--sidebar-text-mute);
    margin-bottom: 6px;
    line-height: 1.4;
  }
  .upload-status.ok  { color: #4ADE80; }
  .upload-status.err { color: #FCA5A5; }

  .field { margin-bottom: 12px; }
  .field label {
    display: block;
    font-size: 11.5px;
    color: var(--sidebar-text-mute);
    margin-bottom: 5px;
  }
  .field select {
    width: 100%;
    background: var(--sidebar-bg-2);
    border: 1px solid var(--sidebar-border);
    color: var(--sidebar-text);
    padding: 7px 10px;
    border-radius: 5px;
    font-size: 12.5px;
    font-family: inherit;
    cursor: pointer;
    appearance: none;
    background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath fill='%239CA3AF' d='M5 6L0 0h10z'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 10px center;
    padding-right: 26px;
  }

  .sidebar-bottom { padding: 14px 18px 18px; }
  .apply-btn {
    width: 100%;
    background: var(--lg-red);
    color: #fff;
    border: none;
    padding: 11px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    font-family: inherit;
    letter-spacing: 0.3px;
    transition: background 0.15s;
  }
  .apply-btn:hover { background: var(--lg-red-dark); }

  /* =================== MAIN =================== */
  main.main { padding: 24px 28px 32px; overflow-x: auto; }

  .page-head {
    display: flex; align-items: baseline; gap: 12px;
    padding-bottom: 14px;
    border-bottom: 2px solid var(--lg-red);
    margin-bottom: 22px;
    flex-wrap: wrap;
  }
  .page-head h1 {
    font-size: 22px; font-weight: 800;
    color: var(--text); letter-spacing: -0.4px;
  }
  .page-head .meta {
    margin-left: auto;
    font-size: 12px; color: var(--text-mute);
  }
  .page-head .meta strong { color: var(--text-soft); font-weight: 600; margin-left: 4px; }
  .page-head .meta span { margin-left: 12px; }

  .row {
    display: grid;
    gap: 16px;
    margin-bottom: 16px;
  }
  .row.top { grid-template-columns: 340px 1fr; }

  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 20px 22px;
    box-shadow: var(--shadow-sm);
  }
  .card-head {
    display: flex; align-items: center; gap: 8px;
    margin-bottom: 16px;
  }
  .card-head .num {
    width: 20px; height: 20px;
    background: var(--text);
    color: #fff;
    border-radius: 50%;
    font-size: 11px; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
  }
  .card-head h2 {
    font-size: 14.5px; font-weight: 700; color: var(--text);
    letter-spacing: -0.2px;
  }
  .card-head .legend {
    margin-left: auto;
    display: flex; gap: 12px; align-items: center;
    font-size: 11px; color: var(--text-soft);
  }
  .card-head .legend .item { display: inline-flex; align-items: center; gap: 5px; }
  .card-head .legend .chip {
    width: 10px; height: 10px; border-radius: 2px; display: inline-block;
  }
  .card-head .legend .chip.g { background: var(--status-green); }
  .card-head .legend .chip.o { background: var(--status-orange); }
  .card-head .legend .chip.r { background: var(--status-red); }

  /* === Card 1 : 전체 진행률 === */
  .sub-title {
    font-size: 12.5px; font-weight: 600;
    color: var(--text-soft);
    margin-bottom: 10px;
  }
  .rate-box {
    background: #FAFAFB;
    border: 1px solid var(--border-soft);
    border-radius: 6px;
    padding: 16px 18px;
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 18px;
  }
  .rate-box .pct {
    font-size: 32px; font-weight: 800;
    color: var(--lg-red);
    letter-spacing: -0.8px;
    line-height: 1;
  }
  .rate-box .meta {
    font-size: 11.5px; color: var(--text-mute);
    text-align: right;
    line-height: 1.6;
  }
  .rate-box .meta strong { color: var(--text-soft); font-weight: 600; }

  .summary-list { display: flex; flex-direction: column; gap: 9px; }
  .summary-row {
    display: flex; align-items: center; justify-content: space-between;
  }
  .pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 5px 12px;
    border-radius: 14px;
    font-size: 11.5px; font-weight: 600;
    color: #fff;
    min-width: 72px;
    justify-content: center;
  }
  .pill.g { background: var(--status-green); }
  .pill.o { background: var(--status-orange); }
  .pill.r { background: var(--status-red); }
  .summary-row .count {
    font-size: 12px; color: var(--text-soft);
  }
  .summary-row .count strong { color: var(--text); font-weight: 700; font-size: 13px; }

  /* === Card 2 : 5 task progress === */
  .task-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 12px;
  }
  .task-cell {
    border: 1px solid var(--border);
    border-radius: 7px;
    padding: 13px 14px;
    background: #FBFBFC;
  }
  .task-cell .t-no {
    font-size: 10.5px; font-weight: 700;
    color: var(--text-mute);
    margin-bottom: 4px;
    letter-spacing: 0.3px;
  }
  .task-cell .t-name {
    font-size: 12.5px; font-weight: 700;
    color: var(--text);
    line-height: 1.35;
    min-height: 34px;
    margin-bottom: 12px;
  }
  .task-cell .t-pct {
    font-size: 22px; font-weight: 800;
    color: var(--text);
    line-height: 1;
    margin-bottom: 8px;
    letter-spacing: -0.5px;
  }
  .task-cell .t-pct .unit { font-size: 12px; color: var(--text-mute); font-weight: 500; margin-left: 2px; }
  .task-cell .t-bar {
    height: 6px;
    background: #E9EAEC;
    border-radius: 3px;
    overflow: hidden;
  }
  .task-cell .t-bar .fill {
    height: 100%;
    border-radius: 3px;
    transition: width 0.3s;
  }
  .task-cell.g .t-pct { color: var(--status-green); }
  .task-cell.g .t-bar .fill { background: var(--status-green); }
  .task-cell.o .t-pct { color: var(--status-orange); }
  .task-cell.o .t-bar .fill { background: var(--status-orange); }
  .task-cell.r .t-pct { color: var(--status-red); }
  .task-cell.r .t-bar .fill { background: var(--status-red); }

  /* === Card 3 : matrix heatmap === */
  table.heatmap {
    width: 100%;
    border-collapse: separate;
    border-spacing: 5px;
    font-size: 12.5px;
  }
  table.heatmap th {
    font-weight: 600; font-size: 11px;
    color: var(--text-soft);
    padding: 9px 5px;
    text-align: center;
    background: #F8F8FA;
    border-radius: 5px;
    line-height: 1.35;
  }
  table.heatmap th.row-label {
    text-align: left; background: transparent; color: var(--text);
    font-weight: 600; min-width: 130px;
    padding-left: 6px;
    font-size: 12.5px;
  }
  table.heatmap td {
    text-align: center;
    padding: 10px 5px;
    border-radius: 5px;
    font-weight: 700;
    font-size: 12.5px;
  }
  td.t-green  { background: var(--status-green);  color: #fff; }
  td.t-orange { background: var(--status-orange); color: #fff; }
  td.t-red    { background: var(--status-red);    color: #fff; }

  table.heatmap tfoot td {
    background: transparent !important;
    color: var(--text-soft); font-weight: 700;
    border-top: 1px solid var(--border);
    padding-top: 12px;
    font-size: 11.5px;
  }
  table.heatmap tfoot td:first-child { text-align: left; }

  @media (max-width: 1100px) {
    .app { grid-template-columns: 1fr; }
    aside.sidebar { display: none; }
    .row.top { grid-template-columns: 1fr; }
    .task-grid { grid-template-columns: repeat(2, 1fr); }
  }
</style>
</head>
<body>

<div class="app">

  <!-- ============================ SIDEBAR ============================ -->
  <aside class="sidebar">
    <div class="brand">
      <div class="logo">LG</div>
      <div class="title">팀장 REINVENT과제<br/>실행/점검</div>
    </div>

    <div class="section">
      <div class="section-head">DATA</div>
      <label class="upload-btn" for="fileInput">
        <span class="ico">📂</span>
        <span>엑셀 파일 업로드</span>
        <input type="file" id="fileInput" accept=".xlsx,.xls" />
      </label>
      <div class="upload-status" id="uploadStatus">현재 샘플 데이터</div>
    </div>

    <div class="section">
      <div class="section-head">FILTER</div>
      <div class="field">
        <label>본부</label>
        <select id="filterHq"><option value="">전체</option></select>
      </div>
      <div class="field">
        <label>과제</label>
        <select id="filterTask">
          <option value="">전체</option>
          <option value="0">① 교육 이수</option>
          <option value="1">② Cascade W/S</option>
          <option value="2">③ 과제 정의서</option>
          <option value="3">④ 중간 우수사례</option>
          <option value="4">⑤ 최종 결과 보고</option>
        </select>
      </div>
    </div>

    <div class="section">
      <div class="section-head">기준 정보</div>
      <div class="field">
        <label>기준일</label>
        <div style="color:#fff;font-size:12.5px;" id="refDate">2026-05-19</div>
      </div>
      <div class="field">
        <label>업데이트 주기</label>
        <div style="color:#fff;font-size:12.5px;">주 1회</div>
      </div>
      <div class="field">
        <label>담당</label>
        <div style="color:#fff;font-size:12.5px;">EX&amp;G팀</div>
      </div>
    </div>

    <div class="sidebar-bottom">
      <button class="apply-btn" id="applyBtn">적용</button>
    </div>
  </aside>

  <!-- ============================ MAIN ============================ -->
  <main class="main">

    <div class="page-head">
      <h1>팀장 REINVENT과제 실행/점검</h1>
      <span class="meta">
        <span>대상<strong>전사 1,600명 팀장</strong></span>
        <span>과제<strong>5개</strong></span>
      </span>
    </div>

    <!-- ============ Row 1 : 전체 진행률 + 5개 과제 ============ -->
    <div class="row top">

      <!-- Card 1 : 전체 진행률 개요 -->
      <section class="card">
        <div class="card-head">
          <div class="num">1</div>
          <h2>전체 진행률 개요</h2>
        </div>

        <div class="sub-title">1) 전사 평균 진행률</div>
        <div class="rate-box">
          <div class="pct"><span id="overallPct">68</span>%</div>
          <div class="meta">
            5개 과제<br/>
            <strong id="hqCount">10</strong>개 본부 평균
          </div>
        </div>

        <div class="sub-title">2) 본부별 현황</div>
        <div class="summary-list">
          <div class="summary-row">
            <span class="pill g">양호</span>
            <span class="count"><strong id="cntGood">4</strong> 개 본부 <span style="color:var(--text-mute)">· 80% 이상</span></span>
          </div>
          <div class="summary-row">
            <span class="pill o">보통</span>
            <span class="count"><strong id="cntMid">4</strong> 개 본부 <span style="color:var(--text-mute)">· 50–79%</span></span>
          </div>
          <div class="summary-row">
            <span class="pill r">지연</span>
            <span class="count"><strong id="cntLate">2</strong> 개 본부 <span style="color:var(--text-mute)">· 50% 미만</span></span>
          </div>
        </div>
      </section>

      <!-- Card 2 : 5개 과제별 진행률 -->
      <section class="card">
        <div class="card-head">
          <div class="num">2</div>
          <h2>과제별 전사 진행률</h2>
          <div class="legend">
            <span class="item"><span class="chip g"></span>80↑</span>
            <span class="item"><span class="chip o"></span>50–79</span>
            <span class="item"><span class="chip r"></span>50↓</span>
          </div>
        </div>

        <div class="task-grid" id="taskGrid">
          <!-- filled by JS -->
        </div>
      </section>

    </div>

    <!-- ============ Row 2 : 본부 × 과제 매트릭스 ============ -->
    <section class="card">
      <div class="card-head">
        <div class="num">3</div>
        <h2>본부 × 과제 진행률 매트릭스 (%)</h2>
        <div class="legend">
          <span class="item"><span class="chip g"></span>80↑</span>
          <span class="item"><span class="chip o"></span>50–79</span>
          <span class="item"><span class="chip r"></span>50↓</span>
        </div>
      </div>

      <table class="heatmap">
        <thead>
          <tr>
            <th class="row-label">본부</th>
            <th>① 교육<br/>이수</th>
            <th>② Cascade<br/>W/S</th>
            <th>③ 과제<br/>정의서</th>
            <th>④ 중간<br/>우수사례</th>
            <th>⑤ 최종 결과<br/>보고</th>
            <th>평균</th>
          </tr>
        </thead>
        <tbody id="matrixBody"></tbody>
        <tfoot id="matrixFoot"></tfoot>
      </table>
    </section>

  </main>
</div>

<script>
  /* ==================== TASKS & 본부 (2026년 조직 기준) ==================== */
  const TASKS = [
    { no: '①', name: '교육 이수' },
    { no: '②', name: 'Cascade W/S' },
    { no: '③', name: '과제 정의서' },
    { no: '④', name: '중간 우수사례' },
    { no: '⑤', name: '최종 결과 보고' },
  ];

  // 2026년 LG전자 조직 기준 본부명
  const SAMPLE_HQS = [
    { name: 'HS사업본부',       tasks: [96, 92, 88, 82, 68] },
    { name: 'MS사업본부',       tasks: [94, 90, 85, 78, 62] },
    { name: 'VS사업본부',       tasks: [92, 86, 82, 74, 58] },
    { name: 'ES사업본부',       tasks: [90, 84, 78, 70, 55] },
    { name: 'BS사업본부',       tasks: [88, 80, 74, 66, 50] },
    { name: '한국영업본부',     tasks: [88, 78, 72, 64, 48] },
    { name: '글로벌마케팅부문', tasks: [85, 76, 68, 60, 45] },
    { name: '고객가치혁신부문', tasks: [82, 74, 65, 56, 42] },
    { name: 'CTO부문',          tasks: [80, 70, 62, 52, 38] },
    { name: 'CHO부문',          tasks: [68, 58, 48, 38, 25] },
  ];

  /* ==================== UTILS ==================== */
  function tier(pct) {
    if (pct >= 80) return 'g';
    if (pct >= 50) return 'o';
    return 'r';
  }
  function tdClass(pct) {
    if (pct >= 80) return 't-green';
    if (pct >= 50) return 't-orange';
    return 't-red';
  }

  /* ==================== RENDER ==================== */
  function render(hqs) {
    // 1) Overall avg + 양호/보통/지연 count
    const avgPerHq = hqs.map(h => Math.round(h.tasks.reduce((a,b)=>a+b,0) / h.tasks.length));
    const overall = avgPerHq.length ? Math.round(avgPerHq.reduce((a,b)=>a+b,0) / avgPerHq.length) : 0;
    const good = avgPerHq.filter(v => v >= 80).length;
    const late = avgPerHq.filter(v => v <  50).length;
    const mid  = avgPerHq.length - good - late;

    document.getElementById('overallPct').textContent = overall;
    document.getElementById('hqCount').textContent    = hqs.length;
    document.getElementById('cntGood').textContent    = good;
    document.getElementById('cntMid').textContent     = mid;
    document.getElementById('cntLate').textContent    = late;

    // 2) 5개 과제 카드
    const taskAvg = TASKS.map((_, ti) => {
      const sum = hqs.reduce((a, h) => a + (h.tasks[ti] || 0), 0);
      return hqs.length ? Math.round(sum / hqs.length) : 0;
    });
    const grid = document.getElementById('taskGrid');
    grid.innerHTML = TASKS.map((t, i) => {
      const v = taskAvg[i];
      const c = tier(v);
      return `<div class="task-cell ${c}">
        <div class="t-no">${t.no} 과제</div>
        <div class="t-name">${t.name}</div>
        <div class="t-pct">${v}<span class="unit">%</span></div>
        <div class="t-bar"><div class="fill" style="width:${v}%"></div></div>
      </div>`;
    }).join('');

    // 3) 본부 × 과제 매트릭스
    const tbody = document.getElementById('matrixBody');
    tbody.innerHTML = hqs.map((h, idx) => {
      const cells = h.tasks.map(v => `<td class="${tdClass(v)}">${v}</td>`).join('');
      const avg = avgPerHq[idx];
      return `<tr><th class="row-label">${h.name}</th>${cells}<td class="${tdClass(avg)}"><strong>${avg}</strong></td></tr>`;
    }).join('');

    document.getElementById('matrixFoot').innerHTML =
      `<tr><td>전사 평균</td>${taskAvg.map(v => `<td>${v}%</td>`).join('')}<td>${overall}%</td></tr>`;

    // Populate 본부 filter
    const sel = document.getElementById('filterHq');
    const current = sel.value;
    sel.innerHTML = '<option value="">전체</option>' +
      hqs.map(h => `<option value="${h.name}">${h.name}</option>`).join('');
    sel.value = current;

    // ref date
    const d = new Date();
    document.getElementById('refDate').textContent =
      `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  /* ==================== FILTER ==================== */
  let currentHqs = SAMPLE_HQS;

  function applyFilter() {
    const hq   = document.getElementById('filterHq').value;
    const task = document.getElementById('filterTask').value;

    let filtered = currentHqs.map(h => ({ ...h, tasks: [...h.tasks] }));
    if (hq) filtered = filtered.filter(h => h.name === hq);
    if (task !== '') {
      const ti = Number(task);
      filtered = filtered.map(h => ({ name: h.name, tasks: h.tasks.map((v, i) => i === ti ? v : 0) }));
    }
    render(filtered.length ? filtered : currentHqs);
  }

  /* ==================== UPLOAD & PARSE ==================== */
  function setStatus(msg, kind) {
    const el = document.getElementById('uploadStatus');
    el.textContent = msg;
    el.className = 'upload-status ' + (kind || '');
  }

  function parseWorkbook(wb) {
    const sheetName = wb.SheetNames.find(n => n.includes('대상자') || n.includes('명단')) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(5, aoa.length); i++) {
      const row = aoa[i].map(v => String(v || ''));
      if (row.some(c => c === '본부') && row.some(c => c.includes('제출여부'))) {
        headerRowIdx = i; break;
      }
    }
    if (headerRowIdx < 0) throw new Error('헤더(본부 / 제출여부)를 찾지 못했습니다.');

    const header = aoa[headerRowIdx].map(v => String(v || '').trim());
    const idIdx  = header.findIndex(h => h.includes('사번'));
    const hqIdx  = header.findIndex(h => h === '본부');
    const submitCols = [];
    header.forEach((h, idx) => { if (h.includes('제출여부')) submitCols.push(idx); });
    if (submitCols.length < 5) throw new Error('"제출여부" 컬럼 5개 필요 (현재 ' + submitCols.length + '개)');
    const taskCols = submitCols.slice(0, 5);

    const byHq = new Map();
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row || row.length === 0) continue;
      const id = idIdx >= 0 ? String(row[idIdx] || '').trim() : '';
      const hq = String(row[hqIdx] || '').trim();
      if (!hq) continue;
      if (idIdx >= 0 && !id) continue;

      let entry = byHq.get(hq);
      if (!entry) { entry = { name: hq, totals: [0,0,0,0,0], submitted: [0,0,0,0,0] }; byHq.set(hq, entry); }
      for (let t = 0; t < 5; t++) {
        const val = String(row[taskCols[t]] || '').trim();
        if (val === '●' || val.toLowerCase() === 'o' || val === '제출' || val === 'TRUE' || val === '완료') {
          entry.totals[t]++; entry.submitted[t]++;
        } else if (val !== '') {
          entry.totals[t]++;
        }
      }
    }

    const hqs = [];
    for (const [name, e] of byHq.entries()) {
      const tasks = e.submitted.map((s, i) => e.totals[i] > 0 ? Math.round((s / e.totals[i]) * 100) : 0);
      hqs.push({ name, tasks });
    }
    return hqs;
  }

  function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) { setStatus('Excel(.xlsx) 파일만 가능', 'err'); return; }
    setStatus('분석 중...');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const hqs = parseWorkbook(wb);
        if (!hqs.length) { setStatus('본부 정보 없음', 'err'); return; }
        currentHqs = hqs;
        applyFilter();
        setStatus(`반영 완료 · ${hqs.length}개 본부`, 'ok');
      } catch (err) {
        console.error(err);
        setStatus('읽기 실패: ' + err.message, 'err');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /* ==================== EVENTS ==================== */
  document.getElementById('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));
  document.getElementById('applyBtn').addEventListener('click', applyFilter);

  // initial render
  render(currentHqs);
</script>

</body>
</html>
