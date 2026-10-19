(() => {
'use strict';

// ── 공통 유틸 ───────────────────────────────────────────────
const $ = (s, r = document) => r.querySelector(s);
const el = (tag, props = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v; else if (k === 'style') e.style.cssText = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat(Infinity)) if (c != null && c !== false) e.append(c instanceof Node ? c : String(c));
  return e;
};
const NS = 'http://www.w3.org/2000/svg';
const sv = (tag, attrs = {}) => { const e = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(attrs)) if (v != null) e.setAttribute(k, v); return e; };
const dots = (...parts) => parts.flat().filter((p) => p != null && p !== '').flatMap((p, i) => (i ? [' · ', p] : [p]));
const f1 = (v) => (v == null || !isFinite(v) ? '—' : d3.format(',.1f')(v));
const f0 = (v) => (v == null || !isFinite(v) ? '—' : d3.format(',')(Math.round(v)));
const pct = (v) => (v == null || !isFinite(v) ? '—' : d3.format('.0%')(v));
const fnum = (v) => (v == null || v === '' ? '—' : typeof v === 'number' ? d3.format(',')(+v.toFixed(2)) : v);
const ko = (a, b) => String(a ?? '').localeCompare(String(b ?? ''), 'ko');
const sum = (rs, k) => d3.sum(rs, (r) => (typeof k === 'function' ? k(r) : r[k]) ?? 0);
const isoDay = d3.utcFormat('%Y-%m-%d');
const today = () => d3.timeFormat('%Y-%m-%d')(new Date());
const addDays = (iso, n) => isoDay(d3.utcDay.offset(new Date(iso + 'T00:00:00Z'), n));
const mondayOf = (iso) => addDays(iso, -((new Date(iso + 'T00:00:00Z').getUTCDay() + 6) % 7));
const daysBetween = (a, b) => Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 864e5);
const pDate = d3.utcParse('%Y-%m-%d');
const ls = {
  get: (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* 저장소를 못 쓰면 매번 다시 입력 */ } },
};
const deltaText = (v) => { if (v == null || !isFinite(v)) return '—'; const r = Math.round(v * 10) / 10; return `${r > 0 ? '▲ ' : r < 0 ? '▼ ' : ''}${d3.format(',.1f')(Math.abs(r))}`; };
const deltaEl = (v, unit = '') => {
  if (v == null || !isFinite(v)) return el('span', { class: 'delta flat' }, '—');
  const r = Math.round(v * 10) / 10;
  return el('span', { class: 'delta ' + (r > 0 ? 'up' : r < 0 ? 'down' : 'flat') }, deltaText(v) + unit);
};
const statusOf = (rate) => (rate == null ? { k: '', t: '목표 전' } : rate >= 1 ? { k: 'ok', t: '양호' } : rate >= 0.8 ? { k: 'warn', t: '주의' } : { k: 'bad', t: '미흡' });
const ratePill = (rate) => { const s = statusOf(rate); return el('span', { class: 'pill ' + s.k }, rate == null ? s.t : `${s.t} ${pct(rate)}`); };
const swc = (cls, style) => el('i', { class: 'sw ' + cls, style });

// ── 양식 정의 ───────────────────────────────────────────────
const SEC = ['01 조직·과제 정보', '02 필수 지표·KPI 선택', '03 과제 진행 현황', '04 최종 KPI 목표·확보·실현', '05 최종 FTE 목표·확보', '06 S1~S5 단계별 누적 확보 목표', '07 주요 추진 일정', '08 Process-PKMS', '09 FTE 실현·재배치'];
const STAGES = ['후보 발굴·타당성', '개발·구현', '현장 실증·검증', '정식 업무 적용', '표준화·확산'];
const MGMT = { P: '생산성', Q: '품질', C: '원가', D: '납기', S: '안전', M: '사기' };
const PR = { 상: 0, 중: 1, 하: 2 };
// [키, 엑셀 컬럼명(양식과 동일), 구분 번호, 입력 형식, 선택지]
// 형식: t 글자 · ta 긴 글 · n 숫자 · d 날짜 · sel 선택 · dl 선택+직접입력 · unit KPI 단위 · auto 자동 계산 · unitcol 엑셀의 단위 표시 칸
const FIELDS = [
  ['team', '* 팀', 0, 'dl'], ['group', '* 그룹', 0, 'dl'], ['cat', '* 과제명(대분류)', 0, 'dl'], ['task', '* 과제명(중분류)', 0, 't'],
  ['kind', '* 과제구분', 0, 'sel', ['AX', 'FA']], ['agent', '* Agent 개발 포함 여부', 0, 'sel', ['O', 'X']], ['scope', '* 적용범위', 0, 'dl'],
  ['line', '* 적용범위(라인/공정)', 0, 't'], ['tgt_dept', '* Target 담당 부서명', 0, 'dl'], ['tgt_ids', '* Target 담당자 ID', 0, 't'],
  ['dev_dept', '* 개발 부서명', 0, 'dl'], ['dev_ids', '* 개발부서 담당자 ID', 0, 't'],
  ['mgmt', '* 경영관리지표', 1, 'sel', Object.keys(MGMT)], ['kpi', '* KPI', 1, 'dl'], ['kpi_etc', '* [조건] 기타 KPI', 1, 't'],
  ['kpi_etc_unit', '* [조건] 기타 KPI 단위', 1, 't'], ['unit', 'KPI 단위 [자동]', 1, 'unit'],
  ['stage', '* 진행단계', 2, 'sel', ['S1', 'S2', 'S3', 'S4', 'S5']], ['next_plan', '* [조건] 다음 단계 추진계획', 2, 'ta'], ['priority', '우선순위', 2, 'sel', ['상', '중', '하']],
  ['_u1', 'KPI AS-IS 단위 [자동]', 3, 'unitcol'], ['kpi_asis', '* [P외] KPI AS-IS 현재 수준', 3, 'n'], ['kpi_asis_basis', '* [P외] KPI AS-IS 산정근거', 3, 'ta'],
  ['kpi_dir', '* [P외] KPI 개선방향', 3, 'sel', ['증가', '감소']], ['_u2', 'KPI TO-BE 단위 [자동]', 3, 'unitcol'], ['kpi_tobe', '* [P외] KPI TO-BE 목표 수준', 3, 'n'],
  ['kpi_tobe_basis', '* [P외] KPI TO-BE 산정근거', 3, 'ta'], ['_u3', 'KPI 발굴 효과 단위 [자동]', 3, 'unitcol'], ['kpi_found', 'KPI 발굴 효과 [자동]', 3, 'auto'],
  ['_u4', 'KPI 확보 효과 단위 [자동]', 3, 'unitcol'], ['kpi_secured', 'KPI 확보 효과(현재까지)', 3, 'n'], ['kpi_secured_basis', '* [조건] KPI 확보 효과 산정근거', 3, 'ta'],
  ['_u5', 'KPI 실현 효과 단위 [자동]', 3, 'unitcol'], ['kpi_realized', 'KPI 실현 효과', 3, 'n'], ['kpi_realized_basis', '* [조건] KPI 실현 효과 산정근거', 3, 'ta'],
  ['asis_hr', '* AS-IS 1인당 건당 소요시간(시간)', 4, 'n'], ['asis_cnt', '* AS-IS 연간 처리건수', 4, 'n'], ['asis_ppl', '* AS-IS 투입인원', 4, 'n'], ['asis_basis', '* AS-IS 공수 산정근거', 4, 'ta'],
  ['tobe_hr', '* TO-BE 1인당 건당 소요시간(시간)', 4, 'n'], ['tobe_cnt', '* TO-BE 연간 처리건수', 4, 'n'], ['tobe_ppl', '* TO-BE 투입인원', 4, 'n'], ['tobe_basis', '* TO-BE 공수 산정근거', 4, 'ta'],
  ['asis_mh', 'AS-IS 연간 공수(MH) [자동]', 4, 'auto'], ['tobe_mh', 'TO-BE 연간 공수(MH) [자동]', 4, 'auto'], ['base_fte', 'Baseline FTE [자동]', 4, 'auto'], ['found_fte', '발굴 FTE [자동]', 4, 'auto'],
  ['secured_fte', '확보 FTE(현재까지)', 4, 'n'], ['secured_basis', '* [조건] 확보 FTE 산정근거', 4, 'ta'],
  ['s1_date', '* S1 목표일정', 5, 'd'], ['s2_date', '* S2 목표일정', 5, 'd'], ['s3_date', '* S3 목표일정', 5, 'd'], ['s3_fte', '* S3 누적 확보 목표 효과(FTE)', 5, 'n'],
  ['_u6', 'S3 KPI 단위 [자동]', 5, 'unitcol'], ['s3_kpi', '* [P외] S3 누적 확보 목표 효과(KPI)', 5, 'n'], ['s4_date', '* S4 목표일정', 5, 'd'], ['s4_fte', '* S4 누적 확보 목표 효과(FTE)', 5, 'n'],
  ['_u7', 'S4 KPI 단위 [자동]', 5, 'unitcol'], ['s4_kpi', '* [P외] S4 누적 확보 목표 효과(KPI)', 5, 'n'], ['s5_date', '* S5 목표일정', 5, 'd'], ['s5_fte', '* S5 누적 확보 목표 효과(FTE)', 5, 'n'],
  ['_u8', 'S5 KPI 단위 [자동]', 5, 'unitcol'], ['s5_kpi', '* [P외] S5 누적 확보 목표 효과(KPI)', 5, 'n'],
  ['plan_start', '* 계획 시작일', 6, 'd'], ['plan_end', '* 계획 종료일', 6, 'd'], ['act_start', '* [조건] 실적 시작일', 6, 'd'], ['act_end', '* [조건] 실적 종료일', 6, 'd'],
  ['pkms', '* Process-PKMS 등록여부', 7, 'sel', ['등록', '미등록']], ['pkms_name', '* [조건] Process-PKMS 과제명', 7, 't'],
  ['re_team', '팀내 업무 재배치 인원(명)', 8, 'n'], ['re_team_date', '팀내 업무 재배치 일정', 8, 'd'], ['re_line', '신규라인 재배치 인원(명)', 8, 'n'],
  ['re_line_date', '신규라인 재배치 일정', 8, 'd'], ['re_etc', '기타 실현 인원(휴직/파견 등)(명)', 8, 'n'], ['re_etc_date', '기타 실현 일정', 8, 'd'],
];
const FORM_FIELDS = FIELDS.filter((f) => f[3] !== 'unitcol');
const FORM_KEYS = FORM_FIELDS.filter((f) => f[3] !== 'auto').map((f) => f[0]);
const label = Object.fromEntries(FIELDS.map(([k, h]) => [k, h.replace(/^\*\s*/, '')]));
const shortName = (h) => h.replace(/^\*\s*/, '').replace(/\[(조건|P외)\]\s*/, '').replace(/\s*\[자동\]/, '');
const norm = (h) => String(h ?? '').replace(/^\s*\*\s*/, '').replace(/\s+/g, '');
const byHeader = new Map(FIELDS.map(([k, h]) => [norm(h), k]));

// ── 값 정리·계산 ─────────────────────────────────────────────
const isoDate = (v) => {
  if (v == null || v === '') return null;
  if (v instanceof Date) return d3.timeFormat('%Y-%m-%d')(new Date(v.getTime() + 6 * 3600e3));
  if (typeof v === 'number') return isoDay(new Date(Date.UTC(1899, 11, 30) + v * 864e5));
  const m = String(v).trim().match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : null;
};
const num = (v) => { if (v == null || v === '') return null; const n = +String(v).replace(/,/g, ''); return isNaN(n) ? null : n; };
const txt = (v) => (v == null || String(v).trim() === '' ? null : String(v).trim());

// 행 하나를 정리하고 양식의 [자동] 항목을 같은 식으로 다시 계산
function normalize(r) {
  const o = { id: txt(r.id) };
  for (const [k, , , t] of FIELDS) {
    if (t === 'unitcol' || t === 'auto') continue;
    o[k] = t === 'n' ? num(r[k]) : t === 'd' ? isoDate(r[k]) : txt(r[k]);
  }
  const st = o.stage && String(o.stage).match(/[1-5]/);
  o.stage = st ? 'S' + st[0] : null;
  o.mgmt = o.mgmt ? o.mgmt.toUpperCase().slice(0, 1) : null;
  o.priority = o.priority in PR ? o.priority : null;
  const mh = (h, c, p) => (h != null && c != null && p != null ? h * c * p : null);
  o.asis_mh = mh(o.asis_hr, o.asis_cnt, o.asis_ppl);
  o.tobe_mh = mh(o.tobe_hr, o.tobe_cnt, o.tobe_ppl);
  o.base_fte = o.asis_mh != null ? o.asis_mh / 2000 : null;
  o.found_fte = o.asis_mh != null && o.tobe_mh != null ? (o.asis_mh - o.tobe_mh) / 2000 : null;
  o.unit = o.kpi === '기타' ? (o.kpi_etc_unit ?? o.unit) : (o.unit ?? (o.mgmt === 'P' ? 'FTE' : null));
  o.kpi_found = o.mgmt !== 'P' && o.kpi_asis != null && o.kpi_tobe != null ? (o.kpi_dir === '증가' ? o.kpi_tobe - o.kpi_asis : o.kpi_asis - o.kpi_tobe) : null;
  return o;
}

// 기준일까지 도래한 가장 늦은 단계(S5→S4→S3)의 누적 확보 목표
const targetAt = (o, day) => { const s = [5, 4, 3].find((k) => o[`s${k}_date`] && o[`s${k}_date`] <= day); return s ? (o[`s${s}_fte`] ?? 0) : 0; };

function derive(o, asof) {
  const idx = o.stage ? +o.stage.slice(1) : 0;
  const done = idx === 5 && !!o.act_end;
  const due = idx ? o[`s${idx}_date`] : null;
  const delayed = !done && !!due && due < asof;
  const target_now = targetAt(o, asof);
  const realized = (o.re_team ?? 0) + (o.re_line ?? 0) + (o.re_etc ?? 0);
  const lagging = target_now > 0 && (o.secured_fte ?? 0) < target_now * 0.8;
  const late_days = delayed ? daysBetween(due, asof) : 0;
  const issues = validate(o, idx);
  const missN = issues.filter((i) => i.level === '누락').length;
  const alerts = [];
  if (delayed) alerts.push({ k: 'delay', t: `지연 ${late_days}일` });
  if (lagging) alerts.push({ k: 'lag', t: '목표 미달' });
  if (missN) alerts.push({ k: 'miss', t: `누락 ${missN}` });
  const status = done ? '완료' : delayed ? '지연' : o.plan_start && o.plan_start > asof && !o.act_start ? '착수 전' : '진행';
  return { ...o, idx, done, due, delayed, target_now, realized, lagging, late_days, issues, alerts, status,
    prog: o.found_fte ? (o.secured_fte ?? 0) / o.found_fte : null,
    kpi_prog: o.kpi_found ? (o.kpi_secured ?? 0) / o.kpi_found : null };
}

function validate(o, idx) {
  const out = [];
  const miss = (k, rule) => { if (o[k] == null || o[k] === '') out.push({ key: k, level: '누락', field: label[k], rule }); };
  for (const [k, h, , t] of FIELDS) {
    if (!h.startsWith('*') || h.includes('[조건]') || h.includes('[P외]') || t === 'auto' || t === 'unitcol') continue;
    miss(k, '공통 필수');
  }
  if (o.kpi === '기타') { miss('kpi_etc', 'KPI 기타 선택 시'); miss('kpi_etc_unit', 'KPI 기타 선택 시'); }
  if (idx >= 1 && idx <= 4) miss('next_plan', 'S1~S4 진행 시');
  if (o.mgmt && o.mgmt !== 'P') for (const k of ['kpi_asis', 'kpi_asis_basis', 'kpi_dir', 'kpi_tobe', 'kpi_tobe_basis', 's3_kpi', 's4_kpi', 's5_kpi']) miss(k, 'P 외 지표');
  if (o.kpi_secured != null) miss('kpi_secured_basis', 'KPI 확보 효과 입력 시');
  if (o.kpi_realized != null) miss('kpi_realized_basis', 'KPI 실현 효과 입력 시');
  if (o.secured_fte != null) miss('secured_basis', '확보 FTE 입력 시');
  if (o.pkms === '등록') miss('pkms_name', 'PKMS 등록 시');
  if (idx >= 2) miss('act_start', 'S2 이상 진행(착수) 시');
  if (o.act_end && idx !== 5) out.push({ key: 'act_end', level: '확인', field: label.act_end, rule: '실적 종료일이 있으나 S5가 아님' });
  const ds = [1, 2, 3, 4, 5].map((s) => o[`s${s}_date`]);
  for (let i = 1; i < 5; i++) if (ds[i - 1] && ds[i] && ds[i] < ds[i - 1]) out.push({ key: `s${i + 1}_date`, level: '확인', field: `S${i + 1} 목표일정`, rule: `S${i} 목표일보다 빠름` });
  const fs = [o.s3_fte, o.s4_fte, o.s5_fte];
  if ((fs[0] != null && fs[1] != null && fs[1] < fs[0]) || (fs[1] != null && fs[2] != null && fs[2] < fs[1])) out.push({ key: 's4_fte', level: '확인', field: '누적 확보 목표(FTE)', rule: 'S3→S4→S5 누적값이 줄어듦' });
  if (o.s5_fte != null && o.found_fte != null && o.s5_fte > o.found_fte + 0.01) out.push({ key: 's5_fte', level: '확인', field: label.s5_fte, rule: '발굴 FTE보다 큼' });
  if (o.secured_fte != null && o.found_fte != null && o.secured_fte > o.found_fte + 0.01) out.push({ key: 'secured_fte', level: '확인', field: label.secured_fte, rule: '발굴 FTE보다 큼' });
  const rel = (o.re_team ?? 0) + (o.re_line ?? 0) + (o.re_etc ?? 0);
  if (rel > Math.floor((o.secured_fte ?? 0) + 1e-9)) out.push({ key: 're_team', level: '확인', field: '실현 인원', rule: '확보 FTE보다 많음' });
  for (const [n, d] of [['re_team', 're_team_date'], ['re_line', 're_line_date'], ['re_etc', 're_etc_date']]) if (o[n] && !o[d]) out.push({ key: d, level: '누락', field: label[d], rule: '재배치 인원 입력 시' });
  return out;
}

// ── 상태 ────────────────────────────────────────────────────
const SAMPLE = JSON.parse($('#sample-data').textContent || '[]');
const SAMPLE_HIST = JSON.parse($('#sample-history').textContent || '[]');
const SAMPLE_ASOF = SAMPLE_HIST.length ? addDays(SAMPLE_HIST[SAMPLE_HIST.length - 1].end, 1) : today();
const state = {
  mode: 'demo', server: false, editable: false,      // mode: server 저장 데이터 · demo 예시 · local 서버 없이 연 화면의 내 파일
  store: [], version: 0, updated: null, updatedBy: null, serverHistory: [], serverWeekStart: null,
  rows: [], source: '', history: [], weekStart: null,
  asof: SAMPLE_ASOF, team: '', group: '', kind: '', prio: '',
  tab: 'overview', view: 'list', q: '', stage: null, sort: { key: 'priority', dir: 1 }, rule: null, open: new Set(), changes: [],
};
let ALL = [], VIEW = [];
const kindOf = (r) => r.kind ?? '미지정';
const prioOf = (r) => r.priority ?? '미지정';
const match = (r) => (!state.team || r.team === state.team) && (!state.group || r.group === state.group)
  && (!state.kind || kindOf(r) === state.kind) && (!state.prio || prioOf(r) === state.prio);
const prioSort = (a, b) => (PR[a.priority] ?? 3) - (PR[b.priority] ?? 3) || b.alerts.length - a.alerts.length || (b.found_fte ?? 0) - (a.found_fte ?? 0);
const canEdit = () => state.server && state.editable;

function prevSnap() {
  if (!state.weekStart) return null;
  const c = state.history.filter((h) => h.end < state.weekStart);
  return c.length ? c[c.length - 1] : null;
}
// 최근 8주 각 주 일요일 시점의 스냅샷 (그 주에 저장이 없으면 직전 스냅샷을 이어 씀)
function weekSeries() {
  if (!state.weekStart) return [];
  const out = [];
  for (let i = 8; i >= 1; i--) {
    const end = addDays(state.weekStart, -1 - 7 * (i - 1));
    const c = state.history.filter((h) => h.end <= end);
    out.push({ end, tasks: c.length ? c[c.length - 1].tasks : null });
  }
  return out;
}

function compute() {
  const ps = prevSnap();
  const prev = ps ? new Map(ps.tasks.map((t) => [t.id, t.secured_fte ?? 0])) : null;
  ALL = state.rows.map((r) => {
    const d = derive(r, state.asof);
    d.delta = prev ? (d.secured_fte ?? 0) - (prev.get(d.id) ?? 0) : null;
    return d;
  });
  VIEW = ALL.filter(match);
}

// 조직(팀·그룹) 단위 집계. fn은 VIEW 안에서 그 조직의 과제를 고르는 조건
function statFor(fn) {
  const cur = VIEW.filter(fn), ps = prevSnap();
  const prevT = ps ? ps.tasks.filter(match).filter(fn) : null;
  const target = sum(cur, 'target_now'), secured = sum(cur, 'secured_fte'), found = sum(cur, 'found_fte');
  const prevSec = prevT ? sum(prevT, 'secured_fte') : null;
  const prevTarget = prevT ? sum(prevT, (t) => targetAt(t, ps.end)) : null;
  return {
    rows: cur, n: cur.length, found, base: sum(cur, 'base_fte'), target, secured, realized: sum(cur, 'realized'),
    rate: target ? secured / target : null, prog: found ? secured / found : null,
    delta: prevSec == null ? null : secured - prevSec, prevRate: prevTarget ? prevSec / prevTarget : null,
    delayed: cur.filter((r) => r.delayed).length, lagging: cur.filter((r) => r.lagging).length, miss: cur.filter((r) => r.alerts.some((a) => a.k === 'miss')).length,
    alertN: cur.filter((r) => r.alerts.length).length, high: cur.filter((r) => r.priority === '상').length, done: cur.filter((r) => r.done).length,
    stages: [1, 2, 3, 4, 5].map((i) => cur.filter((r) => r.idx === i).length),
  };
}
function seriesFor(fn) {
  const pts = weekSeries().map((w) => {
    if (!w.tasks) return { end: w.end, sec: null, tgt: null };
    const t = w.tasks.filter(match).filter(fn);
    return { end: w.end, sec: sum(t, 'secured_fte'), tgt: sum(t, (x) => targetAt(x, w.end)) };
  });
  const cur = VIEW.filter(fn);
  pts.push({ end: state.asof, sec: sum(cur, 'secured_fte'), tgt: sum(cur, 'target_now'), now: true });
  return pts;
}
const teamUnit = (t) => ({ name: t, team: t, fn: (r) => (r.team ?? '미지정') === t });
const groupUnit = (t, g) => ({ name: g, sub: t, team: t, group: g, fn: (r) => (r.team ?? '미지정') === t && (r.group ?? '미지정') === g });
// 화면에 보일 조직 목록: 팀 필터가 있으면 그 팀의 그룹, 없으면 팀
function units() {
  if (state.team) return [...new Set(VIEW.map((r) => r.group ?? '미지정'))].sort(ko).map((g) => groupUnit(state.team, g));
  return [...new Set(VIEW.map((r) => r.team ?? '미지정'))].sort(ko).map(teamUnit);
}

// ── 데이터 불러오기 ─────────────────────────────────────────
async function api(path, opts = {}) {
  const res = await fetch(path, { cache: 'no-store', ...opts });
  let data = null; try { data = await res.json(); } catch (e) { /* 본문 없음 */ }
  return { ok: res.ok, status: res.status, data };
}
const post = (path, body) => api(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });

function applyStore(d) {
  state.store = d.tasks ?? []; state.version = d.version ?? 0; state.updated = d.updated; state.updatedBy = d.updated_by;
  state.editable = d.editable !== false; state.serverWeekStart = d.week_start ?? mondayOf(today());
}
async function loadHistory() {
  try { const h = await api('api/history'); state.serverHistory = h.ok ? h.data.weeks ?? [] : []; } catch (e) { state.serverHistory = []; }
}
async function loadServer(keep) {
  let r; try { r = await api('api/store'); } catch (e) { return false; }
  if (!r.ok || !r.data) return false;
  state.server = true; applyStore(r.data); await loadHistory(); showData(keep);
  return true;
}
function showData(keep) {
  if (state.store.length) setData('server', state.store, state.serverHistory, state.serverWeekStart, today(), keep);
  else setData('demo', SAMPLE, SAMPLE_HIST, mondayOf(SAMPLE_ASOF), SAMPLE_ASOF, keep);
}
function setData(mode, rows, history, weekStart, asof, keep, source = '') {
  keep = keep && state.mode === mode;
  state.mode = mode; state.source = source;
  state.rows = rows.map((r, i) => { const o = normalize(r); o.id = o.id ?? `L${i + 1}`; return o; }).filter((r) => r.task);
  state.history = (history ?? []).slice().sort((a, b) => ko(a.end, b.end)); state.weekStart = weekStart;
  if (!keep) { state.asof = asof; state.team = state.group = state.kind = state.prio = ''; state.stage = null; state.rule = null; }
  render();
  if (SHEETS.length) drawSheet();
}

// ── 마우스를 올렸을 때 뜨는 설명 ─────────────────────────────
const HT = $('#htip');
function moveTip(e) {
  const w = HT.offsetWidth, h = HT.offsetHeight;
  let x = e.clientX + 14, y = e.clientY + 16;
  if (x + w > innerWidth - 8) x = e.clientX - w - 14;
  if (y + h > innerHeight - 8) y = e.clientY - h - 14;
  HT.style.left = Math.max(8, x) + 'px'; HT.style.top = Math.max(8, y) + 'px';
}
function tipOn(node, fn) {
  node.addEventListener('mouseenter', (e) => { HT.replaceChildren(...[fn()].flat(Infinity).filter(Boolean)); HT.hidden = false; moveTip(e); });
  node.addEventListener('mousemove', moveTip);
  node.addEventListener('mouseleave', () => { HT.hidden = true; });
  return node;
}
const tipBox = (title, rows, note) => [el('div', { class: 'h' }, title),
  ...rows.filter(Boolean).map(([sw, l, v]) => el('div', { class: 'r' }, el('span', {}, sw, l), el('b', {}, v))), note ? el('div', { class: 'n' }, note) : null];
const bulletTip = (s, name) => tipBox(name, [
  [swc('found'), '발굴 FTE (바탕)', f1(s.found)], [swc('sec'), '확보 FTE (막대)', f1(s.secured)], [swc('tgt'), '기준일 목표 (세로선)', s.target ? f1(s.target) : '도래 전'],
  [null, '달성률 (확보÷목표)', pct(s.rate)], [null, '진척률 (확보÷발굴)', pct(s.prog)]], '클릭: 조직 상세');
const trendTip = (pts, name) => [el('div', { class: 'h' }, `${name} · 최근 8주 추이`),
  el('div', { class: 'r' }, el('span', {}, swc('line'), '실선: 확보 FTE'), ''), el('div', { class: 'r' }, el('span', {}, swc('dash'), '점선: 기준일 목표'), ''),
  el('table', {}, el('tr', {}, el('th', {}, '주말'), el('th', { class: 'num' }, '확보'), el('th', { class: 'num' }, '목표')),
    pts.map((p) => el('tr', {}, el('td', {}, p.now ? '현재' : p.end.slice(5).replace('-', '/')), el('td', { class: 'num' }, p.sec == null ? '—' : f1(p.sec)), el('td', { class: 'num' }, p.tgt == null ? '—' : f1(p.tgt)))))];
const stageTip = (stages, name) => tipBox(`${name} · 진행단계 분포`, stages.map((n, i) => [swc('', `background:${stColor(i + 1)}`), `S${i + 1} ${STAGES[i]}`, `${n}건`]));
const TS = { delay: ['지연', 'var(--bad)'], lag: ['목표 미달', '#c98a1d'], run: ['정상 진행', 'var(--accent)'], done: ['완료', 'var(--ok)'], pre: ['착수 전', 'var(--line-2)'] };
const tileState = (r) => (r.delayed ? 'delay' : r.lagging ? 'lag' : r.done ? 'done' : r.status === '착수 전' ? 'pre' : 'run');
const taskTip = (r) => tipBox(r.task, [
  [null, '조직', `${r.team ?? '—'} · ${r.group ?? '—'}`], [null, '단계', `${r.stage ?? '—'} ${STAGES[r.idx - 1] ?? ''}`],
  [null, '상태', TS[tileState(r)][0] + (r.delayed ? ` (목표일 ${r.due}, ${r.late_days}일 경과)` : '')],
  [swc('found'), '발굴 FTE', f1(r.found_fte)], [swc('sec'), '확보 FTE', f1(r.secured_fte ?? 0)], [swc('tgt'), '기준일 목표', r.target_now ? f1(r.target_now) : '도래 전'],
  [null, '진척률 (확보÷발굴)', pct(r.prog)], [null, '전주 대비', deltaText(r.delta)], r.realized ? [null, '실현 인원', `${r.realized}명`] : null], '클릭: 상세');

// ── 상세 창(시트): 클릭한 수치의 관련 데이터를 띄움. 창 안에서 다시 누르면 위에 쌓이고 '이전'으로 돌아감 ──
const SHEETS = [];
function openSheet(spec) { SHEETS.push(spec); drawSheet(); }
function closeSheets() { SHEETS.length = 0; drawSheet(); }
function backSheet() { SHEETS.pop(); drawSheet(); }
function drawSheet() {
  const box = $('#sheet'), s = SHEETS[SHEETS.length - 1];
  HT.hidden = true;
  if (!s) { box.hidden = true; box.replaceChildren(); return; }
  compute();
  const call = (v) => (typeof v === 'function' ? v() : v);
  box.replaceChildren(el('div', { class: 'sbox' },
    el('header', { class: 'card-hd' },
      el('div', {}, SHEETS.length > 1 ? el('button', { class: 'link', type: 'button', onclick: backSheet }, '← 이전') : null,
        call(s.crumb) ? el('div', { class: 'crumb' }, call(s.crumb)) : null, el('h2', {}, call(s.title)), call(s.sub) ? el('div', { class: 'sub' }, call(s.sub)) : null),
      el('div', { class: 'dt-acts' }, call(s.actions), el('button', { class: 'btn', type: 'button', onclick: closeSheets }, '닫기'))),
    el('div', { class: 'sbd' }, call(s.body))));
  box.hidden = false;
  box.querySelector('.sbd').scrollTop = 0;
}
$('#sheet').addEventListener('click', (e) => { if (e.target.id === 'sheet') closeSheets(); });

const kpis = (items) => el('div', { class: 'kpis' }, items.filter(Boolean).map(([l, v, unit, s]) => el('div', { class: 'kpi' }, el('div', { class: 'l' }, l), el('div', { class: 'v' }, v, unit ? el('small', {}, unit) : null), s ? el('div', { class: 's' }, s) : null)));
const panel = (title, body) => el('section', { class: 'panel' }, el('h4', {}, title), body);
// 과제 표: 행을 누르면 과제 상세 창. extra = [[머리글, 값 함수, 숫자 여부]]
function taskTable(rows, extra = []) {
  if (!rows.length) return el('p', { class: 'empty' }, '해당 과제가 없습니다.');
  const head = ['우선', '과제명', '팀 · 그룹', '단계', '경보', '발굴 FTE', '기준일 목표', '확보 FTE', '전주 대비', ...extra.map((x) => x[0])];
  const numIdx = new Set([5, 6, 7, 8, ...extra.map((x, i) => (x[2] ? 9 + i : -1))]);
  return el('div', { class: 'tblwrap', style: 'max-height:none' }, el('table', {},
    el('thead', {}, el('tr', {}, head.map((h, i) => el('th', { class: numIdx.has(i) ? 'num' : '' }, h)))),
    el('tbody', {}, rows.map((r) => el('tr', { class: 'clk', onclick: () => openTask(r.id) },
      el('td', {}, prioBadge(r)), el('td', { class: 'task-nm' }, r.task), el('td', {}, `${r.team ?? '—'} · ${r.group ?? '—'}`), el('td', {}, stchip(r)),
      el('td', {}, r.alerts.length ? alertChips(r) : '—'), el('td', { class: 'num' }, f1(r.found_fte)), el('td', { class: 'num' }, r.target_now ? f1(r.target_now) : '—'),
      el('td', { class: 'num' }, el('b', { class: r.lagging ? 'c-bad' : '' }, f1(r.secured_fte ?? 0))), el('td', { class: 'num' }, deltaEl(r.delta)),
      extra.map(([, fn, n]) => el('td', { class: n ? 'num' : '' }, fn(r))))))));
}
function orgTable(us) {
  const st = us.map((u) => ({ u, s: statFor(u.fn) })), mx = d3.max(st, (d) => d.s.found) || 1;
  return el('div', { class: 'tblwrap', style: 'max-height:none' }, el('table', {},
    el('thead', {}, el('tr', {}, ['조직', '과제', 'FTE (확보 / 목표 · 발굴)', '달성률', '전주 대비', '실현(명)', '지연', '경보'].map((h, i) => el('th', { class: i && i !== 2 ? 'num' : '' }, h)))),
    el('tbody', {}, st.map(({ u, s }) => el('tr', { class: 'clk', onclick: () => openOrg(u) },
      el('td', { class: 'task-nm' }, u.name), el('td', { class: 'num' }, s.n), el('td', {}, fteCell(s, mx, u.name)), el('td', { class: 'num' }, ratePill(s.rate)),
      el('td', { class: 'num' }, deltaEl(s.delta)), el('td', { class: 'num' }, f0(s.realized)), el('td', { class: 'num' + (s.delayed ? ' c-bad' : '') }, s.delayed), el('td', { class: 'num' + (s.alertN ? ' c-bad' : '') }, s.alertN))))));
}
function openTasks(title, rows, { sub, extra, top } = {}) {
  openSheet({ title, sub: sub ?? `과제 ${rows.length}건`, body: () => [top ? top() : null, panel('과제', taskTable(rows, extra))] });
}

function openTask(id) {
  openSheet({
    crumb: () => { const r = ALL.find((d) => d.id === id); return r ? dots(r.team, r.group, r.cat) : null; },
    title: () => ALL.find((d) => d.id === id)?.task ?? '삭제된 과제',
    sub: () => { const r = ALL.find((d) => d.id === id); return r ? [prioBadge(r), stchip(r), r.delayed ? null : statusPill(r), alertChips(r)] : null; },
    actions: () => {
      const r = ALL.find((d) => d.id === id), ok = r && state.mode === 'server' && canEdit();
      if (!r) return null;
      return [el('button', { class: 'btn primary', type: 'button', disabled: !ok, title: ok ? '' : editHint(), onclick: () => { closeSheets(); openEditor(r); } }, '수정'),
        el('button', { class: 'btn danger', type: 'button', disabled: !ok, onclick: () => removeTask(r) }, '삭제')];
    },
    body: () => { const r = ALL.find((d) => d.id === id); return r ? taskBody(r) : el('p', { class: 'empty' }, '이 과제는 삭제되었습니다.'); },
  });
}
function taskBody(r) {
  const missing = new Set(r.issues.filter((i) => i.level === '누락').map((i) => i.key));
  const show = (k, t) => { const v = r[k]; return v == null || v === '' ? '—' : t === 'n' || t === 'auto' ? fnum(v) : v; };
  return [
    kpis([['발굴 FTE', f1(r.found_fte), 'FTE', `Baseline ${f1(r.base_fte)}`], ['기준일 목표', r.target_now ? f1(r.target_now) : '—', 'FTE', r.target_now ? null : '아직 도래한 목표 없음'],
      ['확보 FTE', f1(r.secured_fte ?? 0), 'FTE', [r.target_now ? ratePill((r.secured_fte ?? 0) / r.target_now) : null, ' 전주 대비 ', deltaEl(r.delta)]],
      ['진척률', pct(r.prog), null, '확보 ÷ 발굴'], ['실현 인원', f0(r.realized), '명', dots(`팀내 ${r.re_team ?? 0}`, `신규라인 ${r.re_line ?? 0}`, `기타 ${r.re_etc ?? 0}`)],
      r.mgmt !== 'P' ? ['KPI 진척', pct(r.kpi_prog), null, `${r.kpi ?? ''} ${fnum(r.kpi_secured)} / ${fnum(r.kpi_found)} ${r.unit ?? ''}`] : null]),
    r.issues.length ? el('p', { class: 'issues-line' }, '점검: ' + r.issues.map((i) => `${i.field} (${i.rule})`).join(' · ')) : null,
    el('div', { class: 'secs' }, SEC.map((s, si) => el('section', { class: 'sec' }, el('h4', {}, s),
      si === 5 ? el('div', { class: 'ms' }, [1, 2, 3, 4, 5].map((k) => el('div', { class: (r.done || k < r.idx ? 'done' : '') + (k === r.idx ? (r.delayed ? ' late' : ' cur') : '') },
        el('b', {}, `S${k}`), r[`s${k}_date`] ? r[`s${k}_date`].slice(2).replace(/-/g, '.') : '—', k >= 3 ? el('div', {}, `${fnum(r[`s${k}_fte`])} FTE`) : null))) : null,
      el('dl', { class: 'kv' }, FORM_FIELDS.filter((f) => f[2] === si).flatMap(([k, , , t]) => [el('dt', {}, label[k]), el('dd', { class: missing.has(k) ? 'miss' : '' }, missing.has(k) ? '미입력' : show(k, t))]))))),
  ];
}
function openOrg(u) {
  openSheet({
    crumb: () => (u.sub ? u.sub : null), title: u.name,
    sub: () => { const s = statFor(u.fn); return [ratePill(s.rate), ...dots(`과제 ${s.n}건`, el('span', {}, '전주 대비 ', deltaEl(s.delta, ' FTE')))]; },
    body: () => {
      const s = statFor(u.fn), pts = seriesFor(u.fn);
      const trend = el('div', { style: 'padding:12px' }, sparkline(pts, 520, 110), el('div', { class: 'legend', style: 'margin-top:6px' }, el('span', {}, swc('line'), '확보 FTE'), el('span', {}, swc('dash'), '기준일 목표'), el('span', { style: 'color:var(--ink-2)' }, '최근 8주 + 현재 · 마우스: 주별 값')));
      tipOn(trend.firstChild, () => trendTip(pts, u.name));
      const groups = !u.group ? [...new Set(s.rows.map((r) => r.group ?? '미지정'))].sort(ko).map((g) => groupUnit(u.team, g)) : [];
      return [
        kpis([['과제', s.n, '건', dots(`완료 ${s.done}`, `지연 ${s.delayed}`, `우선 상 ${s.high}`)], ['발굴 FTE', f1(s.found), 'FTE', `Baseline ${f1(s.base)}`],
          ['기준일 목표', f1(s.target), 'FTE', `달성률 ${pct(s.rate)} (전주 ${pct(s.prevRate)})`], ['확보 FTE', f1(s.secured), 'FTE', ['전주 대비 ', deltaEl(s.delta)]],
          ['진척률', pct(s.prog), null, '확보 ÷ 발굴'], ['실현 인원', f0(s.realized), '명', `확보 대비 ${pct(s.secured ? s.realized / s.secured : null)}`]]),
        el('div', { class: 'grid2' }, panel('주간 추이', trend), panel('진행단계 분포', stageList(s.stages))),
        groups.length > 1 ? panel('그룹별', orgTable(groups)) : null,
        panel('과제 (우선순위 순)', taskTable(s.rows.slice().sort(prioSort))),
      ];
    },
  });
}
const stageList = (stages) => el('div', { class: 'mgmt', style: 'padding:10px 12px' }, stages.map((n, i) => el('div', { class: 'mg-row', style: 'cursor:default;grid-template-columns:9rem 1fr 3rem' },
  el('b', {}, el('i', { class: 'sw', style: `background:${stColor(i + 1)}` }), `S${i + 1} ${STAGES[i]}`), el('i', { style: `width:${Math.max(1, (n / (d3.max(stages) || 1)) * 100)}%;background:${stColor(i + 1)}` }), el('span', {}, `${n}건`))));

// ── 렌더: 공통 ──────────────────────────────────────────────
const ICON = {
  overview: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z',
  map: 'M3 4h8v7H3zM13 4h8v4h-8zM13 10h8v10h-8zM3 13h8v7H3z',
  org: 'M9 3h6v5H9zM3 16h6v5H3zM15 16h6v5h-6zM12 8v4M6 16v-4h12v4',
  tasks: 'M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01',
  schedule: 'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4',
  check: 'M8.5 12.5l2.5 2.5 4.5-5M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18z',
};
const TABS = [['overview', '종합 현황'], ['map', '과제 맵'], ['org', '조직별 현황'], ['tasks', '과제 관리'], ['schedule', '추진 일정'], ['check', '데이터 점검']];
function buildTabs() {
  $('#tabs').replaceChildren(...TABS.map(([k, t]) => {
    const svg = sv('svg', { viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 1.8, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' });
    svg.append(sv('path', { d: ICON[k] }));
    return el('button', { class: 'tab', role: 'tab', id: `t-${k}`, type: 'button', 'aria-selected': 'false', onclick: () => setTab(k) }, svg, t, el('span', { class: 'cnt', id: `cnt-${k}` }));
  }));
}
function setTab(t) {
  state.tab = t;
  for (const [k] of TABS) { $(`#t-${k}`).setAttribute('aria-selected', String(k === t)); $(`#p-${k}`).hidden = k !== t; }
  try { history.replaceState(null, '', '#' + t); } catch (e) { /* 주소 표시만 못 바꿈 */ }
  if (t === 'check' && state.server) loadChanges();
  render();
}
const editHint = () => (state.server ? (state.mode === 'server' ? '보기 전용으로 실행 중인 서버입니다.' : '예시 데이터는 수정할 수 없습니다. 엑셀 업로드나 과제 추가로 시작하세요.') : '서버(serve.py)로 연 화면에서만 저장할 수 있습니다.');

function render() {
  HT.hidden = true;
  compute();
  renderChrome();
  ({ overview: renderOverview, map: renderMap, org: renderOrg, tasks: renderTasks, schedule: renderGantt, check: renderCheck })[state.tab]();
}

function fillSelect(sel, opts, cur) {
  const sig = opts.map((o) => o.join('=')).join('|');
  if (sel.dataset.sig !== sig) { sel.replaceChildren(...opts.map(([v, t]) => new Option(t, v))); sel.dataset.sig = sig; }
  sel.value = cur;
}
function renderChrome() {
  const src = { server: `저장 데이터 · 최종 수정 ${state.updated ?? '—'}${state.updatedBy ? ` (${state.updatedBy})` : ''}`, demo: '예시 데이터(가상)', local: `내 파일 · ${state.source}` }[state.mode];
  $('#meta').textContent = `${src} · 과제 ${state.rows.length}건 · 기준일 ${state.asof}`;
  const add = $('#b-add'); add.disabled = !canEdit(); add.title = canEdit() ? '' : editHint();
  $('#b-export').disabled = !state.rows.length;

  const teams = [...new Set(ALL.map((r) => r.team).filter(Boolean))].sort(ko);
  const groups = [...new Set(ALL.filter((r) => !state.team || r.team === state.team).map((r) => r.group).filter(Boolean))].sort(ko);
  if (state.group && !groups.includes(state.group)) state.group = '';
  fillSelect($('#f-team'), [['', '전체'], ...teams.map((t) => [t, t])], state.team);
  fillSelect($('#f-group'), [['', '전체'], ...groups.map((g) => [g, g])], state.group);
  fillSelect($('#f-kind'), [['', '전체'], ...[...new Set(ALL.map(kindOf))].sort(ko).map((k) => [k, k])], state.kind);
  fillSelect($('#f-prio'), [['', '전체'], ['상', '상'], ['중', '중'], ['하', '하'], ['미지정', '미지정']], state.prio);
  if ($('#f-asof') !== document.activeElement) $('#f-asof').value = state.asof;
  $('#b-reset').hidden = !(state.team || state.group || state.kind || state.prio);

  for (const k of ['overview', 'map', 'schedule']) $(`#cnt-${k}`).hidden = true;
  $('#cnt-org').textContent = units().length;
  $('#cnt-tasks').textContent = VIEW.length;
  const nIss = VIEW.filter((r) => r.issues.length).length, c = $('#cnt-check');
  c.textContent = nIss; c.hidden = !nIss; c.classList.toggle('bad', nIss > 0);

  const b = $('#banner');
  b.className = 'banner'; b.hidden = false;
  if (state.mode === 'local') b.replaceChildren(...dots(el('b', {}, '내 파일 보기'), state.source, '서버 없이 연 화면이라 저장되지 않음'), el('button', { class: 'btn', type: 'button', onclick: () => setData('demo', SAMPLE, SAMPLE_HIST, mondayOf(SAMPLE_ASOF), SAMPLE_ASOF, false) }, '예시 데이터로 돌아가기'));
  else if (state.mode === 'demo' && state.server) b.replaceChildren(...dots(el('b', {}, '예시 데이터(가상) 표시 중'), '저장된 과제 없음', '[엑셀 업로드] 또는 [+ 과제 추가]로 입력하면 예시는 사라짐'));
  else if (state.mode === 'demo') b.replaceChildren(...dots(el('b', {}, '서버 없이 연 화면'), '예시 데이터(가상) 표시', '저장·공유는 serve.py로 실행 시 가능'));
  else if (!state.editable) { b.classList.add('warn'); b.replaceChildren(...dots(el('b', {}, '보기 전용'), '서버가 --readonly 로 실행 중')); }
  else b.hidden = true;
  takeaway();
}

function takeaway() {
  const box = $('#take'), s = statFor(() => true);
  if (!s.n) { box.textContent = '조건에 맞는 과제가 없습니다.'; return; }
  const scope = [state.team || '전체', state.group, state.kind && `${state.kind} 과제`, state.prio && `우선순위 ${state.prio}`].filter(Boolean).join(' ');
  const st = statusOf(s.rate);
  box.replaceChildren(...[
    `${scope} ${s.n}건 · 발굴 `, el('b', {}, `${f1(s.found)} FTE`), ' 중 ', el('b', {}, `${f1(s.secured)} FTE(${pct(s.prog)})`), ' 확보',
    s.target ? [' · 기준일 목표 대비 ', el('b', { class: 'c-' + (st.k || 'mut') }, pct(s.rate))] : ' · 기준일까지 도래한 목표 없음',
    s.delta != null ? [' · 전주 대비 ', deltaEl(s.delta, ' FTE')] : '',
    s.delayed ? [' · 일정 지연 ', el('b', { class: 'c-bad' }, `${s.delayed}건`)] : ' · 일정 지연 없음',
  ].flat());
}

let toastTimer;
function toast(msg, kind = '') {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast ' + kind; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, kind === 'err' ? 6000 : 3500);
}

const stColor = (i) => `var(--st${i || 1})`;
const stchip = (r) => el('span', { class: 'stchip', style: `border-color:${stColor(r.idx)}`, title: STAGES[r.idx - 1] ?? '' }, r.stage ?? '—');
const prioBadge = (r) => el('span', { class: 'prio ' + (r.priority ? 'p' + r.priority : 'none'), title: r.priority ? `우선순위 ${r.priority}` : '우선순위 미지정' }, r.priority ?? '–');
const alertChips = (r) => r.alerts.map((a) => el('span', { class: 'alert ' + a.k }, a.t));
const statusPill = (r) => el('span', { class: 'pill ' + ({ 완료: 'ok', 지연: 'bad' }[r.status] ?? '') }, r.status);

function sparkline(pts, w = 92, h = 30) {
  const svg = sv('svg', { class: 'spark', width: w, height: h, viewBox: `0 0 ${w} ${h}`, role: 'img', 'aria-label': '주간 추이' });
  const vals = pts.flatMap((p) => [p.sec, p.tgt]).filter((v) => v != null);
  if (!vals.length) return svg;
  const x = d3.scaleLinear([0, pts.length - 1], [3, w - 5]), y = d3.scaleLinear([0, d3.max(vals) || 1], [h - 3, 4]);
  svg.append(sv('path', { d: d3.line().defined((p) => p.tgt != null).curve(d3.curveStepAfter).x((p, i) => x(i)).y((p) => y(p.tgt))(pts), fill: 'none', stroke: 'var(--ink-2)', 'stroke-width': 1.2, 'stroke-dasharray': '3 2' }));
  svg.append(sv('path', { d: d3.line().defined((p) => p.sec != null).x((p, i) => x(i)).y((p) => y(p.sec))(pts), fill: 'none', stroke: 'var(--accent)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  svg.append(sv('circle', { cx: x(pts.length - 1), cy: y(pts[pts.length - 1].sec ?? 0), r: 3, fill: 'var(--accent)' }));
  return svg;
}
function bullet(s, mx) {
  const sc = (v) => `${Math.min(100, (v / (mx || 1)) * 100)}%`;
  return el('div', { class: 'bullet' }, el('div', { class: 'rng', style: `width:${sc(s.found)}` }), el('div', { class: 'bar', style: `width:${sc(s.secured)}` }),
    s.target ? el('div', { class: 'tgt', style: `left:${sc(s.target)}` }) : null);
}
const fteCell = (s, mx, name) => tipOn(el('div', { class: 'fte-cell' }, bullet(s, mx), el('small', {}, '확보 ', el('b', {}, f1(s.secured)), ` / 목표 ${f1(s.target)} · 발굴 ${f1(s.found)}`)), () => bulletTip(s, name));
const legendEl = (items) => items.map((it) => (it === '|' ? el('span', { class: 'sep' }) : typeof it === 'string' ? el('span', { class: 'grp-l' }, it) : el('span', {}, it[0], it[1])));
function diamond(fill, stroke, sw = 1.3) {
  const s = sv('svg', { width: 12, height: 12, viewBox: '0 0 12 12' }); s.append(sv('path', { d: 'M6 1L11 6L6 11L1 6Z', fill, stroke, 'stroke-width': sw }));
  return el('span', { class: 'sw-svg' }, s);
}
function buildLegends() {
  $('#lg-org').replaceChildren(...legendEl([[swc('found'), '발굴'], [swc('sec'), '확보'], [swc('tgt'), '기준일 목표'], '|', [swc('line'), '추이: 확보'], [swc('dash'), '추이: 목표']]));
  $('#lg-curve').replaceChildren(...legendEl([[swc('step'), '누적 목표'], [swc('line'), '확보 실적'], [swc('dash2'), '발굴 총량']]));
  $('#lg-map').replaceChildren(el('div', { class: 'legend' }, ...legendEl(['상태', ...Object.entries(TS).map(([k, [t, c]]) => [el('i', { class: 'st-sw', style: `border-left-color:${c}` }), t]), '|', '막대',
    [swc('sec'), '확보'], [swc('found'), '발굴(전체 길이)'], [swc('tgt'), '기준일 목표'], '|', '단계', [el('span', { class: 'sdots' }, el('i', { class: 'on' }), el('i', { class: 'on' }), el('i', { class: 'cur' }), el('i'), el('i')), ' S1~S5 (테두리: 현재)'],
    '|', [el('span', { class: 'prio p상', style: 'margin-right:5px' }, '상'), '우선순위']])));
  $('#lg-score').replaceChildren(el('div', { class: 'legend' }, ...legendEl(['단계 분포', ...STAGES.map((n, i) => [swc('', `background:${stColor(i + 1)}`), `S${i + 1} ${n}`]), '|', 'FTE',
    [swc('found'), '발굴'], [swc('sec'), '확보'], [swc('tgt'), '기준일 목표'], '|', '추이(8주)', [swc('line'), '확보'], [swc('dash'), '목표']])));
  $('#lg-gantt').replaceChildren(el('div', { class: 'legend' }, ...legendEl(['기간', [swc('found'), '계획'], [swc('sec'), '실적(진행)'], [swc('bad'), '실적(지연)'], [swc('ok'), '실적(완료)'], '|',
    '단계 목표일', [diamond('var(--st3)', 'var(--st5)'), '지난 단계'], [diamond('var(--panel)', 'var(--st5)'), '남은 단계'], [diamond('var(--panel)', 'var(--bad)', 2), '목표일 경과(지연)'], '|',
    [el('i', { class: 'sw tgt' }), '기준일']])));
}

// ── 렌더: 종합 현황 ─────────────────────────────────────────
function renderOverview() { tiles(); orgBars(); flow(); pipe(); curve(); watch(); mgmt(); }

// 숫자 카드를 누르면 뜨는 상세 창
const DRILL = {
  count: () => openTasks('등록 과제 현황', VIEW.slice().sort((a, b) => a.idx - b.idx || prioSort(a, b)), { top: () => {
    const s = statFor(() => true);
    return [kpis([['등록 과제', s.n, '건'], ['완료', s.done, '건'], ['진행', s.n - s.done - s.delayed, '건'], ['지연', s.delayed, '건'], ['우선순위 상', s.high, '건']]),
      el('div', { class: 'grid2' }, panel('진행단계 분포', stageList(s.stages)), panel('조직별', orgTable(units())))];
  } }),
  found: () => openTasks('발굴 FTE 현황', VIEW.slice().sort((a, b) => (b.found_fte ?? 0) - (a.found_fte ?? 0)), {
    extra: [['Baseline FTE', (r) => f1(r.base_fte), 1], ['발굴률', (r) => pct(r.base_fte ? r.found_fte / r.base_fte : null), 1]],
    top: () => { const s = statFor(() => true); return [kpis([['Baseline FTE', f1(s.base), 'FTE', '현재 연간 공수 ÷ 2,000'], ['발굴 FTE', f1(s.found), 'FTE', '(AS-IS − TO-BE 공수) ÷ 2,000'], ['발굴률', pct(s.base ? s.found / s.base : null), null, '발굴 ÷ Baseline']]), panel('조직별', orgTable(units()))]; } }),
  secured: () => openTasks('확보 FTE 현황', VIEW.slice().sort((a, b) => ((b.target_now ?? 0) - (b.secured_fte ?? 0)) - ((a.target_now ?? 0) - (a.secured_fte ?? 0))), {
    sub: '목표 대비 부족이 큰 과제부터',
    extra: [['달성률', (r) => (r.target_now ? ratePill((r.secured_fte ?? 0) / r.target_now) : '목표 전')], ['부족분', (r) => (r.target_now ? f1(Math.max(0, r.target_now - (r.secured_fte ?? 0))) : '—'), 1]],
    top: () => { const s = statFor(() => true); return [kpis([['확보 FTE', f1(s.secured), 'FTE'], ['기준일 목표', f1(s.target), 'FTE', ratePill(s.rate)], ['전주 대비', deltaText(s.delta), 'FTE', `전주 달성률 ${pct(s.prevRate)}`], ['진척률', pct(s.prog), null, '확보 ÷ 발굴']]), panel('조직별', orgTable(units()))]; } }),
  realized: () => openTasks('실현 인원 현황', VIEW.filter((r) => r.realized || (r.secured_fte ?? 0) >= 1).sort((a, b) => b.realized - a.realized || (b.secured_fte ?? 0) - (a.secured_fte ?? 0)), {
    sub: '실현 인원이 있거나 확보 FTE 1 이상인 과제',
    extra: [['실현 계', (r) => (r.realized ? `${r.realized}명` : '—'), 1], ['팀내 재배치', (r) => (r.re_team ? `${r.re_team}명 · ${r.re_team_date ?? '일정 미입력'}` : '—')], ['신규라인', (r) => (r.re_line ? `${r.re_line}명 · ${r.re_line_date ?? '일정 미입력'}` : '—')], ['기타', (r) => (r.re_etc ? `${r.re_etc}명 · ${r.re_etc_date ?? '일정 미입력'}` : '—')]],
    top: () => { const s = statFor(() => true); return kpis([['실현 인원', f0(s.realized), '명', `확보 FTE ${f1(s.secured)} 대비 ${pct(s.secured ? s.realized / s.secured : null)}`], ['팀내 재배치', f0(sum(s.rows, 're_team')), '명'], ['신규라인 재배치', f0(sum(s.rows, 're_line')), '명'], ['기타(휴직/파견 등)', f0(sum(s.rows, 're_etc')), '명']]); } }),
  alerts: () => openTasks('관리 필요 과제', VIEW.filter((r) => r.alerts.length).sort(prioSort), {
    extra: [['사유', (r) => [r.delayed ? `${r.stage} 목표일 ${r.due} 경과` : null, r.lagging ? `확보 ${f1(r.secured_fte ?? 0)} < 목표 ${f1(r.target_now)}의 80%` : null, ...r.issues.filter((i) => i.level === '누락').map((i) => `${shortName(i.field)} 미입력`)].filter(Boolean).join(' · ')]],
    top: () => { const s = statFor(() => true); return kpis([['관리 필요', s.alertN, '건'], ['일정 지연', s.delayed, '건', '현재 단계 목표일 경과'], ['목표 미달', s.lagging, '건', '확보 < 기준일 목표의 80%'], ['입력 누락', s.miss, '건', '필수·조건부 필수 미입력']]); } }),
};

function tiles() {
  const s = statFor(() => true);
  const tile = (key, lab, big, unit, sub, cls = '') => el('button', { class: 'tile ' + cls, type: 'button', onclick: DRILL[key] }, el('span', { class: 'lab' }, lab), el('span', { class: 'big' }, big, unit ? el('small', {}, unit) : null), el('span', { class: 'sub' }, sub));
  $('#tiles').replaceChildren(
    tile('count', '등록 과제', f0(s.n), '건', dots(`완료 ${s.done}`, `진행 ${s.n - s.done}`, `우선순위 상 ${s.high}`)),
    tile('found', '발굴 FTE', f1(s.found), 'FTE', `Baseline ${f1(s.base)} FTE 대비 ${pct(s.base ? s.found / s.base : null)}`),
    tile('secured', '확보 FTE', f1(s.secured), 'FTE', [ratePill(s.rate), `목표 ${f1(s.target)}`, el('span', {}, '전주 대비 ', deltaEl(s.delta))]),
    tile('realized', '실현 인원', f0(s.realized), '명', dots(`확보 대비 ${pct(s.secured ? s.realized / s.secured : null)}`, `팀내 ${f0(sum(s.rows, 're_team'))}`, `신규라인 ${f0(sum(s.rows, 're_line'))}`, `기타 ${f0(sum(s.rows, 're_etc'))}`)),
    tile('alerts', '관리 필요 과제', f0(s.alertN), '건', dots(`지연 ${s.delayed}`, `목표 미달 ${s.lagging}`, `입력 누락 ${s.miss}`), s.alertN ? 'alertile' : ''),
  );
}

function orgBars() {
  const us = units().map((u) => ({ ...u, s: statFor(u.fn), pts: seriesFor(u.fn) }));
  const mx = d3.max(us, (u) => u.s.found) || 1;
  $('#orgbars').replaceChildren(...(us.length ? us.map((u) => {
    const b = tipOn(bullet(u.s, mx), () => bulletTip(u.s, u.name)), sp = tipOn(el('div', {}, sparkline(u.pts)), () => trendTip(u.pts, u.name));
    const rate = tipOn(el('div', { class: 'ob-num' }, ratePill(u.s.rate), el('div', {}, `전주 ${pct(u.s.prevRate)}`)), () => tipBox(`${u.name} · 달성률`, [[null, '이번 주', pct(u.s.rate)], [null, '지난주', pct(u.s.prevRate)]], '확보 FTE ÷ 기준일 목표 · 양호 100%↑ · 주의 80~100% · 미흡 80%↓'));
    const dl = tipOn(el('div', { class: 'ob-num' }, deltaEl(u.s.delta), el('div', {}, '전주 대비')), () => tipBox(`${u.name} · 전주 대비`, [[null, '확보 FTE 증감', deltaText(u.s.delta)]], '지난주 일요일 확보 FTE와 비교'));
    return el('div', { class: 'ob-row', tabindex: 0, role: 'button', onclick: () => openOrg(u), onkeydown: (e) => { if (e.key === 'Enter') openOrg(u); } },
      el('div', { class: 'ob-name' }, el('b', {}, u.name), el('small', {}, `확보 ${f1(u.s.secured)} / 목표 ${f1(u.s.target)}`)), b, rate, dl, sp);
  }) : [el('p', { class: 'empty' }, '조건에 맞는 조직이 없습니다.')]));
}

function flow() {
  const s = statFor(() => true), mx = Math.max(s.base, 1e-9);
  const row = (nm, v, unit, color, drill, extra, fmt = f1) => el('div', { class: 'row clk', role: 'button', tabindex: 0, onclick: drill, title: '클릭: 과제별 값' }, el('span', {}, nm),
    el('div', { class: 'track' }, el('div', { class: 'bar', style: `width:${Math.min(100, (v / mx) * 100)}%;background:${color}` }), extra), el('span', { class: 'val' }, `${fmt(v)} ${unit}`));
  const conv = (t) => el('div', { class: 'row' }, el('span'), el('span', { class: 'conv' }, t), el('span'));
  const tg = s.target ? tipOn(el('div', { class: 'tg', style: `left:${Math.min(100, (s.target / mx) * 100)}%` }), () => tipBox('기준일 목표', [[swc('tgt'), '누적 확보 목표', `${f1(s.target)} FTE`], [null, '달성률', pct(s.rate)]])) : null;
  $('#flow').replaceChildren(
    row('Baseline', s.base, 'FTE', 'var(--line-2)', DRILL.found), conv(`↓ 발굴률 ${pct(s.base ? s.found / s.base : null)}`),
    row('발굴', s.found, 'FTE', 'var(--found)', DRILL.found), conv(`↓ 진척률 ${pct(s.prog)} · 세로선: 기준일 목표 ${f1(s.target)} (달성 ${pct(s.rate)})`),
    row('확보', s.secured, 'FTE', 'var(--accent)', DRILL.secured, tg), conv(`↓ 실현율 ${pct(s.secured ? s.realized / s.secured : null)}`),
    row('실현', s.realized, '명', 'var(--real)', DRILL.realized, null, f0));
}

function pipe() {
  const by = d3.group(VIEW, (r) => r.idx), mx = d3.max([1, 2, 3, 4, 5], (i) => (by.get(i) ?? []).length) || 1;
  $('#pipe').replaceChildren(...[1, 2, 3, 4, 5].map((i) => {
    const g = by.get(i) ?? [], dl = g.filter((r) => r.delayed).length, dn = g.filter((r) => r.done).length;
    return el('button', { class: 'stg', type: 'button', 'aria-label': `S${i} 과제 ${g.length}건 보기`, onclick: () => openTasks(`S${i} ${STAGES[i - 1]} 단계 과제`, g.slice().sort(prioSort), { sub: dots(`과제 ${g.length}건`, `발굴 ${f1(sum(g, 'found_fte'))} FTE`, `지연 ${dl}건`) }) },
      el('span', { class: 'nm' }, el('b', {}, `S${i}`), STAGES[i - 1]),
      el('span', { class: 'col' }, el('div', { style: `height:${(g.length / mx) * 100}%;background:${stColor(i)}` })),
      el('span', {}, el('div', { class: 'n' }, g.length), el('div', { class: 'f' }, `${f1(sum(g, 'found_fte'))} FTE`),
        el('div', { class: 'd' }, dl ? el('span', { class: 'c-bad' }, `지연 ${dl}`) : i === 5 && dn ? el('span', { class: 'c-ok' }, `완료 ${dn}`) : '')));
  }));
}

// 누적 확보 목표(계단)와 주별 확보 실적(선). 기준일에서 목표와 실적의 차이를 표시하고, 월을 누르면 그달 목표가 도래하는 과제를 띄움
function curve() {
  const box = $('#curve'); box.replaceChildren();
  const W = box.clientWidth; if (!W) return;
  const rs = VIEW.filter((r) => r.s3_date || r.s5_date);
  if (!rs.length) { box.append(el('p', { class: 'empty' }, '목표일정이 입력된 과제가 없습니다.')); return; }
  const asof = pDate(state.asof), fmt = d3.utcFormat('%Y-%m-%d');
  const actual = seriesFor(() => true).filter((p) => p.sec != null).map((p) => ({ d: pDate(p.end), v: p.sec }));
  const firstT = d3.min(rs, (r) => pDate(r.s3_date || r.s4_date || r.s5_date));
  const lo = d3.utcMonth.offset(d3.utcMonth.floor(d3.min([firstT, actual[0]?.d ?? asof])), -1);
  const hi = d3.utcMonth.ceil(d3.utcDay.offset(d3.max([...rs.map((r) => pDate(r.s5_date || r.s4_date || r.s3_date)), asof]), 1));
  const months = d3.utcMonth.range(lo, hi).map((m) => {
    const e = fmt(d3.utcDay.offset(d3.utcMonth.offset(m, 1), -1)), s0 = fmt(m);
    const due = rs.filter((r) => [3, 4, 5].some((k) => r[`s${k}_date`] && r[`s${k}_date`] >= s0 && r[`s${k}_date`] <= e));
    return { m, e, v: sum(rs, (r) => targetAt(r, e)), due };
  });
  const s = statFor(() => true);
  const H = 270, mg = { t: 24, r: 70, b: 26, l: 36 };
  const x = d3.scaleUtc([lo, hi], [mg.l, W - mg.r]);
  const y = d3.scaleLinear([0, Math.max(s.found, s.secured) * 1.05 || 1], [H - mg.b, mg.t]).nice();
  const svg = d3.select(box).append('svg').attr('width', '100%').attr('height', H).attr('viewBox', `0 0 ${W} ${H}`);
  svg.append('rect').attr('x', x(asof)).attr('y', mg.t).attr('width', Math.max(0, W - mg.r - x(asof))).attr('height', H - mg.t - mg.b).attr('fill', 'var(--soft)');
  svg.append('text').attr('x', (x(asof) + W - mg.r) / 2).attr('y', mg.t + 14).attr('text-anchor', 'middle').attr('style', 'fill:var(--muted)').text('앞으로 도래할 목표');
  svg.append('g').attr('class', 'gl').selectAll('line').data(y.ticks(4)).join('line').attr('x1', mg.l).attr('x2', W - mg.r).attr('y1', y).attr('y2', y);
  svg.append('g').selectAll('text').data(y.ticks(4)).join('text').attr('x', mg.l - 6).attr('y', y).attr('dy', '0.32em').attr('text-anchor', 'end').text(d3.format(','));
  svg.append('g').attr('class', 'axis').attr('transform', `translate(0,${H - mg.b})`)
    .call(d3.axisBottom(x).ticks(Math.max(2, Math.floor((W - mg.l - mg.r) / 70))).tickFormat(d3.utcFormat('%y.%m')).tickSizeOuter(0)).attr('font-size', null).attr('font-family', null);
  svg.append('line').attr('x1', mg.l).attr('x2', W - mg.r).attr('y1', y(s.found)).attr('y2', y(s.found)).attr('stroke', 'var(--muted)').attr('stroke-dasharray', '4 4');
  svg.append('text').attr('x', W - mg.r + 4).attr('y', y(s.found)).attr('dy', '0.32em').attr('style', 'fill:var(--ink-2)').text(`발굴 ${f1(s.found)}`);
  const step = months.map((d) => ({ x: d.m, v: d.v })).concat([{ x: hi, v: months[months.length - 1].v }]);
  svg.append('path').datum(step).attr('fill', 'none').attr('stroke', 'var(--ink-2)').attr('stroke-width', 2).attr('d', d3.line().curve(d3.curveStepAfter).x((d) => x(d.x)).y((d) => y(d.v)));
  if (actual.length) {
    svg.append('path').datum(actual).attr('fill', 'none').attr('stroke', 'var(--accent)').attr('stroke-width', 2.4).attr('stroke-linejoin', 'round').attr('d', d3.line().x((d) => x(d.d)).y((d) => y(d.v)));
    svg.append('circle').attr('cx', x(asof)).attr('cy', y(s.secured)).attr('r', 5).attr('fill', 'var(--accent)');
  }
  svg.append('line').attr('x1', x(asof)).attr('x2', x(asof)).attr('y1', mg.t - 6).attr('y2', H - mg.b).attr('stroke', 'var(--ink)');
  svg.append('text').attr('x', x(asof)).attr('y', mg.t - 10).attr('text-anchor', 'middle').attr('style', 'fill:var(--ink);font-weight:700').text('기준일');
  // 기준일의 목표·실적 차이
  if (s.target) {
    const gap = s.secured - s.target, col = gap >= 0 ? 'var(--ok)' : 'var(--bad)', gx = x(asof) - 10;
    svg.append('path').attr('d', `M${gx + 4},${y(s.target)}H${gx}V${y(s.secured)}H${gx + 4}`).attr('fill', 'none').attr('stroke', col).attr('stroke-width', 2);
    svg.append('text').attr('x', x(asof) + 8).attr('y', y(s.secured) + (gap >= 0 ? -12 : 20)).attr('class', 'halo').attr('style', `fill:${col};font-weight:700;font-size:12.5px`)
      .text(`목표 대비 ${gap >= 0 ? '초과' : '부족'} ${f1(Math.abs(gap))} FTE`);
  }
  // 월 단위 마우스·클릭 영역
  const band = svg.append('g');
  months.forEach((d) => {
    const x0 = x(d.m), x1 = x(d3.utcMonth.offset(d.m, 1));
    const act = actual.filter((a) => a.d <= pDate(d.e)).pop();
    const rect = band.append('rect').attr('x', x0).attr('y', mg.t).attr('width', Math.max(1, x1 - x0)).attr('height', H - mg.t - mg.b).attr('fill', 'var(--accent)').attr('fill-opacity', 0).style('cursor', d.due.length ? 'pointer' : 'default')
      .on('mouseenter', function () { d3.select(this).attr('fill-opacity', 0.08); }).on('mouseleave', function () { d3.select(this).attr('fill-opacity', 0); })
      .on('click', () => { if (d.due.length) openTasks(`${d3.utcFormat('%Y년 %m월')(d.m)} 목표 도래 과제`, d.due.slice().sort(prioSort), { sub: `이달 S3·S4·S5 목표일이 있는 과제 ${d.due.length}건 · 월말 누적 목표 ${f1(d.v)} FTE`, extra: [['이달 도래 단계', (r) => [3, 4, 5].filter((k) => r[`s${k}_date`] && r[`s${k}_date`].slice(0, 7) === d.e.slice(0, 7)).map((k) => `S${k} ${r[`s${k}_date`]} · 누적 ${fnum(r[`s${k}_fte`])} FTE`).join(', ')]] }); });
    tipOn(rect.node(), () => tipBox(d3.utcFormat('%Y년 %m월')(d.m), [[swc('step'), '월말 누적 목표', `${f1(d.v)} FTE`], act && d.m <= asof ? [swc('line'), '확보 실적(주말 기준)', `${f1(act.v)} FTE`] : null,
      [null, '이달 목표 도래 과제', `${d.due.length}건`]], d.due.length ? '클릭: 도래 과제 목록' : null));
  });
}

function watch() {
  const w = VIEW.filter((r) => r.priority === '상' || r.alerts.length).sort(prioSort).slice(0, 12);
  const t = $('#watch');
  if (!w.length) { t.replaceChildren(el('tbody', {}, el('tr', {}, el('td', { class: 'empty' }, '우선 관리 대상 과제가 없습니다.')))); return; }
  t.replaceChildren(el('thead', {}, el('tr', {}, ['우선', '과제명', '팀 · 그룹', '단계', '경보', '확보 / 기준일 목표', '진척률'].map((h, i) => el('th', { class: i >= 5 ? 'num' : '' }, h)))),
    el('tbody', {}, w.map((r) => el('tr', { class: 'clk', onclick: () => openTask(r.id) },
      el('td', {}, prioBadge(r)), el('td', { class: 'task-nm' }, r.task), el('td', {}, `${r.team ?? '—'} · ${r.group ?? '—'}`), el('td', {}, stchip(r)),
      el('td', {}, r.alerts.length ? alertChips(r) : '—'), el('td', { class: 'num' }, `${f1(r.secured_fte ?? 0)} / ${r.target_now ? f1(r.target_now) : '—'}`), el('td', { class: 'num' }, pct(r.prog))))));
}

function mgmt() {
  const g = d3.group(VIEW, (r) => r.mgmt), mx = d3.max(Object.keys(MGMT), (k) => (g.get(k) ?? []).length) || 1;
  $('#mgmt').replaceChildren(...Object.keys(MGMT).map((k) => {
    const v = g.get(k) ?? [], kp = v.filter((r) => r.kpi_prog != null);
    const open = () => openTasks(`${k} ${MGMT[k]} 지표 과제`, v.slice().sort(prioSort), { extra: k === 'P' ? [] : [['KPI', (r) => (r.kpi === '기타' ? r.kpi_etc ?? '기타' : r.kpi) ?? '—'], ['AS-IS → TO-BE', (r) => `${fnum(r.kpi_asis)} → ${fnum(r.kpi_tobe)} ${r.unit ?? ''}`], ['확보 / 발굴 효과', (r) => `${fnum(r.kpi_secured)} / ${fnum(r.kpi_found)}`, 1], ['KPI 진척', (r) => pct(r.kpi_prog), 1]] });
    return el('div', { class: 'mg-row', role: 'button', tabindex: 0, onclick: open, title: '클릭: 과제 목록' }, el('b', {}, `${k} ${MGMT[k]}`), el('i', { style: `width:${Math.max(1, (v.length / mx) * 100)}%` }),
      el('span', {}, `${v.length}건 · 발굴 ${f1(sum(v, 'found_fte'))}`, k === 'P' ? '' : ` · KPI 진척 ${kp.length ? pct(d3.mean(kp, (r) => r.kpi_prog)) : '—'}`));
  }));
}

// ── 렌더: 과제 맵 (팀 › 그룹 › 과제 박스) ─────────────────────
function stageDots(r) {
  return el('span', { class: 'sdots', title: `${r.stage ?? '—'} ${STAGES[r.idx - 1] ?? ''}` },
    [1, 2, 3, 4, 5].map((k) => el('i', { class: r.done || k < r.idx ? 'on' : k === r.idx ? (r.delayed ? 'cur late' : 'cur') : '' })), el('span', {}, r.stage ?? '—'));
}
function taskTile(r) {
  const st = tileState(r), f = r.found_fte || 0;
  const node = el('div', { class: `tt s-${st}`, role: 'button', tabindex: 0, onclick: () => openTask(r.id), onkeydown: (e) => { if (e.key === 'Enter') openTask(r.id); } },
    el('div', { class: 'tt-top' }, prioBadge(r), el('span', { class: 'tt-name' }, r.task), stageDots(r)),
    el('div', { class: 'tt-bar' }, el('i', { class: 'fill', style: `width:${Math.min(100, (r.prog ?? 0) * 100)}%` }), r.target_now && f ? el('i', { class: 'tick', style: `left:${Math.min(100, (r.target_now / f) * 100)}%` }) : null),
    el('div', { class: 'tt-nums' }, el('span', {}, '확보 ', el('b', {}, f1(r.secured_fte ?? 0)), ` / 발굴 ${f1(f)} FTE`), el('span', {}, '진척 ', el('b', {}, pct(r.prog)))),
    el('div', { class: 'tt-foot' }, el('span', { class: 'tt-state' }, TS[st][0] + (st === 'delay' ? ` ${r.late_days}일` : '')),
      el('span', {}, r.target_now ? `목표 ${f1(r.target_now)}` : '목표 도래 전'), r.delta ? deltaEl(r.delta) : null, r.realized ? el('span', {}, `실현 ${r.realized}명`) : null,
      r.kpi_prog != null ? el('span', {}, `KPI ${pct(r.kpi_prog)}`) : null, r.alerts.some((a) => a.k === 'miss') ? el('span', { class: 'alert miss' }, r.alerts.find((a) => a.k === 'miss').t) : null));
  return tipOn(node, () => taskTip(r));
}
function renderMap() {
  const box = $('#map'), teams = [...new Set(VIEW.map((r) => r.team ?? '미지정'))].sort(ko);
  if (!teams.length) { box.replaceChildren(el('p', { class: 'empty' }, '조건에 맞는 과제가 없습니다.')); return; }
  box.replaceChildren(...teams.map((t) => {
    const u = teamUnit(t), s = statFor(u.fn);
    const groups = [...new Set(s.rows.map((r) => r.group ?? '미지정'))].sort(ko);
    const kpi = el('div', { class: 'mt-kpi' }, tipOn(bullet(s, s.found), () => bulletTip(s, t)), ratePill(s.rate), deltaEl(s.delta, ' FTE'));
    return el('section', { class: 'mteam' },
      el('header', { class: 'mt-hd', role: 'button', tabindex: 0, onclick: () => openOrg(u) },
        el('div', {}, el('h3', {}, t), el('div', { class: 'mt-sub' }, dots(`과제 ${s.n}`, `확보 ${f1(s.secured)} / 목표 ${f1(s.target)} · 발굴 ${f1(s.found)} FTE`, `진척 ${pct(s.prog)}`, s.delayed ? el('span', { class: 'c-bad' }, `지연 ${s.delayed}`) : null))), kpi),
      el('div', { class: 'mt-groups' }, groups.map((g) => {
        const gu = groupUnit(t, g), gs = statFor(gu.fn);
        return el('div', { class: 'mgroup' },
          el('div', { class: 'mg-hd', role: 'button', tabindex: 0, onclick: () => openOrg(gu) }, el('b', {}, g), el('span', { class: 'pc-stat', style: 'font-size:12px' }, `${gs.n}건`, ratePill(gs.rate))),
          gs.rows.slice().sort(prioSort).map(taskTile));
      })));
  }));
}

// ── 렌더: 조직별 현황 ───────────────────────────────────────
function renderOrg() {
  const teams = [...new Set(VIEW.map((r) => r.team ?? '미지정'))].sort(ko);
  if (state.team) state.open.add(state.team);
  const tStats = teams.map((t) => ({ u: teamUnit(t), s: statFor(teamUnit(t).fn) }));
  const mx = d3.max(tStats, (d) => d.s.found) || 1;
  const head = ['조직', '과제 · 단계 분포', 'FTE (확보 / 목표 · 발굴)', '달성률', '전주 대비', '추이(8주)', '실현(명)', '지연', '우선 상', '경보'];
  const NUMS = new Set([3, 4, 6, 7, 8, 9]);
  const row = (cls, nameCell, u, s, scale = mx) => {
    const pts = seriesFor(u.fn);
    return el('tr', { class: cls, onclick: (e) => { if (!e.target.closest('.tog')) openOrg(u); } }, nameCell,
      el('td', {}, el('div', { style: 'display:flex;align-items:center;gap:8px' }, el('b', { style: 'min-width:1.6em;text-align:right' }, s.n),
        tipOn(el('div', { class: 'stbar' }, s.stages.map((n, i) => (n ? el('i', { style: `width:${(n / s.n) * 100}%;background:${stColor(i + 1)}` }) : null))), () => stageTip(s.stages, u.name)))),
      el('td', {}, fteCell(s, scale, u.name)), el('td', { class: 'num' }, ratePill(s.rate)), el('td', { class: 'num' }, deltaEl(s.delta)),
      el('td', {}, tipOn(el('div', {}, sparkline(pts, 100, 28)), () => trendTip(pts, u.name))), el('td', { class: 'num' }, f0(s.realized)),
      el('td', { class: 'num' + (s.delayed ? ' c-bad' : '') }, s.delayed), el('td', { class: 'num' }, s.high), el('td', { class: 'num' + (s.alertN ? ' c-bad' : '') }, s.alertN));
  };
  const body = [];
  for (const { u, s } of tStats) {
    const t = u.team, open = state.open.has(t);
    const tog = el('button', { class: 'tog', type: 'button', 'aria-expanded': String(open), onclick: () => { open ? state.open.delete(t) : state.open.add(t); renderOrg(); } }, el('span', { class: 'car' }, '▸'), t);
    body.push(row('team', el('td', {}, tog), u, s));
    if (open) for (const g of [...new Set(s.rows.map((r) => r.group ?? '미지정'))].sort(ko)) { const gu = groupUnit(t, g); body.push(row('grp', el('td', {}, g), gu, statFor(gu.fn))); }
  }
  if (tStats.length > 1) { const all = { name: '전체', fn: () => true }, s = statFor(all.fn); body.push(row('total', el('td', {}, '전체'), all, s, s.found)); }
  $('#score').replaceChildren(el('thead', {}, el('tr', {}, head.map((h, i) => el('th', { class: NUMS.has(i) ? 'num' : '' }, h)))),
    el('tbody', {}, body.length ? body : el('tr', {}, el('td', { colspan: head.length, class: 'empty' }, '조건에 맞는 과제가 없습니다.'))));
  const allOpen = teams.every((t) => state.open.has(t));
  const bx = $('#b-expand'); bx.textContent = allOpen ? '그룹 모두 접기' : '그룹 모두 펼치기';
  bx.onclick = () => { if (allOpen) state.open.clear(); else teams.forEach((t) => state.open.add(t)); renderOrg(); };

  $('#prio').replaceChildren(...units().map((u) => {
    const s = statFor(u.fn), top = s.rows.slice().sort(prioSort).slice(0, 5);
    return el('article', { class: 'card pc' },
      el('header', { class: 'card-hd', onclick: () => openOrg(u), title: '클릭: 조직 상세' }, el('div', {}, el('h3', {}, u.sub ? `${u.sub} · ${u.name}` : u.name),
        el('div', { class: 'pc-stat', style: 'margin-top:4px' }, dots(`과제 ${s.n}`, `우선 상 ${s.high}`, el('span', { class: s.alertN ? 'c-bad' : '' }, `경보 ${s.alertN}`)))),
        el('div', { class: 'pc-stat' }, ratePill(s.rate), deltaEl(s.delta, ' FTE'))),
      top.length ? el('ul', { class: 'pc-list' }, top.map((r) => tipOn(el('li', { onclick: () => openTask(r.id) },
        prioBadge(r),
        el('div', {}, el('div', { class: 't' }, r.task), el('div', { class: 's' }, u.group ? '' : `${r.group ?? '—'} ·`, stchip(r), alertChips(r))),
        el('div', { class: 'v' }, el('b', {}, f1(r.secured_fte ?? 0)), ` / ${r.target_now ? f1(r.target_now) : '—'}`, el('div', { style: 'font-size:11.5px;color:var(--ink-2)' }, '확보 / 목표'))), () => taskTip(r)))) : el('p', { class: 'empty' }, '과제 없음'));
  }));
}

// ── 렌더: 과제 관리 ─────────────────────────────────────────
const COLS = [
  ['priority', '우선'], ['task', '과제명'], ['team', '팀'], ['group', '그룹'], ['kind', '구분'], ['stage', '단계'], ['status', '상태'], ['alerts', '경보'],
  ['found_fte', '발굴 FTE', 'num'], ['target_now', '기준일 목표', 'num'], ['secured_fte', '확보 FTE', 'num'], ['prog', '진척률', 'num'], ['delta', '전주 대비', 'num'], ['realized', '실현(명)', 'num'], ['issues', '점검', 'num'],
];
const sortVal = (r, k) => (k === 'priority' ? PR[r.priority] ?? 3 : k === 'alerts' ? -r.alerts.length : k === 'issues' ? r.issues.length : k === 'stage' ? r.idx : r[k]);
function renderTasks() {
  $('#viewseg').replaceChildren(...[['list', '목록'], ['xl', '엑셀 형식']].map(([k, t]) => el('button', { type: 'button', 'aria-pressed': String(state.view === k), onclick: () => { state.view = k; render(); } }, t)));
  const q = state.q.trim().toLowerCase();
  $('#stchips').replaceChildren(...[null, 1, 2, 3, 4, 5].map((i) => el('button', { class: 'chip', type: 'button', 'aria-pressed': String(state.stage === i), onclick: () => { state.stage = i; render(); } },
    i ? `S${i}` : '전체', el('b', {}, i ? VIEW.filter((r) => r.idx === i).length : VIEW.length))));
  let v = VIEW.filter((r) => (!state.stage || r.idx === state.stage) && (!q || [r.task, r.cat, r.team, r.group, r.kpi, r.tgt_ids, r.dev_ids, r.dev_dept].join(' ').toLowerCase().includes(q)));
  if (state.view === 'xl') { renderSheetView(v.slice().sort((a, b) => ko(a.team, b.team) || ko(a.group, b.group) || prioSort(a, b))); return; }
  const { key, dir } = state.sort;
  v = v.slice().sort((a, b) => { const x = sortVal(a, key), y = sortVal(b, key); return (x == null) - (y == null) || (typeof x === 'number' ? x - y : ko(x, y)) * dir || prioSort(a, b); });
  const mx = d3.max(VIEW, (r) => r.found_fte) || 1;
  const cell = (r, k) => {
    if (k === 'priority') return prioBadge(r);
    if (k === 'task') return el('span', { class: 'task-nm' }, r.task);
    if (k === 'stage') return stchip(r);
    if (k === 'status') return statusPill(r);
    if (k === 'alerts') return r.alerts.length ? alertChips(r) : '—';
    if (k === 'found_fte') return [el('span', { class: 'cellbar', style: `width:${Math.round(((r.found_fte ?? 0) / mx) * 56)}px` }), f1(r.found_fte)];
    if (k === 'target_now') return r.target_now ? f1(r.target_now) : '—';
    if (k === 'secured_fte') return el('b', { class: r.lagging ? 'c-bad' : '' }, f1(r.secured_fte ?? 0));
    if (k === 'prog') return pct(r.prog);
    if (k === 'delta') return deltaEl(r.delta);
    if (k === 'realized') return r.realized ? f0(r.realized) : '—';
    if (k === 'issues') return r.issues.length ? el('span', { class: 'c-bad' }, `${r.issues.length}건`) : el('span', { class: 'c-ok' }, '정상');
    return r[k] ?? '—';
  };
  $('#tbl').className = '';
  $('#tbl').replaceChildren(
    el('thead', {}, el('tr', {}, COLS.map(([k, h, c]) => el('th', { class: c ?? '', 'aria-sort': key === k ? (dir > 0 ? 'ascending' : 'descending') : 'none' },
      el('button', { type: 'button', onclick: () => { state.sort = { key: k, dir: key === k ? -dir : 1 }; render(); } }, h))))),
    el('tbody', {}, v.length ? v.map((r) => el('tr', { class: 'clk', onclick: () => openTask(r.id) }, COLS.map(([k, , c]) => el('td', { class: c ?? '' }, cell(r, k)))))
      : el('tr', {}, el('td', { colspan: COLS.length, class: 'empty' }, '조건에 맞는 과제가 없습니다.'))));
}
// 엑셀 형식: 업로드 양식과 같은 컬럼 순서(팀·과제명은 왼쪽 고정), 자동 계산 칸은 회색, 미입력 필수 칸은 빨간 바탕
function renderSheetView(v) {
  const rest = FIELDS.filter((f) => f[0] !== 'team' && f[0] !== 'task');
  const spans = SEC.map((s, si) => [s, rest.filter((f) => f[2] === si).length]);
  const val = (r, [k, , , t]) => {
    const x = t === 'unitcol' ? r.unit : r[k];
    return x == null || x === '' ? '' : t === 'n' || t === 'auto' ? fnum(x) : String(x);
  };
  $('#tbl').className = 'xl';
  $('#tbl').replaceChildren(
    el('thead', {}, el('tr', {}, el('th', { class: 'fix1' }, ''), el('th', { class: 'fix2' }, ''), spans.map(([s, n]) => el('th', { colspan: n }, s))),
      el('tr', {}, el('th', { class: 'fix1' }, '팀'), el('th', { class: 'fix2' }, '과제명(중분류)'), rest.map(([, h, , t]) => el('th', { class: t === 'auto' || t === 'unitcol' ? 'auto' : '' }, shortName(h))))),
    el('tbody', {}, v.length ? v.map((r) => {
      const miss = new Set(r.issues.filter((i) => i.level === '누락').map((i) => i.key));
      return el('tr', { class: 'clk', onclick: () => openTask(r.id) }, el('td', { class: 'fix1' }, r.team ?? ''), el('td', { class: 'fix2 task-nm', title: r.task }, r.task),
        rest.map((f) => { const x = val(r, f); return el('td', { class: (miss.has(f[0]) ? 'miss ' : '') + (f[3] === 'ta' ? 'long' : f[3] === 'n' || f[3] === 'auto' ? 'num' : ''), title: f[3] === 'ta' ? x : null }, x); }));
    }) : el('tr', {}, el('td', { colspan: rest.length + 2, class: 'empty' }, '조건에 맞는 과제가 없습니다.'))));
}

// ── 렌더: 추진 일정 ─────────────────────────────────────────
function renderGantt() {
  const box = $('#gantt'); box.replaceChildren();
  const cw = box.parentElement.clientWidth; if (!cw) return;
  const W = Math.max(720, cw), asof = pDate(state.asof);
  const rows = VIEW.filter((r) => r.plan_start && r.plan_end).sort((a, b) => ko(a.team, b.team) || ko(a.group, b.group) || ko(a.plan_start, b.plan_start));
  if (!rows.length) { box.append(el('p', { class: 'empty' }, '계획 일정이 입력된 과제가 없습니다.')); return; }
  const items = []; let lt = null, lg = null;
  for (const r of rows) {
    if (r.team !== lt) { items.push({ head: r.team ?? '미지정', lv: 0 }); lt = r.team; lg = null; }
    if (r.group !== lg) { items.push({ head: r.group ?? '미지정', lv: 1 }); lg = r.group; }
    items.push(r);
  }
  const LW = Math.min(260, W * 0.3), RH = 26, top = 28, H = top + items.length * RH + 8;
  const lo = d3.utcMonth.floor(d3.min(rows, (r) => pDate(r.act_start && r.act_start < r.plan_start ? r.act_start : r.plan_start)));
  const hi = d3.utcMonth.ceil(d3.max([...rows.map((r) => pDate(r.act_end && r.act_end > r.plan_end ? r.act_end : r.plan_end)), asof]));
  const x = d3.scaleUtc([lo, hi], [LW + 8, W - 30]);
  const svg = d3.select(box).append('svg').attr('width', W).attr('height', H).attr('viewBox', `0 0 ${W} ${H}`);
  const ticks = x.ticks(d3.utcMonth.every(W > 1100 ? 2 : 3));
  svg.append('g').attr('class', 'gl').selectAll('line').data(ticks).join('line').attr('x1', x).attr('x2', x).attr('y1', top - 6).attr('y2', H);
  svg.append('g').selectAll('text').data(ticks.filter((t) => Math.abs(x(t) - x(asof)) > 34)).join('text').attr('x', x).attr('y', 15).attr('text-anchor', 'middle').text(d3.utcFormat('%y.%m'));
  items.forEach((it, i) => {
    const y0 = top + i * RH;
    if (it.head) {
      if (it.lv === 0) svg.append('rect').attr('x', 0).attr('y', y0 + 1).attr('width', W).attr('height', RH - 2).attr('fill', 'var(--soft)');
      svg.append('text').attr('x', it.lv ? 14 : 6).attr('y', y0 + RH / 2 + 4).attr('style', `fill:var(--ink);font-weight:700;font-size:${it.lv ? 12 : 13}px`).text(it.head);
      return;
    }
    const r = it, g = svg.append('g').style('cursor', 'pointer').on('click', () => openTask(r.id));
    g.append('rect').attr('x', 0).attr('y', y0).attr('width', W).attr('height', RH).attr('fill', 'transparent');
    const name = g.append('text').attr('x', 28).attr('y', y0 + RH / 2 + 4).attr('style', 'fill:var(--ink)').text(r.task);
    while (name.node().getComputedTextLength() > LW - 32 && name.text().length > 3) name.text(name.text().slice(0, -2) + '…');
    tipOn(name.node(), () => taskTip(r));
    const plan = g.append('rect').attr('x', x(pDate(r.plan_start))).attr('y', y0 + 5).attr('width', Math.max(2, x(pDate(r.plan_end)) - x(pDate(r.plan_start)))).attr('height', RH - 10).attr('rx', 4).attr('fill', 'var(--found)').attr('fill-opacity', 0.75);
    tipOn(plan.node(), () => tipBox(r.task, [[swc('found'), '계획 기간', `${r.plan_start} ~ ${r.plan_end}`]], '클릭: 상세'));
    if (r.act_start) {
      const e = r.act_end ? pDate(r.act_end) : asof, c = r.done ? 'var(--ok)' : r.delayed ? 'var(--bad)' : 'var(--accent)';
      const act = g.append('rect').attr('x', x(pDate(r.act_start))).attr('y', y0 + 10).attr('width', Math.max(2, x(e) - x(pDate(r.act_start)))).attr('height', RH - 20).attr('rx', 3).attr('fill', c);
      tipOn(act.node(), () => tipBox(r.task, [[swc('', `background:${c}`), `실적 기간 (${r.done ? '완료' : r.delayed ? '지연' : '진행 중'})`, `${r.act_start} ~ ${r.act_end ?? '기준일'}`]], '클릭: 상세'));
    }
    [1, 2, 3, 4, 5].forEach((k) => {
      const d = r[`s${k}_date`]; if (!d) return;
      const cx = x(pDate(d)), cy = y0 + RH / 2, past = r.done || k < r.idx, late = k === r.idx && r.delayed;
      const dm = g.append('path').attr('d', `M${cx},${cy - 6}L${cx + 6},${cy}L${cx},${cy + 6}L${cx - 6},${cy}Z`)
        .attr('fill', past ? stColor(3) : 'var(--panel)').attr('stroke', late ? 'var(--bad)' : 'var(--st5)').attr('stroke-width', late ? 2.2 : 1.3);
      tipOn(dm.node(), () => tipBox(`S${k} ${STAGES[k - 1]} 목표일`, [[null, '목표일', d], [null, '상태', past ? '지난 단계' : late ? `현재 단계 · 목표일 ${r.late_days}일 경과` : k === r.idx ? '현재 단계' : '남은 단계'],
        k >= 3 ? [null, '누적 확보 목표', `${fnum(r[`s${k}_fte`])} FTE`] : null], r.task));
    });
  });
  svg.append('line').attr('x1', x(asof)).attr('x2', x(asof)).attr('y1', top - 6).attr('y2', H).attr('stroke', 'var(--ink)');
  svg.append('text').attr('x', x(asof)).attr('y', 15).attr('text-anchor', 'middle').attr('style', 'fill:var(--ink);font-weight:700').text('기준일');
}

// ── 렌더: 데이터 점검 ───────────────────────────────────────
function renderCheck() {
  const all = VIEW.flatMap((r) => r.issues.map((i) => ({ ...i, id: r.id, task: r.task, team: r.team, group: r.group, stage: r.stage, rk: `${i.level}|${i.rule}` })));
  const rules = d3.rollups(all, (v) => v.length, (d) => d.rk).sort((a, b) => b[1] - a[1]);
  $('#rules').replaceChildren(el('button', { class: 'chip', type: 'button', 'aria-pressed': String(!state.rule), onclick: () => { state.rule = null; render(); } }, '전체', el('b', {}, all.length)),
    ...rules.map(([k, n]) => el('button', { class: 'chip', type: 'button', 'aria-pressed': String(state.rule === k), onclick: () => { state.rule = k; render(); } }, k.replace('|', ' · '), el('b', {}, n))));
  const v = all.filter((d) => !state.rule || d.rk === state.rule);
  $('#issues').replaceChildren(el('thead', {}, el('tr', {}, ['팀', '그룹', '과제명', '단계', '구분', '항목', '규칙'].map((h) => el('th', {}, h)))),
    el('tbody', {}, v.length ? v.map((d) => el('tr', { class: 'clk', onclick: () => openTask(d.id) }, el('td', {}, d.team ?? '—'), el('td', {}, d.group ?? '—'), el('td', { class: 'task-nm' }, d.task), el('td', {}, d.stage ?? '—'),
      el('td', {}, el('span', { class: 'pill ' + (d.level === '누락' ? 'bad' : 'warn') }, d.level)), el('td', {}, d.field), el('td', {}, d.rule)))
      : el('tr', {}, el('td', { colspan: 7, class: 'empty' }, '점검에서 걸린 항목이 없습니다.'))));
  const c = $('#c-changes'); c.hidden = !state.server;
  if (state.server) {
    const OP = { add: '추가', save: '수정', delete: '삭제', import: '엑셀 반영' };
    $('#changes').replaceChildren(el('thead', {}, el('tr', {}, ['일시', '입력자', '작업', '내용'].map((h) => el('th', {}, h)))),
      el('tbody', {}, state.changes.length ? state.changes.map((x) => el('tr', {}, el('td', {}, x.at), el('td', {}, x.editor), el('td', {}, OP[x.op] ?? x.op),
        el('td', { class: 'wrap' }, x.op === 'import' ? (x.teams ?? []).map((t) => `${t.team} (추가 ${t.added} · 변경 ${t.changed} · 삭제 ${t.removed})`).join(', ')
          : `${x.team ?? ''} · ${x.task ?? ''}${x.fields?.length ? ` — ${x.fields.map((k) => shortName(FIELDS.find((f) => f[0] === k)?.[1] ?? k)).join(', ')}` : ''}`)))
        : el('tr', {}, el('td', { colspan: 4, class: 'empty' }, '변경 이력이 없습니다.'))));
  }
}
async function loadChanges() {
  try { const r = await api('api/changes'); state.changes = r.ok ? r.data.changes ?? [] : []; } catch (e) { state.changes = []; }
  if (state.tab === 'check') renderCheck();
}

// ── 확인 대화상자 ───────────────────────────────────────────
function modal({ title, notes, body, actions, wide }) {
  const m = $('#modal');
  HT.hidden = true;
  return new Promise((resolve) => {
    const close = (v) => { m.hidden = true; m.replaceChildren(); resolve(v); };
    const btns = actions.map((a) => el('button', { class: 'btn ' + (a.cls ?? ''), type: 'button', onclick: () => close(a.value ?? null) }, a.label));
    m.replaceChildren(el('div', { class: 'mbox' + (wide ? ' wide' : '') },
      el('header', { class: 'card-hd' }, el('div', {}, el('h2', {}, title), notes ? el('ul', { class: 'notes' }, notes.map((n) => el('li', {}, n))) : null)),
      el('div', { class: 'm-bd' }, body), el('div', { class: 'm-ft' }, btns)));
    m.onkeydown = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(null); } };
    m.hidden = false; btns[btns.length - 1].focus();
  });
}
const editorName = () => ls.get('ax-editor-name') ?? '';
const nameInput = () => { const i = el('input', { type: 'text', maxlength: 30, placeholder: '이름(선택)', 'aria-label': '입력자 이름' }); i.value = editorName(); i.addEventListener('input', () => ls.set('ax-editor-name', i.value.trim())); return i; };

// 저장 요청: 서버가 새 데이터 전체를 돌려주면 화면을 갱신
async function commit(body) {
  let r; try { r = await post('api/commit', { ...body, editor: editorName(), base_version: state.version }); } catch (e) { toast('서버에 연결하지 못했습니다. 서버 PC의 serve.py 창이 켜져 있는지 확인해 주세요.', 'err'); return null; }
  if (r.status === 409) { await loadServer(true); toast('다른 사람이 먼저 수정했습니다. 최신 내용을 불러왔으니 다시 시도해 주세요.', 'err'); return null; }
  if (!r.ok) { toast(r.data?.error ?? `저장하지 못했습니다 (HTTP ${r.status}).`, 'err'); return null; }
  applyStore(r.data); await loadHistory(); showData(true); if (state.tab === 'check') loadChanges();
  return r.data;
}

async function removeTask(r) {
  const ok = await modal({ title: '과제 삭제', notes: [`${r.team} · ${r.group ?? ''} · ${r.task}`, '삭제 전 상태는 data/backup에 자동 보관'], body: el('p', {}, '이 과제를 삭제합니다.'),
    actions: [{ label: '취소' }, { label: '삭제', cls: 'danger', value: true }] });
  if (!ok) return;
  if (await commit({ op: 'delete', id: r.id })) { closeSheets(); toast(`'${r.task}'을(를) 삭제했습니다.`); }
}

// ── 과제 입력 패널 ──────────────────────────────────────────
const isBusy = () => !$('#drawer').hidden || !$('#modal').hidden;
function openEditor(task) {
  if (!canEdit()) { toast(editHint(), 'err'); return; }
  closeSheets();
  const base = state.store.map(normalize);
  const opts = (k, f = () => true) => [...new Set(base.filter(f).map((r) => r[k]).filter(Boolean))].sort(ko);
  const draft = task ? { ...task } : { team: state.team || null, group: state.group || null, stage: 'S1', priority: '중', pkms: '미등록', agent: 'X' };
  const dr = $('#drawer'), lists = el('div');
  const mkList = (id, vals) => el('datalist', { id }, vals.map((v) => el('option', { value: v })));
  const refreshLists = (cur) => lists.replaceChildren(
    mkList('dl-team', opts('team')), mkList('dl-group', opts('group', (r) => !cur.team || r.team === cur.team)), mkList('dl-cat', opts('cat', (r) => !cur.team || r.team === cur.team)),
    mkList('dl-scope', [...new Set(['단일 라인', '복수 라인', '공장 전체', '팀 업무', ...opts('scope')])]), mkList('dl-tgt_dept', opts('tgt_dept')), mkList('dl-dev_dept', opts('dev_dept')),
    mkList('dl-kpi', [...opts('kpi', (r) => !cur.mgmt || r.mgmt === cur.mgmt), '기타']));
  const inputs = new Map();
  const field = ([k, h, , t, choices]) => {
    const id = `ff-${k}`, v = draft[k];
    let input;
    if (t === 'auto') input = el('output', { id });
    else if (t === 'ta') input = el('textarea', { id, rows: 2 });
    else if (t === 'sel') { const cs = [...new Set([...choices, ...(v != null && !choices.includes(v) ? [v] : [])])]; input = el('select', { id }, el('option', { value: '' }, '선택'), cs.map((c) => el('option', { value: c }, k === 'stage' ? `${c} ${STAGES[+c.slice(1) - 1]}` : k === 'mgmt' ? `${c} ${MGMT[c]}` : c))); }
    else input = el('input', { id, type: t === 'n' ? 'number' : t === 'd' ? 'date' : 'text', step: t === 'n' ? 'any' : null, list: t === 'dl' ? `dl-${k}` : null });
    if (t !== 'auto') input.value = v ?? '';
    input.dataset.key = k; inputs.set(k, input);
    const tag = h.includes('[조건]') ? '조건부' : h.includes('[P외]') ? 'P 외 필수' : t === 'auto' ? '자동 계산' : null;
    return el('div', { class: 'ff' + (t === 'ta' ? ' wide' : ''), id: `w-${k}` },
      el('label', { for: id }, shortName(h), h.startsWith('*') && !tag ? el('span', { class: 'req', title: '필수' }, '*') : null, tag ? el('span', { class: 'tag' }, tag) : null),
      input, el('span', { class: 'hint', hidden: true }));
  };
  const sumEl = el('span', { class: 'sum' });
  const read = () => { const raw = { id: task?.id ?? null }; for (const [k, inp] of inputs) if (inp.tagName !== 'OUTPUT') raw[k] = inp.value; return normalize(raw); };
  const update = (changedKey) => {
    const cur = read();
    if (changedKey === 'team' || changedKey === 'mgmt' || !changedKey) refreshLists(cur);
    if (changedKey === 'kpi' && !inputs.get('unit').value) { const same = base.find((r) => r.kpi === cur.kpi && r.unit); if (same) { inputs.get('unit').value = same.unit; cur.unit = same.unit; } }
    for (const [k, inp] of inputs) if (inp.tagName === 'OUTPUT') inp.value = fnum(cur[k]);
    const iss = validate(cur, cur.stage ? +cur.stage.slice(1) : 0), byKey = d3.group(iss, (i) => i.key);
    for (const [k] of inputs) {
      const w = $(`#w-${k}`, dr), list = byKey.get(k), hint = w.querySelector('.hint');
      w.classList.toggle('miss', !!list); hint.hidden = !list; hint.textContent = list ? list.map((i) => (i.level === '누락' ? `미입력 · ${i.rule}` : i.rule)).join(' / ') : '';
    }
    const nMiss = iss.filter((i) => i.level === '누락').length, nChk = iss.length - nMiss;
    sumEl.replaceChildren(iss.length ? el('span', { class: 'c-bad' }, dots(nMiss ? `미입력 ${nMiss}개` : null, nChk ? `확인 필요 ${nChk}개` : null), ' · 저장은 가능') : el('span', { class: 'c-ok' }, '입력 규칙 이상 없음'));
  };
  const close = () => { dr.hidden = true; $('#overlay').hidden = true; dr.replaceChildren(); };
  const save = async () => {
    const cur = read();
    if (!cur.team || !cur.task) { toast('팀과 과제명(중분류)은 꼭 입력해야 합니다.', 'err'); inputs.get(cur.team ? 'task' : 'team').focus(); return; }
    const dup = state.store.find((r) => r.team === cur.team && r.task === cur.task && r.id !== cur.id);
    if (dup) { toast(`${cur.team}에 같은 과제명 '${cur.task}'이(가) 이미 있습니다.`, 'err'); inputs.get('task').focus(); return; }
    const res = await commit({ op: 'save', task: cur });
    if (res) { const saved = res.tasks.find((t) => t.team === cur.team && t.task === cur.task); close(); toast(`'${cur.task}'을(를) 저장했습니다.`); if (saved) openTask(saved.id); }
  };
  dr.replaceChildren(
    el('div', { class: 'dr-hd' }, el('h2', {}, task ? `과제 수정 · ${task.task}` : '과제 추가'), el('button', { type: 'button', onclick: close }, '닫기')),
    el('div', { class: 'dr-bd' }, SEC.map((s, si) => el('fieldset', { class: 'fs' }, el('legend', {}, s), el('div', { class: 'fs-grid' }, FORM_FIELDS.filter((f) => f[2] === si).map(field)))), lists),
    el('div', { class: 'dr-ft' }, sumEl, el('div', { class: 'dt-acts' }, el('label', { class: 'who-in' }, '입력자', nameInput()), el('button', { class: 'btn', type: 'button', onclick: close }, '취소'), el('button', { class: 'btn primary', type: 'button', onclick: save }, '저장'))));
  dr.oninput = (e) => { if (e.target.dataset.key) update(e.target.dataset.key); };
  dr.onchange = (e) => { if (e.target.dataset.key) update(e.target.dataset.key); };
  dr.hidden = false; $('#overlay').hidden = false;
  update(null);
  inputs.get(task ? 'task' : 'team').focus();
}

// ── 엑셀 읽기·반영·내려받기 ─────────────────────────────────
function readWorkbook(buf, name) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  for (const sn of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
    const hi = aoa.findIndex((row) => row.some((c) => norm(c) === norm('과제명(중분류)')));
    if (hi < 0) continue;
    const keys = aoa[hi].map((h) => byHeader.get(norm(h)) ?? null);
    const rows = aoa.slice(hi + 1).filter((row) => row.some((c) => c != null && c !== '')).map((row) => Object.fromEntries(keys.map((k, i) => [k, row[i]]).filter(([k]) => k && !k.startsWith('_'))));
    return { rows, sheet: sn, known: keys.filter(Boolean).length, hasPriority: keys.includes('priority'), unknown: aoa[hi].filter((h, i) => h != null && !keys[i]) };
  }
  throw new Error(`'${name}'에서 '과제명(중분류)' 컬럼을 찾지 못했습니다. 양식의 컬럼명 행이 있는지 확인해 주세요.`);
}

// 팀 단위 교체 계획: 엑셀에 있는 팀의 과제만 엑셀 내용으로 바꾸고, 과제명이 같으면 같은 과제로 봄(ID·우선순위 유지)
function planImport(res) {
  const inc = res.rows.map(normalize), valid = inc.filter((r) => r.team && r.task);
  const base = state.store.map(normalize), tasks = [], dups = [];
  const same = (a, b) => FORM_KEYS.every((k) => (a[k] ?? null) === (b[k] ?? null));
  const summary = [...new Set(valid.map((r) => r.team))].sort(ko).map((team) => {
    const ex = new Map(base.filter((r) => r.team === team).map((r) => [r.task, r]));
    const seen = new Set(); let added = 0, changed = 0, kept = 0;
    for (const r of valid.filter((x) => x.team === team)) {
      if (seen.has(r.task)) { dups.push(`${team} · ${r.task}`); continue; }
      seen.add(r.task);
      const old = ex.get(r.task);
      const t = { ...r, id: old?.id ?? null, priority: res.hasPriority ? r.priority : old?.priority ?? null };
      if (!old) added++; else if (same(old, t)) kept++; else changed++;
      tasks.push(t);
    }
    return { team, total: seen.size, added, changed, kept, removed: [...ex.keys()].filter((k) => !seen.has(k)).length };
  });
  return { summary, tasks, dups, skipped: inc.length - valid.length };
}

async function onUpload(f) {
  let res; try { res = readWorkbook(await f.arrayBuffer(), f.name); } catch (err) { toast(err.message, 'err'); return; }
  const plan = planImport(res), apply = canEdit();
  const body = el('div', { style: 'display:grid;gap:12px' },
    el('p', {}, el('b', {}, f.name), ` · ${res.sheet} 시트 · ${res.rows.length}행 · 인식 컬럼 ${res.known}개`),
    el('div', { class: 'tblwrap', style: 'max-height:300px;border:1px solid var(--line);border-radius:8px' }, el('table', {},
      el('thead', {}, el('tr', {}, ['팀', '엑셀 과제', '추가', '변경', '삭제', '동일'].map((h, i) => el('th', { class: i ? 'num' : '' }, h)))),
      el('tbody', {}, plan.summary.map((s) => el('tr', {}, el('td', {}, el('b', {}, s.team)), el('td', { class: 'num' }, s.total), el('td', { class: 'num' }, s.added),
        el('td', { class: 'num' }, s.changed), el('td', { class: 'num' + (s.removed ? ' c-bad' : '') }, s.removed), el('td', { class: 'num' }, s.kept)))))),
    el('ul', { class: 'notes', style: 'display:grid;gap:3px;color:var(--ink-2)' },
      el('li', {}, '엑셀에 있는 팀만 교체, 다른 팀 과제는 유지'),
      el('li', {}, '삭제: 저장 데이터에는 있으나 엑셀에는 없는 과제'),
      el('li', {}, '반영 후에도 웹에서 과제별 수정 가능'),
      res.hasPriority ? null : el('li', {}, '엑셀에 우선순위 컬럼 없음 → 기존 우선순위 유지'),
      plan.dups.length ? el('li', { class: 'c-bad' }, `같은 팀 안 중복 과제명 ${plan.dups.length}건은 첫 행만 반영: ${plan.dups.slice(0, 3).join(', ')}${plan.dups.length > 3 ? ' 외' : ''}`) : null,
      plan.skipped ? el('li', { class: 'c-bad' }, `팀 또는 과제명이 빈 행 ${plan.skipped}건 제외`) : null,
      res.unknown.length ? el('li', {}, `인식하지 못한 컬럼: ${res.unknown.join(', ')}`) : null,
      apply ? null : el('li', {}, state.server ? '보기 전용 서버라 반영 불가' : '서버 없이 연 화면이라 이 화면에서만 표시, 저장 안 됨')),
    apply ? el('label', { class: 'who-in', style: 'display:flex;gap:8px;align-items:center;font-size:13px;font-weight:600' }, '입력자', nameInput()) : null);
  const choice = await modal({ title: '엑셀 업로드 미리보기', notes: ['팀 단위 교체'], body, wide: true,
    actions: [{ label: '취소' }, apply ? (plan.summary.length ? { label: '반영', cls: 'primary', value: 'apply' } : null) : (state.server ? null : { label: '이 화면에서 보기', cls: 'primary', value: 'local' })].filter(Boolean) });
  if (choice === 'local') setData('local', res.rows, [], null, today(), false, f.name);
  if (choice === 'apply') {
    const r = await commit({ op: 'import', teams: plan.summary.map((s) => s.team), tasks: plan.tasks });
    if (r) toast(`${f.name} 반영 완료 · ${plan.summary.map((s) => `${s.team} 추가 ${s.added}·변경 ${s.changed}·삭제 ${s.removed}`).join(' / ')}`);
  }
}

function exportXlsx() {
  const head = FIELDS.map((f) => f[1]);
  const row1 = FIELDS.map((f, i) => (i === 0 || FIELDS[i - 1][2] !== f[2] ? SEC[f[2]] : null));
  const cell = (r, [k, , , t]) => {
    const v = t === 'unitcol' ? r.unit : r[k];
    if (v == null) return null;
    if (t === 'd') { const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); }
    return v;
  };
  const ws = XLSX.utils.aoa_to_sheet([row1, head, ...state.rows.map((r) => FIELDS.map((f) => cell(r, f)))], { cellDates: true, dateNF: 'yyyy-mm-dd' });
  ws['!merges'] = SEC.map((s, si) => { const cols = FIELDS.map((f, i) => (f[2] === si ? i : -1)).filter((i) => i >= 0); return { s: { r: 0, c: cols[0] }, e: { r: 0, c: cols[cols.length - 1] } }; });
  ws['!cols'] = head.map((h) => ({ wch: Math.max(10, Math.min(28, Math.round(h.length * 1.6))) }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '과제');
  XLSX.writeFile(wb, `AX_과제_${today()}.xlsx`);
}

// ── 이벤트·시작 ─────────────────────────────────────────────
$('#b-upload').addEventListener('click', () => $('#f-file').click());
$('#f-file').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) onUpload(f); });
$('#b-export').addEventListener('click', exportXlsx);
$('#b-add').addEventListener('click', () => openEditor(null));
$('#f-team').addEventListener('change', (e) => { state.team = e.target.value; state.group = ''; render(); });
$('#f-group').addEventListener('change', (e) => { state.group = e.target.value; render(); });
$('#f-kind').addEventListener('change', (e) => { state.kind = e.target.value; render(); });
$('#f-prio').addEventListener('change', (e) => { state.prio = e.target.value; render(); });
$('#f-asof').addEventListener('change', (e) => { if (e.target.value && e.target.validity.valid) { state.asof = e.target.value; render(); } });
$('#b-reset').addEventListener('click', () => { state.team = state.group = state.kind = state.prio = ''; render(); });
$('#q').addEventListener('input', (e) => { state.q = e.target.value; render(); });
$('#overlay').addEventListener('click', () => $('#drawer .dr-hd button')?.click());
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape' || !$('#modal').hidden) return;
  if (!$('#drawer').hidden) $('#drawer .dr-hd button').click();
  else if (SHEETS.length) backSheet();
});
let rt; window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => { if (state.tab === 'overview' || state.tab === 'schedule') render(); }, 150); });

// 다른 사람이 저장하면 20초 안에 반영 (입력 중에는 건너뜀)
setInterval(async () => {
  if (!state.server || document.hidden || isBusy() || state.mode === 'local') return;
  let r; try { r = await api('api/version'); } catch (e) { return; }
  if (r.ok && r.data.version !== state.version) { await loadServer(true); toast('다른 사용자가 저장한 내용을 반영했습니다.'); }
}, 20000);

buildTabs(); buildLegends();
loadServer(false).then((ok) => {
  if (!ok) setData('demo', SAMPLE, SAMPLE_HIST, mondayOf(SAMPLE_ASOF), SAMPLE_ASOF, false);
  const h = location.hash.slice(1);
  setTab(TABS.some(([k]) => k === h) ? h : 'overview');
});
})();
