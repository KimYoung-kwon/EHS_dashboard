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
const store = {
  get: (s, k) => { try { return s.getItem(k); } catch (e) { return null; } },
  set: (s, k, v) => { try { v == null ? s.removeItem(k) : s.setItem(k, v); } catch (e) { /* 저장소를 못 쓰면 매번 다시 입력 */ } },
};
const deltaEl = (v, unit = '') => {
  if (v == null || !isFinite(v)) return el('span', { class: 'delta flat' }, '—');
  const r = Math.round(v * 10) / 10, cls = r > 0 ? 'up' : r < 0 ? 'down' : 'flat';
  return el('span', { class: 'delta ' + cls }, `${r > 0 ? '▲ ' : r < 0 ? '▼ ' : ''}${d3.format(',.1f')(Math.abs(r))}${unit}`);
};
const statusOf = (rate) => (rate == null ? { k: '', t: '목표 전' } : rate >= 1 ? { k: 'ok', t: '양호' } : rate >= 0.8 ? { k: 'warn', t: '주의' } : { k: 'bad', t: '미흡' });
const ratePill = (rate) => { const s = statusOf(rate); return el('span', { class: 'pill ' + s.k }, rate == null ? s.t : `${s.t} ${pct(rate)}`); };

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
  mode: 'demo', server: false, editable: false,      // mode: server 저장 데이터 · demo 예시 · local 내 파일 미리보기
  store: [], version: 0, updated: null, updatedBy: null, serverHistory: [], serverWeekStart: null,
  rows: [], source: '', history: [], weekStart: null,
  asof: SAMPLE_ASOF, team: '', group: '', kind: '', prio: '',
  tab: 'overview', q: '', stage: null, sort: { key: 'priority', dir: 1 }, sel: null, rule: null, open: new Set(), changes: [],
};
let ALL = [], VIEW = [];
const kindOf = (r) => r.kind ?? '미지정';
const prioOf = (r) => r.priority ?? '미지정';
const match = (r) => (!state.team || r.team === state.team) && (!state.group || r.group === state.group)
  && (!state.kind || kindOf(r) === state.kind) && (!state.prio || prioOf(r) === state.prio);
const prioSort = (a, b) => (PR[a.priority] ?? 3) - (PR[b.priority] ?? 3) || b.alerts.length - a.alerts.length || (b.found_fte ?? 0) - (a.found_fte ?? 0);

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
  const target = sum(cur, 'target_now'), secured = sum(cur, 'secured_fte');
  const prevSec = prevT ? sum(prevT, 'secured_fte') : null;
  const prevTarget = prevT ? sum(prevT, (t) => targetAt(t, ps.end)) : null;
  return {
    rows: cur, n: cur.length, found: sum(cur, 'found_fte'), base: sum(cur, 'base_fte'), target, secured, realized: sum(cur, 'realized'),
    rate: target ? secured / target : null, delta: prevSec == null ? null : secured - prevSec, prevRate: prevTarget ? prevSec / prevTarget : null,
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
// 화면에 보일 조직 목록: 팀 필터가 있으면 그 팀의 그룹, 없으면 팀
function units() {
  const rows = ALL.filter((r) => (!state.kind || kindOf(r) === state.kind) && (!state.prio || prioOf(r) === state.prio));
  if (state.team) {
    return [...new Set(rows.filter((r) => r.team === state.team).map((r) => r.group ?? '미지정'))].sort(ko)
      .filter((g) => !state.group || g === state.group)
      .map((g) => ({ name: g, sub: state.team, team: state.team, group: g, fn: (r) => r.team === state.team && (r.group ?? '미지정') === g }));
  }
  return [...new Set(rows.map((r) => r.team ?? '미지정'))].sort(ko).map((t) => ({ name: t, team: t, fn: (r) => (r.team ?? '미지정') === t }));
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
  state.editable = !!d.editable; state.serverWeekStart = d.week_start ?? mondayOf(today());
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
  if (!keep) { state.asof = asof; state.team = state.group = state.kind = state.prio = ''; state.stage = null; state.sel = null; state.rule = null; }
  render();
}

// ── 렌더: 공통 ──────────────────────────────────────────────
const ICON = {
  overview: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z',
  org: 'M9 3h6v5H9zM3 16h6v5H3zM15 16h6v5h-6zM12 8v4M6 16v-4h12v4',
  tasks: 'M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01',
  schedule: 'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4',
  check: 'M8.5 12.5l2.5 2.5 4.5-5M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18z',
};
const TABS = [['overview', '종합 현황'], ['org', '조직별 현황'], ['tasks', '과제 관리'], ['schedule', '추진 일정'], ['check', '데이터 점검']];
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
const getEditor = () => { try { return JSON.parse(store.get(sessionStorage, 'ax-editor')); } catch (e) { return null; } };
const canEdit = () => state.server && state.editable;

function render() {
  compute();
  renderChrome();
  ({ overview: renderOverview, org: renderOrg, tasks: renderTasks, schedule: renderGantt, check: renderCheck })[state.tab]();
}

function fillSelect(sel, opts, cur) {
  const sig = opts.map((o) => o.join('=')).join('|');
  if (sel.dataset.sig !== sig) { sel.replaceChildren(...opts.map(([v, t]) => new Option(t, v))); sel.dataset.sig = sig; }
  sel.value = cur;
}
function renderChrome() {
  const src = { server: `저장 데이터 · 최종 수정 ${state.updated ?? '—'}${state.updatedBy ? ` (${state.updatedBy})` : ''}`, demo: '예시 데이터(가상)', local: `내 파일 미리보기 · ${state.source}` }[state.mode];
  $('#meta').textContent = `${src} · 과제 ${state.rows.length}건 · 기준일 ${state.asof}`;
  const ed = getEditor(), who = $('#who');
  who.hidden = !(state.server && ed);
  if (ed) who.replaceChildren(`편집: ${ed.name}`, el('button', { type: 'button', onclick: () => { store.set(sessionStorage, 'ax-editor', null); render(); toast('편집을 잠갔습니다.'); } }, '잠금'));
  const add = $('#b-add');
  add.disabled = !canEdit();
  add.title = canEdit() ? '' : state.server ? '편집 비밀번호가 설정되지 않아 보기 전용입니다.' : '서버(serve.py)로 연 화면에서만 저장할 수 있습니다.';
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

  $('#cnt-overview').hidden = true; $('#cnt-schedule').hidden = true;
  $('#cnt-org').textContent = units().length;
  $('#cnt-tasks').textContent = VIEW.length;
  const nIss = VIEW.filter((r) => r.issues.length).length, c = $('#cnt-check');
  c.textContent = nIss; c.hidden = !nIss; c.classList.toggle('bad', nIss > 0);

  const b = $('#banner');
  b.className = 'banner'; b.hidden = false;
  if (state.mode === 'local') b.replaceChildren(...dots(el('b', {}, '내 파일 미리보기'), state.source, '다른 사람에게는 보이지 않음'),
    state.server ? el('button', { class: 'btn', type: 'button', onclick: () => showData(false) }, '저장 데이터로 돌아가기') : el('button', { class: 'btn', type: 'button', onclick: () => setData('demo', SAMPLE, SAMPLE_HIST, mondayOf(SAMPLE_ASOF), SAMPLE_ASOF, false) }, '예시 데이터로 돌아가기'));
  else if (state.mode === 'demo' && state.server) b.replaceChildren(...dots(el('b', {}, '예시 데이터(가상) 표시 중'), '저장된 과제 없음', canEdit() ? '[엑셀 업로드] 또는 [+ 과제 추가]로 실제 데이터 입력 시 예시는 사라짐' : '편집 비밀번호 설정 후 입력 가능'));
  else if (state.mode === 'demo') b.replaceChildren(...dots(el('b', {}, '서버 없이 연 화면'), '예시 데이터(가상) 표시', '저장·공유는 serve.py로 실행 시 가능'));
  else if (!state.editable) { b.classList.add('warn'); b.replaceChildren(...dots(el('b', {}, '보기 전용'), '서버 PC에서 py serve.py --set-password 로 편집 비밀번호 설정 필요')); }
  else b.hidden = true;

  takeaway();
}

function takeaway() {
  const box = $('#take'), s = statFor(() => true);
  if (!s.n) { box.textContent = '조건에 맞는 과제가 없습니다.'; return; }
  const scope = [state.team || '전체', state.group, state.kind && `${state.kind} 과제`, state.prio && `우선순위 ${state.prio}`].filter(Boolean).join(' ');
  const st = statusOf(s.rate);
  box.replaceChildren(...[
    `${scope} ${s.n}건 · 발굴 `, el('b', {}, `${f1(s.found)} FTE`), ' 중 ', el('b', {}, `${f1(s.secured)} FTE(${pct(s.found ? s.secured / s.found : null)})`), ' 확보',
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

function tooltip(box) {
  let t = box.querySelector(':scope > .tip');
  if (!t) { t = el('div', { class: 'tip', hidden: true }); box.append(t); }
  return {
    show(ev, parts) {
      t.replaceChildren(...parts); t.hidden = false;
      const b = box.getBoundingClientRect(), w = t.offsetWidth, h = t.offsetHeight;
      let x = ev.clientX - b.left + 14, y = ev.clientY - b.top + 14;
      if (x + w > b.width) x = ev.clientX - b.left - w - 14;
      if (y + h > box.scrollHeight) y = ev.clientY - b.top - h - 14;
      t.style.left = Math.max(0, x) + 'px'; t.style.top = Math.max(0, y) + 'px';
    },
    hide() { t.hidden = true; },
  };
}
const stColor = (i) => `var(--st${i || 1})`;
const stchip = (r) => el('span', { class: 'stchip', style: `border-color:${stColor(r.idx)}`, title: STAGES[r.idx - 1] ?? '' }, r.stage ?? '—');
const prioBadge = (r) => el('span', { class: 'prio ' + (r.priority ? 'p' + r.priority : 'none'), title: r.priority ? `우선순위 ${r.priority}` : '우선순위 미지정' }, r.priority ?? '–');
const alertChips = (r) => r.alerts.map((a) => el('span', { class: 'alert ' + a.k }, a.t));
const statusPill = (r) => el('span', { class: 'pill ' + ({ 완료: 'ok', 지연: 'bad' }[r.status] ?? '') }, r.status);

function sparkline(pts, w = 120, h = 32) {
  const svg = sv('svg', { class: 'spark', width: w, height: h, viewBox: `0 0 ${w} ${h}`, 'aria-hidden': 'true' });
  const vals = pts.flatMap((p) => [p.sec, p.tgt]).filter((v) => v != null);
  if (!vals.length) return svg;
  const x = d3.scaleLinear([0, pts.length - 1], [3, w - 5]), y = d3.scaleLinear([0, d3.max(vals) || 1], [h - 3, 4]);
  const line = (k) => d3.line().defined((p) => p[k] != null).x((p, i) => x(i)).y((p) => y(p[k]));
  svg.append(sv('path', { d: d3.line().defined((p) => p.tgt != null).curve(d3.curveStepAfter).x((p, i) => x(i)).y((p) => y(p.tgt))(pts), fill: 'none', stroke: 'var(--ink-2)', 'stroke-width': 1, 'stroke-dasharray': '3 2' }));
  svg.append(sv('path', { d: line('sec')(pts), fill: 'none', stroke: 'var(--accent)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  const last = pts[pts.length - 1];
  svg.append(sv('circle', { cx: x(pts.length - 1), cy: y(last.sec ?? 0), r: 3, fill: 'var(--accent)' }));
  return svg;
}
function bullet(s, mx) {
  const sc = (v) => `${Math.min(100, (v / (mx || 1)) * 100)}%`;
  return el('div', { class: 'bullet', title: `발굴 ${f1(s.found)} · 확보 ${f1(s.secured)} · 기준일 목표 ${f1(s.target)} FTE` },
    el('div', { class: 'rng', style: `width:${sc(s.found)}` }), el('div', { class: 'bar', style: `width:${sc(s.secured)}` }),
    s.target ? el('div', { class: 'tgt', style: `left:${sc(s.target)}` }) : null);
}

// ── 렌더: 종합 현황 ─────────────────────────────────────────
function renderOverview() { tiles(); orgBars(); flow(); pipe(); curve(); watch(); mgmt(); }

function tiles() {
  const s = statFor(() => true);
  const tile = (lab, big, unit, sub, cls = '') => el('div', { class: 'tile ' + cls }, el('span', { class: 'lab' }, lab), el('span', { class: 'big' }, big, unit ? el('small', {}, unit) : null), el('span', { class: 'sub' }, sub));
  $('#tiles').replaceChildren(
    tile('등록 과제', f0(s.n), '건', dots(`완료 ${s.done}`, `진행 ${s.n - s.done}`, `우선순위 상 ${s.high}`)),
    tile('발굴 FTE', f1(s.found), 'FTE', `Baseline ${f1(s.base)} FTE 대비 ${pct(s.base ? s.found / s.base : null)}`),
    tile('확보 FTE', f1(s.secured), 'FTE', [ratePill(s.rate), `목표 ${f1(s.target)}`, el('span', {}, '전주 대비 ', deltaEl(s.delta))]),
    tile('실현 인원', f0(s.realized), '명', dots(`확보 대비 ${pct(s.secured ? s.realized / s.secured : null)}`, `팀내 ${f0(sum(s.rows, 're_team'))}`, `신규라인 ${f0(sum(s.rows, 're_line'))}`, `기타 ${f0(sum(s.rows, 're_etc'))}`)),
    tile('관리 필요 과제', f0(s.alertN), '건', dots(`지연 ${s.delayed}`, `목표 미달 ${s.lagging}`, `입력 누락 ${s.miss}`), s.alertN ? 'alertile' : ''),
  );
}

function orgBars() {
  const us = units().map((u) => ({ ...u, s: statFor(u.fn), pts: seriesFor(u.fn) }));
  const mx = d3.max(us, (u) => u.s.found) || 1;
  $('#orgbars').replaceChildren(
    el('div', { class: 'ob-head' }, el('span', {}, state.team ? `${state.team} 그룹` : '팀'), el('span', {}, '목표 달성'), el('span', { style: 'text-align:right' }, '달성률'), el('span', { style: 'text-align:right' }, '전주 대비'), el('span', {}, '추이(8주)')),
    ...us.map((u) => el('div', { class: 'ob-row', tabindex: 0, role: 'button', onclick: () => goOrg(u), onkeydown: (e) => { if (e.key === 'Enter') goOrg(u); } },
      el('div', { class: 'ob-name' }, el('b', {}, u.name), el('small', {}, dots(`과제 ${u.s.n}`, `확보 ${f1(u.s.secured)} / 목표 ${f1(u.s.target)}`))),
      bullet(u.s, mx),
      el('div', { class: 'ob-num' }, ratePill(u.s.rate), el('div', { style: 'font-size:12px;color:var(--ink-2);margin-top:2px' }, `전주 ${pct(u.s.prevRate)}`)),
      el('div', { class: 'ob-num' }, deltaEl(u.s.delta), el('div', { style: 'font-size:12px;color:var(--ink-2)' }, 'FTE')),
      sparkline(u.pts))),
    ...(us.length ? [] : [el('p', { class: 'empty' }, '조건에 맞는 조직이 없습니다.')]));
}
function goOrg(u) { state.team = u.team; state.group = u.group ?? ''; state.open.add(u.team); setTab('org'); }

function flow() {
  const s = statFor(() => true), mx = Math.max(s.base, 1e-9);
  const row = (nm, v, unit, color, extra, fmt = f1) => el('div', { class: 'row' }, el('span', {}, nm),
    el('div', { class: 'track' }, el('div', { class: 'bar', style: `width:${Math.min(100, (v / mx) * 100)}%;background:${color}` }), extra), el('span', { class: 'val' }, `${fmt(v)} ${unit}`));
  const conv = (t) => el('div', { class: 'row' }, el('span'), el('span', { class: 'conv' }, t), el('span'));
  $('#flow').replaceChildren(
    row('Baseline', s.base, 'FTE', 'var(--line-2)'), conv(`↓ 발굴률 ${pct(s.base ? s.found / s.base : null)}`),
    row('발굴', s.found, 'FTE', 'var(--found)'), conv(`↓ 확보율 ${pct(s.found ? s.secured / s.found : null)} · 기준일 목표 ${f1(s.target)} 대비 ${pct(s.rate)}`),
    row('확보', s.secured, 'FTE', 'var(--accent)', s.target ? el('div', { class: 'tg', style: `left:${Math.min(100, (s.target / mx) * 100)}%` }) : null),
    conv(`↓ 실현율 ${pct(s.secured ? s.realized / s.secured : null)}`),
    row('실현', s.realized, '명', 'var(--real)', null, f0));
}

function pipe() {
  const by = d3.group(VIEW, (r) => r.idx), mx = d3.max([1, 2, 3, 4, 5], (i) => (by.get(i) ?? []).length) || 1;
  $('#pipe').replaceChildren(...[1, 2, 3, 4, 5].map((i) => {
    const g = by.get(i) ?? [], dl = g.filter((r) => r.delayed).length, dn = g.filter((r) => r.done).length;
    return el('button', { class: 'stg', type: 'button', 'aria-label': `S${i} 과제 ${g.length}건 보기`, onclick: () => { state.stage = i; setTab('tasks'); } },
      el('span', { class: 'nm' }, el('b', {}, `S${i}`), STAGES[i - 1]),
      el('span', { class: 'col' }, el('div', { style: `height:${(g.length / mx) * 100}%;background:${stColor(i)}` })),
      el('span', {}, el('div', { class: 'n' }, g.length), el('div', { class: 'f' }, `${f1(sum(g, 'found_fte'))} FTE`),
        el('div', { class: 'd' }, dl ? el('span', { class: 'c-bad' }, `지연 ${dl}`) : i === 5 && dn ? el('span', { class: 'c-ok' }, `완료 ${dn}`) : '')));
  }));
}

function curve() {
  const box = $('#curve'); box.replaceChildren();
  const W = box.clientWidth; if (!W) return;
  const rs = VIEW.filter((r) => r.s5_date);
  if (!rs.length) { box.append(el('p', { class: 'empty' }, '목표일정이 입력된 과제가 없습니다.')); return; }
  const asof = pDate(state.asof), fmt = d3.utcFormat('%Y-%m-%d');
  const lo = d3.utcMonth.floor(d3.min(rs, (r) => pDate(r.plan_start || r.s1_date) ?? asof));
  const hi = d3.utcMonth.ceil(d3.utcDay.offset(d3.max([...rs.map((r) => pDate(r.s5_date)), asof]), 1));
  const target = d3.utcMonth.range(lo, d3.utcMonth.offset(hi, 1)).map((m) => {
    const e = fmt(d3.utcDay.offset(d3.utcMonth.offset(m, 1), -1));
    return { m, v: sum(rs, (r) => targetAt(r, e)) };
  });
  const actual = seriesFor(() => true).filter((p) => p.sec != null).map((p) => ({ d: pDate(p.end), v: p.sec }));
  const s = statFor(() => true);
  const H = 290, m = { t: 26, r: 18, b: 28, l: 44 };
  const x = d3.scaleUtc([lo, hi], [m.l, W - m.r]);
  const y = d3.scaleLinear([0, Math.max(s.found, s.secured) * 1.08 || 1], [H - m.b, m.t]).nice();
  const svg = d3.select(box).append('svg').attr('width', '100%').attr('height', H).attr('viewBox', `0 0 ${W} ${H}`);
  svg.append('g').attr('class', 'gl').selectAll('line').data(y.ticks(5)).join('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y).attr('y2', y);
  svg.append('g').selectAll('text').data(y.ticks(5)).join('text').attr('x', m.l - 8).attr('y', y).attr('dy', '0.32em').attr('text-anchor', 'end').text(d3.format(','));
  svg.append('g').attr('class', 'axis').attr('transform', `translate(0,${H - m.b})`)
    .call(d3.axisBottom(x).ticks(Math.max(2, Math.floor((W - m.l - m.r) / 72))).tickFormat(d3.utcFormat('%y.%m')).tickSizeOuter(0)).attr('font-size', null).attr('font-family', null);
  svg.append('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y(s.found)).attr('y2', y(s.found)).attr('stroke', 'var(--muted)').attr('stroke-dasharray', '4 4');
  svg.append('text').attr('x', W - m.r).attr('y', y(s.found) - 6).attr('text-anchor', 'end').text(`발굴 총량 ${f1(s.found)} FTE`);
  const step = target.map((d) => ({ x: d.m, v: d.v })).concat([{ x: hi, v: target[target.length - 1].v }]);
  svg.append('path').datum(step).attr('fill', 'var(--accent)').attr('fill-opacity', 0.06).attr('d', d3.area().curve(d3.curveStepAfter).x((d) => x(d.x)).y0(y(0)).y1((d) => y(d.v)));
  svg.append('path').datum(step).attr('fill', 'none').attr('stroke', 'var(--ink-2)').attr('stroke-width', 1.5).attr('d', d3.line().curve(d3.curveStepAfter).x((d) => x(d.x)).y((d) => y(d.v)));
  svg.append('line').attr('x1', x(asof)).attr('x2', x(asof)).attr('y1', m.t - 10).attr('y2', H - m.b).attr('stroke', 'var(--ink)');
  svg.append('text').attr('x', x(asof)).attr('y', m.t - 14).attr('text-anchor', x(asof) > W - 60 ? 'end' : 'middle').attr('style', 'fill:var(--ink);font-weight:700').text('기준일');
  if (actual.length) {
    svg.append('path').datum(actual).attr('fill', 'none').attr('stroke', 'var(--accent)').attr('stroke-width', 2.2).attr('stroke-linejoin', 'round').attr('d', d3.line().x((d) => x(d.d)).y((d) => y(d.v)));
    svg.append('g').selectAll('circle').data(actual).join('circle').attr('cx', (d) => x(d.d)).attr('cy', (d) => y(d.v)).attr('r', (d, i) => (i === actual.length - 1 ? 5 : 2.6)).attr('fill', 'var(--accent)');
  }
  const st = statusOf(s.rate), col = { ok: 'var(--ok)', warn: 'var(--warn)', bad: 'var(--bad)' }[st.k] ?? 'var(--ink)', left = x(asof) > W * 0.6;
  svg.append('text').attr('x', x(asof) + (left ? -12 : 12)).attr('y', y(s.secured) - 10).attr('text-anchor', left ? 'end' : 'start').attr('style', `fill:${col};font-weight:700;font-size:12.5px`)
    .text(`확보 ${f1(s.secured)} FTE · 목표 ${f1(s.target)} 대비 ${pct(s.rate)}`);
  const tip = tooltip(box), dot = svg.append('circle').attr('r', 4).attr('fill', 'var(--ink-2)').style('display', 'none');
  svg.append('rect').attr('x', m.l).attr('y', m.t).attr('width', W - m.l - m.r).attr('height', H - m.t - m.b).attr('fill', 'transparent')
    .on('mousemove', (ev) => {
      const d = target[Math.max(0, d3.bisector((t) => t.m).right(target, x.invert(d3.pointer(ev)[0])) - 1)];
      dot.style('display', null).attr('cx', x(d.m)).attr('cy', y(d.v));
      tip.show(ev, [el('div', { class: 't' }, d3.utcFormat('%Y년 %m월')(d.m)), el('div', {}, el('b', {}, `누적 확보 목표 ${f1(d.v)} FTE`)), el('div', { class: 't' }, `발굴 대비 ${pct(s.found ? d.v / s.found : null)}`)]);
    })
    .on('mouseleave', () => { tip.hide(); dot.style('display', 'none'); });
}

function watch() {
  const w = VIEW.filter((r) => r.priority === '상' || r.alerts.length).sort(prioSort).slice(0, 12);
  const t = $('#watch');
  if (!w.length) { t.replaceChildren(el('tbody', {}, el('tr', {}, el('td', { class: 'empty' }, '우선 관리 대상 과제가 없습니다.')))); return; }
  t.replaceChildren(el('thead', {}, el('tr', {}, ['우선', '과제명', '팀 · 그룹', '단계', '경보', '확보 / 목표'].map((h, i) => el('th', { class: i === 5 ? 'num' : '' }, h)))),
    el('tbody', {}, w.map((r) => el('tr', { class: 'clk', onclick: () => openDetail(r.id) },
      el('td', {}, prioBadge(r)), el('td', { class: 'task-nm' }, r.task), el('td', {}, `${r.team ?? '—'} · ${r.group ?? '—'}`), el('td', {}, stchip(r)),
      el('td', {}, r.alerts.length ? alertChips(r) : '—'), el('td', { class: 'num' }, r.target_now ? `${f1(r.secured_fte ?? 0)} / ${f1(r.target_now)}` : `${f1(r.secured_fte ?? 0)} / —`)))));
}

function mgmt() {
  const g = d3.group(VIEW, (r) => r.mgmt), mx = d3.max(Object.keys(MGMT), (k) => (g.get(k) ?? []).length) || 1;
  $('#mgmt').replaceChildren(...Object.keys(MGMT).map((k) => {
    const v = g.get(k) ?? [], kp = v.filter((r) => r.kpi_prog != null);
    return el('div', { class: 'mg-row' }, el('b', {}, `${k} ${MGMT[k]}`), el('i', { style: `width:${Math.max(1, (v.length / mx) * 100)}%` }),
      el('span', {}, `${v.length}건 · 발굴 ${f1(sum(v, 'found_fte'))}`, k === 'P' ? '' : ` · KPI 진척 ${kp.length ? pct(d3.mean(kp, (r) => r.kpi_prog)) : '—'}`));
  }));
}

// ── 렌더: 조직별 현황 ───────────────────────────────────────
function renderOrg() {
  const teams = [...new Set(VIEW.map((r) => r.team ?? '미지정'))].sort(ko);
  if (state.team) state.open.add(state.team);
  const tStats = teams.map((t) => ({ t, s: statFor((r) => (r.team ?? '미지정') === t) }));
  const mx = d3.max(tStats, (d) => d.s.found) || 1;
  const head = ['조직', '과제', '단계 분포', '발굴 FTE', '기준일 목표', '확보 FTE', '목표 달성', '달성률', '전주 대비', '추이(8주)', '실현(명)', '지연', '우선 상', '경보'];
  const NUMS = new Set([1, 3, 4, 5, 7, 8, 10, 11, 12, 13]);
  const cells = (name, s, pts, scale = mx) => [
    name, el('td', { class: 'num' }, s.n),
    el('td', {}, el('div', { class: 'stbar', title: s.stages.map((n, i) => `S${i + 1} ${n}건`).join(' · ') }, s.stages.map((n, i) => (n ? el('i', { style: `width:${(n / s.n) * 100}%;background:${stColor(i + 1)}` }) : null)))),
    el('td', { class: 'num' }, f1(s.found)), el('td', { class: 'num' }, f1(s.target)), el('td', { class: 'num' }, el('b', {}, f1(s.secured))),
    el('td', {}, bullet(s, scale)), el('td', { class: 'num' }, ratePill(s.rate)), el('td', { class: 'num' }, deltaEl(s.delta)),
    el('td', {}, sparkline(pts, 110, 28)), el('td', { class: 'num' }, f0(s.realized)),
    el('td', { class: 'num' + (s.delayed ? ' c-bad' : '') }, s.delayed), el('td', { class: 'num' }, s.high), el('td', { class: 'num' + (s.alertN ? ' c-bad' : '') }, s.alertN),
  ];
  const body = [];
  for (const { t, s } of tStats) {
    const open = state.open.has(t);
    const tog = el('button', { class: 'tog', type: 'button', 'aria-expanded': String(open), onclick: () => { open ? state.open.delete(t) : state.open.add(t); renderOrg(); } }, el('span', { class: 'car' }, '▸'), t);
    body.push(el('tr', { class: 'team' }, cells(el('td', {}, tog), s, seriesFor((r) => (r.team ?? '미지정') === t))));
    if (open) {
      for (const g of [...new Set(s.rows.map((r) => r.group ?? '미지정'))].sort(ko)) {
        const fn = (r) => (r.team ?? '미지정') === t && (r.group ?? '미지정') === g;
        body.push(el('tr', { class: 'grp' }, cells(el('td', {}, g), statFor(fn), seriesFor(fn))));
      }
    }
  }
  if (tStats.length > 1) { const all = statFor(() => true); body.push(el('tr', { class: 'total' }, cells(el('td', {}, '전체'), all, seriesFor(() => true), all.found))); }
  $('#score').replaceChildren(el('thead', {}, el('tr', {}, head.map((h, i) => el('th', { class: NUMS.has(i) ? 'num' : '' }, h)))),
    el('tbody', {}, body.length ? body : el('tr', {}, el('td', { colspan: head.length, class: 'empty' }, '조건에 맞는 과제가 없습니다.'))));
  const allOpen = teams.every((t) => state.open.has(t));
  const bx = $('#b-expand'); bx.textContent = allOpen ? '그룹 모두 접기' : '그룹 모두 펼치기';
  bx.onclick = () => { if (allOpen) state.open.clear(); else teams.forEach((t) => state.open.add(t)); renderOrg(); };

  // 조직별 우선 과제 카드: 팀 필터가 있으면 그룹 단위
  $('#prio').replaceChildren(...units().map((u) => {
    const s = statFor(u.fn), top = s.rows.slice().sort(prioSort).slice(0, 5);
    return el('article', { class: 'card pc' },
      el('header', { class: 'card-hd' }, el('div', {}, el('h3', {}, u.sub ? `${u.sub} · ${u.name}` : u.name),
        el('div', { class: 'pc-stat', style: 'margin-top:4px' }, dots(`과제 ${s.n}`, `우선 상 ${s.high}`, el('span', { class: s.alertN ? 'c-bad' : '' }, `경보 ${s.alertN}`)))),
        el('div', { class: 'pc-stat' }, ratePill(s.rate), deltaEl(s.delta, ' FTE'))),
      top.length ? el('ul', { class: 'pc-list' }, top.map((r) => el('li', { onclick: () => openDetail(r.id), title: '상세 보기' },
        prioBadge(r),
        el('div', {}, el('div', { class: 't' }, r.task), el('div', { class: 's' }, state.team ? '' : `${r.group ?? '—'} ·`, stchip(r), alertChips(r))),
        el('div', { class: 'v' }, el('b', {}, f1(r.secured_fte ?? 0)), ` / ${r.target_now ? f1(r.target_now) : '—'}`, el('div', { style: 'font-size:11.5px;color:var(--ink-2)' }, '확보 / 목표'))))) : el('p', { class: 'empty' }, '과제 없음'),
      el('div', { class: 'pc-foot' }, el('button', { class: 'link', type: 'button', onclick: () => { state.team = u.team; state.group = u.group ?? ''; setTab('tasks'); } }, `전체 ${s.n}건 보기 →`)));
  }));
}

// ── 렌더: 과제 관리 ─────────────────────────────────────────
const COLS = [
  ['priority', '우선'], ['task', '과제명'], ['team', '팀'], ['group', '그룹'], ['kind', '구분'], ['stage', '단계'], ['status', '상태'], ['alerts', '경보'],
  ['found_fte', '발굴 FTE', 'num'], ['target_now', '기준일 목표', 'num'], ['secured_fte', '확보 FTE', 'num'], ['delta', '전주 대비', 'num'], ['realized', '실현(명)', 'num'], ['issues', '점검', 'num'],
];
const sortVal = (r, k) => (k === 'priority' ? PR[r.priority] ?? 3 : k === 'alerts' ? -r.alerts.length : k === 'issues' ? r.issues.length : k === 'stage' ? r.idx : r[k]);
function renderTasks() {
  const q = state.q.trim().toLowerCase();
  $('#stchips').replaceChildren(...[null, 1, 2, 3, 4, 5].map((i) => el('button', { class: 'chip', type: 'button', 'aria-pressed': String(state.stage === i), onclick: () => { state.stage = i; render(); } },
    i ? `S${i}` : '전체', el('b', {}, i ? VIEW.filter((r) => r.idx === i).length : VIEW.length))));
  let v = VIEW.filter((r) => (!state.stage || r.idx === state.stage) && (!q || [r.task, r.cat, r.team, r.group, r.kpi, r.tgt_ids, r.dev_ids, r.dev_dept].join(' ').toLowerCase().includes(q)));
  const { key, dir } = state.sort;
  v = v.slice().sort((a, b) => {
    const x = sortVal(a, key), y = sortVal(b, key);
    return (x == null) - (y == null) || (typeof x === 'number' ? x - y : ko(x, y)) * dir || prioSort(a, b);
  });
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
    if (k === 'delta') return deltaEl(r.delta);
    if (k === 'realized') return r.realized ? f0(r.realized) : '—';
    if (k === 'issues') return r.issues.length ? el('span', { class: 'c-bad' }, `${r.issues.length}건`) : el('span', { class: 'c-ok' }, '정상');
    return r[k] ?? '—';
  };
  $('#tbl').replaceChildren(
    el('thead', {}, el('tr', {}, COLS.map(([k, h, c]) => el('th', { class: c ?? '', 'aria-sort': key === k ? (dir > 0 ? 'ascending' : 'descending') : 'none' },
      el('button', { type: 'button', onclick: () => { state.sort = { key: k, dir: key === k ? -dir : 1 }; render(); } }, h))))),
    el('tbody', {}, v.length ? v.map((r) => el('tr', { class: 'clk' + (state.sel === r.id ? ' sel' : ''), onclick: () => openDetail(r.id) }, COLS.map(([k, , c]) => el('td', { class: c ?? '' }, cell(r, k)))))
      : el('tr', {}, el('td', { colspan: COLS.length, class: 'empty' }, '조건에 맞는 과제가 없습니다.'))));
  renderDetail();
}

function openDetail(id) {
  state.sel = id;
  if (state.tab !== 'tasks') setTab('tasks'); else render();
  requestAnimationFrame(() => $('#detail').scrollIntoView({ behavior: 'smooth', block: 'start' }));
}

function renderDetail() {
  const box = $('#detail'), r = ALL.find((d) => d.id === state.sel);
  box.hidden = !r; if (!r) return;
  const missing = new Set(r.issues.filter((i) => i.level === '누락').map((i) => i.key));
  const show = (k, t) => { const v = r[k]; return v == null || v === '' ? '—' : t === 'n' || t === 'auto' ? fnum(v) : v; };
  const editable = state.mode === 'server' && canEdit();
  box.replaceChildren(
    el('header', { class: 'card-hd' }, el('div', { class: 'dt-hd', style: 'width:100%' },
      el('div', {}, el('div', { class: 'crumb' }, dots(r.team, r.group, r.cat)), el('h2', {}, r.task),
        el('div', { class: 'pc-stat', style: 'margin-top:6px' }, prioBadge(r), stchip(r), r.delayed ? null : statusPill(r), alertChips(r), `확보 ${f1(r.secured_fte ?? 0)} FTE`, el('span', {}, '전주 대비 ', deltaEl(r.delta)))),
      el('div', { class: 'dt-acts' },
        el('button', { class: 'btn primary', type: 'button', disabled: !editable, title: editable ? '' : '저장 데이터를 보는 중이고 편집 권한이 있을 때만 수정할 수 있습니다.', onclick: () => openEditor(r) }, '수정'),
        el('button', { class: 'btn danger', type: 'button', disabled: !editable, onclick: () => removeTask(r) }, '삭제'),
        el('button', { class: 'btn', type: 'button', onclick: () => { state.sel = null; render(); } }, '닫기')))),
    el('div', { class: 'card-bd' },
      r.issues.length ? el('p', { class: 'issues-line' }, '점검: ' + r.issues.map((i) => `${i.field} (${i.rule})`).join(' · ')) : null,
      el('div', { class: 'secs' }, SEC.map((s, si) => el('section', { class: 'sec' }, el('h4', {}, s),
        si === 5 ? el('div', { class: 'ms' }, [1, 2, 3, 4, 5].map((k) => el('div', { class: (r.done || k < r.idx ? 'done' : '') + (k === r.idx ? (r.delayed ? ' late' : ' cur') : '') },
          el('b', {}, `S${k}`), r[`s${k}_date`] ? r[`s${k}_date`].slice(2).replace(/-/g, '.') : '—', k >= 3 ? el('div', {}, `${fnum(r[`s${k}_fte`])} FTE`) : null))) : null,
        el('dl', { class: 'kv' }, FORM_FIELDS.filter((f) => f[2] === si).flatMap(([k, , , t]) => [el('dt', {}, label[k]), el('dd', { class: missing.has(k) ? 'miss' : '' }, missing.has(k) ? '미입력' : show(k, t))])))))));
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
  const tip = tooltip(box);
  items.forEach((it, i) => {
    const y0 = top + i * RH;
    if (it.head) {
      if (it.lv === 0) svg.append('rect').attr('x', 0).attr('y', y0 + 1).attr('width', W).attr('height', RH - 2).attr('fill', 'var(--soft)');
      svg.append('text').attr('x', it.lv ? 14 : 6).attr('y', y0 + RH / 2 + 4).attr('style', `fill:var(--ink);font-weight:700;font-size:${it.lv ? 12 : 13}px`).text(it.head);
      return;
    }
    const r = it, g = svg.append('g').style('cursor', 'pointer').on('click', () => openDetail(r.id));
    g.append('rect').attr('x', 0).attr('y', y0).attr('width', W).attr('height', RH).attr('fill', 'transparent');
    const name = g.append('text').attr('x', 28).attr('y', y0 + RH / 2 + 4).attr('style', 'fill:var(--ink)').text(r.task);
    while (name.node().getComputedTextLength() > LW - 32 && name.text().length > 3) name.text(name.text().slice(0, -2) + '…');
    g.append('rect').attr('x', x(pDate(r.plan_start))).attr('y', y0 + 5).attr('width', Math.max(2, x(pDate(r.plan_end)) - x(pDate(r.plan_start)))).attr('height', RH - 10).attr('rx', 4).attr('fill', 'var(--found)').attr('fill-opacity', 0.7);
    if (r.act_start) {
      const e = r.act_end ? pDate(r.act_end) : asof, c = r.done ? 'var(--ok)' : r.delayed ? 'var(--bad)' : 'var(--accent)';
      g.append('rect').attr('x', x(pDate(r.act_start))).attr('y', y0 + 10).attr('width', Math.max(2, x(e) - x(pDate(r.act_start)))).attr('height', RH - 20).attr('rx', 3).attr('fill', c);
    }
    [1, 2, 3, 4, 5].forEach((k) => {
      const d = r[`s${k}_date`]; if (!d) return;
      const cx = x(pDate(d)), cy = y0 + RH / 2, past = r.done || k < r.idx, late = k === r.idx && r.delayed;
      g.append('path').attr('d', `M${cx},${cy - 5}L${cx + 5},${cy}L${cx},${cy + 5}L${cx - 5},${cy}Z`)
        .attr('fill', past ? stColor(k) : 'var(--panel)').attr('stroke', late ? 'var(--bad)' : 'var(--st5)').attr('stroke-width', late ? 2 : 1.2);
    });
    g.on('mousemove', (ev) => tip.show(ev, [el('div', { class: 't' }, `${r.team} · ${r.group ?? ''} · ${r.stage} ${STAGES[r.idx - 1] ?? ''}`), el('div', {}, el('b', {}, r.task)),
      el('div', {}, `계획 ${r.plan_start} ~ ${r.plan_end}`), el('div', {}, `실적 ${r.act_start ?? '미착수'} ~ ${r.act_end ?? (r.act_start ? '진행 중' : '')}`),
      el('div', { class: r.delayed ? 'c-bad' : 't' }, r.delayed ? `${r.stage} 목표일 ${r.due} 경과 (${r.late_days}일)` : `현재 단계 목표일 ${r.due ?? '—'}`)])).on('mouseleave', () => tip.hide());
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
    el('tbody', {}, v.length ? v.map((d) => el('tr', { class: 'clk', onclick: () => openDetail(d.id) }, el('td', {}, d.team ?? '—'), el('td', {}, d.group ?? '—'), el('td', { class: 'task-nm' }, d.task), el('td', {}, d.stage ?? '—'),
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

// ── 대화상자 ────────────────────────────────────────────────
function modal({ title, notes, body, actions, wide }) {
  const m = $('#modal');
  return new Promise((resolve) => {
    const err = el('p', { class: 'err', hidden: true });
    const setErr = (t) => { err.textContent = t ?? ''; err.hidden = !t; };
    const close = (v) => { m.hidden = true; m.replaceChildren(); document.removeEventListener('keydown', onKey); resolve(v); };
    const onKey = (e) => { if (e.key === 'Escape') close(null); };
    const btns = actions.map((a) => el('button', { class: 'btn ' + (a.cls ?? ''), type: 'button', onclick: async (e) => {
      if (!a.run) { close(a.value ?? null); return; }
      const btn = e.currentTarget; btn.disabled = true;
      const v = await a.run(setErr); btn.disabled = false;
      if (v !== undefined) close(v);
    } }, a.label));
    m.replaceChildren(el('div', { class: 'mbox' + (wide ? ' wide' : '') },
      el('header', { class: 'card-hd' }, el('div', {}, el('h2', {}, title), notes ? el('ul', { class: 'notes' }, notes.map((n) => el('li', {}, n))) : null)),
      el('div', { class: 'm-bd' }, body, err), el('div', { class: 'm-ft' }, btns)));
    m.onkeydown = (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') btns[btns.length - 1].click(); };
    m.hidden = false; document.addEventListener('keydown', onKey);
    (m.querySelector('input') ?? btns[btns.length - 1]).focus();
  });
}

async function ensureEditor() {
  const saved = getEditor(); if (saved) return saved;
  const name = el('input', { id: 'm-name', type: 'text', maxlength: 30, autocomplete: 'name' });
  const pw = el('input', { id: 'm-pw', type: 'password', autocomplete: 'current-password' });
  name.value = store.get(localStorage, 'ax-editor-name') ?? '';
  return modal({
    title: '편집 권한 확인', notes: ['입력자 이름: 변경 이력에 기록', '비밀번호: 서버 PC에서 정한 편집 비밀번호'],
    body: el('div', { class: 'fs-grid', style: 'padding:0' }, el('div', { class: 'ff' }, el('label', { for: 'm-name' }, '입력자 이름'), name), el('div', { class: 'ff' }, el('label', { for: 'm-pw' }, '편집 비밀번호'), pw)),
    actions: [{ label: '취소' }, { label: '확인', cls: 'primary', run: async (setErr) => {
      if (!name.value.trim()) { setErr('입력자 이름을 적어 주세요.'); return undefined; }
      let r; try { r = await post('api/auth', { password: pw.value }); } catch (e) { setErr('서버에 연결하지 못했습니다.'); return undefined; }
      if (!r.ok) { setErr(r.data?.error ?? '비밀번호가 맞지 않습니다.'); return undefined; }
      const ed = { name: name.value.trim(), pw: pw.value };
      store.set(sessionStorage, 'ax-editor', JSON.stringify(ed)); store.set(localStorage, 'ax-editor-name', ed.name);
      renderChrome(); return ed;
    } }],
  });
}

// 저장 요청: 서버가 새 데이터 전체를 돌려주면 화면을 갱신
async function commit(body) {
  const ed = await ensureEditor(); if (!ed) return null;
  let r; try { r = await post('api/commit', { ...body, editor: ed.name, password: ed.pw, base_version: state.version }); } catch (e) { toast('서버에 연결하지 못했습니다. 서버 PC의 serve.py 창이 켜져 있는지 확인해 주세요.', 'err'); return null; }
  if (r.status === 401) { store.set(sessionStorage, 'ax-editor', null); renderChrome(); toast(r.data?.error ?? '편집 비밀번호가 맞지 않습니다. 다시 확인해 주세요.', 'err'); return null; }
  if (r.status === 409) { await loadServer(true); toast('다른 사람이 먼저 수정했습니다. 최신 내용을 불러왔으니 다시 시도해 주세요.', 'err'); return null; }
  if (!r.ok) { toast(r.data?.error ?? `저장하지 못했습니다 (HTTP ${r.status}).`, 'err'); return null; }
  applyStore(r.data); await loadHistory(); showData(true); if (state.tab === 'check') loadChanges();
  return r.data;
}

async function removeTask(r) {
  const ok = await modal({ title: '과제 삭제', notes: [`${r.team} · ${r.group ?? ''} · ${r.task}`, '삭제 후에는 변경 이력에만 기록이 남음'], body: el('p', {}, '이 과제를 삭제합니다.'),
    actions: [{ label: '취소' }, { label: '삭제', cls: 'danger', value: true }] });
  if (!ok) return;
  if (await commit({ op: 'delete', id: r.id })) { state.sel = null; render(); toast(`'${r.task}'을(를) 삭제했습니다.`); }
}

// ── 과제 입력 패널 ──────────────────────────────────────────
const isBusy = () => !$('#drawer').hidden || !$('#modal').hidden;
async function openEditor(task) {
  if (!canEdit()) { toast('서버(serve.py)로 열고 편집 비밀번호가 설정된 경우에만 저장할 수 있습니다.', 'err'); return; }
  if (!(await ensureEditor())) return;
  const base = state.store.map(normalize);
  const opts = (k, f = () => true) => [...new Set(base.filter(f).map((r) => r[k]).filter(Boolean))].sort(ko);
  const draft = task ? { ...task } : { team: state.team || null, group: state.group || null, stage: 'S1', priority: '중', pkms: '미등록', agent: 'X' };
  const dr = $('#drawer'), lists = el('div');
  const mkList = (id, vals) => el('datalist', { id }, vals.map((v) => el('option', { value: v })));
  const refreshLists = (cur) => lists.replaceChildren(
    mkList('dl-team', opts('team')), mkList('dl-group', opts('group', (r) => !cur.team || r.team === cur.team)), mkList('dl-cat', opts('cat', (r) => !cur.team || r.team === cur.team)),
    mkList('dl-scope', ['단일 라인', '복수 라인', '공장 전체', '팀 업무', ...opts('scope')]), mkList('dl-tgt_dept', opts('tgt_dept')), mkList('dl-dev_dept', opts('dev_dept')),
    mkList('dl-kpi', [...opts('kpi', (r) => !cur.mgmt || r.mgmt === cur.mgmt), '기타']));
  const inputs = new Map();
  const field = ([k, h, , t, choices]) => {
    const id = `ff-${k}`, v = draft[k];
    let input;
    if (t === 'auto') input = el('output', { id, for: '' });
    else if (t === 'ta') input = el('textarea', { id, rows: 2 });
    else if (t === 'sel') { const cs = [...new Set([...(choices ?? []), ...(v != null && !choices.includes(v) ? [v] : [])])]; input = el('select', { id }, el('option', { value: '' }, '선택'), cs.map((c) => el('option', { value: c }, k === 'stage' ? `${c} ${STAGES[+c.slice(1) - 1]}` : k === 'mgmt' ? `${c} ${MGMT[c]}` : c))); }
    else input = el('input', { id, type: t === 'n' ? 'number' : t === 'd' ? 'date' : 'text', step: t === 'n' ? 'any' : null, list: t === 'dl' ? `dl-${k}` : null });
    if (t !== 'auto') input.value = v ?? '';
    input.dataset.key = k; inputs.set(k, input);
    const tag = h.includes('[조건]') ? '조건부' : h.includes('[P외]') ? 'P 외 필수' : t === 'auto' ? '자동 계산' : null;
    return el('div', { class: 'ff' + (t === 'ta' ? ' wide' : ''), id: `w-${k}` },
      el('label', { for: id }, shortName(h), h.startsWith('*') && !tag ? el('span', { class: 'req', title: '필수' }, '*') : null, tag ? el('span', { class: 'tag' }, tag) : null),
      input, el('span', { class: 'hint', hidden: true }));
  };
  const sumEl = el('span', { class: 'sum' });
  const read = () => {
    const raw = { id: task?.id ?? null };
    for (const [k, inp] of inputs) if (inp.tagName !== 'OUTPUT') raw[k] = inp.value;
    return normalize(raw);
  };
  const update = (changedKey) => {
    const cur = read();
    if (changedKey === 'team' || changedKey === 'mgmt' || !changedKey) refreshLists(cur);
    if (changedKey === 'kpi' && !inputs.get('unit').value) {
      const same = base.find((r) => r.kpi === cur.kpi && r.unit);
      if (same) { inputs.get('unit').value = same.unit; cur.unit = same.unit; }
    }
    for (const [k, inp] of inputs) if (inp.tagName === 'OUTPUT') inp.value = fnum(cur[k]);
    const iss = validate(cur, cur.stage ? +cur.stage.slice(1) : 0);
    const byKey = d3.group(iss, (i) => i.key);
    for (const [k] of inputs) {
      const w = $(`#w-${k}`, dr), list = byKey.get(k), hint = w.querySelector('.hint');
      w.classList.toggle('miss', !!list); hint.hidden = !list; hint.textContent = list ? list.map((i) => (i.level === '누락' ? `미입력 · ${i.rule}` : i.rule)).join(' / ') : '';
    }
    const nMiss = iss.filter((i) => i.level === '누락').length, nChk = iss.length - nMiss;
    sumEl.replaceChildren(iss.length ? el('span', { class: 'c-bad' }, dots(nMiss ? `미입력 ${nMiss}개` : null, nChk ? `확인 필요 ${nChk}개` : null), ' · 저장은 가능') : el('span', { class: 'c-ok' }, '입력 규칙 이상 없음'));
  };
  const close = () => { dr.hidden = true; $('#overlay').hidden = true; dr.replaceChildren(); document.removeEventListener('keydown', onKey); };
  const onKey = (e) => { if (e.key === 'Escape' && $('#modal').hidden) close(); };
  const save = async () => {
    const cur = read();
    if (!cur.team || !cur.task) { toast('팀과 과제명(중분류)은 꼭 입력해야 합니다.', 'err'); inputs.get(cur.team ? 'task' : 'team').focus(); return; }
    const dup = state.store.find((r) => r.team === cur.team && r.task === cur.task && r.id !== cur.id);
    if (dup) { toast(`${cur.team}에 같은 과제명 '${cur.task}'이(가) 이미 있습니다.`, 'err'); inputs.get('task').focus(); return; }
    const res = await commit({ op: 'save', task: cur });
    if (res) { const saved = res.tasks.find((t) => t.team === cur.team && t.task === cur.task); close(); toast(`'${cur.task}'을(를) 저장했습니다.`); if (saved) openDetail(saved.id); }
  };
  dr.replaceChildren(
    el('div', { class: 'dr-hd' }, el('h2', {}, task ? `과제 수정 · ${task.task}` : '과제 추가'), el('button', { type: 'button', onclick: close }, '닫기')),
    el('div', { class: 'dr-bd' }, SEC.map((s, si) => el('fieldset', { class: 'fs' }, el('legend', {}, s), el('div', { class: 'fs-grid' }, FORM_FIELDS.filter((f) => f[2] === si).map(field)))), lists),
    el('div', { class: 'dr-ft' }, sumEl, el('div', { class: 'dt-acts' }, el('button', { class: 'btn', type: 'button', onclick: close }, '취소'), el('button', { class: 'btn primary', type: 'button', onclick: save }, '저장'))));
  dr.addEventListener('input', (e) => { if (e.target.dataset.key) update(e.target.dataset.key); });
  dr.addEventListener('change', (e) => { if (e.target.dataset.key) update(e.target.dataset.key); });
  dr.hidden = false; $('#overlay').hidden = false; document.addEventListener('keydown', onKey);
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

// 팀 단위 교체 계획: 엑셀에 있는 팀의 과제만 엑셀 내용으로 바꾸고, 과제명이 같으면 같은 과제로 봄
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
  const plan = planImport(res), canApply = canEdit();
  const body = el('div', { style: 'display:grid;gap:12px' },
    el('p', {}, el('b', {}, f.name), ` · ${res.sheet} 시트 · ${res.rows.length}행 · 인식 컬럼 ${res.known}개`),
    el('div', { class: 'tblwrap', style: 'max-height:300px;border:1px solid var(--line);border-radius:8px' }, el('table', {},
      el('thead', {}, el('tr', {}, ['팀', '엑셀 과제', '추가', '변경', '삭제', '동일'].map((h, i) => el('th', { class: i ? 'num' : '' }, h)))),
      el('tbody', {}, plan.summary.map((s) => el('tr', {}, el('td', {}, el('b', {}, s.team)), el('td', { class: 'num' }, s.total), el('td', { class: 'num' }, s.added),
        el('td', { class: 'num' }, s.changed), el('td', { class: 'num' + (s.removed ? ' c-bad' : '') }, s.removed), el('td', { class: 'num' }, s.kept)))))),
    el('ul', { class: 'notes', style: 'display:grid;gap:3px;color:var(--ink-2)' },
      el('li', {}, '엑셀에 있는 팀만 교체, 다른 팀 과제는 유지'),
      el('li', {}, '삭제: 저장 데이터에는 있으나 엑셀에는 없는 과제'),
      res.hasPriority ? null : el('li', {}, '엑셀에 우선순위 컬럼 없음 → 기존 우선순위 유지'),
      plan.dups.length ? el('li', { class: 'c-bad' }, `같은 팀 안 중복 과제명 ${plan.dups.length}건은 첫 행만 반영: ${plan.dups.slice(0, 3).join(', ')}${plan.dups.length > 3 ? ' 외' : ''}`) : null,
      plan.skipped ? el('li', { class: 'c-bad' }, `팀 또는 과제명이 빈 행 ${plan.skipped}건 제외`) : null,
      res.unknown.length ? el('li', {}, `인식하지 못한 컬럼: ${res.unknown.join(', ')}`) : null,
      canApply ? null : el('li', {}, state.server ? '편집 비밀번호가 없어 서버 반영 불가 · 내 화면 미리보기만 가능' : '서버 없이 연 화면 · 내 화면 미리보기만 가능')));
  const choice = await modal({ title: '엑셀 업로드 미리보기', notes: ['팀 단위 교체 방식', '서버 반영 시 접속자 모두에게 표시'], body, wide: true,
    actions: [{ label: '취소' }, { label: '내 화면에서만 보기', value: 'local' }, canApply && plan.summary.length ? { label: '서버에 반영', cls: 'primary', value: 'apply' } : null].filter(Boolean) });
  if (choice === 'local') { setData('local', res.rows, [], null, today(), false, f.name); toast(`${f.name}을(를) 내 화면에서만 보여 줍니다.`); }
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
$('#overlay').addEventListener('click', () => { if ($('#modal').hidden) $('#drawer').querySelector('.dr-hd button')?.click(); });
let rt; window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => { if (state.tab === 'overview' || state.tab === 'schedule') render(); }, 150); });

// 다른 사람이 저장하면 20초 안에 반영 (입력 중이거나 내 파일 미리보기 중에는 건너뜀)
setInterval(async () => {
  if (!state.server || document.hidden || isBusy() || state.mode === 'local') return;
  let r; try { r = await api('api/version'); } catch (e) { return; }
  if (r.ok && r.data.version !== state.version) { await loadServer(true); toast('다른 사용자가 저장한 내용을 반영했습니다.'); }
}, 20000);

buildTabs();
loadServer(false).then((ok) => {
  if (!ok) setData('demo', SAMPLE, SAMPLE_HIST, mondayOf(SAMPLE_ASOF), SAMPLE_ASOF, false);
  const h = location.hash.slice(1);
  setTab(TABS.some(([k]) => k === h) ? h : 'overview');
});
})();
