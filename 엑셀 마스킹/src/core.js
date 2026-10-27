/* 엑셀 마스킹 핵심 로직 — 화면과 분리해 자동 시험이 가능하게 한다. AI·네트워크 코드 없음(규칙 방식). */
(function (root) {
  "use strict";

  const PATTERN_ORDER = ["email", "rrn", "biz", "phone", "acct"];
  const RE = {
    email: /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/,
    rrn: /\d{6}-[1-4]\d{6}/,
    biz: /\d{3}-\d{2}-\d{5}/,
    phone: /(?:\+82[- ]?)?0\d{1,2}-\d{3,4}-\d{4}/,
    acct: /\d{2,6}-\d{2,6}-\d{3,8}(?:-\d{1,4})?/,
  };
  const WORD = { person: "담당자", phone: "전화", email: "메일", rrn: "주민번호", biz: "사업자번호", acct: "계좌", own: "비공개" };
  const TITLES = "과장|대리|팀장|부장|차장|사원|주임|실장|본부장|사장|대표이사";

  const rx = (s) => s.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
  const norm = (s) => String(s).replace(/\s+/g, "").toLowerCase();
  const fuzzy = (name) => [...name.replace(/\s+/g, "")].map(rx).join("\\s*");
  const full = (k) => new RegExp("^(?:" + RE[k].source + ")$");
  /* "찾을말 → 바꿀말" 한 줄을 둘로 나눈다. 화살표가 없으면 바꿀말은 없음 */
  function splitArrow(raw) {
    const m = String(raw).split(/\s*(?:→|->|=>)\s*/);
    const find = (m[0] || "").trim();
    const alias = m.length > 1 ? m.slice(1).join(" ").trim() : "";
    return { find, alias };
  }
  /* 메일의 앞/뒤 조각(minsu.kim@ · @novatech.co.kr)은 주소 전체로 넓혀서 찾고, * 는 "아무 글자나"로 읽는다 (*@* = 모든 메일, 010-* = 010-으로 시작하는 번호) */
  const EXT_TAIL = "[\\w-]+(?:\\.[\\w-]+)+", EXT_HEAD = "[\\w.+-]+";
  function nameRe(o) {
    const n = o.n;
    const parts = n.split("*");
    // 별(*) 바로 앞이 @ 이면 메일 뒷부분, 바로 뒤가 @ 이면 메일 앞부분, 그 밖에는 글자·숫자·-
    const body = parts.map((part, i) => {
      const core = fuzzy(part);
      if (i === parts.length - 1) return core;
      const star = part.endsWith("@") ? EXT_TAIL : parts[i + 1].startsWith("@") ? EXT_HEAD : "(?:[\\w-]|[,.](?=\\d))*";
      return core + star;
    }).join("");
    return (n.startsWith("@") ? EXT_HEAD : "") + body + (n.endsWith("@") ? EXT_TAIL : "");
  }

  function aliasOf(type, n, fmt) {
    const br = fmt === "br";
    if (type === "company") {
      let L = "";
      let i = n - 1;
      do { L = String.fromCharCode(65 + (i % 26)) + L; i = Math.floor(i / 26) - 1; } while (i >= 0);
      return br ? "[회사" + L + "]" : L + "사";
    }
    return br ? "[" + WORD[type] + n + "]" : WORD[type] + n;
  }

  /* 가리개: 같은 이름은 같은 별명(띄어쓰기·대소문자 무시), 긴 이름부터 찾기 */
  function createMasker(opts) {
    const on = Object.assign({ company: true, person: true, phone: true, email: true, rrn: true, biz: true, acct: false, own: true }, opts.on || {});
    const fmt = opts.fmt || "plain";
    const reserved = opts.reserved || "";
    const map = new Map();
    const cnt = {};
    const names = [];
    const seen = new Set();
    for (const type of ["company", "person", "own"]) {
      if (!on[type]) continue;
      for (const raw of (opts.lists && opts.lists[type]) || []) {
        const sp = type === "own" ? splitArrow(raw) : { find: String(raw).trim(), alias: "" };
        const n = sp.find;
        const key = type + "|" + norm(n);
        if (!n || !n.replace(/[*\s]/g, "") || seen.has(key)) continue;
        seen.add(key);
        const o = { n, type, k: norm(n), alias: sp.alias };
        o.ext = type === "own" && (n.startsWith("@") || n.endsWith("@") || n.includes("*"));
        if (o.ext) o.full = new RegExp("^(?:" + nameRe(o) + ")$", "i");
        names.push(o);
      }
    }
    for (const it of opts.seed || []) {
      if (!it || !it.alias || !it.orig || !on[it.type]) continue;
      if (["company", "person", "own"].includes(it.type)) {
        const key = it.type + "|" + norm(it.orig);
        if (!seen.has(key)) { seen.add(key); names.push({ n: it.orig, type: it.type, k: norm(it.orig) }); }
      }
    }
    names.sort((a, b) => b.k.length - a.k.length);
    const nameType = new Map();
    names.forEach((o) => { if (!nameType.has(o.k)) nameType.set(o.k, o.type); });
    const kinds = PATTERN_ORDER.filter((k) => on[k]);
    // 내가 정한 말은 번호 모양(메일 등)보다 먼저 찾는다
    const ownNames = names.filter((o) => o.type === "own"), otherNames = names.filter((o) => o.type !== "own");
    // 거래처·사람 이름은 수천 개일 수 있어 정규식 하나에 다 넣으면 느려진다 → 공백 뺀 글에서 길이별로 잘라 목록(Set)에서 찾는다
    const parts = ownNames.map(nameRe).concat(kinds.map((k) => RE[k].source));
    const big = parts.length ? new RegExp(parts.join("|"), "gi") : null;
    const otherKeys = new Map();
    otherNames.forEach((o) => { if (o.k && !otherKeys.has(o.k)) otherKeys.set(o.k, o); });
    const otherLens = [...new Set([...otherKeys.keys()].map((k) => k.length))].sort((a, b) => b - a);
    const ownExact = new Map(), ownExt = [];
    ownNames.forEach((o) => { if (o.ext) ownExt.push(o); else if (!ownExact.has(o.k)) ownExact.set(o.k, o); });
    // str 에서 pos 이후 첫 이름(공백·대소문자 무시, 긴 이름 우선)의 {index, s}. 없으면 null
    function nextOther(str, pos) {
      if (!otherLens.length) return null;
      const t = [], at = [];
      for (let i = pos; i < str.length; i++) {
        const ch = str[i];
        if (/\s/.test(ch)) continue;
        const lc = ch.toLowerCase();
        for (let u = 0; u < lc.length; u++) { t.push(lc[u]); at.push(i); }
      }
      const tt = t.join("");
      for (let i = 0; i < tt.length; i++) {
        for (const L of otherLens) {
          if (i + L > tt.length) continue;
          if (otherKeys.has(tt.substr(i, L))) return { index: at[i], s: str.slice(at[i], at[i + L - 1] + 1) };
        }
      }
      return null;
    }

    for (const it of opts.seed || []) {
      if (!it || !it.alias || !it.orig) continue;
      const key = it.type + "|" + norm(it.orig);
      if (map.has(key)) continue;
      map.set(key, { alias: it.alias, orig: it.orig, type: it.type });
      cnt[it.type] = (cnt[it.type] || 0) + 1;
    }
    // 별명이 원문에 이미 있는 글자와 겹치는지. 큰 파일에서 매번 통째로 찾지 않도록 한 번만 목록을 만든다
    let clash = null;
    const collides = (a) => {
      if (!reserved) return false;
      if (!clash) {
        clash = new Set();
        const words = [...new Set(Object.values(WORD))];
        for (const m of reserved.matchAll(/[A-Z]+사/g)) for (let i = 0; i < m[0].length - 1; i++) clash.add(m[0].slice(i));
        for (const m of reserved.matchAll(/\[회사[A-Z]+\]/g)) clash.add(m[0]);
        for (const w of words) {
          for (const m of reserved.matchAll(new RegExp(rx(w) + "\\d+", "g"))) for (let k = w.length + 1; k <= m[0].length; k++) clash.add(m[0].slice(0, k));
          for (const m of reserved.matchAll(new RegExp("\\[" + rx(w) + "\\d+\\]", "g"))) clash.add(m[0]);
        }
      }
      return clash.has(a);
    };
    const used = new Map();
    function take(type, orig, custom) {
      const key = type + "|" + norm(orig);
      let e = map.get(key);
      if (!e) {
        let a;
        if (custom) {
          // 같은 바꿀말을 여럿이 쓰면 2, 3… 을 붙여 되돌릴 때 헷갈리지 않게 한다
          const n = (used.get(custom) || 0) + 1;
          used.set(custom, n);
          a = n === 1 ? custom : custom + n;
        } else do { cnt[type] = (cnt[type] || 0) + 1; a = aliasOf(type, cnt[type], fmt); } while (collides(a));
        e = { alias: a, orig: orig.trim(), type };
        map.set(key, e);
      }
      return e.alias;
    }
    function text(str) {
      const segs = [];
      if (!big && !otherLens.length) return [{ s: String(str || "") }];
      if (!str) return [{ s: "" }];
      str = String(str);
      let last = 0, pos = 0, other, otherAt = -1;
      big && (big.lastIndex = 0);
      while (pos <= str.length) {
        let m = null;
        if (big) { big.lastIndex = pos; m = big.exec(str); }
        if (otherAt < pos) { other = nextOther(str, pos); otherAt = other ? other.index : Infinity; }
        let s, index;
        if (m && m.index <= otherAt) { if (m[0] === "") { pos = m.index + 1; continue; } s = m[0]; index = m.index; }
        else if (other && otherAt !== Infinity) { s = other.s; index = other.index; }
        else break;
        const nk = norm(s);
        const hit = ownExact.get(nk) || ownExt.find((o) => o.full.test(s));
        const type = hit ? "own" : nameType.get(nk) || kinds.find((k) => full(k).test(s));
        pos = index + s.length;
        if (!type) continue;
        if (index > last) segs.push({ s: str.slice(last, index) });
        segs.push({ m: take(type, s, hit && hit.alias), type, orig: s });
        last = pos;
      }
      if (last < str.length) segs.push({ s: str.slice(last) });
      return segs;
    }
    // "그대로" 칸처럼 다른 가리기는 끄고, 내가 정한 말만 적용
    const ownRe = ownNames.length ? new RegExp(ownNames.map(nameRe).join("|"), "gi") : null;
    function textOwn(str) {
      if (!ownRe || !str) return [{ s: String(str || "") }];
      str = String(str);
      const segs = [];
      let last = 0, m;
      ownRe.lastIndex = 0;
      while ((m = ownRe.exec(str))) {
        if (m[0] === "") { ownRe.lastIndex++; continue; }
        const nk = norm(m[0]);
        const hit = ownExact.get(nk) || ownExt.find((o) => o.full.test(m[0]));
        if (!hit) continue;
        if (m.index > last) segs.push({ s: str.slice(last, m.index) });
        segs.push({ m: take("own", m[0], hit.alias), type: "own", orig: m[0] });
        last = m.index + m[0].length;
      }
      if (last < str.length) segs.push({ s: str.slice(last) });
      return segs;
    }
    const items = () => [...map.values()];
    return { text, textOwn, hasOwn: ownNames.length > 0, whole: (type, str) => take(type, str), items, on, names, kinds };
  }

  const plain = (segs) => segs.map((x) => (x.m != null ? x.m : x.s)).join("");

  /* 되돌리기: 별명표(items)로 별명 → 원래 이름. 긴 별명부터, 숫자로 끝나면 뒤에 숫자가 더 붙은 건 제외 */
  function restoreText(items, str) {
    if (!items || !items.length || !str) return { segs: [{ s: String(str || "") }], count: 0 };
    const byAlias = new Map(items.map((i) => [i.alias, i.orig]));
    const aliases = [...byAlias.keys()].sort((a, b) => b.length - a.length);
    const re = new RegExp(aliases.map((a) => rx(a) + (/\d$/.test(a) ? "(?!\\d)" : "")).join("|"), "g");
    const segs = [];
    let last = 0, m, count = 0;
    while ((m = re.exec(str))) {
      if (m.index > last) segs.push({ s: str.slice(last, m.index) });
      segs.push({ m: byAlias.get(m[0]), alias: m[0] });
      last = m.index + m[0].length;
      count++;
    }
    if (last < str.length) segs.push({ s: str.slice(last) });
    return { segs, count };
  }

  /* 엑셀 칸 처리 종류: company 거래처로 / person 사람이름으로 / text 글 속에서 찾아 가리기 / keep 그대로 / del 지우기 */
  function detectAction(header) {
    const h = String(header == null ? "" : header).toLowerCase();
    if (/주민|사업자|계좌|연락|전화|휴대|핸드폰|tel|phone|mobile|메일|mail/.test(h)) return "text";
    if (/거래처|회사|업체|공급|고객|상호|납품|협력/.test(h)) return "company";
    if (/담당|성명|이름|대표|직원|작성자|name/.test(h)) return "person";
    if (/메모|비고|내용|특이|설명|의견|comment|note|remark/.test(h)) return "text";
    return "keep";
  }

  /* 엑셀이 "표 범위"를 100만 행까지 잡아 두는 경우가 있어서, 실제로 값이 있는 칸만으로 범위를 다시 잡는다 */
  const usedRange = (X, ws) => {
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    for (const k in ws) {
      if (k.charCodeAt(0) === 33) continue; // "!ref" 같은 설정 항목
      const a = X.utils.decode_cell(k);
      if (a.r < r0) r0 = a.r; if (a.r > r1) r1 = a.r;
      if (a.c < c0) c0 = a.c; if (a.c > c1) c1 = a.c;
    }
    return r1 < 0 ? { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } } : { s: { r: r0, c: c0 }, e: { r: r1, c: c1 } };
  };
  const isStr = (c) => c && (c.t === "s" || (c.t === "z" ? false : typeof c.v === "string"));

  function analyze(X, wb) {
    return wb.SheetNames.map((name, i) => {
      const ws = wb.Sheets[name];
      const hidden = !!(wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[i] && wb.Workbook.Sheets[i].Hidden);
      const ref = usedRange(X, ws);
      const cols = [];
      for (let c = ref.s.c; c <= ref.e.c; c++) {
        const hc = ws[X.utils.encode_cell({ r: ref.s.r, c })];
        const header = hc ? String(hc.w != null ? hc.w : hc.v) : "";
        const sample = [];
        for (let r = ref.s.r + 1; r <= ref.e.r && sample.length < 3; r++) {
          const cell = ws[X.utils.encode_cell({ r, c })];
          if (cell && cell.v != null && cell.v !== "") sample.push(String(cell.w != null ? cell.w : cell.v));
        }
        cols.push({ c, header: header || X.utils.encode_col(c) + "열", sample, action: detectAction(header) });
      }
      return { name, hidden, r0: ref.s.r, c0: ref.s.c, rows: ref.e.r - ref.s.r, cols };
    });
  }

  function cleanSheetName(name, used) {
    let n = String(name).replace(/[\[\]:*?\/\\]/g, "_").trim().slice(0, 31) || "시트";
    let base = n, i = 2;
    while (used.has(n.toLowerCase())) { n = base.slice(0, 28) + "_" + i++; }
    used.add(n.toLowerCase());
    return n;
  }

  /* 통합 문서 가리기. plan = {sheets:[{name,use,cols:{열번호:처리}}]}, opts = createMasker 옵션(+own, lists) */
  /* 일을 잘게 나눠 하는 버전. 몇 행마다 {phase,done,total}을 내보내 화면이 진행률을 그릴 수 있게 한다 */
  function* maskGen(X, wb, plan, opts, limit) {
    limit = limit || 200;
    const lists = { company: [...((opts.lists && opts.lists.company) || [])], person: [...((opts.lists && opts.lists.person) || [])], own: [...((opts.lists && opts.lists.own) || [])] };
    const used = plan.sheets.filter((p) => p.use && wb.Sheets[p.name]);
    const rowsOf = (ws) => { const rg = usedRange(X, ws); return rg.e.r - rg.s.r + 1; };
    const total = used.reduce((a, p) => a + rowsOf(wb.Sheets[p.name]), 0) * 2 || 1;
    const STEP = 200;
    let done = 0;
    let reserved = "";
    // 1) 거래처·사람 이름 칸의 값을 모아 이름 목록에 보태고, 겹침 방지용 글을 모은다
    for (const p of used) {
      const ws = wb.Sheets[p.name];
      const ref = usedRange(X, ws);
      reserved += "\n" + p.name;
      for (let r = ref.s.r; r <= ref.e.r; r++) {
        if (++done % STEP === 0) yield { phase: "이름 모으는 중", done, total };
        for (let c = ref.s.c; c <= ref.e.c; c++) {
          const cell = ws[X.utils.encode_cell({ r, c })];
          if (!isStr(cell)) continue;
          reserved += "\n" + cell.v;
          const act = p.cols[c];
          if (r > ref.s.r && (act === "company" || act === "person") && cell.v.trim()) lists[act].push(cell.v.trim());
        }
      }
    }
    done = total / 2;
    const masker = createMasker({ on: opts.on, fmt: opts.fmt, lists, reserved });
    const on = masker.on;
    opts.on_ = on;
    const report = { changed: 0, sheets: [], removed: { hiddenSheets: 0, skippedSheets: 0, comments: 0, links: 0, formulas: 0 } };
    const out = X.utils.book_new();
    const usedNames = new Set();
    const strings = [];

    for (let i = 0; i < wb.SheetNames.length; i++) {
      const name = wb.SheetNames[i];
      const p = plan.sheets.find((s) => s.name === name);
      const hidden = !!(wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[i] && wb.Workbook.Sheets[i].Hidden);
      if (!p || !p.use) { report.removed.skippedSheets++; if (hidden) report.removed.hiddenSheets++; continue; }
      const ws = wb.Sheets[name];
      const ref = usedRange(X, ws);
      const delCols = [];
      for (let c = ref.s.c; c <= ref.e.c; c++) if (p.cols[c] === "del") delCols.push(c);
      const newC = (c) => c - delCols.filter((d) => d < c).length;
      const ns = {};
      const headOf = (c) => { const h = ws[X.utils.encode_cell({ r: ref.s.r, c })]; return h && h.v != null ? String(h.v) : ""; };
      const prev = [];
      let maxR = ref.s.r, maxC = ref.s.c - 1;
      for (let r = ref.s.r; r <= ref.e.r; r++) {
        if (++done % STEP === 0) yield { phase: "이름 가리는 중", done, total };
        const prow = [];
        for (let c = ref.s.c; c <= ref.e.c; c++) {
          const cell = ws[X.utils.encode_cell({ r, c })];
          let act = p.cols[c] || "keep";
          if ((act === "company" || act === "person") && !opts.on_[act]) act = "keep";
          const disp = cell ? String(cell.w != null ? cell.w : cell.v == null ? "" : cell.v) : "";
          if (cell && cell.c) report.removed.comments++;
          if (cell && cell.l) report.removed.links++;
          if (act === "del") { prow.push({ v: disp, del: true }); continue; }
          if (!cell) { prow.push({ v: "" }); continue; }
          let nv = null;
          if (isStr(cell) && cell.v !== "") {
            if (r > ref.s.r && (act === "company" || act === "person")) nv = masker.whole(act, cell.v);
            else if (act === "text" || r === ref.s.r) {
              const segs = masker.text(cell.v);
              if (segs.some((x) => x.m != null)) nv = plain(segs);
            } else if (act === "keep" && masker.hasOwn) {
              const segs = masker.textOwn(cell.v);
              if (segs.some((x) => x.m != null)) nv = plain(segs);
            }
          }
          // 숫자·날짜 칸은 보이는 모양(₩22,000 등)에 "내가 정한 말"이 걸릴 때만 글자로 바꾼다
          else if (!isStr(cell) && (act === "text" || act === "keep") && masker.hasOwn && disp) {
            const segs = masker.textOwn(disp);
            if (segs.some((x) => x.m != null)) nv = plain(segs);
          }
          const nc = nv != null ? { t: "s", v: nv } : Object.assign({}, cell);
          if (nv == null) { delete nc.c; delete nc.l; if (nc.f) report.removed.formulas++; delete nc.f; delete nc.F; }
          const addr = X.utils.encode_cell({ r, c: newC(c) });
          ns[addr] = nc;
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, newC(c));
          if (nv != null) report.changed++;
          const outStr = nv != null ? nv : isStr(cell) ? cell.v : "";
          if (outStr) strings.push({ s: outStr, w: { sheet: name, col: X.utils.encode_col(c), head: r === ref.s.r ? "" : headOf(c), row: r + 1 } });
          prow.push({ v: nv != null ? nv : disp, ch: nv != null });
        }
        if (prev.length < limit) prev.push(prow);
      }
      ns["!ref"] = X.utils.encode_range({ s: { r: ref.s.r, c: ref.s.c }, e: { r: maxR, c: Math.max(maxC, ref.s.c) } });
      if (ws["!cols"]) ns["!cols"] = ws["!cols"].filter((_, idx) => !delCols.includes(idx));
      const segsN = masker.text(name);
      const newName = cleanSheetName(plain(segsN), usedNames);
      strings.push({ s: newName, w: { sheet: name, tab: true } });
      if (segsN.some((x) => x.m != null)) report.changed++;
      X.utils.book_append_sheet(out, ns, newName);
      report.sheets.push({ name: newName, origName: name, preview: prev, totalRows: ref.e.r - ref.s.r + 1, c0: ref.s.c });
    }
    out.Props = {};
    report.items = masker.items();
    report.unique = report.items.length;
    yield { phase: "남은 이름 점검 중", done: total, total, last: true };
    report.warnings = scan(strings, masker);
    return { wb: out, report, masker };
  }

  /* 한 번에 끝까지 (시험·작은 파일용) */
  function maskWorkbook(X, wb, plan, opts, limit) {
    const g = maskGen(X, wb, plan, opts, limit);
    for (;;) { const r = g.next(); if (r.done) return r.value; }
  }

  /* 조금씩 나눠서 (큰 파일용). 약 40ms마다 쉬면서 화면을 갱신하게 한다. token.cancelled 가 true 면 멈추고 null */
  function maskWorkbookAsync(X, wb, plan, opts, limit, onProgress, token) {
    return new Promise((resolve, reject) => {
      const g = maskGen(X, wb, plan, opts, limit);
      const step = () => {
        try {
          const t0 = Date.now();
          for (;;) {
            if (token && token.cancelled) { resolve(null); return; }
            const r = g.next();
            if (r.done) { resolve(r.value); return; }
            if (Date.now() - t0 > 40) { if (onProgress) onProgress(r.value); setTimeout(step, 0); return; }
          }
        } catch (e) { reject(e); }
      };
      step();
    });
  }

  /* 가린 뒤 남은 것 점검: 번호 모양, 목록 이름, 이름 같은 글자 */
  function scan(strings, masker) {
    // strings: 글 한 줄 또는 {s:글, w:위치}. 위치가 있으면 어느 시트·칸·몇 행인지 같이 알려 준다
    const items = strings.map((x) => (typeof x === "string" ? { s: x } : x));
    const all = items.map((x) => x.s).join("\n");
    const out = [];
    const where = (re) => {
      const g = new Map();
      for (const it of items) {
        if (!it.w) continue;
        re.lastIndex = 0;
        if (!re.test(it.s)) continue;
        const key = it.w.sheet + "|" + (it.w.tab ? "탭" : it.w.col);
        if (!g.has(key)) g.set(key, { w: it.w, rows: [] });
        if (!it.w.tab) g.get(key).rows.push(it.w.row);
      }
      return [...g.values()].slice(0, 6).map(({ w, rows }) => w.tab ? w.sheet + " (시트 이름)" : w.sheet + " · " + w.col + "열" + (w.head ? "(" + w.head + ")" : "(제목 줄)") + " · " + rows.slice(0, 4).join(", ") + "행" + (rows.length > 4 ? " 외 " + (rows.length - 4) + "곳" : ""));
    };
    const labels = { phone: "전화번호", email: "이메일", rrn: "주민등록번호", biz: "사업자등록번호", acct: "계좌번호처럼 생긴 글자" };
    for (const k of ["email", "rrn", "biz", "phone", "acct"]) {
      if (k === "acct" && !masker.on.acct) continue;
      const found = all.match(new RegExp(RE[k].source, "g"));
      if (found) out.push({ kind: k, text: labels[k] + " 모양이 " + found.length + "개 그대로 남았어요", examples: [...new Set(found)].slice(0, 4), where: where(new RegExp(RE[k].source)) });
    }
    // 이름이 수천 개여도 느려지지 않게: 칸마다 공백을 없앤 글에서 이름 길이별로 잘라 목록(Set)에서 찾는다
    const keyOf = new Map();
    masker.names.forEach((o) => { if (o.k && !keyOf.has(o.k)) keyOf.set(o.k, o); });
    const lens = [...new Set([...keyOf.keys()].map((k) => k.length))];
    const hitKeys = new Set();
    for (const it of items) {
      if (!it.s || hitKeys.size === keyOf.size) continue;
      const t = norm(it.s);
      for (const L of lens) for (let i = 0; i + L <= t.length; i++) { const k = t.substr(i, L); if (keyOf.has(k)) hitKeys.add(k); }
    }
    const left = [...hitKeys].map((k) => keyOf.get(k));
    if (left.length) out.push({ kind: "name", text: "가리기로 한 이름이 남았어요", examples: [...new Set(left.map((o) => o.n))].slice(0, 6), where: where(new RegExp(left.map((o) => fuzzy(o.n)).join("|"), "i")) });
    const ko = all.match(new RegExp("[가-힣]{2,4}\\s?(?:" + TITLES + ")", "g"));
    if (ko) out.push({ kind: "guess", text: "이름 같아 보이는 글자가 남았어요(목록에 없는 이름일 수 있어요)", examples: [...new Set(ko)].slice(0, 6), where: where(new RegExp("[가-힣]{2,4}\\s?(?:" + TITLES + ")")) });
    return out;
  }

  /* 별명표 파일 */
  function exportMap(items) {
    return JSON.stringify({ app: "이름가리개", version: 1, created: new Date().toISOString(), note: "이 파일에는 진짜 이름이 들어 있어요. AI에 올리지 마세요.", items }, null, 2);
  }
  function importMap(json) {
    const o = typeof json === "string" ? JSON.parse(json) : json;
    if (!o || o.app !== "이름가리개" || !Array.isArray(o.items)) throw new Error("엑셀 마스킹 별명표 파일이 아니에요");
    return o.items.filter((i) => i && i.alias && i.orig);
  }

  const api = { RE, createMasker, restoreText, maskWorkbook, maskWorkbookAsync, analyze, scan, detectAction, exportMap, importMap, plain, rx };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.MaskCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
