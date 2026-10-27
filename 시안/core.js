/* 월 실적 도우미 — 계산 부분 (시안 A·B 공통). 인터넷으로 아무것도 보내지 않습니다. */
(function (g) {
  const pad = (n) => String(n).padStart(2, '0');
  const won = (n) => Math.round(n).toLocaleString('ko-KR');
  const pct = (n, d = 1) => (n * 100).toFixed(d) + '%';

  const SHEETS = {
    sales:    { label: '매출·수출입', types: ['date', 'text', 'text', 'text', 'num', 'num'], detect: (h) => h.includes('거래처') },
    purchase: { label: '구매',        types: ['date', 'text', 'text', 'num', 'num', 'num'],  detect: (h) => h.includes('공급처') },
    defect:   { label: '생산·불량',   types: ['date', 'line', 'text', 'num', 'num'],         detect: (h) => h.includes('라인') },
  };
  const DEFAULTS = { trim: true, date: true, num: true, line: true, fill: true, dup: true };
  const RULE_LABEL = {
    trim: '이름 앞뒤 빈칸 지우기',
    date: '날짜를 한 가지 모양으로',
    num:  '금액의 "원"·쉼표 없애기',
    line: '라인 이름 통일 (c라인 → C라인)',
    fill: '빈 금액은 수량×단가로 채우기',
    dup:  '똑같은 줄은 한 번만 세기',
  };

  function serialToISO(n) {
    const d = new Date(Date.UTC(1899, 11, 30) + n * 864e5);
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }
  function cleanDate(v) {
    if (/^\d{5}$/.test(v)) return serialToISO(+v);
    const m = v.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/);
    return m ? m[1] + '-' + pad(m[2]) + '-' + pad(m[3]) : v;
  }
  const toNum = (v) => { const n = Number(String(v).trim()); return isNaN(n) ? 0 : n; };
  const monthOf = (v) => (/^\d{4}-\d{2}/.test(v) ? v.slice(0, 7) : '날짜 모양 다름');

  function parse(text) {
    const lines = text.replace(/\r/g, '').split('\n').filter((l) => l.trim() !== '');
    if (!lines.length) return null;
    const header = lines[0].split('\t').map((s) => s.trim());
    const kind = Object.keys(SHEETS).find((k) => SHEETS[k].detect(header));
    if (!kind) return null;
    const rows = lines.slice(1).map((l) => {
      const c = l.split('\t');
      while (c.length < header.length) c.push('');
      return c.slice(0, header.length);
    });
    return { kind, header, rows };
  }

  /* 정리: 어디를 어떻게 바꿨는지 기억해 둔다 */
  function clean(text, opts) {
    opts = Object.assign({}, DEFAULTS, opts || {});
    const p = parse(text);
    if (!p) return null;
    const types = SHEETS[p.kind].types;
    const counts = { trim: 0, date: 0, num: 0, line: 0, fill: 0, dup: 0 };
    const seen = new Set();
    const rows = p.rows.map((raw) => {
      const cells = raw.map((r, i) => {
        let v = r;
        const kinds = [];
        if (opts.trim && v !== v.trim()) { v = v.trim(); kinds.push('trim'); }
        const t = types[i];
        if (t === 'date' && opts.date) { const nv = cleanDate(v.trim()); if (nv !== v) { v = nv; kinds.push('date'); } }
        if (t === 'num' && opts.num) { const nv = v.replace(/[,원\s]/g, ''); if (nv !== v) { v = nv; kinds.push('num'); } }
        if (t === 'line' && opts.line) {
          const nv = v.replace(/\s+/g, '').replace(/^[a-z]/, (m) => m.toUpperCase());
          if (nv !== v) { v = nv; kinds.push('line'); }
        }
        return { raw: r, v, kinds };
      });
      if (p.kind === 'purchase' && opts.fill && cells[5].v === '') {
        cells[5].v = String(toNum(cells[3].v) * toNum(cells[4].v));
        cells[5].kinds.push('fill');
      }
      cells.forEach((c) => c.kinds.forEach((k) => counts[k]++));
      const key = cells.map((c) => c.v).join('|');
      const dup = opts.dup && seen.has(key);
      seen.add(key);
      if (dup) counts.dup++;
      return { cells, dup };
    });
    return { kind: p.kind, header: p.header, rows, counts, kept: rows.filter((r) => !r.dup).length, total: rows.length, opts };
  }

  const group = (list, keyFn, valFns) => {
    const m = new Map();
    list.forEach((x) => {
      const k = keyFn(x);
      if (!m.has(k)) m.set(k, {});
      const o = m.get(k);
      Object.keys(valFns).forEach((n) => { o[n] = (o[n] || 0) + valFns[n](x); });
    });
    return m;
  };

  /* 집계: 정리된 표 → 월별·구분별 합계 */
  function summarize(res) {
    const rows = res.rows.filter((r) => !r.dup).map((r) => r.cells.map((c) => c.v));
    const out = { kind: res.kind, kept: rows.length };
    if (res.kind === 'sales') {
      const mm = group(rows, (r) => monthOf(r[0]), { 수출: (r) => (r[2] === '수출' ? toNum(r[5]) : 0), 내수: (r) => (r[2] === '내수' ? toNum(r[5]) : 0) });
      out.months = [...mm.keys()].sort().map((m) => ({ m, 수출: mm.get(m).수출, 내수: mm.get(m).내수, 합계: mm.get(m).수출 + mm.get(m).내수 }));
      out.total = out.months.reduce((a, x) => ({ 수출: a.수출 + x.수출, 내수: a.내수 + x.내수, 합계: a.합계 + x.합계 }), { 수출: 0, 내수: 0, 합계: 0 });
      const by = (idx) => [...group(rows, (r) => r[idx], { v: (r) => toNum(r[5]) }).entries()].map(([k, o]) => ({ k, v: o.v })).sort((a, b) => b.v - a.v);
      out.clients = by(1);
      out.items = by(3);
    } else if (res.kind === 'purchase') {
      const mm = group(rows, (r) => monthOf(r[0]), { v: (r) => toNum(r[5]) });
      out.months = [...mm.keys()].sort().map((m) => ({ m, 합계: mm.get(m).v }));
      out.total = out.months.reduce((a, x) => a + x.합계, 0);
      const by = (idx) => [...group(rows, (r) => r[idx], { v: (r) => toNum(r[5]) }).entries()].map(([k, o]) => ({ k, v: o.v })).sort((a, b) => b.v - a.v);
      out.suppliers = by(1);
      out.materials = by(2);
    } else {
      const f = { prod: (r) => toNum(r[3]), bad: (r) => toNum(r[4]) };
      const lm = group(rows, (r) => r[1], f);
      out.lines = [...lm.entries()].map(([k, o]) => ({ k, prod: o.prod, bad: o.bad, rate: o.prod ? o.bad / o.prod : 0 })).sort((a, b) => a.k.localeCompare(b.k));
      const mm = group(rows, (r) => monthOf(r[0]), f);
      out.months = [...mm.keys()].sort().map((m) => ({ m, prod: mm.get(m).prod, bad: mm.get(m).bad, rate: mm.get(m).prod ? mm.get(m).bad / mm.get(m).prod : 0 }));
      const t = out.months.reduce((a, x) => ({ prod: a.prod + x.prod, bad: a.bad + x.bad }), { prod: 0, bad: 0 });
      out.total = { prod: t.prod, bad: t.bad, rate: t.prod ? t.bad / t.prod : 0 };
      out.worst = out.lines.slice().sort((a, b) => b.rate - a.rate)[0];
    }
    return out;
  }

  const mlabel = (m) => (/^\d{4}-\d{2}$/.test(m) ? Number(m.slice(5)) + '월' : m);

  /* AI에게 보낼 질문: 이름은 별명으로 바꾸고, 집계표만 담는다 */
  function buildPrompt(S) {
    const alias = (list, prefix) => { const m = {}; list.forEach((x, i) => (m[x.k] = prefix + String.fromCharCode(65 + i))); return m; };
    const lines = [];
    lines.push('당신은 제조업 경영지원팀의 데이터 분석 도우미입니다.');
    lines.push('아래는 7~9월 월별 집계표입니다. 거래처·공급처 이름은 별명으로 바꿨습니다.');
    lines.push('');
    if (S.sales) {
      const a = alias(S.sales.clients, '거래처');
      lines.push('[매출·수출입 월별 합계(원)]');
      S.sales.months.forEach((x) => lines.push(`- ${mlabel(x.m)}: 수출 ${won(x.수출)} / 내수 ${won(x.내수)} / 합계 ${won(x.합계)}`));
      lines.push('[거래처별 매출(원)]');
      S.sales.clients.forEach((x) => lines.push(`- ${a[x.k]}: ${won(x.v)}`));
      lines.push('');
    }
    if (S.defect) {
      lines.push('[라인별 불량률]');
      S.defect.lines.forEach((x) => lines.push(`- ${x.k}: 생산 ${won(x.prod)}개 중 불량 ${won(x.bad)}개 (${pct(x.rate)})`));
      lines.push('');
    }
    lines.push('요청:');
    lines.push('1. 숫자에서 눈에 띄는 변화나 이상한 점 3가지를 쉬운 말로 알려 주세요.');
    lines.push('2. 각 항목마다 어느 숫자를 보고 판단했는지 함께 적어 주세요.');
    lines.push('3. 모르는 것은 추측하지 말고 "자료만으로는 알 수 없음"이라고 써 주세요.');
    return lines.join('\n');
  }

  /* 경영진 보고 초안: 문장마다 어떤 계산에서 나온 숫자인지 함께 돌려준다 */
  function buildReport(S) {
    const items = [];
    if (S.sales && S.sales.months.length) {
      const ms = S.sales.months, last = ms[ms.length - 1], prev = ms[ms.length - 2];
      items.push({ text: `3분기(7~9월) 매출은 ${won(S.sales.total.합계)}원이며, 수출 ${pct(S.sales.total.수출 / S.sales.total.합계, 0)}·내수 ${pct(S.sales.total.내수 / S.sales.total.합계, 0)}입니다.`, calc: `수출 ${won(S.sales.total.수출)} + 내수 ${won(S.sales.total.내수)} = ${won(S.sales.total.합계)}` });
      if (prev) {
        const d = last.합계 / prev.합계 - 1;
        items.push({ text: `${mlabel(last.m)} 매출은 ${won(last.합계)}원으로 ${mlabel(prev.m)}보다 ${pct(Math.abs(d))} ${d >= 0 ? '늘었습니다' : '줄었습니다'}.`, calc: `${won(last.합계)} ÷ ${won(prev.합계)} − 1 = ${(d * 100).toFixed(1)}%` });
      }
      const top = S.sales.clients[0];
      items.push({ text: `매출이 가장 큰 거래처 비중은 ${pct(top.v / S.sales.total.합계, 0)}입니다. 한 곳에 몰려 있는지 확인이 필요합니다.`, calc: `${won(top.v)} ÷ ${won(S.sales.total.합계)} = ${(top.v / S.sales.total.합계 * 100).toFixed(1)}%` });
    }
    if (S.purchase && S.purchase.months.length) {
      const ms = S.purchase.months, last = ms[ms.length - 1], prev = ms[ms.length - 2];
      items.push({ text: `구매 금액은 3분기 ${won(S.purchase.total)}원이고, ${mlabel(last.m)}은 ${won(last.합계)}원입니다${prev ? ` (${mlabel(prev.m)} ${won(prev.합계)}원)` : ''}.`, calc: `월별 합계를 더함 = ${won(S.purchase.total)}` });
      const top = S.purchase.materials[0];
      items.push({ text: `가장 많이 산 자재는 ${top.k}(${won(top.v)}원)입니다.`, calc: `자재별 합계 중 최대 = ${won(top.v)}` });
    }
    if (S.defect && S.defect.lines.length) {
      const w = S.defect.worst;
      items.push({ text: `전체 불량률은 ${pct(S.defect.total.rate)}이고, ${w.k}이 ${pct(w.rate)}로 가장 높습니다.`, calc: `전체 ${won(S.defect.total.bad)} ÷ ${won(S.defect.total.prod)}, ${w.k} ${won(w.bad)} ÷ ${won(w.prod)}` });
    }
    return items;
  }

  /* 그래프 (SVG). 색은 화면의 CSS 변수를 따른다 */
  function chartBars(months, keys, colors, o) {
    o = Object.assign({ w: 640, h: 240, unit: 1e6, suffix: '백만', stack: true }, o || {});
    const L = 56, R = 12, T = 30, B = 34, W = o.w - L - R, H = o.h - T - B;
    const tot = (m) => (o.stack ? keys.reduce((a, k) => a + m[k], 0) : Math.max(...keys.map((k) => m[k])));
    const max = Math.max(...months.map(tot), 1);
    const step = Math.pow(10, Math.floor(Math.log10(max / o.unit))) * o.unit;
    const top = Math.ceil(max / step) * step;
    const y = (v) => T + H - (v / top) * H;
    let s = `<svg viewBox="0 0 ${o.w} ${o.h}" role="img" aria-label="월별 막대 그래프" style="width:100%;height:auto;display:block">`;
    for (let v = 0; v <= top + 1; v += step) {
      s += `<line x1="${L}" x2="${L + W}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)" stroke-width="1"/>`;
      s += `<text x="${L - 8}" y="${y(v) + 4}" text-anchor="end" font-size="12" fill="var(--sub)">${v / o.unit}${v === 0 ? '' : ''}</text>`;
    }
    s += `<text x="0" y="12" font-size="12" fill="var(--sub)">단위: ${o.suffix}원</text>`;
    const gw = W / months.length;
    months.forEach((m, i) => {
      const bw = Math.min(64, gw * (o.stack ? 0.5 : 0.7 / keys.length));
      let acc = 0;
      keys.forEach((k, j) => {
        const v = m[k];
        const x = o.stack ? L + gw * i + (gw - bw) / 2 : L + gw * i + (gw - bw * keys.length) / 2 + bw * j;
        const y0 = o.stack ? y(acc + v) : y(v);
        const hh = o.stack ? y(acc) - y(acc + v) : y(0) - y(v);
        s += `<rect class="bar" x="${x}" y="${y0}" width="${bw - (o.stack ? 0 : 2)}" height="${Math.max(hh, 0)}" fill="${colors[j]}"><title>${mlabel(m.m)} ${k}: ${won(v)}원</title></rect>`;
        acc += v;
      });
      s += `<text x="${L + gw * i + gw / 2}" y="${o.h - 12}" text-anchor="middle" font-size="13" fill="var(--text)">${mlabel(m.m)}</text>`;
    });
    return s + '</svg>';
  }

  function chartHBars(list, valKey, labelKey, color, o) {
    o = Object.assign({ w: 640, rowH: 34, fmt: (v) => v, max: null }, o || {});
    const L = 84, R = 70, W = o.w - L - R;
    const max = o.max || Math.max(...list.map((x) => x[valKey]), 1);
    const h = list.length * o.rowH + 8;
    let s = `<svg viewBox="0 0 ${o.w} ${h}" role="img" aria-label="가로 막대 그래프" style="width:100%;height:auto;display:block">`;
    list.forEach((x, i) => {
      const y = 4 + i * o.rowH, bw = (x[valKey] / max) * W;
      s += `<text x="${L - 10}" y="${y + o.rowH / 2 + 4}" text-anchor="end" font-size="14" fill="var(--text)">${x[labelKey]}</text>`;
      s += `<rect class="bar" x="${L}" y="${y + 6}" width="${bw}" height="${o.rowH - 14}" fill="${typeof color === 'function' ? color(x) : color}"><title>${x[labelKey]}: ${o.fmt(x[valKey])}</title></rect>`;
      s += `<text x="${L + bw + 8}" y="${y + o.rowH / 2 + 4}" font-size="13" fill="var(--sub)" class="mono">${o.fmt(x[valKey])}</text>`;
    });
    return s + '</svg>';
  }

  g.SNC = { SHEETS, DEFAULTS, RULE_LABEL, parse, clean, summarize, buildPrompt, buildReport, chartBars, chartHBars, won, pct, mlabel };
})(window);
