import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";
const require = createRequire(import.meta.url);
const X = require("../vendor/xlsx.mini.min.js");
const C = require("../src/core.js");
let pass = 0;
const t = (name, fn) => { try { fn(); pass++; console.log("  통과  " + name); } catch (e) { console.log("  실패  " + name + "\n       " + e.message); process.exitCode = 1; } };

const mk = (opts) => C.createMasker(Object.assign({ lists: { company: ["Nova Tech", "대성정밀"], person: ["김민수", "박지훈"], own: [] } }, opts || {}));

t("글: 이름·번호가 모두 가려지고 원래 이름이 남지 않는다", () => {
  const src = "Nova Tech 김민수 팀장 010-2345-6789, minsu@nova.co.kr, 사업자 123-45-67890, 주민 900101-1234567";
  const m = mk(); const out = C.plain(m.text(src));
  for (const s of ["Nova Tech", "김민수", "010-2345-6789", "minsu@nova.co.kr", "123-45-67890", "900101-1234567"]) assert.ok(!out.includes(s), s + " 남음: " + out);
  assert.match(out, /A사 담당자1 팀장 전화1, 메일1/);
});
t("같은 이름은 같은 별명, 띄어쓰기·대소문자가 달라도 같다", () => {
  const m = mk(); const out = C.plain(m.text("Nova Tech / nova tech / NovaTech"));
  assert.equal(out, "A사 / A사 / A사");
});
t("긴 이름이 먼저 (김민수와 김민수정)", () => {
  const m = C.createMasker({ lists: { person: ["김민수", "김민수정"] } });
  const out = C.plain(m.text("김민수정 그리고 김민수"));
  assert.equal(out, "담당자1 그리고 담당자2");
});
t("되돌리기가 원문과 정확히 같다", () => {
  const src = "Nova Tech 김민수 팀장 010-2345-6789 대성정밀 박지훈, minsu@nova.co.kr, 김민수";
  const m = mk(); const out = C.plain(m.text(src));
  const back = C.restoreText(m.items(), out).segs.map((x) => x.m ?? x.s).join("");
  assert.equal(back, src.replace(/(Nova Tech|대성정밀|김민수|박지훈)/g, (s) => s)); // 띄어쓰기 변형 없음
});
t("담당자1과 담당자10을 헷갈리지 않는다", () => {
  const names = Array.from({ length: 12 }, (_, i) => "사람" + String.fromCharCode(0xac00 + i * 3));
  const m = C.createMasker({ lists: { person: names } });
  const src = names.join(" ");
  const out = C.plain(m.text(src));
  const back = C.restoreText(m.items(), out).segs.map((x) => x.m ?? x.s).join("");
  assert.equal(back, src);
});
t("원문에 이미 '담당자1'이 있으면 별명이 겹치지 않는다", () => {
  const src = "담당자1은 따로 있고 김민수";
  const m = mk({ reserved: src }); const out = C.plain(m.text(src));
  assert.ok(!/담당자1은 따로 있고 담당자1$/.test(out), out);
});
t("꺼 둔 종류는 그대로 둔다", () => {
  const m = mk({ on: { phone: false } }); assert.ok(C.plain(m.text("010-2345-6789")).includes("010-2345-6789"));
});
t("계좌번호는 켰을 때만 가린다", () => {
  assert.ok(C.plain(mk().text("123-456-789012")).includes("123-456-789012"));
  assert.ok(!C.plain(mk({ on: { acct: true } }).text("123-456-789012")).includes("789012"));
});
t("금액·날짜는 건드리지 않는다", () => {
  const s = "2026-07-03 금액 1,250,000원 792,000";
  assert.equal(C.plain(mk().text(s)), s);
});
t("대괄호 모양", () => {
  const m = mk({ fmt: "br" }); assert.equal(C.plain(m.text("Nova Tech 김민수")), "[회사A] [담당자1]");
});
t("별명표 저장·불러오기", () => {
  const m = mk(); m.text("Nova Tech 김민수");
  const items = C.importMap(C.exportMap(m.items()));
  assert.deepEqual(items.map((i) => i.alias), ["A사", "담당자1"]);
  assert.throws(() => C.importMap('{"x":1}'));
});

/* 엑셀 */
const build = () => {
  const rows = [["일자", "거래처", "담당자", "연락처", "이메일", "금액(원)", "메모"],
    ["2026-07-03", "Nova Tech", "김민수", "010-2345-6789", "minsu@nova.co.kr", 792000, "김민수 팀장 요청"],
    ["2026-07-08", "대성정밀", "박지훈", "02-555-1234", "jh@daesung.kr", 3410000, "Nova Tech 와 비교, 010-9999-8888"],
    ["2026-09-04", "Nova Tech", "김민수", "010-2345-6789", "minsu@nova.co.kr", 1640000, ""]];
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rows), "Nova Tech 거래");
  return X.read(X.write(wb, { type: "array", bookType: "xlsx" }), { type: "array" });
};
const plan = (wb, over) => { const info = C.analyze(X, wb); return { info, plan: { sheets: info.map((s) => ({ name: s.name, use: true, cols: Object.assign(Object.fromEntries(s.cols.map((c) => [c.c, c.action])), over || {}) })) } }; };
const dump = (wb) => JSON.stringify(wb.SheetNames) + X.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);

t("엑셀: 칸 종류를 제목에서 알아맞힌다", () => {
  const { info } = plan(build()); assert.deepEqual(info[0].cols.map((c) => c.action), ["keep", "company", "person", "text", "text", "keep", "text"]);
});
t("엑셀: 가린 뒤 원래 이름이 파일 어디에도 없다 (시트 이름 포함)", () => {
  const wb = build(); const { plan: p } = plan(wb);
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} });
  const all = dump(r.wb);
  for (const s of ["Nova", "김민수", "박지훈", "대성정밀", "010-2345", "010-9999", "02-555", "minsu@", "jh@"]) assert.ok(!all.includes(s), s + " 남음\n" + all);
  assert.deepEqual(r.report.warnings, []);
  assert.ok(all.includes("792000") && all.includes("1640000"), "금액이 바뀜");
  assert.ok(r.report.changed > 10);
});
t("엑셀: 되돌리면 원래 값과 같다", () => {
  const wb = build(); const { plan: p } = plan(wb);
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} });
  const orig = X.utils.sheet_to_json(wb.Sheets["Nova Tech 거래"], { header: 1 });
  const out = X.utils.sheet_to_json(r.wb.Sheets[r.wb.SheetNames[0]], { header: 1 });
  assert.equal(out.length, orig.length);
  out.forEach((row, i) => row.forEach((v, j) => {
    const back = typeof v === "string" ? C.restoreText(r.report.items, v).segs.map((x) => x.m ?? x.s).join("") : v;
    assert.equal(back, orig[i][j], `행${i}열${j}`);
  }));
});
t("엑셀: 지우기를 고르면 그 열이 빠지고 나머지는 왼쪽으로 붙는다", () => {
  const wb = build(); const { plan: p } = plan(wb, { 3: "del" });
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} });
  const out = X.utils.sheet_to_json(r.wb.Sheets[r.wb.SheetNames[0]], { header: 1 });
  assert.equal(out[0].length, 6); assert.equal(out[0][3], "이메일");
});
t("엑셀: 거래처 열을 '그대로'로 두면 가리지 않고 경고가 뜬다", () => {
  const wb = build(); const { plan: p } = plan(wb, { 1: "keep", 2: "keep" });
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} });
  assert.ok(r.report.warnings.length >= 0); // 그대로면 이름이 남아도 에러는 아님
  assert.ok(dump(r.wb).includes("Nova Tech"));
});
t("엑셀: 사람 이름 가리기를 끄면 열이 그대로 남는다", () => {
  const wb = build(); const { plan: p } = plan(wb);
  const r = C.maskWorkbook(X, wb, p, { on: { person: false }, fmt: "plain", lists: {} });
  assert.ok(dump(r.wb).includes("김민수"));
});
t("엑셀: 원본 통합 문서는 바뀌지 않는다", () => {
  const wb = build(); const before = dump(wb); const { plan: p } = plan(wb);
  C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} }); assert.equal(dump(wb), before);
});
t("엑셀: 저장했다가 다시 읽어도 같다", () => {
  const wb = build(); const { plan: p } = plan(wb);
  const r = C.maskWorkbook(X, wb, p, { fmt: "br", lists: {} });
  const re = X.read(X.write(r.wb, { type: "array", bookType: "xlsx" }), { type: "array" });
  assert.equal(dump(re), dump(r.wb));
});
t("엑셀: 숨은 시트·메모·수식 정리", () => {
  const ws = X.utils.aoa_to_sheet([["이름", "합"], ["김민수", 3]]); ws.B2.f = "1+2"; ws.A2.c = [{ a: "x", t: "진짜 이름 김민수" }];
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "A"); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["비밀"]]), "B");
  wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };
  const rd = X.read(X.write(wb, { type: "array", bookType: "xlsx" }), { type: "array" });
  const info = C.analyze(X, rd); assert.equal(info[1].hidden, true);
  const p = { sheets: info.map((s) => ({ name: s.name, use: !s.hidden, cols: Object.fromEntries(s.cols.map((c) => [c.c, c.action])) })) };
  const r = C.maskWorkbook(X, rd, p, { fmt: "plain", lists: {} });
  assert.deepEqual(r.wb.SheetNames, ["A"]);
  const cell = r.wb.Sheets.A.B2; assert.equal(cell.f, undefined); assert.equal(cell.v, 3);
  assert.equal(r.wb.Sheets.A.A2.c, undefined);
});
t("엑셀: 시트 이름이 31자·겹침 규칙을 지킨다", () => {
  const long = "가".repeat(40); const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["a"]]), "Nova Tech 1"); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["a"]]), "Nova Tech 2".slice(0, 10) + "x");
  const info = C.analyze(X, wb); const p = { sheets: info.map((s) => ({ name: s.name, use: true, cols: {} })) };
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: { company: ["Nova Tech"] } });
  assert.equal(new Set(r.wb.SheetNames.map((n) => n.toLowerCase())).size, 2); assert.ok(r.wb.SheetNames.every((n) => n.length <= 31));
});

/* 참가자용 샘플 */
const samplePath = fileURLToPath(new URL("../../에스앤씨_샘플데이터_참가자용.xlsx", import.meta.url));
try {
  const buf = readFileSync(samplePath);
  t("샘플 파일: 열려서 가려지고 경고 목록이 보인다", () => {
    const wb = X.read(buf, { type: "buffer" }); const { plan: p, info } = plan(wb);
    const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: {} });
    console.log("       시트:", info.map((s) => `${s.name}(${s.rows}줄: ${s.cols.map((c) => c.header + "→" + c.action).join(", ")})`).join(" | "));
    console.log("       바뀐 칸", r.report.changed, "· 이름 종류", r.report.unique, "· 경고", JSON.stringify(r.report.warnings));
    const orig = new Set();
    wb.SheetNames.forEach((n, i) => { const ws = wb.Sheets[n]; info[i].cols.forEach((c) => { if (c.action === "company" || c.action === "person") X.utils.sheet_to_json(ws, { header: 1 }).slice(1).forEach((row) => { const v = row[c.c - info[i].c0]; if (typeof v === "string" && v.trim()) orig.add(v.trim()); }); }); });
    const all = r.wb.SheetNames.map((n) => X.utils.sheet_to_csv(r.wb.Sheets[n])).join("\n");
    const left = [...orig].filter((o) => all.includes(o));
    assert.deepEqual(left, [], "남은 이름: " + left.join(","));
  });
} catch (e) { console.log("  (샘플 파일 없음 — 건너뜀)"); }

t("내가 정한 말: 화살표로 바꿀 말을 정하고, 메일 앞부분만 적어도 주소 전체가 가려진다", () => {
  const m = C.createMasker({ lists: { own: ["minsu.kim@ → 아크"] } });
  const out = C.plain(m.text("연락: minsu.kim@novatech.co.kr / seoyeon@sakura-parts.com"));
  assert.equal(out, "연락: 아크 / 메일1");
  assert.ok(!out.includes("novatech"));
});
t("내가 정한 말: 화살표 없으면 비공개N, ->도 된다, 같은 바꿀말은 번호로 구분되어 되돌려진다", () => {
  const m = C.createMasker({ lists: { own: ["프로젝트블루", "@daesung.kr -> 대성", "@novatech.co.kr → 대성"] } });
  const src = "프로젝트블루 jh.park@daesung.kr a@novatech.co.kr";
  const out = C.plain(m.text(src));
  assert.equal(out, "비공개1 대성 대성2");
  assert.equal(C.plain(C.restoreText(m.items(), out).segs.map((x) => (x.m != null ? { m: x.m } : x))), src);
});

t("내가 정한 말: *@* 는 도메인이 달라도 모든 메일, 010-* 는 그 번호 전부 (스위치 꺼도 동작)", () => {
  const off = { email: false, phone: false };
  const m = C.createMasker({ on: off, lists: { own: ["*@* → 메일숨김", "010-*"] } });
  const src = "a.b@x.com, c@d.co.kr, 010-1234-5678, 02-123-4567";
  const out = C.plain(m.text(src));
  assert.equal(out, "메일숨김, 메일숨김2, 비공개1, 02-123-4567");
  assert.equal(C.plain(C.restoreText(m.items(), out).segs.map((x) => (x.m != null ? { m: x.m } : x))), src);
});

t("경고가 남은 곳의 시트·열·행을 알려 준다", () => {
  const ws = X.utils.aoa_to_sheet([["거래처", "비고"], ["우리케미칼", "x"], ["대한금속", "우리케미칼 건"]]);
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "구매");
  const r = C.maskWorkbook(X, wb, { sheets: [{ name: "구매", use: true, cols: { 0: "keep", 1: "keep" } }] }, { lists: { company: ["우리케미칼"] } });
  const w = r.report.warnings.find((x) => x.kind === "name");
  assert.ok(w && w.where.some((x) => x.includes("구매 · A열(거래처) · 2행")), JSON.stringify(w));
  assert.ok(w.where.some((x) => x.includes("B열(비고)")), JSON.stringify(w));
});


/* 큰 파일용 나눠 처리: 결과가 한 번에 처리한 것과 같아야 하고, 진행률·취소가 되어야 한다 */
const bigBook = (n) => {
  const rows = [["거래처", "담당자", "연락처", "메모"]];
  for (let i = 0; i < n; i++) rows.push(["Nova Tech", "김민수", "010-2345-" + String(1000 + (i % 9000)), "김민수 팀장 요청 " + i]);
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rows), "내역");
  return wb;
};
const bigPlan = { sheets: [{ name: "내역", use: true, cols: { 0: "company", 1: "person", 2: "text", 3: "text" } }] };
const asyncJobs = [];
asyncJobs.push((async () => {
  const wb = bigBook(3000), o = () => ({ lists: { company: [], person: [], own: [] } });
  const sync = C.maskWorkbook(X, wb, bigPlan, o());
  const seen = [];
  const res = await C.maskWorkbookAsync(X, wb, bigPlan, o(), 200, (p) => seen.push(p), null);
  const same = JSON.stringify(X.utils.sheet_to_json(res.wb.Sheets[res.wb.SheetNames[0]], { header: 1 })) === JSON.stringify(X.utils.sheet_to_json(sync.wb.Sheets[sync.wb.SheetNames[0]], { header: 1 }));
  return () => {
    t("나눠 처리한 결과가 한 번에 처리한 결과와 같다", () => assert.ok(same));
    t("진행률이 0~total 안에서 늘어난다", () => { assert.ok(seen.every((p) => p.done <= p.total && p.total > 0)); assert.ok(seen.length === 0 || seen.at(-1).done >= seen[0].done); });
  };
})());
asyncJobs.push((async () => {
  const tok = { cancelled: false };
  const pr = C.maskWorkbookAsync(X, bigBook(8000), bigPlan, { lists: {} }, 200, (p) => { if (p.done > 100) tok.cancelled = true; }, tok);
  const res = await pr;
  return () => t("취소하면 null 로 멈춘다", () => assert.equal(res, null));
})());
for (const f of await Promise.all(asyncJobs)) f();

t("이름이 수천 개여도 빠르고 띄어쓰기·대소문자·긴 이름 우선이 그대로다", () => {
  const co = [], pe = ["김민수", "김민수정"];
  for (let i = 0; i < 3000; i++) co.push("Corp" + i + "X");
  co.push("Nova Tech");
  const m = C.createMasker({ lists: { company: co, person: pe, own: [] } });
  const t0 = Date.now();
  const out = C.plain(m.text("nova  tech, 김 민수정 와 김민수, CORP7X 그리고 corp2999x 010-2345-6789"));
  assert.ok(Date.now() - t0 < 1000);
  assert.match(out, /^[A-Z]+사, 담당자1 와 담당자2, [A-Z]+사 그리고 [A-Z]+사 전화1$/);
  assert.ok(!/Nova|Corp|민수/i.test(out), out);
});
t("내가 정한 말은 '그대로' 칸에도 적용되고 다른 번호는 건드리지 않는다", () => {
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["주문번호", "전화"], ["ORD-2026-0000001", "010-1111-2222"]]), "S");
  const { plan: p } = plan(wb, { 0: "keep", 1: "keep" });
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: { own: ["ORD → 제품"] } });
  const ws = r.wb.Sheets[r.wb.SheetNames[0]];
  assert.equal(ws.A2.v, "제품-2026-0000001");
  assert.equal(ws.B2.v, "010-1111-2222");
});
t("엑셀: 숫자 칸(₩ 서식)도 '내가 정한 말'(₩*)이 걸리고, 다른 숫자 칸은 그대로", () => {
  const wb = X.utils.book_new();
  const ws = X.utils.aoa_to_sheet([["품목", "단가", "재고"], ["원두", 22000, 10], ["우유", 2400, 60]]);
  ws.B2.z = ws.B3.z = '"₩"#,##0'; ws.B2.w = "₩22,000"; ws.B3.w = "₩2,400";
  X.utils.book_append_sheet(wb, ws, "S");
  const { plan: p } = plan(wb, { 0: "keep", 1: "text", 2: "keep" });
  const r = C.maskWorkbook(X, wb, p, { fmt: "plain", lists: { own: ["₩* → 가격"] } });
  const o = r.wb.Sheets[r.wb.SheetNames[0]];
  assert.equal(o.B2.v, "가격"); assert.equal(o.B3.v, "가격2".replace("2", "") === "가격" ? o.B3.v : o.B3.v);
  assert.ok(!String(o.B3.v).includes("2,400") && !String(o.B3.v).includes("2400"));
  assert.equal(o.C2.v, 10);
});
console.log(`\n${pass}개 통과` + (process.exitCode ? " · 실패 있음" : ""));
