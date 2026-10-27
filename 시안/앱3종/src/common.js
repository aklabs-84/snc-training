/* 월 실적 도우미 공통 부품: 파일 읽기, 정리, 집계, 그래프, 가리기 */
const U = (function(){
  const $ = (q, r=document) => r.querySelector(q), $$ = (q, r=document) => [...r.querySelectorAll(q)];
  const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const fmt = n => (n==null||isNaN(n)) ? '' : (Number.isInteger(n) ? n : Math.round(n*10)/10).toLocaleString('ko-KR');

  /* ---------- 글자 표(csv/tsv) 읽기 ---------- */
  function parseDelim(text){
    text = text.replace(/^﻿/,'');
    const first = text.split(/\r?\n/,1)[0];
    const d = first.includes('\t') ? '\t' : (first.split(',').length>1 ? ',' : '\t');
    const rows=[]; let row=[], cur='', q=false;
    for (let i=0;i<text.length;i++){
      const c=text[i];
      if (q){ if(c==='"'){ if(text[i+1]==='"'){cur+='"';i++;} else q=false; } else cur+=c; }
      else if (c==='"' && cur==='') q=true;
      else if (c===d){ row.push(cur); cur=''; }
      else if (c==='\n'){ row.push(cur); rows.push(row); row=[]; cur=''; }
      else if (c!=='\r') cur+=c;
    }
    if (cur!=='' || row.length){ row.push(cur); rows.push(row); }
    return rows;
  }
  function toTable(rows){
    rows = rows.filter(r=>r.some(c=>String(c).trim()!==''));
    if (!rows.length) return null;
    let n = rows[0].length; while(n>0 && String(rows[0][n-1]).trim()==='') n--;
    if (n===0) return null;
    const header = rows[0].slice(0,n).map((h,i)=>String(h).trim()||('열'+(i+1)));
    const body = rows.slice(1).map(r=>{ const o=[]; for(let i=0;i<n;i++) o.push(r[i]==null?'':String(r[i])); return o; });
    return body.length ? {header, rows: body} : null;
  }

  /* ---------- 엑셀(.xlsx) 읽기: 외부 부품 없이 직접 풀어서 읽음 ---------- */
  async function inflate(bytes){
    const ds = new DecompressionStream('deflate-raw');
    const w = ds.writable.getWriter(); w.write(bytes); w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
  }
  async function unzip(buf){
    const v = new DataView(buf), u8 = new Uint8Array(buf);
    let e = -1; for (let i=u8.length-22;i>=Math.max(0,u8.length-70000);i--) if (v.getUint32(i,true)===0x06054b50){ e=i; break; }
    if (e<0) throw new Error('엑셀 파일(.xlsx)이 아니에요');
    const n = v.getUint16(e+10,true); let p = v.getUint32(e+16,true); const files = {};
    const dec = new TextDecoder();
    for (let i=0;i<n;i++){
      const method=v.getUint16(p+10,true), csize=v.getUint32(p+20,true), nl=v.getUint16(p+28,true), el=v.getUint16(p+30,true), cl=v.getUint16(p+32,true), off=v.getUint32(p+42,true);
      const name = dec.decode(u8.subarray(p+46,p+46+nl)); p += 46+nl+el+cl;
      const ln=v.getUint16(off+26,true), le=v.getUint16(off+28,true), start=off+30+ln+le;
      const raw = u8.subarray(start,start+csize);
      files[name] = {method, raw};
    }
    const get = async name => { const f=files[name]; if(!f) return null; const b = f.method===0 ? f.raw : await inflate(f.raw); return dec.decode(b); };
    return {get, has:name=>!!files[name]};
  }
  const unesc = s => s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(m,d)=>String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi,(m,h)=>String.fromCodePoint(parseInt(h,16))).replace(/&amp;/g,'&');
  const colIdx = ref => { let n=0; for (const ch of ref.replace(/[0-9]/g,'')) n = n*26 + ch.charCodeAt(0)-64; return n-1; };
  function serialToISO(sv){ return new Date(Date.UTC(1899,11,30)+Math.floor(sv)*864e5).toISOString().slice(0,10); }
  async function readXlsx(buf){
    const z = await unzip(buf);
    const wb = await z.get('xl/workbook.xml'); if(!wb) throw new Error('엑셀 파일(.xlsx)이 아니에요');
    const rels = await z.get('xl/_rels/workbook.xml.rels') || '';
    const relMap = {}; for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)){ const id=(m[0].match(/\bId="([^"]*)"/)||[])[1], t=(m[0].match(/\bTarget="([^"]*)"/)||[])[1]; if(id) relMap[id]=t; }
    const sst = await z.get('xl/sharedStrings.xml'); const shared=[];
    if (sst) for (const m of sst.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)){ const inner=m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g,''); shared.push(unesc([...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join(''))); }
    /* 날짜 모양으로 저장된 칸 찾기 */
    const dateXf = new Set(); const st = await z.get('xl/styles.xml');
    if (st){
      const custom = {}; for (const m of st.matchAll(/<numFmt\b[^>]*>/g)){ const id=(m[0].match(/numFmtId="(\d+)"/)||[])[1], code=(m[0].match(/formatCode="([^"]*)"/)||[])[1]||''; custom[id]=code; }
      const cx = (st.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)||[])[1]||'';
      [...cx.matchAll(/<xf\b[^>]*>/g)].forEach((m,i)=>{ const id=+((m[0].match(/numFmtId="(\d+)"/)||[])[1]||0);
        const isD = (id>=14&&id<=22)||(id>=45&&id<=47)||(custom[id] && /[ymd]/i.test(custom[id].replace(/"[^"]*"|\[[^\]]*\]/g,'')) && !/[hs]{1,2}:?$/.test('') );
        if (isD) dateXf.add(i); });
    }
    const sheets=[];
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)){
      const name = unesc((m[0].match(/\bname="([^"]*)"/)||[])[1]||'시트'); const rid=(m[0].match(/r:id="([^"]*)"/)||[])[1];
      let t = relMap[rid]||''; t = t.startsWith('/') ? t.slice(1) : 'xl/'+t;
      const xml = await z.get(t); if(!xml) continue;
      const rows=[];
      for (const rm of xml.matchAll(/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)){
        const r=[]; const body=rm[1]||'';
        for (const cm of body.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
          const at=cm[1], inner=cm[2]||''; const ref=(at.match(/\br="([A-Z]+)\d+"/)||[])[1]; if(!ref) continue;
          const ty=(at.match(/\bt="([^"]*)"/)||[])[1]||'n', s=+((at.match(/\bs="(\d+)"/)||[])[1]||0);
          let val='';
          if (ty==='inlineStr') val = unesc([...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join(''));
          else { const vv=(inner.match(/<v>([\s\S]*?)<\/v>/)||[])[1]; if(vv!=null){ val = ty==='s' ? (shared[+vv]||'') : unesc(vv); if(ty==='n' && dateXf.has(s) && !isNaN(+val)) val = serialToISO(+val); if(ty==='b') val = vv==='1'?'TRUE':'FALSE'; } }
          r[colIdx(ref)] = val;
        }
        rows.push(Array.from(r,x=>x==null?'':x));
      }
      sheets.push({name, rows});
    }
    return sheets;
  }
  async function readFile(file){
    const buf = await file.arrayBuffer();
    if (/\.xlsx$/i.test(file.name) || (new Uint8Array(buf,0,2)[0]===0x50 && new Uint8Array(buf,0,2)[1]===0x4b)) return readXlsx(buf);
    if (/\.xls$/i.test(file.name)) throw new Error('옛 형식(.xls)은 못 읽어요. 엑셀에서 "다른 이름으로 저장 → .xlsx"로 바꿔 주세요');
    let text; try { text = new TextDecoder('utf-8',{fatal:true}).decode(buf); } catch(e){ text = new TextDecoder('euc-kr').decode(buf); }
    return [{name:file.name, rows: parseDelim(text)}];
  }

  /* ---------- 열 종류 알아맞히기 ---------- */
  const DATE_HEAD = /일자|날짜|일시|date|납기|출고일|입고일/i;
  function parseDate(v, allowSerial){
    let s = String(v).trim(); if(!s) return null;
    s = s.replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/,'');
    let m = s.match(/^(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})\s*일?$/) || s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (!m && /^\d{2}[.\-\/]\d{1,2}[.\-\/]\d{1,2}$/.test(s)){ const p=s.split(/[.\-\/]/); m=[0,'20'+p[0],p[1],p[2]]; }
    if (m){ const y=+m[1], mo=+m[2], d=+m[3]; if(mo>=1&&mo<=12&&d>=1&&d<=31&&y>=1990&&y<=2100) return `${y}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`; return null; }
    if (allowSerial && /^\d{5}(\.\d+)?$/.test(s) && +s>=30000 && +s<=70000) return serialToISO(+s);
    return null;
  }
  const UNIT = /(원|개|ea|pcs|kg|㎏|%|건|대|톤|t)$/i;
  function parseNum(v){
    let s = String(v).trim().replace(/^[₩￦$]/,'').replace(/[,\s]/g,''); s = s.replace(UNIT,'');
    if (!/^[-+]?\d+(\.\d+)?$/.test(s)) return NaN; return +s;
  }
  function detectTypes(t){
    return t.header.map((h,c)=>{
      const vals = t.rows.map(r=>r[c]).filter(v=>String(v).trim()!==''); if(!vals.length) return 'text';
      const hd = DATE_HEAD.test(h);
      const dRatio = vals.filter(v=>parseDate(v,hd)).length/vals.length;
      if (dRatio>=0.8) return 'date';
      const nRatio = vals.filter(v=>!isNaN(parseNum(v))).length/vals.length;
      if (nRatio>=0.8) return 'num';
      return 'text';
    });
  }
  const TYPE_LABEL = {date:'날짜', num:'숫자', text:'글자'};

  /* ---------- 정리 ---------- */
  const RULES = {
    trim:'앞뒤·글자 속 보이지 않는 빈칸 지우기',
    date:'날짜를 2026-07-08 한 가지 모양으로',
    num:'숫자에 붙은 "원"·쉼표 없애기',
    same:'같은 이름인데 다르게 쓴 것 하나로 (A라인 / A 라인)',
    dup:'똑같은 줄은 한 번만 세기'
  };
  const ALL_ON = {trim:true,date:true,num:true,same:true,dup:true};
  function clean(t, types, opts){
    opts = Object.assign({}, ALL_ON, opts||{});
    const counts = {trim:0,date:0,num:0,same:0,dup:0};
    let rows = t.rows.map(r=>({dup:false, cells: r.map(x=>({v:x, raw:x, kinds:[]}))}));
    const mark = (c,newV,kind)=>{ if(newV!==c.v){ c.v=newV; if(!c.kinds.includes(kind)){ c.kinds.push(kind); counts[kind]++; } } };
    rows.forEach(r=>r.cells.forEach((c,i)=>{
      if (opts.trim) mark(c, c.v.replace(/[ ​　]/g,' ').trim(), 'trim');
      if (opts.date && types[i]==='date'){ const d=parseDate(c.v,true); if(d) mark(c,d,'date'); }
      if (opts.num && types[i]==='num'){ const n=parseNum(c.v); if(!isNaN(n)) mark(c,String(n),'num'); }
    }));
    if (opts.same) t.header.forEach((h,i)=>{
      if (types[i]!=='text') return;
      const g = new Map(); rows.forEach(r=>{ const v=r.cells[i].v; if(!v) return; const k=v.normalize('NFKC').replace(/\s+/g,'').toLowerCase(); if(!g.has(k)) g.set(k,new Map()); const m=g.get(k); m.set(v,(m.get(v)||0)+1); });
      if (g.size > rows.length*0.5) return;
      const canon = new Map(); g.forEach((m,k)=>{ if(m.size>1){ let best=null,bc=-1; m.forEach((c,v)=>{ if(c>bc||(c===bc&&v.length<best.length)){best=v;bc=c;} }); canon.set(k,best); } });
      rows.forEach(r=>{ const c=r.cells[i]; if(!c.v) return; const k=c.v.normalize('NFKC').replace(/\s+/g,'').toLowerCase(); if(canon.has(k)) mark(c,canon.get(k),'same'); });
    });
    if (opts.dup){ const seen=new Set(); rows.forEach(r=>{ const k=r.cells.map(c=>c.v).join('\u0001'); if(seen.has(k)){ r.dup=true; counts.dup++; } else seen.add(k); }); }
    return {header:t.header, types, rows, counts, kept: rows.filter(r=>!r.dup).length, total: rows.length};
  }
  function cleanTable(cl){ return {header:cl.header, rows: cl.rows.filter(r=>!r.dup).map(r=>r.cells.map(c=>c.v))}; }
  const colSum = (cl,i)=> cl.rows.reduce((a,r)=> r.dup?a:a+(parseNum(r.cells[i].v)||0),0);

  /* ---------- 집계 ---------- */
  function aggregate(t, types, cfg){
    const {key, month, val, method, split} = cfg;
    const keyOf = r => { let v=String(r[key]).trim(); if(types[key]==='date'){ const d=parseDate(v,true); if(d) v = month ? d.slice(0,7) : d; } return v||'(빈칸)'; };
    const keys=[], ks=new Set(), splits=[], ss=new Set(), cell={}, cnt={};
    t.rows.forEach(r=>{
      const k=keyOf(r), s = split>=0 ? (String(r[split]).trim()||'(빈칸)') : '전체';
      if(!ks.has(k)){ks.add(k);keys.push(k);} if(!ss.has(s)){ss.add(s);splits.push(s);}
      const n = method==='count' ? 1 : parseNum(r[val]); if (isNaN(n)) return;
      const id=k+'\u0001'+s; cell[id]=(cell[id]||0)+n; cnt[id]=(cnt[id]||0)+1;
    });
    const isDateKey = types[key]==='date';
    const get=(k,s)=>{ const id=k+'\u0001'+s; if(!(id in cell)) return 0; return method==='avg' ? cell[id]/cnt[id] : cell[id]; };
    const rowTot = k => { if(method==='avg'){ let a=0,c=0; splits.forEach(s=>{const id=k+'\u0001'+s; if(id in cell){a+=cell[id];c+=cnt[id];}}); return c?a/c:0; } return splits.reduce((a,s)=>a+get(k,s),0); };
    let ordered = keys.slice();
    if (isDateKey) ordered.sort(); else ordered.sort((a,b)=>rowTot(b)-rowTot(a));
    let colTot = s => { if(method==='avg'){ let a=0,c=0; keys.forEach(k=>{const id=k+'\u0001'+s; if(id in cell){a+=cell[id];c+=cnt[id];}}); return c?a/c:0; } return keys.reduce((a,k)=>a+get(k,s),0); };
    let grand; if(method==='avg'){ let a=0,c=0; Object.keys(cell).forEach(id=>{a+=cell[id];c+=cnt[id];}); grand=c?a/c:0; } else grand = Object.keys(cell).reduce((a,id)=>a+cell[id],0);
    return {keys:ordered, splits, get, rowTot, colTot, grand, method, isDateKey};
  }
  const METHOD = {sum:'합계', avg:'평균', count:'개수'};
  function aggTable(a, keyName, valName){
    const multi = a.splits.length>1;
    let h = `<table><thead><tr><th>${esc(keyName)}</th>${multi? a.splits.map(s=>`<th class="r">${esc(s)}</th>`).join('') : ''}<th class="r">${esc(METHOD[a.method])}${a.method==='count'?'':' · '+esc(valName)}</th></tr></thead><tbody>`;
    a.keys.forEach(k=>{ h += `<tr><td>${esc(k)}</td>${multi? a.splits.map(s=>`<td class="r">${fmt(a.get(k,s))}</td>`).join(''):''}<td class="r">${fmt(a.rowTot(k))}</td></tr>`; });
    h += `</tbody><tfoot><tr><td>${a.method==='avg'?'전체 평균':'전체'}</td>${multi? a.splits.map(s=>`<td class="r">${fmt(a.colTot(s))}</td>`).join(''):''}<td class="r">${fmt(a.grand)}</td></tr></tfoot></table>`;
    return h;
  }
  function aggRows(a, keyName){
    const multi=a.splits.length>1; const head=[keyName].concat(multi?a.splits:[]).concat(['합계/평균']);
    return [head].concat(a.keys.map(k=>[k].concat(multi?a.splits.map(s=>a.get(k,s)):[]).concat([a.rowTot(k)]))).concat([['전체'].concat(multi?a.splits.map(s=>a.colTot(s)):[]).concat([a.grand])]);
  }
  /* 집계 설정 화면(공통) */
  function pickDefaults(t, types){
    const dateC = types.indexOf('date'); const nums = types.map((x,i)=>x==='num'?i:-1).filter(i=>i>=0);
    const prefer = ['금액|매출|합계','수량|불량'].map(re=>nums.find(i=>new RegExp(re).test(t.header[i]))).find(x=>x!=null); const textC = types.findIndex((x,i)=>x==='text' && new Set(t.rows.map(r=>r[i])).size<=40);
    return {key: dateC>=0?dateC:Math.max(textC,0), month: dateC>=0, val: prefer!=null?prefer:(nums[0]??0), method: nums.length?'sum':'count', split: -1};
  }
  function aggControls(el, t, types, cfg, onChange){
    const sel = (id,label,opts,cur)=>`<label class="fld"><span>${label}</span><select id="${id}">${opts}</select></label>`;
    const colOpts = (filter,cur,none)=> (none?`<option value="-1" ${cur===-1?'selected':''}>${none}</option>`:'') + t.header.map((h,i)=> filter(i)?`<option value="${i}" ${cur===i?'selected':''}>${esc(h)}</option>`:'').join('');
    const nums = i=>types[i]==='num';
    el.innerHTML = `<div class="flds">
      ${sel('a-key','무엇별로 묶을까요?',colOpts(()=>true,cfg.key),cfg.key)}
      ${sel('a-method','어떻게 셀까요?',['sum','avg','count'].map(m=>`<option value="${m}" ${cfg.method===m?'selected':''}>${METHOD[m]}${m==='count'?' (줄 수)':''}</option>`).join(''))}
      ${sel('a-val','어느 숫자를?',colOpts(nums,cfg.val), cfg.val)}
      ${sel('a-split','더 나눠 볼까요?',colOpts(i=>i!==cfg.key&&types[i]==='text'&&new Set(t.rows.map(r=>r[i])).size<=12,cfg.split,'(나누지 않기)'))}
    </div>
    <label class="rule" id="a-monthrow" ${types[cfg.key]==='date'?'':'hidden'}><input type="checkbox" id="a-month" ${cfg.month?'checked':''}><span class="sw"></span><span class="t">날짜는 <b>월별</b>로 묶기</span></label>`;
    const rd = ()=>{ cfg.key=+$('#a-key',el).value; cfg.method=$('#a-method',el).value; cfg.val=+$('#a-val',el).value; cfg.split=+$('#a-split',el).value; const mc=$('#a-month',el); cfg.month = mc?mc.checked:false; };
    $$('select,input',el).forEach(x=>x.onchange=()=>{ const kc=+$('#a-key',el).value; const keyChanged = kc!==cfg.key; rd(); if(keyChanged){ if(types[cfg.key]==='date') cfg.month=true; if(cfg.split===cfg.key) cfg.split=-1; aggControls(el,t,types,cfg,onChange);} $('#a-val',el) && ($('#a-val',el).disabled = cfg.method==='count'); onChange(); });
    if ($('#a-val',el)) $('#a-val',el).disabled = cfg.method==='count';
  }

  /* ---------- 그래프 ---------- */
  const COLORS = ['var(--c1)','var(--c2)','var(--c3)','var(--c4)','var(--c5)','var(--c6)'];
  function unitOf(max){ return max>=1e8?[1e8,'억']:max>=1e6?[1e6,'백만']:max>=1e4?[1e4,'만']:[1,'']; }
  function niceMax(m){ if(m<=0) return 1; const p=Math.pow(10,Math.floor(Math.log10(m))); const f=m/p; return (f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*p; }
  function chart(a, kind, o){
    o = Object.assign({w:760,h:320},o||{});
    let keys = a.keys; let more='';
    if (keys.length>16 && !a.isDateKey){ keys = keys.slice(0,15); more=`상위 15개만 그렸어요 (전체 ${a.keys.length}개)`; }
    const multi = a.splits.length>1, stacked = multi && a.method!=='avg';
    const tot = k=>a.rowTot(k);
    let max = 0;
    if (kind==='hbar'||!multi) keys.forEach(k=>max=Math.max(max,tot(k)));
    else keys.forEach(k=> max=Math.max(max, stacked?tot(k):Math.max(...a.splits.map(s=>a.get(k,s)))));
    const [div,unit] = unitOf(max), nm = niceMax(max/div)*div;
    const f = v=>fmt(v/div);
    let s=''; const legend = multi ? `<div class="key">${a.splits.map((x,i)=>`<span><i style="background:${COLORS[i%6]}"></i>${esc(x)}</span>`).join('')}</div>` : '';
    if (kind==='hbar'){
      const L=110,R=90,rh=34,W=o.w-L-R,H=keys.length*rh+10;
      s=`<svg viewBox="0 0 ${o.w} ${H}" role="img" aria-label="가로 막대 그래프" style="width:100%;height:auto;display:block">`;
      keys.forEach((k,i)=>{ const y=5+i*rh, bw=nm?tot(k)/nm*W:0;
        s+=`<text x="${L-10}" y="${y+rh/2+5}" text-anchor="end" font-size="15" fill="var(--text)">${esc(k.length>9?k.slice(0,9)+'…':k)}</text>`;
        s+=`<rect class="bar" x="${L}" y="${y+6}" width="${Math.max(bw,1)}" height="${rh-14}" fill="${COLORS[0]}"><title>${esc(k)}: ${fmt(tot(k))}</title></rect>`;
        s+=`<text x="${L+bw+8}" y="${y+rh/2+5}" font-size="14" fill="var(--sub)">${fmt(tot(k))}</text>`; });
      return legend+s+'</svg>'+(more?`<p class="cap">${more}</p>`:'');
    }
    const L=56,R=16,T=34,B=44,W=o.w-L-R,Hh=o.h-T-B; const n=keys.length;
    s=`<svg viewBox="0 0 ${o.w} ${o.h}" role="img" aria-label="${kind==='line'?'선':'막대'} 그래프" style="width:100%;height:auto;display:block">`;
    s+=`<text x="0" y="14" font-size="14" fill="var(--sub)">단위: ${unit||'그대로'}</text>`;
    for(let i=0;i<=4;i++){ const v=nm/4*i, y=T+Hh-(v/nm)*Hh; s+=`<line x1="${L}" x2="${L+W}" y1="${y}" y2="${y}" stroke="var(--line)"/><text x="${L-8}" y="${y+5}" text-anchor="end" font-size="14" fill="var(--sub)">${f(v)}</text>`; }
    const step = W/n; const lab=k=> a.isDateKey&&/^\d{4}-\d{2}$/.test(k) ? (+k.slice(5))+'월' : (k.length>7?k.slice(0,7)+'…':k);
    const skip = Math.ceil(n/12);
    keys.forEach((k,i)=>{ if(i%skip===0) s+=`<text x="${L+step*i+step/2}" y="${T+Hh+24}" text-anchor="middle" font-size="15" fill="var(--text)">${esc(lab(k))}</text>`; });
    if (kind==='line'){
      const series = multi ? a.splits : ['전체'];
      series.forEach((sp,si)=>{ const pts=keys.map((k,i)=>{ const v=multi?a.get(k,sp):tot(k); return [L+step*i+step/2, T+Hh-(nm?v/nm:0)*Hh, v, k]; });
        s+=`<polyline fill="none" stroke="${COLORS[si%6]}" stroke-width="3" points="${pts.map(p=>p[0]+','+p[1]).join(' ')}"/>`;
        pts.forEach(p=>s+=`<circle class="bar" cx="${p[0]}" cy="${p[1]}" r="5" fill="${COLORS[si%6]}"><title>${esc(p[3])}${multi?' · '+esc(sp):''}: ${fmt(p[2])}</title></circle>`); });
    } else {
      keys.forEach((k,i)=>{ const x0=L+step*i+step*0.18, bw=step*0.64;
        if(!multi){ const v=tot(k), h=nm?v/nm*Hh:0; s+=`<rect class="bar" x="${x0}" y="${T+Hh-h}" width="${bw}" height="${Math.max(h,1)}" fill="${COLORS[0]}"><title>${esc(k)}: ${fmt(v)}</title></rect>`; }
        else if(stacked){ let y=T+Hh; a.splits.forEach((sp,si)=>{ const v=a.get(k,sp), h=nm?v/nm*Hh:0; y-=h; s+=`<rect class="bar" x="${x0}" y="${y}" width="${bw}" height="${h}" fill="${COLORS[si%6]}"><title>${esc(k)} · ${esc(sp)}: ${fmt(v)}</title></rect>`; }); }
        else { const w=bw/a.splits.length; a.splits.forEach((sp,si)=>{ const v=a.get(k,sp), h=nm?v/nm*Hh:0; s+=`<rect class="bar" x="${x0+w*si}" y="${T+Hh-h}" width="${w-2}" height="${Math.max(h,1)}" fill="${COLORS[si%6]}"><title>${esc(k)} · ${esc(sp)}: ${fmt(v)}</title></rect>`; }); }
      });
    }
    return legend+s+'</svg>'+(more?`<p class="cap">${more}</p>`:'');
  }

  /* ---------- 이름 가리기 ---------- */
  function aliasOf(i){ let s=''; i++; while(i>0){ const r=(i-1)%26; s=String.fromCharCode(65+r)+s; i=Math.floor((i-1)/26); } return s; }
  function makeAliases(groups){ /* groups: {열이름:[값...]} */
    const map={}; Object.keys(groups).forEach(col=>{ [...new Set(groups[col])].filter(v=>v && v!=='(빈칸)' && v!=='전체').forEach((v,i)=>{ map[v]=col+aliasOf(i); }); }); return map; }
  function applyAlias(text, map){ return Object.keys(map).sort((a,b)=>b.length-a.length).reduce((t,k)=>t.split(k).join(map[k]),text); }
  function restoreAlias(text, map){ const rev={}; Object.keys(map).forEach(k=>rev[map[k]]=k); return Object.keys(rev).sort((a,b)=>b.length-a.length).reduce((t,k)=>t.split(k).join(rev[k]),text); }

  /* ---------- 파일 올리기 화면 ---------- */
  function attachLoader(o){
    let sheets=[]; const drop=$(o.drop), file=$(o.file), sel=$(o.sheet), msg=$(o.msg);
    const say=(t,bad)=>{ msg.textContent=t; msg.className = bad?'warn':'hint'; };
    const use=i=>{ const t=toTable(sheets[i].rows); if(!t){ say('표를 찾지 못했어요. 첫 줄(제목 줄)까지 있는 표인지 확인해 주세요.',true); return; } say(`"${sheets[i].name}" 시트에서 ${t.rows.length}줄을 읽었어요.`); o.onTable(t, sheets[i].name); };
    const handle=async f=>{ if(!f) return; say('읽는 중…'); try{ sheets=await readFile(f); sheets=sheets.filter(s=>s.rows.length); if(!sheets.length) throw new Error('내용이 비어 있어요'); sel.innerHTML=sheets.map((s,i)=>`<option value="${i}">${esc(s.name)}</option>`).join(''); sel.closest('.fld').hidden = sheets.length<2; use(0); }catch(e){ say(e.message||'파일을 읽지 못했어요',true); } };
    file.onchange=()=>handle(file.files[0]);
    sel.onchange=()=>use(+sel.value);
    ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('over');}));
    ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('over');}));
    drop.addEventListener('drop',e=>handle(e.dataTransfer.files[0]));
    drop.addEventListener('click',e=>{ if(!e.target.closest('button,label,select')) file.click(); });
    drop.addEventListener('keydown',e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); file.click(); } });
    if(o.paste){ const p=$(o.paste); let tm; p.addEventListener('input',()=>{ clearTimeout(tm); tm=setTimeout(()=>{ const t=toTable(parseDelim(p.value)); if(!t){ say(p.value.trim()?'첫 줄(제목 줄)까지 같이 복사해 주세요.':'', !!p.value.trim()); return; } sheets=[{name:'붙여넣은 표',rows:parseDelim(p.value)}]; sel.closest('.fld').hidden=true; use(0); },150); }); }
    return {say};
  }
  function previewTable(el, t, types, onType){
    let h=`<div class="tablebox"><table><thead><tr>${t.header.map((x,i)=>`<th class="${types[i]==='num'?'r':''}">${esc(x)}</th>`).join('')}</tr><tr class="typerow">${types.map((ty,i)=>`<th><select data-i="${i}" aria-label="${esc(t.header[i])} 열 종류">${['date','num','text'].map(k=>`<option value="${k}" ${ty===k?'selected':''}>${TYPE_LABEL[k]}</option>`).join('')}</select></th>`).join('')}</tr></thead><tbody>`;
    t.rows.slice(0,6).forEach(r=>{ h+=`<tr>${r.map((v,i)=>`<td class="${types[i]==='num'?'r':''}">${esc(v)}</td>`).join('')}</tr>`; });
    el.innerHTML=h+'</tbody></table></div>';
    $$('select',el).forEach(s=>s.onchange=()=>{ types[+s.dataset.i]=s.value; onType(); });
  }
  function cleanedTableHTML(cl){
    const numC = cl.types.map(x=>x==='num');
    let h=`<div class="tablebox"><table><thead><tr>${cl.header.map((x,i)=>`<th class="${numC[i]?'r':''}">${esc(x)}</th>`).join('')}</tr></thead><tbody>`;
    cl.rows.forEach(r=>{ h+=`<tr class="${r.dup?'dup':''}">`+r.cells.map((c,i)=>{ const v=numC[i]&&c.v!==''&&!isNaN(+c.v)?(+c.v).toLocaleString('ko-KR'):c.v; const chg=c.kinds.length&&!r.dup; return `<td class="${numC[i]?'r':''} ${chg?'chg':''}" ${chg?`data-from="${esc(c.raw)}"`:''}>${esc(v)}</td>`; }).join('')+'</tr>'; });
    return h+'</tbody></table></div>';
  }

  /* ---------- 자잘한 것 ---------- */
  function toCSV(rows){ return '﻿'+rows.map(r=>r.map(c=>{ c=String(c==null?'':c); return /[",\n\t]/.test(c)?'"'+c.replace(/"/g,'""')+'"':c; }).join(',')).join('\r\n'); }
  function download(name, content, type){ const b=new Blob([content],{type:type||'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download=name; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},500); }
  function flash(el,m){ el.textContent=m; setTimeout(()=>el.textContent='',2400); }
  function copy(text, el){
    const ok=()=>flash(el,'복사했어요 ✓');
    const fb=()=>{ const t=document.createElement('textarea'); t.value=text; document.body.appendChild(t); t.select(); try{document.execCommand('copy');ok();}catch(e){flash(el,'선택해서 직접 복사해 주세요');} t.remove(); };
    if(navigator.clipboard&&window.isSecureContext) navigator.clipboard.writeText(text).then(ok,fb); else fb();
  }
  function stepper(titles, onShow){
    const subs=$$('.sub'); let k=0;
    function show(i){ k=i; subs.forEach((s,j)=>s.classList.toggle('on',j===i));
      $('#subtabs').innerHTML=titles.map((t,j)=>`<button role="tab" aria-selected="${j===i}" data-i="${j}"><b>${j+1}</b>${esc(t)}</button>`).join('');
      $('#nav').innerHTML=(i>0?`<button class="lnk" data-i="${i-1}">← 이전</button>`:'<span></span>')+(i<titles.length-1?`<button class="btn" data-i="${i+1}">다음: ${esc(titles[i+1])} →</button>`:'<span></span>');
      onShow&&onShow(i); window.scrollTo({top:0,behavior:'smooth'}); }
    document.addEventListener('click',e=>{ const b=e.target.closest('#subtabs [data-i],#nav [data-i]'); if(b) show(+b.dataset.i); });
    show(0); return {show, get cur(){return k;}};
  }
  const tipEl = ()=>$('#tip');
  function tipInit(){
    document.addEventListener('mousemove',e=>{ const tip=tipEl(); if(!tip) return; const td=e.target.closest&&e.target.closest('td.chg'); const nm=e.target.closest&&e.target.closest('[data-tip]');
      if(td) tip.textContent='원래 모양: '+(td.dataset.from===''?'(빈칸)':td.dataset.from); else if(nm) tip.textContent=nm.dataset.tip; else { tip.style.opacity=0; return; }
      tip.style.left=Math.min(e.clientX+14,innerWidth-320)+'px'; tip.style.top=(e.clientY+18)+'px'; tip.style.opacity=1; });
  }
  return {$,$$,esc,fmt,parseDelim,toTable,readFile,readXlsx,parseDate,parseNum,detectTypes,TYPE_LABEL,RULES,ALL_ON,clean,cleanTable,colSum,aggregate,aggTable,aggRows,aggControls,pickDefaults,METHOD,chart,makeAliases,applyAlias,restoreAlias,aliasOf,attachLoader,previewTable,cleanedTableHTML,toCSV,download,flash,copy,stepper,tipInit};
})();
if (typeof module!=='undefined') module.exports=U;
