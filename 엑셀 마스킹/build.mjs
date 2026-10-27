// src/app.html + vendor/xlsx + src/core.js 를 한 장짜리 엑셀 마스킹.html 로 합친다.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
const d = dirname(fileURLToPath(import.meta.url));
const rd = (p) => readFileSync(join(d, p), "utf8");
// "</script" 이 글 안에 있으면 HTML이 끊기므로 막는다
const safe = (s) => s.replace(/<\/(script)/gi, "<\\/$1");
// replace(문자열, 문자열) 은 $& 같은 특수 표기를 해석하므로 함수로 넣는다
const html = rd("src/app.html")
  .replace("/*XLSX*/", () => safe(rd("vendor/xlsx.mini.min.js")))
  .replace("/*CORE*/", () => safe(rd("src/core.js")));
writeFileSync(join(d, "엑셀 마스킹.html"), html);
console.log("만들었어요: 엑셀 마스킹.html", (html.length / 1024).toFixed(0) + "KB");
