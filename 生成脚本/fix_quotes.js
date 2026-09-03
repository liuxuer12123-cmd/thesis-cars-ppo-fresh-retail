// 修复 gen_thesis_v2.js 中文本字符串内部的 ASCII 双引号 -> 中文全角引号 “ ”
// 规则：逐字符扫描，维护 in_string 状态。
//  - 非字符串中遇到 " → 定界开始
//  - 字符串中遇到 "：若其后紧跟 , ) ] } 空格 或行尾 → 定界结束；否则为内部引号，交替替换为 “ ”
// 跳过注释行（以 //、*、/* 开头）
const fs = require("fs");
const FILE = "/Users/liuxinyi18/.codeflicker/workshop/gen_thesis_v2.js";
const CLOSERS = new Set([",", ")", "]", "}", " ", "	"]);

const lines = fs.readFileSync(FILE, "utf8").split("\n");
let converted = 0;
const out = lines.map((line) => {
  const trimmed = line.trimStart();
  if (trimmed.startsWith("//") || trimmed.startsWith("*") || trimmed.startsWith("/*")) return line;
  const chars = [...line];
  let inStr = false, innerOpen = false;
  const res = [];
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (c !== '"') { res.push(c); continue; }
    if (!inStr) { inStr = true; innerOpen = false; res.push(c); continue; }
    const next = chars[i + 1];
    const isCloser = next === undefined || CLOSERS.has(next);
    if (isCloser) { inStr = false; res.push(c); continue; }
    // inner quote
    res.push(innerOpen ? "”" : "“");
    innerOpen = !innerOpen;
    converted++;
  }
  return res.join("");
});
fs.writeFileSync(FILE, out.join("\n"));
console.log(`转换内部引号数量: ${converted}`);
