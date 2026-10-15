// md2docx.js — 把指定目录下所有 .md 转成同名 .docx(楷体 15pt,美化排版)
// 用法:node md2docx.js [目录]     (缺省为当前目录;含子目录 logs/ 则一并转换)
// 依赖:npm install docx
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        WidthType, ShadingType, HeadingLevel, AlignmentType, BorderStyle,
        Header, Footer, PageNumber } = require("docx");

const BASE = path.resolve(process.argv[2] || ".");
const FONT = { ascii: "楷体", eastAsia: "楷体", hAnsi: "楷体" };
const C = { ink: "1F2937", acc: "E8590C", mut: "6B7280", hair: "D1D5DB", tint: "FDEEE4", grid: "E5E7EB" };
const BODY = 30, H1 = 40, H2 = 32; // half-points: 15pt / 20pt / 16pt
const LINE = { line: 312, lineRule: "auto" };
const bd = (sz, col) => ({ style: BorderStyle.SINGLE, size: sz, color: col });

function collectMd(dir) {
  let out = [];
  for (const f of fs.readdirSync(dir, { withFileTypes: true })) {
    if (f.isDirectory()) {
      if (f.name === "node_modules" || f.name === "docs") continue;
      const sub = path.join(dir, f.name);
      if (fs.readdirSync(sub).some(x => x.endsWith(".md"))) out = out.concat(collectMd(sub));
    } else if (f.name.endsWith(".md")) {
      out.push(path.join(dir, f.name));
    }
  }
  return out;
}
const files = collectMd(BASE);

function inlineRuns(text, extra = {}) {
  const runs = [];
  for (const part of text.split(/(\*\*[^*]+\*\*)/)) {
    if (!part) continue;
    if (part.startsWith("**") && part.endsWith("**")) {
      runs.push(new TextRun({ text: part.slice(2, -2), bold: true, font: FONT, size: BODY, color: C.ink, ...extra }));
    } else {
      runs.push(new TextRun({ text: part.replace(/`([^`]*)`/g, "$1"), font: FONT, size: BODY, color: C.ink, ...extra }));
    }
  }
  return runs.length ? runs : [new TextRun({ text: "", font: FONT, size: BODY, ...extra })];
}

function makeCell(text, isHeader) {
  return new TableCell({
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    shading: isHeader ? { type: ShadingType.CLEAR, fill: C.tint, color: "auto" } : undefined,
    children: [new Paragraph({
      spacing: { line: 300, lineRule: "auto" },
      children: inlineRuns(text, isHeader ? { bold: true, color: C.acc } : {}),
    })],
  });
}

function mdToDocx(mdPath, docxPath) {
  const lines = fs.readFileSync(mdPath, "utf8").split(/\r?\n/);
  const children = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (/^\s*$/.test(line) || /^---+\s*$/.test(line)) { i++; continue; }
    const hm = line.match(/^(#{1,4})\s+(.*)$/);
    if (hm) {
      const lv = hm[1].length, text = hm[2].trim();
      if (lv === 1) {
        children.push(new Paragraph({
          spacing: { before: 0, after: 200, ...LINE },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.hair, space: 6 } },
          children: [new TextRun({ text, bold: true, font: FONT, size: H1, color: C.ink })],
        }));
      } else if (lv === 2) {
        children.push(new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 280, after: 120, ...LINE },
          children: [new TextRun({ text, bold: true, font: FONT, size: H2, color: C.acc })],
        }));
      } else {
        children.push(new Paragraph({
          heading: lv === 3 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4,
          spacing: { before: 200, after: 100, ...LINE },
          children: [new TextRun({ text, bold: true, font: FONT, size: BODY, color: C.ink })],
        }));
      }
      i++; continue;
    }
    if (/^\|/.test(line)) {
      const rows = [];
      while (i < lines.length && /^\|/.test(lines[i])) {
        const cells = lines[i].split("|").slice(1, -1).map(c => c.trim());
        if (cells.length && !cells.every(c => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      if (rows.length) {
        children.push(new Table({
          rows: rows.map((cells, ri) => new TableRow({
            tableHeader: ri === 0, cantSplit: true,
            children: cells.map(c => makeCell(c, ri === 0)),
          })),
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: bd(4, C.grid), bottom: bd(4, C.grid), left: bd(4, C.grid), right: bd(4, C.grid),
            insideHorizontal: bd(4, C.grid), insideVertical: bd(4, C.grid),
          },
        }));
        children.push(new Paragraph({ spacing: { after: 140 }, children: [] }));
      }
      continue;
    }
    if (/^[-*]\s+/.test(line)) {
      children.push(new Paragraph({
        spacing: { ...LINE, after: 70 }, bullet: { level: 0 },
        children: inlineRuns(line.replace(/^[-*]\s+/, "")),
      }));
      i++; continue;
    }
    if (/^>\s?/.test(line)) {
      children.push(new Paragraph({
        spacing: { ...LINE, after: 100 },
        indent: { left: 360 },
        border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.acc, space: 12 } },
        shading: { type: ShadingType.CLEAR, fill: "FFF8F4" },
        children: inlineRuns(line.replace(/^>\s?/, ""), { color: C.mut }),
      }));
      i++; continue;
    }
    children.push(new Paragraph({
      spacing: { ...LINE, after: 80 }, alignment: AlignmentType.JUSTIFIED,
      children: inlineRuns(line),
    }));
    i++;
  }
  const doc = new Document({
    title: path.basename(mdPath, ".md"),
    styles: { default: { document: { run: { font: FONT, size: BODY, color: C.ink } } } },
    sections: [{
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1700, right: 1700 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.hair, space: 2 } },
            children: [new TextRun({ text: "AI Life OS", font: FONT, size: 18, color: C.mut })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "- ", font: FONT, size: 18, color: C.mut }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: C.mut }),
              new TextRun({ text: " -", font: FONT, size: 18, color: C.mut }),
            ],
          })],
        }),
      },
      children,
    }],
  });
  return Packer.toBuffer(doc).then(buf => fs.writeFileSync(docxPath, buf));
}

(async () => {
  let ok = 0; const fail = [];
  for (const f of files) {
    try { await mdToDocx(f, f.replace(/\.md$/, ".docx")); ok++; }
    catch (e) { fail.push(path.basename(f) + ": " + (e.code || e.message)); }
  }
  console.log(`md2docx: converted ${ok}/${files.length} in "${BASE}"` + (fail.length ? "; FAILED: " + fail.join(" | ") : ""));
})();
