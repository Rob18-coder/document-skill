import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, PageBreak, LevelFormat } from "docx";
import { writeFileSync } from "fs";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";

const argv = yargs(hideBin(process.argv))
  .option("spec", { type: "string", demandOption: true, description: "Path to JSON specification file" })
  .parseSync();

function parseAlignment(val) {
  const map = { left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFY };
  return map[val] || AlignmentType.LEFT;
}

function createTextRun(content, fmt = {}) {
  return new TextRun({
    text: content,
    bold: fmt.bold || false,
    italics: fmt.italic || false,
    size: fmt.size ? fmt.size * 2 : 22,
    color: fmt.color || "000000",
    font: fmt.name || "Calibri",
  });
}

function createParagraph(content, fmt = {}) {
  return new Paragraph({
    alignment: parseAlignment(fmt.alignment),
    children: [createTextRun(content, fmt)],
    spacing: { line: 276 },
  });
}

function createHeading(content, level, fmt = {}) {
  return new Paragraph({
    text: content,
    heading: level,
    alignment: parseAlignment(fmt.alignment),
    children: [createTextRun(content, fmt)],
    spacing: { before: 200, after: 100 },
  });
}

function createTable(data, header = true, fmt = {}) {
  const rows = data.map((row, i) => {
    const cells = row.map((cell, j) => new TableCell({
      children: [new Paragraph({ children: [createTextRun(String(cell), { bold: header && i === 0 })] })],
      width: { size: 100 / (data[0]?.length || 1), type: WidthType.PERCENTAGE },
    }));
    return new TableRow({ children: cells });
  });
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

function createList(items, ordered = false, fmt = {}) {
  return items.map((item, i) => new Paragraph({
    children: [createTextRun(item, fmt)],
    numbering: { reference: ordered ? "ordered" : "bullet", level: 0 },
    indent: { left: 720 },
  }));
}

async function buildDocument(spec) {
  const docConfig = spec.document || {};
  const margins = docConfig.margins || { top: 1.0, bottom: 1.0, left: 1.0, right: 1.0 };
  const orientation = docConfig.orientation === "landscape" ? "landscape" : "portrait";

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: Math.round(margins.top * 1440),
            bottom: Math.round(margins.bottom * 1440),
            left: Math.round(margins.left * 1440),
            right: Math.round(margins.right * 1440),
          },
          size: orientation === "landscape" ? { orientation: "landscape" } : undefined,
        },
      },
      children: [],
    }],
    numbering: {
      config: [
        { reference: "ordered", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT }] },
        { reference: "bullet", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT }] },
      ],
    },
  });

  const body = doc.documentWrapper.document.body;

  for (const sec of docConfig.sections || []) {
    switch (sec.type) {
      case "heading":
        body.root.push(createHeading(sec.content, sec.level || 1, sec.formatting));
        break;
      case "paragraph":
        body.root.push(createParagraph(sec.content, sec.formatting));
        break;
      case "table":
        body.root.push(createTable(sec.data, sec.header !== false, sec.formatting));
        break;
      case "list":
        for (const p of createList(sec.items, sec.ordered, sec.formatting)) {
          body.root.push(p);
        }
        break;
      case "page_break":
        body.root.push(new Paragraph({ children: [new PageBreak()] }));
        break;
    }
  }

  const buffer = await Packer.toBuffer(doc);
  writeFileSync(spec.output_path, buffer);
  console.log(`Document generated: ${spec.output_path}`);
}

const fs = await import("fs/promises");
const specContent = await fs.readFile(argv.spec, "utf-8");
const spec = JSON.parse(specContent);
await buildDocument(spec).catch(err => { console.error(err); process.exit(1); });