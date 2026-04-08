const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, ImageRun, BorderStyle, TableRow, TableCell,
  Table, WidthType, ShadingType, ExternalHyperlink, Footer, PageNumber
} = require("docx");

const GUIDE_DIR = path.join(__dirname, "vejledninger");
const OUT_DIR = path.join(__dirname, "vejledninger", "docx");
const LOGO_PATH = path.join(__dirname, "assets", "logo.png");

const COLORS = {
  primary: "1F4E79",
  accent: "2E75B6",
  lightBg: "E8F0FE",
  text: "333333",
  muted: "666666",
  link: "0563C1",
  white: "FFFFFF",
};

function parseFrontmatter(content) {
  const match = content.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
  if (!match) return { title: "", body: content };
  const fm = match[1];
  const titleMatch = fm.match(/title:\s*"?(.+?)"?\s*$/m);
  return { title: titleMatch ? titleMatch[1] : "", body: match[2].trim() };
}

function parseMarkdownToBlocks(md) {
  const lines = md.split("\n");
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Skip Jekyll link at bottom
    if (line.includes("{{ site.baseurl }}")) { i++; continue; }
    if (line.trim() === "---") { i++; continue; }

    // Table
    if (line.includes("|") && i + 1 < lines.length && lines[i + 1].match(/^\|[\s-|]+\|$/)) {
      const tableLines = [];
      while (i < lines.length && lines[i].includes("|")) {
        const trimmed = lines[i].trim();
        if (!trimmed.match(/^\|[\s-|]+\|$/)) {
          tableLines.push(trimmed);
        }
        i++;
      }
      blocks.push({ type: "table", rows: tableLines.map(l =>
        l.split("|").filter(c => c.trim() !== "").map(c => c.trim())
      )});
      continue;
    }

    // Heading
    const hMatch = line.match(/^(#{1,3})\s+(.+)/);
    if (hMatch) {
      blocks.push({ type: "heading", level: hMatch[1].length, text: hMatch[2] });
      i++; continue;
    }

    // Numbered list item
    const numMatch = line.match(/^\d+\.\s+(.+)/);
    if (numMatch) {
      blocks.push({ type: "numbered", text: numMatch[1] });
      i++; continue;
    }

    // Bullet item
    const bulletMatch = line.match(/^[-*]\s+(.+)/);
    if (bulletMatch) {
      blocks.push({ type: "bullet", text: bulletMatch[1] });
      i++; continue;
    }

    // Empty line
    if (line.trim() === "") { i++; continue; }

    // Regular paragraph
    blocks.push({ type: "paragraph", text: line.trim() });
    i++;
  }
  return blocks;
}

function formatInlineText(text, opts = {}) {
  const runs = [];
  // Split on **bold**, `code`, and [link](url) patterns
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))/g);
  for (const part of parts) {
    if (!part) continue;
    const boldMatch = part.match(/^\*\*(.+)\*\*$/);
    const codeMatch = part.match(/^`(.+)`$/);
    const linkMatch = part.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
    if (boldMatch) {
      runs.push(new TextRun({ text: boldMatch[1], bold: true, font: "Calibri", size: opts.size || 22, color: opts.color || COLORS.text }));
    } else if (codeMatch) {
      runs.push(new TextRun({ text: codeMatch[1], font: "Consolas", size: opts.size || 20, color: COLORS.accent, shading: { type: ShadingType.CLEAR, fill: "F0F0F0" } }));
    } else if (linkMatch) {
      runs.push(new TextRun({ text: linkMatch[1], font: "Calibri", size: opts.size || 22, color: COLORS.link, underline: {} }));
    } else {
      runs.push(new TextRun({ text: part, font: "Calibri", size: opts.size || 22, color: opts.color || COLORS.text }));
    }
  }
  return runs;
}

function blocksToParagraphs(blocks) {
  const paragraphs = [];
  let numberedCounter = 0;

  for (const block of blocks) {
    switch (block.type) {
      case "heading":
        numberedCounter = 0;
        if (block.level === 1) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: block.text, font: "Calibri", size: 36, bold: true, color: COLORS.primary })],
            spacing: { before: 240, after: 120 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.accent } },
          }));
        } else if (block.level === 2) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: block.text, font: "Calibri", size: 28, bold: true, color: COLORS.accent })],
            spacing: { before: 200, after: 80 },
          }));
        } else {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: block.text, font: "Calibri", size: 24, bold: true, color: COLORS.primary })],
            spacing: { before: 160, after: 60 },
          }));
        }
        break;

      case "numbered":
        numberedCounter++;
        paragraphs.push(new Paragraph({
          children: [
            new TextRun({ text: `${numberedCounter}. `, font: "Calibri", size: 22, bold: true, color: COLORS.accent }),
            ...formatInlineText(block.text),
          ],
          spacing: { before: 40, after: 40 },
          indent: { left: 360 },
        }));
        break;

      case "bullet":
        numberedCounter = 0;
        paragraphs.push(new Paragraph({
          children: [
            new TextRun({ text: "  \u2022  ", font: "Calibri", size: 22, color: COLORS.accent }),
            ...formatInlineText(block.text),
          ],
          spacing: { before: 40, after: 40 },
          indent: { left: 360 },
        }));
        break;

      case "table":
        numberedCounter = 0;
        if (block.rows.length > 0) {
          const rows = block.rows.map((cells, idx) => new TableRow({
            children: cells.map(cell => new TableCell({
              children: [new Paragraph({
                children: [new TextRun({
                  text: cell.replace(/\*\*/g, ""),
                  font: "Calibri",
                  size: 20,
                  bold: idx === 0,
                  color: idx === 0 ? COLORS.white : COLORS.text,
                })],
                spacing: { before: 40, after: 40 },
              })],
              shading: idx === 0
                ? { type: ShadingType.CLEAR, fill: COLORS.primary }
                : { type: ShadingType.CLEAR, fill: idx % 2 === 1 ? COLORS.lightBg : COLORS.white },
              margins: { top: 40, bottom: 40, left: 80, right: 80 },
            })),
          }));
          paragraphs.push(new Table({
            rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
          }));
          paragraphs.push(new Paragraph({ spacing: { before: 80, after: 80 } }));
        }
        break;

      case "paragraph":
        numberedCounter = 0;
        paragraphs.push(new Paragraph({
          children: formatInlineText(block.text),
          spacing: { before: 60, after: 60 },
        }));
        break;
    }
  }
  return paragraphs;
}

async function generateDocx(mdFile) {
  const content = fs.readFileSync(path.join(GUIDE_DIR, mdFile), "utf8");
  const { title, body } = parseFrontmatter(content);
  const blocks = parseMarkdownToBlocks(body);

  const logoData = fs.readFileSync(LOGO_PATH);

  const headerParagraphs = [
    new Paragraph({
      children: [
        new ImageRun({
          data: logoData,
          transformation: { width: 120, height: 40 },
          type: "png",
        }),
        new TextRun({ text: "    Emmerske Efterskole \u2014 M365 Kursus", font: "Calibri", size: 18, color: COLORS.muted }),
      ],
      spacing: { after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: COLORS.lightBg } },
    }),
  ];

  const titleParagraph = new Paragraph({
    children: [new TextRun({ text: title || mdFile.replace(".md", ""), font: "Calibri", size: 44, bold: true, color: COLORS.primary })],
    spacing: { before: 200, after: 40 },
    alignment: AlignmentType.LEFT,
  });

  const subtitleParagraph = new Paragraph({
    children: [new TextRun({ text: "Vejledning til medarbejdere", font: "Calibri", size: 22, italics: true, color: COLORS.muted })],
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.accent } },
  });

  const bodyParagraphs = blocksToParagraphs(blocks);

  const footerParagraph = new Paragraph({
    children: [
      new TextRun({ text: "Emmerske Efterskole \u2014 M365 Kursus  |  Side ", font: "Calibri", size: 16, color: COLORS.muted }),
    ],
    alignment: AlignmentType.CENTER,
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 720, bottom: 720, right: 720, left: 720 },
        },
      },
      headers: {
        default: { options: { children: headerParagraphs } },
      },
      footers: {
        default: { options: { children: [footerParagraph] } },
      },
      children: [titleParagraph, subtitleParagraph, ...bodyParagraphs],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  const outName = mdFile.replace(".md", ".docx");
  const outPath = path.join(OUT_DIR, outName);
  fs.writeFileSync(outPath, buffer);
  console.log(`  OK: ${outName}`);
}

async function main() {
  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

  const mdFiles = fs.readdirSync(GUIDE_DIR)
    .filter(f => f.endsWith(".md"))
    .sort();

  console.log(`Generating ${mdFiles.length} Word documents...`);
  for (const mdFile of mdFiles) {
    await generateDocx(mdFile);
  }
  console.log("Done!");
}

main().catch(err => { console.error(err); process.exit(1); });
