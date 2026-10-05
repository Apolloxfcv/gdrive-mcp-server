import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

/**
 * Conversion Markdown -> .docx (sous-ensemble courant) : titres #..######, listes a puces
 * (-, *, +), listes numerotees (texte conserve), tableaux | a | b |, **gras**, *italique*,
 * `code`, lignes vides = separation de paragraphes. Tout est genere en memoire.
 */

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

function inlineRuns(text: string): TextRun[] {
  const runs: TextRun[] = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0;
  for (const m of text.matchAll(re)) {
    if (m.index! > last) runs.push(new TextRun(text.slice(last, m.index)));
    const tok = m[0];
    if (tok.startsWith("**")) runs.push(new TextRun({ text: tok.slice(2, -2), bold: true }));
    else if (tok.startsWith("`")) runs.push(new TextRun({ text: tok.slice(1, -1), font: "Consolas" }));
    else runs.push(new TextRun({ text: tok.slice(1, -1), italics: true }));
    last = m.index! + tok.length;
  }
  if (last < text.length) runs.push(new TextRun(text.slice(last)));
  return runs.length ? runs : [new TextRun("")];
}

const isTableRow = (l: string) => /^\s*\|.*\|\s*$/.test(l);
const isSeparator = (l: string) => /^\s*\|[\s:|-]+\|\s*$/.test(l);
const cells = (l: string) =>
  l
    .trim()
    .replace(/^\||\|$/g, "")
    .split("|")
    .map((c) => c.trim());

export async function markdownToDocx(markdown: string): Promise<Buffer> {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const children: (Paragraph | Table)[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    if (isTableRow(line)) {
      const rows: string[][] = [];
      while (i < lines.length && isTableRow(lines[i])) {
        if (!isSeparator(lines[i])) rows.push(cells(lines[i]));
        i++;
      }
      i--;
      const width = Math.max(...rows.map((r) => r.length));
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: rows.map(
            (r, ri) =>
              new TableRow({
                children: Array.from({ length: width }, (_, ci) =>
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: ri === 0 ? [new TextRun({ text: r[ci] ?? "", bold: true })] : inlineRuns(r[ci] ?? ""),
                      }),
                    ],
                  })
                ),
              })
          ),
        })
      );
      children.push(new Paragraph({ children: [] }));
      continue;
    }

    const h = /^(#{1,6})\s+(.*)$/.exec(line);
    if (h) {
      children.push(new Paragraph({ heading: HEADINGS[h[1].length - 1], children: inlineRuns(h[2]) }));
      continue;
    }
    const b = /^(\s*)[-*+]\s+(.*)$/.exec(line);
    if (b) {
      children.push(
        new Paragraph({ bullet: { level: Math.min(Math.floor(b[1].length / 2), 8) }, children: inlineRuns(b[2]) })
      );
      continue;
    }
    children.push(new Paragraph({ children: inlineRuns(line.trim()) }));
  }

  const doc = new Document({ sections: [{ children: children.length ? children : [new Paragraph({ children: [] })] }] });
  return Buffer.from(await Packer.toBuffer(doc));
}
