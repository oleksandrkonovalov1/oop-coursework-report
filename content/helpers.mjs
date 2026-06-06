// Хелпери оформлення за ДСТУ 3008:2015 (адаптовано з генератора звітів ЛР):
// Times New Roman 14 пт, інтервал 1.5, береги 30/15/20/20 мм, A4.
import {
  Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel,
  BorderStyle, ShadingType,
} from "docx";
import { readFileSync } from "fs";

export const MM_TO_DXA = 56.693;
export const FONT = "Times New Roman";
export const FONT_CODE = "Courier New";
export const BODY_SIZE = 14 * 2;   // 14 пт у напівпунктах
export const CODE_SIZE = 10 * 2;   // 10 пт для фрагментів коду
export const LINE_15 = 360;        // полуторний міжрядковий інтервал
export const FIRST_LINE = Math.round(12.5 * MM_TO_DXA); // абзацний відступ

export const margins = {
  top: Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left: Math.round(30 * MM_TO_DXA),
  right: Math.round(15 * MM_TO_DXA),
};

export const run = (text, opts = {}) =>
  new TextRun({ text, font: FONT, size: BODY_SIZE, ...opts });

export const centered = (runs, spacing = {}) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });

export const emptyLine = () => centered(run(""));

export const body = (text, opts = {}) =>
  new Paragraph({
    alignment: opts.alignment ?? AlignmentType.JUSTIFIED,
    indent: { firstLine: opts.firstLine ?? FIRST_LINE },
    spacing: {
      before: opts.before ?? 0, after: opts.after ?? 0,
      line: LINE_15, lineRule: "auto",
    },
    keepNext: opts.keepNext ?? false,
    children: [run(text)],
  });

// Заголовок структурної одиниці чи розділу: ВЕЛИКИМИ, напівжирно, по центру,
// з нової сторінки; стиль Heading 1 — потрапляє до змісту
export const h1 = (title, opts = {}) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: opts.pageBreakBefore ?? true,
    keepNext: true,
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360, line: LINE_15, lineRule: "auto" },
    children: [run(title.toUpperCase(), { bold: true })],
  });

// Заголовок підрозділу: з абзацного відступу, напівжирно; Heading 2 — у змісті
export const h2 = (title) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    indent: { firstLine: FIRST_LINE },
    spacing: { before: 360, after: 360, line: LINE_15, lineRule: "auto" },
    children: [run(title, { bold: true })],
  });

// Псевдозаголовок, що НЕ потрапляє до змісту (РЕФЕРАТ, ЗМІСТ)
export const plainHeading = (title, opts = {}) =>
  new Paragraph({
    pageBreakBefore: opts.pageBreakBefore ?? true,
    keepNext: true,
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360, line: LINE_15, lineRule: "auto" },
    children: [run(title.toUpperCase(), { bold: true })],
  });

export const codeParagraph = (text, { keepLines = false, keepNext = false } = {}) =>
  new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    indent: { left: 283 },
    shading: { fill: "F2F2F2", color: "auto", type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: "4472C4", space: 4 } },
    keepLines, keepNext,
    children: [new TextRun({ text: text || " ", font: FONT_CODE, size: CODE_SIZE })],
  });

export const codeBlock = (text) => {
  const lines = text.split("\n");
  const keep = lines.length <= 25;
  return lines.map((line, i) =>
    codeParagraph(line, { keepLines: keep, keepNext: keep && i < lines.length - 1 }));
};

// Рисунок з підписом за ДСТУ: зображення по центру, під ним «Рисунок N – Назва» (коротке тире, як у зразку методички)
export const figure = (path, { width, height, number, caption }) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    keepNext: true,
    spacing: { before: 120, after: 0, line: LINE_15, lineRule: "auto" },
    children: [new ImageRun({
      data: readFileSync(path),
      transformation: { width, height },
      type: "png",
    })],
  }),
  centered(run(`Рисунок ${number} – ${caption}`), { before: 60, after: 240 }),
];
