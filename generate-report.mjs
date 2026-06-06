#!/usr/bin/env node
// Генератор пояснювальної записки курсової роботи «Довідник абітурієнта».
// Структура — за методичкою курсової (НЕ лаб-звіт): титульний аркуш (дод. Г),
// аркуш завдання з календарним планом (дод. Д), реферат (дод. Е), зміст,
// вступ, 1 Опис вимог, 2 Проєктування програми, висновки, джерела, додаток А.
// Запуск: pnpm report (з кореня репозиторію).

import {
  Document, Packer, Paragraph, TextRun, Header,
  AlignmentType, PageNumber, TabStopType, LeaderType,
  Table, TableRow, TableCell, WidthType, VerticalAlign,
} from "docx";
import { writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import {
  FONT, BODY_SIZE, LINE_15, FIRST_LINE, margins, MM_TO_DXA,
  run, centered, emptyLine, body, h1, h2, plainHeading, figure,
} from "./content/helpers.mjs";
import { introParagraphs } from "./content/intro.mjs";
import { requirementsIntro, scenariosIntro, scenarios } from "./content/scenarios.mjs";
import { functionsContent } from "./content/functions.mjs";
import { designContent } from "./content/design.mjs";
import { conclusionsParagraphs } from "./content/conclusions.mjs";
import { sources } from "./content/sources.mjs";
import { tocEntries } from "./content/toc-data.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));

// Обсяг записки для реферату — звірено з PDF після рендеру (Step 5)
const PAGES_TOTAL = "26";
const FIGURES_TOTAL = "6";
const SOURCES_TOTAL = "6";

// ─── ТИТУЛЬНИЙ АРКУШ (додаток Г методички) ──────────────────────────

const rightLine = (text) =>
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_15, lineRule: "auto" },
    children: [run(text)],
  });

const titlePage = [
  centered(run("Міністерство освіти і науки України")),
  centered(run("Харківський національний університет радіоелектроніки")),
  emptyLine(),
  centered(run("Кафедра програмної інженерії")),
  emptyLine(), emptyLine(),
  centered(run("КУРСОВА РОБОТА", { bold: true })),
  centered(run("Пояснювальна записка")),
  centered(run("з дисципліни «Об'єктно-орієнтоване програмування»")),
  centered(run("ДОВІДНИК АБІТУРІЄНТА")),
  emptyLine(), emptyLine(),
  rightLine("Виконав:"),
  rightLine("здобувач 1 року навчання,"),
  rightLine("групи ПЗПІ-25-6,"),
  rightLine("Олександр КОНОВАЛОВ"),
  emptyLine(),
  rightLine("Керівник:"),
  rightLine("ст. викл. Юлія ЧЕРЕПАНОВА"),
  emptyLine(),
  rightLine("Комісія: проф. Володимир БОНДАРЄВ"),
  rightLine("ст. викл. Юлія ЧЕРЕПАНОВА"),
  rightLine("ст. викл. Віталій ЛЯПОТА"),
  emptyLine(),
  centered(run("Харків 2026")),
];

// ─── АРКУШ ЗАВДАННЯ + КАЛЕНДАРНИЙ ПЛАН (додаток Д) ──────────────────

const cell = (text, { width, align = AlignmentType.LEFT, bold = false } = {}) =>
  new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: align,
      spacing: { after: 0, line: 240, lineRule: "auto" },
      children: [run(text, { bold })],
    })],
  });

const planStages = [
  ["1", "Видача і затвердження теми курсової роботи", "25.03 – 01.04.2026"],
  ["2", "Формулювання мети роботи", "01.04 – 08.04.2026"],
  ["3", "Опис вимог до програми", "08.04 – 15.04.2026"],
  ["4", "Проєктування програми", "15.04 – 22.04.2026"],
  ["5", "Кодування програми", "22.04 – 06.05.2026"],
  ["6", "Підготовка до захисту", "13.05 – 20.05.2026"],
  ["7", "Захист курсової роботи", "20.05 – 30.05.2026"],
];

const planTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        cell("№", { width: 7, align: AlignmentType.CENTER, bold: true }),
        cell("Назва етапу", { width: 48, align: AlignmentType.CENTER, bold: true }),
        cell("Термін виконання", { width: 27, align: AlignmentType.CENTER, bold: true }),
        cell("Примітка", { width: 18, align: AlignmentType.CENTER, bold: true }),
      ],
    }),
    ...planStages.map(([n, stage, term], i) => new TableRow({
      children: [
        cell(n, { align: AlignmentType.CENTER }),
        cell(stage),
        cell(term, { align: AlignmentType.CENTER }),
        // останній етап (захист) ще не відбувся на момент подання записки
        cell(i === planStages.length - 1 ? "" : "Виконано", { align: AlignmentType.CENTER }),
      ],
    })),
  ],
});

const taskSheet = [
  new Paragraph({
    pageBreakBefore: true,
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_15, lineRule: "auto" },
    children: [run("Харківський національний університет радіоелектроніки")],
  }),
  emptyLine(),
  body("Кафедра: програмної інженерії", { firstLine: 0 }),
  body("Дисципліна: «Об'єктно-орієнтоване програмування»", { firstLine: 0 }),
  body("Рівень вищої освіти: перший (бакалаврський)", { firstLine: 0 }),
  body("Спеціальність: F2 Інженерія програмного забезпечення", { firstLine: 0 }),
  body("Освітня програма: «Програмна інженерія»", { firstLine: 0 }),
  emptyLine(),
  centered(run("ЗАВДАННЯ", { bold: true })),
  centered(run("на курсову роботу")),
  centered(run("здобувача вищої освіти групи ПЗПІ-25-6 Коновалова Олександра")),
  emptyLine(),
  body("1. Тема роботи: «Довідник абітурієнта»."),
  body("2. Термін подання здобувачем завершеної роботи: 30 травня 2026 р."),
  body(
    "3. Вихідні дані до роботи: база вузів: найменування, адреса, перелік " +
    "спеціальностей, конкурс минулого року за кожною спеціальністю (денною, " +
    "вечірньою, заочною формами), розмір оплати при договірному навчанні. " +
    "Вибір за різними критеріями: все щодо обраного вузу; все щодо обраної " +
    "спеціальності, пошук мінімального конкурсу з даної спеціальності та інше.",
  ),
  body(
    "4. Перелік питань, що їх належить розробити: вступ, опис вимог, " +
    "проєктування програми, висновки.",
  ),
  body("5. Календарний план виконання роботи:", { keepNext: true, after: 120 }),
  planTable,
  emptyLine(),
  body("Дата видачі завдання: «21» лютого 2026 р.", { firstLine: 0 }),
  emptyLine(),
  body("Здобувач вищої освіти  ________________  Олександр КОНОВАЛОВ", { firstLine: 0 }),
  emptyLine(),
  body("Керівник роботи  ________________  ст. викл. Юлія ЧЕРЕПАНОВА", { firstLine: 0 }),
];

// ─── РЕФЕРАТ (додаток Е) ─────────────────────────────────────────────

const abstract = [
  plainHeading("РЕФЕРАТ", { pageBreakBefore: false }),
  body(
    `Пояснювальна записка до курсової роботи: ${PAGES_TOTAL} с., ` +
    `${FIGURES_TOTAL} рис., 1 додаток, ${SOURCES_TOTAL} джерел.`,
  ),
  emptyLine(),
  body(
    "ДОВІДНИК, АБІТУРІЄНТ, ВУЗ, СПЕЦІАЛЬНІСТЬ, КОНКУРС, МОВА C#, ООП, .NET, " +
    "REACT",
  ),
  emptyLine(),
  body(
    "Мета роботи — розробити програму-довідник абітурієнта для ведення " +
    "власної бази вузів та їхніх спеціальностей і виконання запитів, що " +
    "допомагають обрати місце вступу.",
  ),
  body(
    "У результаті виконання роботи розроблено вебзастосунок, який забезпечує " +
    "ведення колекцій двох пов'язаних сутностей — вузів і спеціальностей — з " +
    "перевіркою введених даних, пошук вузів за назвою та адресою, перегляд " +
    "усіх відомостей щодо обраного вузу, перегляд пропозицій усіх вузів за " +
    "обраною спеціальністю з фільтром за вартістю навчання, пошук " +
    "мінімального конкурсу за обраною формою навчання, а також автоматичне " +
    "збереження і завантаження бази з файлів JSON. Цілісність даних " +
    "підтримується каскадним видаленням та підтвердженням незворотних дій; " +
    "коректність керуючої логіки підтверджено модульними тестами.",
  ),
  body(
    "У процесі розробки використано: платформу .NET 10, каркас ASP.NET Core " +
    "Web API, мову C#, бібліотеку Newtonsoft.Json, бібліотеки React і " +
    "Tailwind CSS, мову TypeScript.",
  ),
];

// ─── ЗМІСТ ───────────────────────────────────────────────────────────

// Статичний зміст (а не Word-поле TableOfContents): видимий у будь-якому
// переглядачі — LibreOffice, превʼю кафедрального диска, друк — без «оновити поле».
// Номери сторінок вимірюються update-toc.mjs з фінального PDF.
const CONTENT_RIGHT = 11906 - margins.left - margins.right; // правий край тексту, DXA
const tocLine = (e) =>
  new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_RIGHT, leader: LeaderType.DOT }],
    indent: e.indent ? { left: Math.round(5 * MM_TO_DXA) } : undefined,
    spacing: { line: LINE_15 },
    children: [run(e.text), run("\t" + e.page)],
  });
const tableOfContents = [
  plainHeading("ЗМІСТ"),
  ...tocEntries.map(tocLine),
];

// ─── ВСТУП ───────────────────────────────────────────────────────────

const introSection = [h1("ВСТУП"), ...introParagraphs.map((t) => body(t))];

// ─── 1 ОПИС ВИМОГ ────────────────────────────────────────────────────

const renderScenario = (s, i) => {
  const out = [
    body(`Сценарій ${i} — ${s.title}.`, { keepNext: true, before: 120 }),
    body(`Передумова: ${s.precondition}`, { keepNext: true }),
    body("Основний сценарій:", { keepNext: true }),
  ];
  s.main.forEach((step, j) => {
    const last = j === s.main.length - 1;
    out.push(body(`${j + 1}) ${step}${last ? "." : ";"}`, { keepNext: !last }));
  });
  s.alt.forEach((a) => out.push(body(a)));
  return out;
};

const renderFunctionItem = (item) =>
  typeof item === "string"
    ? [body(item)]
    : figure(join(__dirname, "sketches", item.image), {
        width: 500, height: 380,
        number: item.number, caption: item.caption,
      });

const requirementsSection = [
  h1("1 ОПИС ВИМОГ"),
  ...requirementsIntro.map((t) => body(t)),
  h2("1.1 Сценарії використання"),
  ...scenariosIntro.map((t) => body(t)),
  ...scenarios.flatMap((s, i) => renderScenario(s, i + 1)),
  h2("1.2 Функції програми"),
  ...functionsContent.flatMap(renderFunctionItem),
];

// ─── 2 ПРОЄКТУВАННЯ ПРОГРАМИ ─────────────────────────────────────────

const designSection = [h1("2 ПРОЄКТУВАННЯ ПРОГРАМИ"), ...designContent];

// ─── ВИСНОВКИ, ДЖЕРЕЛА, ДОДАТОК А ────────────────────────────────────

const conclusionsSection = [
  h1("ВИСНОВКИ"),
  ...conclusionsParagraphs.map((t) => body(t)),
];

const sourcesSection = [
  h1("ПЕРЕЛІК ДЖЕРЕЛ ПОСИЛАННЯ"),
  ...sources.map((s, i) => body(`${i + 1}. ${s}`)),
];

const appendixSection = [
  h1("ДОДАТОК А"),
  centered(run("Посилання на код програми")),
  emptyLine(),
  body(
    "Вихідний код програми розміщено у репозиторії GitHub: " +
    "https://github.com/oleksandrkonovalov1/oop-coursework.",
  ),
];

// ─── ДОКУМЕНТ ────────────────────────────────────────────────────────

const pageNumberHeader = new Header({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: BODY_SIZE })],
  })],
});

const emptyHeader = new Header({ children: [new Paragraph({ children: [] })] });

const pageSetup = {
  page: {
    size: { width: 11906, height: 16838 }, // A4
    margin: { ...margins, header: 708, footer: 708 },
  },
};

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: BODY_SIZE, language: { value: "uk-UA" } },
        paragraph: { spacing: { after: 0, line: LINE_15, lineRule: "auto" } },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT, color: "000000" },
        paragraph: {
          spacing: { before: 0, after: 360, line: LINE_15, lineRule: "auto" },
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT, color: "000000" },
        paragraph: {
          spacing: { before: 360, after: 360, line: LINE_15, lineRule: "auto" },
          indent: { firstLine: FIRST_LINE },
          outlineLevel: 1,
        },
      },
    ],
  },
  sections: [
    {
      // Титульний аркуш і аркуш завдання: рахуються, але не нумеруються
      properties: pageSetup,
      headers: { default: emptyHeader },
      children: [...titlePage, ...taskSheet],
    },
    {
      // Решта записки: номер сторінки праворуч угорі (нумерація триває: 3, 4…)
      properties: pageSetup,
      headers: { default: pageNumberHeader },
      children: [
        ...abstract,
        ...tableOfContents,
        ...introSection,
        ...requirementsSection,
        ...designSection,
        ...conclusionsSection,
        ...sourcesSection,
        ...appendixSection,
      ],
    },
  ],
});

const outputPath = join(__dirname, "2026_ПІ_ООП_ПЗПІ-25-6_Коновалов_О_О.docx");
const buffer = await Packer.toBuffer(doc);
writeFileSync(outputPath, buffer);
console.log(`Created: ${outputPath}`);
