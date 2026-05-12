#!/usr/bin/env node

/**
 * DOCX report generator for Lab 1 — Algorithms & Data Structures
 * Topic: Studying time characteristics of algorithms
 *
 * Usage: npm run report
 */

import {
  Document, Packer, Paragraph, TextRun, Header,
  AlignmentType, PageNumber, HeadingLevel,
  PageBreak, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign,
} from "docx";
import { readFileSync, writeFileSync, readdirSync, statSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// ─── Constants ───────────────────────────────────────────────────────

const MM_TO_DXA = 56.693;
const PT_TO_HALF_PT = 2;

const FONT = "Times New Roman";
const FONT_CODE = "Courier New";
const BODY_SIZE = 14 * PT_TO_HALF_PT;
const CODE_SIZE = 10 * PT_TO_HALF_PT;
const TITLE_SIZE = 14 * PT_TO_HALF_PT;
const LINE_SPACING_15 = 360;

const margins = {
  top: Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left: Math.round(30 * MM_TO_DXA),
  right: Math.round(15 * MM_TO_DXA),
};

// ─── Helpers (from lab-report.mjs template) ──────────────────────────

function titleRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: TITLE_SIZE,
    bold: opts.bold ?? false,
    ...opts,
  });
}

function bodyRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: BODY_SIZE,
    ...opts,
  });
}

function centeredParagraph(runs, spacing = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function emptyLine() {
  return centeredParagraph(titleRun(""));
}

function sectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [
      new TextRun({
        text: `${number} ${title}`.toUpperCase(),
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function subsectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      new TextRun({
        text: `${number} ${title}`,
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function bodyParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(text)],
  });
}

function codeParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    children: [
      new TextRun({
        text,
        font: FONT_CODE,
        size: CODE_SIZE,
      }),
    ],
  });
}

function listingCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      bodyRun(`Лістинг ${number} — ${title}`),
    ],
  });
}

function tableCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      bodyRun(`Таблиця ${number} — ${title}`),
    ],
  });
}

// ─── Code block helper ───────────────────────────────────────────────

function codeBlock(text) {
  return text.split("\n").map((line) => codeParagraph(line));
}

// ─── Source code reader ──────────────────────────────────────────────

function readSourceFiles() {
  const srcDir = join(__dirname, "..", "..", "src");
  const files = [];

  function walk(dir, prefix = "") {
    const entries = readdirSync(dir, { withFileTypes: true }).sort((a, b) =>
      a.name.localeCompare(b.name),
    );
    for (const entry of entries) {
      const fullPath = join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath, prefix ? `${prefix}/${entry.name}` : entry.name);
      } else if (entry.name.endsWith(".ts")) {
        const relativeName = prefix ? `${prefix}/${entry.name}` : entry.name;
        files.push({
          name: relativeName,
          content: readFileSync(fullPath, "utf-8"),
        });
      }
    }
  }

  walk(srcDir);
  return files;
}

// ─── Benchmark results table ─────────────────────────────────────────

const BORDER = {
  style: BorderStyle.SINGLE,
  size: 1,
  color: "000000",
};

const TABLE_BORDERS = {
  top: BORDER,
  bottom: BORDER,
  left: BORDER,
  right: BORDER,
  insideHorizontal: BORDER,
  insideVertical: BORDER,
};

const CELL_MARGINS = {
  top: 40,
  bottom: 40,
  left: 80,
  right: 80,
};

function headerCell(text) {
  return new TableCell({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
        children: [bodyRun(text, { bold: true })],
      }),
    ],
    shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL_MARGINS,
  });
}

function dataCell(text, alignment = AlignmentType.CENTER) {
  return new TableCell({
    children: [
      new Paragraph({
        alignment,
        spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
        children: [bodyRun(text)],
      }),
    ],
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL_MARGINS,
  });
}

function buildBenchmarkTable() {
  const data = [
    ["1 000", "0.0006", "0.0000", "0.0001"],
    ["10 000", "0.0058", "0.0001", "0.0001"],
    ["100 000", "0.0595", "0.0001", "0.0002"],
    ["1 000 000", "0.5814", "0.0001", "0.0002"],
    ["10 000 000", "5.9384", "0.0001", "0.0002"],
  ];

  const headerRow = new TableRow({
    children: [
      headerCell("n"),
      headerCell("Linear (ms)"),
      headerCell("Binary iter (ms)"),
      headerCell("Binary rec (ms)"),
    ],
    tableHeader: true,
  });

  const dataRows = data.map(
    ([n, lin, binI, binR]) =>
      new TableRow({
        children: [
          dataCell(n),
          dataCell(lin),
          dataCell(binI),
          dataCell(binR),
        ],
      }),
  );

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: TABLE_BORDERS,
  });
}

// ─── TITLE PAGE ──────────────────────────────────────────────────────

const titlePageParagraphs = [
  centeredParagraph(titleRun("Міністерство освіти і науки України")),
  centeredParagraph(titleRun("Харківський національний університет радіоелектроніки")),
  emptyLine(),
  centeredParagraph(titleRun("Кафедра програмної інженерії")),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("ЗВІТ", { bold: true })),
  centeredParagraph(titleRun("з лабораторної роботи № 1")),
  centeredParagraph(titleRun("з дисципліни «Алгоритми та структури даних»")),
  centeredParagraph(titleRun("на тему: «Вивчення часових характеристик алгоритмів»")),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("Варіант 1")),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Виконав: ст. гр. ПЗПІ-25-6")],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Коновалов О. О.")],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Перевірив: Олійник О. О.")],
  }),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("Харків — 2026")),
];

// ─── BODY SECTIONS ───────────────────────────────────────────────────

const bodyParagraphs = [
  new Paragraph({ children: [new PageBreak()] }),

  // 1 МЕТА РОБОТИ
  sectionHeading("1", "Мета роботи"),
  bodyParagraph(
    "Розвиток навичок теоретичної та практичної оцінки тимчасових характеристик алгоритмів.",
  ),

  // 2 ЗАВДАННЯ
  sectionHeading("2", "Завдання"),
  bodyParagraph(
    "Порівняти алгоритми лінійного пошуку та двійкового пошуку, представленого циклами та рекурсією. Для обраного завдання скласти алгоритми, знайти асимптотичні оцінки, скласти програми, для різних обсягів вхідних даних отримати час виконання та порівняти теоретичні оцінки з практичними.",
  ),

  // 3 ХІД РОБОТИ
  sectionHeading("3", "Хід роботи"),

  // 3.1 Алгоритми пошуку
  subsectionHeading("3.1", "Алгоритми пошуку"),
  bodyParagraph(
    "Нижче наведено псевдокод трьох алгоритмів пошуку, які порівнюються у цій лабораторній роботі.",
  ),

  listingCaption("3.1", "Лінійний пошук"),
  ...codeBlock(
    `function linearSearch(arr, target):
    for i = 0 to n-1:
        if arr[i] == target:
            return i
    return -1`,
  ),

  listingCaption("3.2", "Двійковий пошук (ітеративний)"),
  ...codeBlock(
    `function binarySearchIterative(arr, target):
    low = 0, high = n-1
    while low <= high:
        mid = floor((low + high) / 2)
        if arr[mid] == target: return mid
        if arr[mid] < target: low = mid + 1
        else: high = mid - 1
    return -1`,
  ),

  listingCaption("3.3", "Двійковий пошук (рекурсивний)"),
  ...codeBlock(
    `function binarySearchRecursive(arr, target, low, high):
    if low > high: return -1
    mid = floor((low + high) / 2)
    if arr[mid] == target: return mid
    if arr[mid] < target:
        return binarySearchRecursive(arr, target, mid+1, high)
    return binarySearchRecursive(arr, target, low, mid-1)`,
  ),

  // 3.2 Теоретичний аналіз складності
  subsectionHeading("3.2", "Теоретичний аналіз складності"),
  bodyParagraph(
    "Для кожного алгоритму проведено аналіз часової складності шляхом підрахунку вартості та кількості виконань кожної операції.",
  ),

  bodyParagraph("Лінійний пошук (найгірший випадок — елемент не знайдено):"),
  ...codeBlock(
    `Рядок 1: for i = 0 to n-1       — вартість c1, виконується n разів
Рядок 2:     if arr[i] == target  — вартість c2, виконується n разів
Рядок 3: return -1                — вартість c4, виконується 1 раз

T(n) = c1·n + c2·n + c4 = Θ(n)`,
  ),

  bodyParagraph("Двійковий пошук ітеративний (найгірший випадок):"),
  ...codeBlock(
    `Рядок 1: low = 0, high = n-1             — вартість c1, виконується 1 раз
Рядок 2: while low <= high                — вартість c2, виконується ⌊log₂n⌋+1 разів
Рядок 3:     mid = floor((low+high)/2)    — вартість c3, виконується ⌊log₂n⌋ разів
Рядок 4:     if arr[mid] == target         — вартість c4, виконується ⌊log₂n⌋ разів
Рядок 5:     if arr[mid] < target          — вартість c5, виконується ⌊log₂n⌋ разів
Рядок 6:         low = mid+1 / high=mid-1  — вартість c6, виконується ⌊log₂n⌋ разів
Рядок 7: return -1                         — вартість c7, виконується 1 раз

T(n) = c1 + (c2+c3+c4+c5+c6)·⌊log₂n⌋ + c7 = Θ(log n)`,
  ),

  // 3.3 Аналіз рекурсивного алгоритму
  subsectionHeading("3.3", "Аналіз рекурсивного алгоритму"),
  bodyParagraph(
    "Рекурсивний варіант двійкового пошуку має рекурентне співвідношення:",
  ),
  ...codeBlock("T(n) = T(n/2) + Θ(1)"),
  bodyParagraph(
    "Застосуємо основну теорему (Master Theorem): a = 1, b = 2, f(n) = Θ(1).",
  ),
  bodyParagraph(
    "Оскільки log_b(a) = log₂(1) = 0, маємо f(n) = Θ(n^0) = Θ(1), що відповідає випадку 2 основної теореми.",
  ),
  bodyParagraph("Отже: T(n) = Θ(log n)."),

  // 4 РЕЗУЛЬТАТИ
  sectionHeading("4", "Результати"),
  bodyParagraph(
    "Програму було запущено для масивів різного розміру. Результати вимірювання часу виконання наведено у таблиці 4.1.",
  ),
  tableCaption("4.1", "Результати вимірювання часу виконання алгоритмів пошуку"),
  buildBenchmarkTable(),
  bodyParagraph(
    "Як видно з таблиці, час лінійного пошуку зростає лінійно зі збільшенням n, тоді як двійковий пошук залишається майже сталим, що відповідає логарифмічній складності.",
  ),

  // 5 ВИСНОВКИ
  sectionHeading("5", "Висновки"),
  bodyParagraph(
    "Практичні виміри підтвердили теоретичні оцінки: час лінійного пошуку зростає пропорційно n, тоді як час двійкового пошуку зростає логарифмічно. Рекурсивний та ітеративний варіанти двійкового пошуку показали близьку швидкодію, що відповідає однаковій асимптотичній оцінці Θ(log n).",
  ),
];

// ─── ДОДАТОК А — Вихідний код програми ──────────────────────────────

const appendixParagraphs = [
  new Paragraph({ children: [new PageBreak()] }),
  sectionHeading("", "Додаток А"),
  centeredParagraph(bodyRun("Вихідний код програми", { bold: true })),
  emptyLine(),
];

const sourceFiles = readSourceFiles();
let listingCounter = 1;
for (const file of sourceFiles) {
  appendixParagraphs.push(
    listingCaption(`А.${listingCounter}`, file.name),
  );
  appendixParagraphs.push(...codeBlock(file.content.trimEnd()));
  appendixParagraphs.push(emptyLine());
  listingCounter++;
}

// ─── DOCUMENT ────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: FONT,
          size: BODY_SIZE,
          language: { value: "uk-UA" },
        },
        paragraph: {
          spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 1,
        },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { ...margins, header: 708, footer: 708 },
        },
        titlePage: true,
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  children: [PageNumber.CURRENT],
                  font: FONT,
                  size: BODY_SIZE,
                }),
              ],
            }),
          ],
        }),
      },
      children: [
        ...titlePageParagraphs,
        ...bodyParagraphs,
        ...appendixParagraphs,
      ],
    },
  ],
});

// ─── Output ──────────────────────────────────────────────────────────

const outputPath = join(__dirname, "Звіт_ЛР1_Коновалов_ПЗПІ-25-6.docx");
const buffer = await Packer.toBuffer(doc);
writeFileSync(outputPath, buffer);
console.log(`Created: ${outputPath}`);
