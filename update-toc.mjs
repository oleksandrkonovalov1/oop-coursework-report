// Вимірює номери сторінок заголовків у фінальному PDF і вписує їх у content/toc-data.mjs.
// Запуск з кореня репо: node reports/coursework/update-toc.mjs
// Двопрохідний: генерація DOCX → PDF → pdftotext по сторінках → оновлення toc-data → повторна генерація.
import { execSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const dir = dirname(fileURLToPath(import.meta.url));
const docx = join(dir, "2026_ПІ_ООП_ПЗПІ-25-6_Коновалов_О_О.docx");
const pdf = "/tmp/toc-measure.pdf";

const generate = () => execSync(`node ${join(dir, "generate-report.mjs")}`, { stdio: "inherit" });
const toPdf = () => {
  execSync(`soffice --headless --convert-to pdf --outdir /tmp "${docx}"`, { stdio: "pipe" });
  execSync(`mv "/tmp/${docx.split("/").pop().replace(/\.docx$/, ".pdf")}" ${pdf}`);
};
const pageCount = () => Number(execSync(`pdfinfo ${pdf}`).toString().match(/Pages:\s+(\d+)/)[1]);
const pageText = (n) => execSync(`pdftotext -f ${n} -l ${n} ${pdf} -`).toString();

for (const pass of [1, 2]) {
  generate();
  toPdf();
  const total = pageCount();
  const texts = [];
  for (let i = 1; i <= total; i++) texts.push(pageText(i));

  const dataPath = join(dir, "content/toc-data.mjs");
  let data = readFileSync(dataPath, "utf8");
  let changed = false;
  // Заголовок шукаємо З КІНЦЯ документа для додатка/висновків? Ні — перше входження
  // ПІСЛЯ сторінки змісту (зміст сам містить ці рядки). Зміст — сторінка 5.
  const TOC_PAGE = 5;
  for (const m of data.matchAll(/\{ text: "(.*?)", search: "(.*?)", indent: (\d), page: (\d+) \}/g)) {
    const [full, text, search, indent, oldPage] = m;
    const found = texts.findIndex((t, i) => i + 1 > TOC_PAGE && t.includes(search)) + 1;
    if (!found) { console.error(`НЕ ЗНАЙДЕНО: «${search}»`); process.exitCode = 1; continue; }
    if (Number(oldPage) !== found) {
      data = data.replace(full, `{ text: "${text}", search: "${search}", indent: ${indent}, page: ${found} }`);
      changed = true;
      console.log(`«${text}» → с. ${found}`);
    }
  }
  writeFileSync(dataPath, data);
  if (!changed) { console.log(`Прохід ${pass}: номери стабільні (${total} с.) — готово`); break; }
  if (pass === 2) console.warn("УВАГА: номери змінились і на 2-му проході — перевір пагінацію вручну");
}
