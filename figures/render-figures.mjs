import pkg from "playwright";
const { chromium } = pkg;
const tree = `oop-coursework/
└── apps/
    ├── api/            серверна частина — ASP.NET Core Web API (C#)
    │   ├── Models/         класи предметної галузі
    │   ├── Storage/        сховище даних (JsonDataStore)
    │   ├── Services/       керуюча логіка (DirectoryService)
    │   ├── Contracts/      типи запитів і відповідей API
    │   ├── Controllers/    HTTP-контролери
    │   └── Data/           JSON-файли бази даних
    ├── api.tests/      модульні тести (xUnit)
    └── web/            клієнтська частина — React, TypeScript`;

const code = `/// <summary>
/// Пошук мінімального конкурсу з даної спеціальності за обраною
/// формою навчання. Вузи, де форма не ведеться (null), пропускаються.
/// </summary>
public MinCompetitionResult? GetMinCompetition(string name, StudyForm form)
{
    var best = _store.Specialties
        .Where(s => string.Equals(s.Name, name.Trim(),
            StringComparison.OrdinalIgnoreCase))
        .Where(s => s.Competition.ByForm(form).HasValue)
        .OrderBy(s => s.Competition.ByForm(form)!.Value)
        .FirstOrDefault();
    return best is null
        ? null
        : new MinCompetitionResult(GetUniversity(best.UniversityId),
            best, form, best.Competition.ByForm(form)!.Value);
}`;

const esc = (s) => s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const html = (text) => `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body { margin:0; background:#fff; }
  pre { display:inline-block; margin:0; background:#fff;
        border-left:3px solid #4472C4; padding:14px 22px;
        font-family:'Courier New',monospace; font-size:16px; line-height:1.55;
        color:#1a1a1a; white-space:pre; }
</style></head><body><pre id="t">${esc(text)}</pre></body></html>`;

const b = await chromium.launch();
for (const [name, text] of [["structure", tree], ["min-competition", code]]) {
  const p = await b.newPage({ deviceScaleFactor: 2 });
  await p.setContent(html(text));
  const el = await p.$("#t");
  await el.screenshot({ path: `./${name}.png` });
  await p.close();
}
await b.close();
console.log("done");
