#!/usr/bin/env node
/**
 * CLI-обёртка: запускает Python-скрипт заполнения дневника практики ПГТУ.
 * Требуется Python 3 и python-docx (pip install -r requirements.txt).
 *
 * Использование:
 *   npx dnevnik-praktiki-pgtu data.json "Дневник_Иванов.docx"
 */
const { spawnSync } = require("child_process");
const path = require("path");

const script = path.join(__dirname, "..", "skills", "dnevnik-praktiki", "scripts", "fill_diary.py");
const args = process.argv.slice(2);

if (args.includes("-h") || args.includes("--help") || args.length === 0) {
  console.log(
    "Дневник практики ПГТУ\n\n" +
    "  npx dnevnik-praktiki-pgtu <data.json> [output.docx]\n\n" +
    "data.json — данные студента (см. skills/dnevnik-praktiki/assets/data_example.json).\n" +
    "Требуется Python 3 и python-docx: pip install -r requirements.txt"
  );
  process.exit(args.length === 0 ? 1 : 0);
}

let py = spawnSync("python3", [script, ...args], { stdio: "inherit" });
if (py.error && py.error.code === "ENOENT") {
  py = spawnSync("python", [script, ...args], { stdio: "inherit" });
}
if (py.error) {
  console.error("Не найден Python. Установите Python 3 и python-docx.");
  process.exit(1);
}
process.exit(py.status === null ? 1 : py.status);
