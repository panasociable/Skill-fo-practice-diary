#!/usr/bin/env python3
"""Заполняет дневник практики ПГТУ из JSON с данными студента.

Запуск:
    python3 fill_diary.py data.json "Дневник_Иванов.docx"

Поля JSON (см. data_example.json). Обязательные: fio, kafedra, spec,
forma, kurs, gruppa, mesto, start_date, end_date, zadanie.
Даты в формате ГГГГ-ММ-ДД. План работ (8 пунктов) и его даты
рассчитываются автоматически по срокам практики.
"""
import json
import os
import re
import sys
from datetime import date, timedelta
from docx import Document

MONTHS_GEN = ["", "января","февраля","марта","апреля","мая","июня",
              "июля","августа","сентября","октября","ноября","декабря"]
REQUIRED_FIELDS = (
    "fio", "kafedra", "spec", "kurs", "gruppa", "mesto",
    "start_date", "end_date", "zadanie",
)
PLACEHOLDER_RE = re.compile(r"\{\{[^}]+}}")

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "template.docx")

def fail(message, code=1):
    print("Ошибка:", message, file=sys.stderr)
    raise SystemExit(code)

def parse(value, field_name):
    if not isinstance(value, str):
        fail("поле %s должно быть строкой с датой в формате ГГГГ-ММ-ДД" % field_name)
    try:
        y,m,d = map(int, value.split("-"))
        return date(y,m,d)
    except ValueError:
        fail("поле %s должно быть датой в формате ГГГГ-ММ-ДД" % field_name)

def validate(data):
    missing = [k for k in REQUIRED_FIELDS if data.get(k) in (None, "")]
    if missing:
        fail("не хватает обязательных полей: " + ", ".join(missing))

    start = parse(data["start_date"], "start_date")
    end = parse(data["end_date"], "end_date")
    if end < start:
        fail("end_date не может быть раньше start_date")
    return start, end

def split_periods(start, end, n=8):
    """Делит срок практики на n этапов, формат 'ДД.ММ-ДД.ММ'."""
    total = (end - start).days + 1
    out = []
    for i in range(n):
        a_offset = min((total*i)//n, total - 1)
        b_offset = min(max((total*(i+1))//n - 1, a_offset), total - 1)
        a = start + timedelta(days=a_offset)
        b = start + timedelta(days=b_offset)
        if b < a: b = a
        out.append("%02d.%02d-%02d.%02d" % (a.day, a.month, b.day, b.month))
    return out

def set_text(p, new):
    if not p.runs:
        p.add_run(new); return
    p.runs[0].text = new
    for r in p.runs[1:]:
        r.text = ""

def replace_all(doc, mapping):
    def fix(p):
        if "{{" not in p.text: return
        t = p.text
        for k,v in mapping.items():
            t = t.replace(k, v)
        set_text(p, t)
    for p in doc.paragraphs: fix(p)
    for tb in doc.tables:
        for r in tb.rows:
            for c in r.cells:
                for p in c.paragraphs: fix(p)

def find_placeholders(doc):
    found = set()
    for p in doc.paragraphs:
        found.update(PLACEHOLDER_RE.findall(p.text))
    for tb in doc.tables:
        for r in tb.rows:
            for c in r.cells:
                found.update(PLACEHOLDER_RE.findall(c.text))
    return sorted(found)

def main():
    if len(sys.argv) < 2:
        fail('использование: python fill_diary.py data.json "Дневник_Иванов.docx"', code=2)

    try:
        with open(sys.argv[1], encoding="utf-8") as f:
            data = json.load(f)
    except OSError as exc:
        fail("не удалось прочитать JSON-файл: %s" % exc)
    except json.JSONDecodeError as exc:
        fail("некорректный JSON: %s" % exc)
    if not isinstance(data, dict):
        fail("JSON должен быть объектом с полями дневника")
    out = sys.argv[2] if len(sys.argv) > 2 else "Дневник_практики.docx"

    s, e = validate(data)
    periods = split_periods(s, e)

    g = data.get  # короткий доступ с дефолтами
    m = {
        "{{VID}}":    g("vid", "производственная"),
        "{{TIP}}":    g("tip", ""),
        "{{FIO}}":    data["fio"],
        "{{KAFEDRA}}": data["kafedra"],
        "{{SPEC}}":   data["spec"],
        "{{FORMA}}":  g("forma", "очная"),
        "{{KURS}}":   str(data["kurs"]),
        "{{GRUPPA}}": data["gruppa"],
        "{{MESTO}}":  data["mesto"],
        "{{D1}}": str(s.day), "{{M1}}": MONTHS_GEN[s.month], "{{Y1}}": str(s.year),
        "{{D2}}": str(e.day), "{{M2}}": MONTHS_GEN[e.month], "{{Y2}}": str(e.year),
        "{{YEAR}}": str(s.year),
        "{{INSTR_VUZ}}": g("instr_vuz", ""),
        "{{INSTR_ORG}}": g("instr_org", ""),
        "{{ZADANIE}}": data["zadanie"],
    }
    # Приказ / договор (опционально)
    def datebits(key, pre):
        v = data.get(key)
        if v:
            d = parse(v, key)
            m["{{%s_D}}"%pre]=str(d.day); m["{{%s_M}}"%pre]=MONTHS_GEN[d.month]; m["{{%s_Y}}"%pre]=str(d.year)
        else:
            m["{{%s_D}}"%pre]="___"; m["{{%s_M}}"%pre]="________"; m["{{%s_Y}}"%pre]="20__"
    m["{{PRIKAZ_N}}"] = g("prikaz_n", "______")
    m["{{DOG_N}}"]    = g("dog_n", "______")
    datebits("prikaz_date","PRIKAZ"); datebits("dog_date","DOG")
    for i,p in enumerate(periods,1):
        m["{{P%d}}"%i] = p

    doc = Document(TEMPLATE)
    replace_all(doc, m)
    leftovers = find_placeholders(doc)
    if leftovers:
        fail("в шаблоне остались незаполненные плейсхолдеры: " + ", ".join(leftovers))
    doc.save(out)
    print("Готово:", out)

if __name__ == "__main__":
    main()
