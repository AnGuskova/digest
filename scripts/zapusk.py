#!/usr/bin/env python3
"""Подготовка дайджеста через API любой модели в формате OpenAI Chat Completions.

Подходит для OpenAI, DeepSeek, Anthropic (режим совместимости с OpenAI),
OpenRouter, YandexGPT, локальных Ollama и LM Studio и других провайдеров
с тем же форматом запроса. Сторонние библиотеки не нужны.

Настройка: переменные окружения или файл .env в корне репозитория
(образец: .env.example):
    AI_BASE_URL     адрес API, например https://api.deepseek.com
    AI_API_KEY      ключ API (для локальной модели можно не задавать)
    AI_MODEL        имя модели
    AI_AUTH_SCHEME  схема авторизации, по умолчанию Bearer

Запуск:
    python3 scripts/zapusk.py --period "III квартал 2026" [--data-sreza 05.10.2026]
        [--adresat внутренний|внешний] [--imya 3кв-2026] файл1.md файл2.txt ...

Входные файлы: текст (.md, .txt, .csv). PDF и DOCX сначала сохраните как текст.
Результат: workspace/outputs/ГГГГ-ММ-ДД_дайджест_<имя>.md, затем проверка proverka.py.
"""

import argparse
import datetime
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

KOREN = Path(__file__).resolve().parent.parent


def zagruzit_env():
    env = KOREN / ".env"
    if not env.exists():
        return
    for s in env.read_text(encoding="utf-8").splitlines():
        s = s.strip()
        if s and not s.startswith("#") and "=" in s:
            k, v = s.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"'))


def prochitat(put):
    return Path(put).read_text(encoding="utf-8")


def sistemnyy_prompt():
    chasti = [
        ("AGENTS.md", prochitat(KOREN / "AGENTS.md")),
        ("templates/daydzhest.md", prochitat(KOREN / "templates/daydzhest.md")),
        (
            "nastroyki/rabochie-gruppy.md",
            prochitat(KOREN / "nastroyki/rabochie-gruppy.md"),
        ),
        (
            "nastroyki/dorozhnaya-karta.md",
            prochitat(KOREN / "nastroyki/dorozhnaya-karta.md"),
        ),
    ]
    return "\n\n".join(f"===== Файл {imya} =====\n{tekst}" for imya, tekst in chasti)


def zapros_polzovatelya(a, fayly):
    shapka = (
        "Подготовь дайджест Центра компетенций по Открытым API по правилам AGENTS.md и шаблону templates/daydzhest.md.\n"
        f"Период: {a.period}\nДата среза: {a.data_sreza}\nАдресат: {a.adresat}\n"
        "Источники: только материалы ниже. Других источников у тебя нет; чего нет в материалах, "
        "в дайджест не входит, а вопрос идёт в «Открытые вопросы».\n"
        "Верни только файл дайджеста в Markdown, без пояснений до и после.\n"
    )
    materialy = "\n\n".join(
        f"===== Материал: {Path(f).name} =====\n{prochitat(f)}" for f in fayly
    )
    return shapka + "\n" + materialy


def vyzov(sistema, polzovatel):
    baza = os.environ.get("AI_BASE_URL", "").rstrip("/")
    model = os.environ.get("AI_MODEL", "")
    if not baza or not model:
        sys.exit(
            "Не заданы AI_BASE_URL и AI_MODEL: см. .env.example и docs/nastroyka-ii.md"
        )
    url = baza if baza.endswith("/chat/completions") else baza + "/chat/completions"
    zagolovki = {"Content-Type": "application/json"}
    kluch = os.environ.get("AI_API_KEY", "")
    if kluch:
        zagolovki["Authorization"] = (
            f"{os.environ.get('AI_AUTH_SCHEME', 'Bearer')} {kluch}"
        )
    telo = {
        "model": model,
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": sistema},
            {"role": "user", "content": polzovatel},
        ],
    }
    zapros = urllib.request.Request(
        url, data=json.dumps(telo).encode("utf-8"), headers=zagolovki
    )
    try:
        with urllib.request.urlopen(zapros, timeout=600) as otvet:
            dannye = json.load(otvet)
    except urllib.error.HTTPError as e:
        sys.exit(f"Ошибка API {e.code}: {e.read().decode('utf-8', 'replace')[:1000]}")
    return dannye["choices"][0]["message"]["content"]


def vyrezat_markdown(tekst):
    """Если модель обернула файл в блок ```markdown, берётся содержимое блока."""
    m = re.search(r"```(?:markdown|md)?\s*\n(.*\n)```", tekst, re.S)
    return m.group(1) if m and "# Открытые API" in m.group(1) else tekst


def main():
    zagruzit_env()
    p = argparse.ArgumentParser(description="Дайджест через API модели")
    p.add_argument("--period", required=True, help="например «III квартал 2026»")
    p.add_argument("--data-sreza", default=datetime.date.today().strftime("%d.%m.%Y"))
    p.add_argument("--adresat", default="внутренний", choices=["внутренний", "внешний"])
    p.add_argument("--imya", help="период для имени файла: 3кв-2026, 1пг-2026, 2026")
    p.add_argument("fayly", nargs="+", help="материалы за период (текстовые файлы)")
    a = p.parse_args()

    otvet = vyrezat_markdown(vyzov(sistemnyy_prompt(), zapros_polzovatelya(a, a.fayly)))

    data = datetime.datetime.strptime(a.data_sreza, "%d.%m.%Y").strftime("%Y-%m-%d")
    imya = a.imya or re.sub(r"\s+", "-", a.period.strip())
    put = KOREN / "workspace" / "outputs" / f"{data}_дайджест_{imya}.md"
    put.parent.mkdir(parents=True, exist_ok=True)
    put.write_text(otvet, encoding="utf-8")
    print(f"Сохранено: {put.relative_to(KOREN)}")

    rez = subprocess.run(
        [sys.executable, str(KOREN / "scripts" / "proverka.py"), str(put)]
    )
    if rez.returncode:
        print("Проверка нашла нарушения: исправьте файл вручную или повторите запуск.")
    return rez.returncode


if __name__ == "__main__":
    sys.exit(main())
