#!/usr/bin/env python3
"""Формальная проверка файла дайджеста по AGENTS.md.

Проверяет структуру, объём читательской части, трассировку, метки,
даты актуальности регуляторных документов и технические слова.
Смысл (верность цитат и источников) скрипт не проверяет.

Запуск:
    python3 scripts/proverka.py <файл или папка> [...]

Код выхода 1, если найдены ошибки.
"""

import re
import sys
from pathlib import Path

PREDEL = 3500  # знаков с пробелами в читательской части (AGENTS.md, п. 3)
ZAGOLOVOK = re.compile(r"^#\s+Открытые API: дайджест", re.M)
RAZDELITEL = re.compile(r"^---\s*$", re.M)
DATA = r"\d{2}\.\d{2}\.\d{4}"
METKA = re.compile(r"\[(ФАКТ|ПРЕДЛОЖЕНИЕ|ПРЕДПОЛОЖЕНИЕ|ИСТОЧНИК НЕ НАЙДЕН)\]")
SSYLKA = re.compile(r"\[[^\[\]]*\d{2}\.\d{2}\.\d{4}[^\[\]]*\]")
AKTUALNO = re.compile(r"актуально на\s+" + DATA, re.I)

# Регуляторные документы, для которых нужна дата актуальности (AGENTS.md, п. 2.5).
REGULYATORNYE = re.compile(
    r"СТО БР [А-ЯA-Z0-9][А-ЯA-Z.\-0-9]*[А-ЯA-Z0-9]"
    r"|ГОСТ(?: Р)? \d[\d.\-]*\d"
    r"|ISO(?:/IEC)? \d+(?:-\d+)?"
    r"|RFC \d+"
    r"|(?:Указани[еяю]|Положени[еяю]) Банка России(?: от "
    + DATA
    + r")?(?: № ?[\d\-А-Яа-я]+)?"
    r"|Федеральн\w+ закон\w*(?: от " + DATA + r")?(?: № ?[\d\-А-Яа-я]+)?"
)

# Технические подробности, которых не должно быть в читательской части (AGENTS.md, п. 4.4).
TEHNICHESKIE = re.compile(
    r"\b(?:mTLS|PKCE|OAuth|OIDC|OpenID|JWS|JWE|JWT|DPoP|FAPI|HTTP|JSON|XML|REST|endpoint|scope)\b"
    r"|эндпоинт|скоуп|код(?:ы|ов)? ошиб|`[^`]+`",
    re.I,
)


class Otchet:
    def __init__(self):
        self.oshibki = []
        self.preduprezhdeniya = []

    def err(self, mesto, tekst):
        self.oshibki.append(f"ОШИБКА  {mesto}: {tekst}")

    def warn(self, mesto, tekst):
        self.preduprezhdeniya.append(f"ВНИМАНИЕ {mesto}: {tekst}")


def razdely(tekst):
    """Заголовки любого уровня: [(номер строки, уровень, текст)]."""
    rez = []
    v_kode = False
    for n, s in enumerate(tekst.splitlines(), 1):
        if s.strip().startswith("```"):
            v_kode = not v_kode
            continue
        m = re.match(r"^(#{1,6})\s+(.*?)\s*$", s)
        if m and not v_kode:
            rez.append((n, len(m.group(1)), m.group(2)))
    return rez


def telo_razdela(tekst, nachalo):
    """Строки после заголовка в строке nachalo до следующего заголовка."""
    stroki = tekst.splitlines()
    rez = []
    for s in stroki[nachalo:]:
        if re.match(r"^#{1,6}\s", s):
            break
        rez.append(s)
    return rez


def yacheyki(s):
    s = s.strip()
    if not (s.startswith("|") and s.endswith("|")):
        return None
    return [y.strip() for y in s.strip("|").split("|")]


def tablitsa_posle(tekst, nachalo):
    """Первая таблица в разделе: (заголовок, [(номер строки, ячейки)])."""
    stroki = tekst.splitlines()
    zag, ryady = None, []
    for n in range(nachalo, len(stroki)):
        s = stroki[n]
        if re.match(r"^#{1,6}\s", s):
            break
        ya = yacheyki(s)
        if ya is None:
            if zag:
                break
            continue
        if zag is None:
            zag = ya
        elif all(re.fullmatch(r":?-{3,}:?", y) for y in ya if y):
            continue
        else:
            ryady.append((n + 1, ya))
    return zag, ryady


def proverit(put, tekst, o):
    if not ZAGOLOVOK.search(tekst):
        o.warn(
            put,
            "нет заголовка «# Открытые API: дайджест …»: файл не похож на дайджест, проверка пропущена",
        )
        return

    m = RAZDELITEL.search(tekst)
    if not m:
        o.err(
            put,
            "нет разделителя «---» между читательской и служебной частью (AGENTS.md, п. 3)",
        )
        return
    chitatel, sluzh = tekst[: m.start()], tekst[m.end() :]
    smeshchenie = tekst[: m.end()].count("\n")

    # Объём.
    if len(chitatel) > PREDEL:
        o.err(
            put,
            f"объём читательской части {len(chitatel)} знаков, предел {PREDEL} (AGENTS.md, п. 3)",
        )

    # Заполнители шаблона.
    for n, s in enumerate(tekst.splitlines(), 1):
        if re.search(r"<(?:заполнить|ДД\.ММ\.ГГГГ|период|N)\b[^>]*>|<…>", s):
            o.err(f"{put}:{n}", "остался заполнитель шаблона")

    # Читательская часть: без меток, ссылок и технических подробностей.
    bez_kommentariev = re.sub(
        r"<!--.*?-->", lambda x: "\n" * x.group(0).count("\n"), chitatel, flags=re.S
    )
    for n, s in enumerate(bez_kommentariev.splitlines(), 1):
        if METKA.search(s):
            o.err(
                f"{put}:{n}",
                "метка в читательской части: метки только в служебной части (AGENTS.md, п. 3)",
            )
        if SSYLKA.search(s):
            o.err(
                f"{put}:{n}",
                "ссылка на источник в читательской части: ссылки только в трассировке (AGENTS.md, п. 3)",
            )
        for t in TEHNICHESKIE.finditer(s):
            o.warn(
                f"{put}:{n}",
                f"возможна техническая подробность «{t.group(0)}» (AGENTS.md, п. 4.4)",
            )
        slov = len(re.findall(r"\w+", s))
        if s.lstrip().startswith(("- ", "* ")) and slov > 30:
            o.warn(
                f"{put}:{n}",
                f"строка из {slov} слов, рекомендуется до 25 (AGENTS.md, п. 2.7)",
            )

    # Главное за период: первый раздел второго уровня, 2-3 пункта.
    zag = [z for z in razdely(chitatel) if z[1] == 2]
    if not zag or not zag[0][2].startswith("Главное за период"):
        o.err(
            put,
            "первый раздел читательской части должен быть «Главное за период» (AGENTS.md, п. 6.6)",
        )
    else:
        punkty = [
            s
            for s in telo_razdela(chitatel, zag[0][0])
            if re.match(r"^\s*[-*]\s+\S", s)
        ]
        if not 2 <= len(punkty) <= 3:
            o.err(
                f"{put}:{zag[0][0]}",
                f"«Главное за период»: пунктов {len(punkty)}, нужно 2-3 (AGENTS.md, п. 6.6)",
            )
    punktov = len(
        [s for s in bez_kommentariev.splitlines() if re.match(r"^\s*[-*]\s+\S", s)]
    )

    # Служебная часть: обязательные разделы и порядок.
    sl = razdely(sluzh)
    imena = [z[2].lower() for z in sl]

    def nomer(nachalo):
        return next((i for i, x in enumerate(imena) if x.startswith(nachalo)), None)

    for nazv in (
        "параметры",
        "трассировка",
        "не вошло",
        "открытые вопросы",
        "источники",
        "проверка",
    ):
        if nomer(nazv) is None:
            o.err(put, f"в служебной части нет раздела «{nazv}» (AGENTS.md, п. 3)")
    i_ov, i_ist = nomer("открытые вопросы"), nomer("источники")
    if i_ov is not None and i_ist is not None and i_ov > i_ist:
        o.err(put, "«Открытые вопросы» должны стоять перед «Источники»")

    # Трассировка: у каждой строки метка [ФАКТ] и источник с датой.
    i_tr = nomer("трассировка")
    if i_tr is not None:
        zag_t, ryady = tablitsa_posle(sluzh, sl[i_tr][0])
        if not zag_t:
            o.err(put, "в разделе «Трассировка» нет таблицы")
        else:
            k_metka = next(
                (k for k, z in enumerate(zag_t) if z.lower().startswith("метка")), None
            )
            k_ist = next(
                (k for k, z in enumerate(zag_t) if z.lower().startswith("источник")),
                None,
            )
            k_data = next(
                (k for k, z in enumerate(zag_t) if "период" in z.lower()), None
            )
            zapolnennye = 0
            for n, ya in ryady:
                if not any(y for y in ya[1:]):
                    continue
                zapolnennye += 1
                mesto = f"{put}:{n + smeshchenie}"
                if k_metka is not None and (
                    k_metka >= len(ya) or "[ФАКТ]" not in ya[k_metka]
                ):
                    o.err(
                        mesto,
                        "в трассировке строка без метки [ФАКТ]: в читательскую часть входят только факты (AGENTS.md, п. 4.1)",
                    )
                if k_ist is not None and (
                    k_ist >= len(ya) or not re.search(DATA, ya[k_ist])
                ):
                    o.err(
                        mesto,
                        "у источника нет даты [документ, место, ДД.ММ.ГГГГ] (AGENTS.md, п. 2.1)",
                    )
                if (
                    k_data is not None
                    and k_data < len(ya)
                    and ya[k_data].lower().startswith("нет")
                ):
                    o.err(
                        mesto,
                        "событие вне периода в читательской части (AGENTS.md, п. 4.2)",
                    )
            if zapolnennye < punktov:
                o.warn(
                    put,
                    f"строк трассировки {zapolnennye}, пунктов в читательской части {punktov}: "
                    "у каждой строки дайджеста должна быть строка трассировки (AGENTS.md, п. 3)",
                )

    # Регуляторные документы: «актуально на» хотя бы в одном упоминании (AGENTS.md, п. 2.5).
    pervoe, proverennye = {}, set()
    for n, s in enumerate(tekst.splitlines(), 1):
        doki = {re.sub(r"\s+", " ", x.group(0)) for x in REGULYATORNYE.finditer(s)}
        if AKTUALNO.search(s):
            proverennye |= doki
        for d in doki:
            pervoe.setdefault(d, n)
    for d, n in sorted(pervoe.items(), key=lambda x: x[1]):
        if d not in proverennye:
            o.err(
                f"{put}:{n}",
                f"«{d}»: нет «актуально на ДД.ММ.ГГГГ» ни в одном упоминании (AGENTS.md, п. 2.5)",
            )


def fayly(argumenty):
    for a in argumenty:
        p = Path(a)
        if p.is_dir():
            yield from sorted(x for x in p.rglob("*.md") if x.name != "README.md")
        else:
            yield p


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    o = Otchet()
    proshlo = 0
    for p in fayly(argv):
        try:
            tekst = p.read_text(encoding="utf-8")
        except OSError as e:
            o.err(str(p), f"не удалось прочитать: {e}")
            continue
        proverit(str(p), tekst, o)
        proshlo += 1
    for s in o.preduprezhdeniya + o.oshibki:
        print(s)
    print(
        f"Проверено файлов: {proshlo}. Ошибок: {len(o.oshibki)}. Предупреждений: {len(o.preduprezhdeniya)}."
    )
    return 1 if o.oshibki else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
