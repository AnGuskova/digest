#!/usr/bin/env python3
"""Формальная проверка черновика дайджеста по AGENTS.md.

Проверяет структуру, объём, трассировку, метки и статусы, сверку действий
регулятора с его сайтом, технические подробности и номера версий, сроки
внутренних замечаний, проценты голосования, пометки «нет данных» и термины
по справочнику. Смысл (верность цитат и источников) скрипт не проверяет.

Запуск:
    python3 scripts/proverka.py <файл или папка> [...]

Код выхода 1, если найдены ошибки. Предупреждения выхода не меняют,
но каждое стоит посмотреть.
"""

import re
import sys
from pathlib import Path

KOREN = Path(__file__).resolve().parent.parent
PREDEL = 6000  # знаков с пробелами в читательской части (AGENTS.md, п. 3)
SLOV_PUNKT = 30  # предупреждение: пункт карточки длиннее (рекомендуется до 25)
SLOV_ITOG = 50  # предупреждение: текст ключевого итога длиннее
SLOV_ZAG = 12  # предупреждение: заголовок ключевого итога длиннее
MARKER_NET = "В предоставленных данных информации нет"

ZAGOLOVOK = re.compile(r"^#\s+Открытые API: дайджест", re.M)
RAZDELITEL = re.compile(r"^---\s*$", re.M)
DATA = r"\d{2}\.\d{2}\.\d{4}"
METKA = re.compile(r"\[(ФАКТ|ПРЕДЛОЖЕНИЕ|ПРЕДПОЛОЖЕНИЕ|ИСТОЧНИК НЕ НАЙДЕН)\]")
SSYLKA = re.compile(r"\[[^\[\]^]*\d{2}\.\d{2}\.\d{4}[^\[\]]*\]")
AKTUALNO = re.compile(r"актуально на\s+" + DATA, re.I)
STATUSY = (
    "решение", "утвержден", "утверждён", "опубликован", "согласован", "одобрен",
    "на согласовании", "проект", "на рассмотрении", "предложен", "пилот",
    "тестирование", "не согласован", "разъяснение", "событие", "направлен",
)

# Регуляторные документы, для которых нужна дата актуальности (AGENTS.md, п. 2.5).
REGULYATORNYE = re.compile(
    r"СТО БР [А-ЯA-Z0-9][А-ЯA-Z.\-0-9]*[А-ЯA-Z0-9]"
    r"|ГОСТ(?: Р)? \d[\d.\-]*\d"
    r"|ISO(?:/IEC)? \d+(?:-\d+)?"
    r"|RFC \d+"
    r"|(?:Указани[еяю]|Положени[еяю]) Банка России(?: от " + DATA + r")?(?: № ?[\d\-А-Яа-я]+)?"
    r"|Федеральн\w+ закон\w*(?: от " + DATA + r")?(?: № ?[\d\-А-Яа-я]+)?"
)

# Действия регулятора: проверяются по официальному сайту (AGENTS.md, п. 2.5).
DEYSTVIE_REGULYATORA = re.compile(
    r"(?:Банк\w*\s+России|Минцифры(?:\s+России)?|Минфин\w*)\s+"
    r"(?:утвердил\w*|опубликовал\w*|рекомендовал\w*|ввел\w*|ввёл\w*|издал\w*|принял\w*)"
    r"|(?:утвержден\w*|утверждён\w*|опубликован\w*|рекомендован\w*)\s+Банком\s+России",
    re.I,
)
OFITSIALNYE = re.compile(r"cbr\.ru|pravo\.gov\.ru|digital\.gov\.ru|minfin\.gov\.ru|\.gov\.ru", re.I)

# Технические подробности (AGENTS.md, п. 4.6).
TEHNICHESKIE = re.compile(
    r"\b(?:mTLS|TLS|SSL|PKCE|OAuth|OIDC|OpenID|JWS|JWE|JWT|JWKS|DPoP|FAPI|HTTP|JSON|XML|REST|endpoint|scope)\b"
    r"|\bX-[A-Za-z][\w-]+"
    r"|(?<![\w.])/[a-z][a-z0-9_-]+(?:/[\w-]+)*\b"
    r"|эндпоинт|скоуп|код(?:ы|ов)? ошиб|`[^`]+`",
    re.I,
)
VERSIYA = re.compile(r"\b[vV]\d+(?:\.\d+)+\b|\bверси\w*\s+\d+(?:\.\d+)+", re.I)
NACHALO_S_VERSII = re.compile(r"^\s*[-*]\s+(?:\*\*[^*]+\*\*\s*)?(?:[Вв]ерси\w*\s+)?[vV]?\d+\.\d+(?!\d|\.\d{4})")

# Проценты голосования (AGENTS.md, п. 4.4).
PROTSENTY = re.compile(
    r"\d+(?:[.,]\d+)?\s?%[^.;]*(?:голос|«за»|порог)|(?:голос|«за»|порог)[^.;]*\d+(?:[.,]\d+)?\s?%",
    re.I,
)
# Сроки внутренних замечаний (AGENTS.md, п. 4.5).
SROK_ZAMECHANIY = re.compile(
    r"(?:замечани\w*|позици\w*|отзыв\w*|предложени\w*)[^.;]{0,40}?(?:до|не позднее)\s+" + DATA
    + r"|принима\w+\s+до\s+" + DATA,
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


# ---------------------------------------------------------------- справочные файлы


def yacheyki(s):
    s = s.strip()
    if not (s.startswith("|") and s.endswith("|")):
        return None
    return [y.strip() for y in s.strip("|").split("|")]


def tablitsy(put):
    """Все строки таблиц файла: [ячейки], без заголовков и разделителей."""
    if not put.exists():
        return []
    rez, zag = [], False
    for s in put.read_text(encoding="utf-8").splitlines():
        ya = yacheyki(s)
        if ya is None:
            zag = False
            continue
        if not zag:
            zag = True
            continue
        if all(re.fullmatch(r":?-{3,}:?", y) for y in ya if y):
            continue
        rez.append(ya)
    return rez


def spravochnik():
    """[(канонические формы, [варианты «не пишем»])]."""
    rez = []
    for ya in tablitsy(KOREN / "nastroyki" / "spravochnik-terminov.md"):
        if len(ya) < 2 or not ya[1]:
            continue
        pishem = re.sub(r"\([^)]*\)", "", ya[0])
        kanon = [x.strip() for x in re.split(r",", pishem) if x.strip()]
        kanon += [x.strip() for x in re.findall(r"\(([^)]*)\)", ya[0])]
        kanon.append(ya[0].strip())
        ne = [x.strip() for x in ya[1].split(";") if x.strip()]
        rez.append((kanon, ne))
    return rez


def nazvaniya_kartochek():
    """{номер карточки: название для дайджеста} из rabochie-gruppy.md."""
    for put in (KOREN / "workspace" / "rabochie-gruppy.md", KOREN / "nastroyki" / "rabochie-gruppy.md"):
        rez = {}
        for ya in tablitsy(put):
            if len(ya) >= 3 and ya[0].isdigit() and not ya[2].startswith("<"):
                rez.setdefault(int(ya[0]), ya[2])
        if rez:
            return rez
    return {}


# ---------------------------------------------------------------- разбор черновика


def razdely(tekst):
    """Заголовки любого уровня: [(номер строки, уровень, текст)]."""
    rez, v_kode = [], False
    for n, s in enumerate(tekst.splitlines(), 1):
        if s.strip().startswith("```"):
            v_kode = not v_kode
            continue
        m = re.match(r"^(#{1,6})\s+(.*?)\s*$", s)
        if m and not v_kode:
            rez.append((n, len(m.group(1)), m.group(2)))
    return rez


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


def telo_razdela(tekst, nachalo):
    rez = []
    for s in tekst.splitlines()[nachalo:]:
        if re.match(r"^#{1,6}\s", s):
            break
        rez.append(s)
    return rez


def norm(s):
    s = re.sub(r"\*\*|`|\[\^\w+\]", "", s)
    s = re.sub(r"^\s*[-*]\s+", "", s)
    return re.sub(r"\s+", " ", s).strip(" .…«»\"").lower()


def slov(s):
    return len(re.findall(r"\w+", s))


# ---------------------------------------------------------------- проверки


def proverit(put, tekst, o):
    if not ZAGOLOVOK.search(tekst):
        o.warn(put, "нет заголовка «# Открытые API: дайджест …»: файл не похож на дайджест, проверка пропущена")
        return

    m = RAZDELITEL.search(tekst)
    if not m:
        o.err(put, "нет разделителя «---» между читательской и служебной частью (AGENTS.md, п. 3)")
        return
    chitatel, sluzh = tekst[: m.start()], tekst[m.end():]
    smeshchenie = tekst[: m.end()].count("\n")
    chist = re.sub(r"<!--.*?-->", lambda x: "\n" * x.group(0).count("\n"), chitatel, flags=re.S)
    stroki = chist.splitlines()

    # Объём.
    if len(chitatel) > PREDEL:
        o.err(put, f"объём читательской части {len(chitatel)} знаков, предел {PREDEL} (AGENTS.md, п. 3)")

    # Заполнители шаблона.
    for n, s in enumerate(stroki, 1):
        if re.search(r"<(?!!--)[^<>\n]+>", s):
            o.err(f"{put}:{n}", "остался заполнитель шаблона <…>")
    for n, s in enumerate(sluzh.splitlines(), smeshchenie + 1):
        if re.search(r"<(?:заполнить|ДД\.ММ\.ГГГГ|период|N)\b[^>]*>|<…>", s):
            o.err(f"{put}:{n}", "остался заполнитель шаблона")

    # Построчные проверки читательской части.
    for n, s in enumerate(stroki, 1):
        snoska = bool(re.match(r"^\[\^\w+\]:", s))
        if METKA.search(s):
            o.err(f"{put}:{n}", "метка в читательской части: метки только в служебной части (AGENTS.md, п. 3)")
        if SSYLKA.search(s):
            o.err(f"{put}:{n}", "ссылка на источник в читательской части: ссылки только в трассировке (AGENTS.md, п. 3)")
        if PROTSENTY.search(s):
            o.err(f"{put}:{n}", "проценты голосования не указываются: пишется «не согласованы» (AGENTS.md, п. 4.4)")
        if snoska:
            continue
        for t in TEHNICHESKIE.finditer(s):
            o.warn(f"{put}:{n}", f"техническая подробность «{t.group(0)}»: замените смыслом для читателя (AGENTS.md, п. 4.6)")
        for t in VERSIYA.finditer(s):
            o.warn(f"{put}:{n}", f"номер версии «{t.group(0)}» в тексте: перенесите в сноску или напишите, что изменилось (AGENTS.md, п. 4.6)")
        if NACHALO_S_VERSII.match(s):
            o.warn(f"{put}:{n}", "строка начинается с номера версии (AGENTS.md, раздел 5, проверка 1)")
        for t in SROK_ZAMECHANIY.finditer(s):
            # Исключение п. 4.5: срок объясняет, когда документ уйдёт регулятору.
            if re.search(r"направ\w+\s+(?:в|на)\s+(?:Банк\w* России|Минцифры|регулятор)", s[t.end():]):
                continue
            o.warn(f"{put}:{n}", f"срок внутренних замечаний «{t.group(0)}»: нужен, только если объясняет, когда документ уйдёт регулятору (AGENTS.md, п. 4.5)")

    # Справочник терминов.
    for n, s in enumerate(stroki, 1):
        if re.match(r"^\[\^\w+\]:", s):
            continue
        for kanon, ne in spravochnik():
            ostatok = s
            for k in sorted(kanon, key=len, reverse=True):
                ostatok = ostatok.replace(k, " ")
            for v in ne:
                if re.search(r"(?<![\w-])" + re.escape(v) + r"(?![\w-])", ostatok):
                    o.warn(f"{put}:{n}", f"«{v}»: по справочнику пишется «{kanon[0]}» (nastroyki/spravochnik-terminov.md)")

    # Ключевые итоги.
    zag2 = [z for z in razdely(chist) if z[1] == 2]
    imena2 = [z[2].lower() for z in zag2]
    if not zag2 or not imena2[0].startswith("ключевые итоги"):
        o.err(put, "первый раздел читательской части должен быть «## Ключевые итоги» (templates/daydzhest.md)")
    else:
        punkty = [s for s in telo_razdela(chist, zag2[0][0]) if re.match(r"^\s*[-*]\s+\S", s)]
        if not 2 <= len(punkty) <= 3:
            o.err(f"{put}:{zag2[0][0]}", f"«Ключевые итоги»: пунктов {len(punkty)}, нужно 2-3 (AGENTS.md, п. 3)")
        for s in punkty:
            mz = re.match(r"^\s*[-*]\s+\*\*([^*]+)\*\*\s*(.*)$", s)
            if not mz:
                o.err(put, f"ключевой итог без заголовка «- **Заголовок.** Текст»: «{s.strip()[:50]}…»")
                continue
            if slov(mz.group(1)) > SLOV_ZAG:
                o.warn(put, f"заголовок ключевого итога длиннее {SLOV_ZAG} слов: «{mz.group(1)[:50]}…»")
            if slov(mz.group(2)) > SLOV_ITOG:
                o.warn(put, f"текст ключевого итога длиннее {SLOV_ITOG} слов: «{mz.group(2)[:50]}…»")

    # Основные направления.
    i_napr = next((i for i, x in enumerate(imena2) if x.startswith("основные направления")), None)
    punktov = 0
    if i_napr is None:
        o.err(put, "нет раздела «## Основные направления» (templates/daydzhest.md)")
    else:
        nachalo = zag2[i_napr][0]
        konets = zag2[i_napr + 1][0] - 1 if i_napr + 1 < len(zag2) else len(stroki)
        nazv = nazvaniya_kartochek()
        kartochki, tek = [], None
        for n in range(nachalo, konets):
            s = stroki[n]
            z3 = re.match(r"^###\s+(.*?)\s*$", s)
            if z3:
                mn = re.match(r"^(\d+)\.\s*(.+)$", z3.group(1))
                if not mn:
                    o.err(f"{put}:{n + 1}", "заголовок карточки должен быть «### N. Название»")
                    tek = None
                    continue
                tek = {"n": int(mn.group(1)), "zag": mn.group(2), "stroka": n + 1, "punkty": [], "uch": False}
                kartochki.append(tek)
                continue
            if tek is None:
                continue
            if re.match(r"^\s*\**Участники", s):
                tek["uch"] = True
            elif re.match(r"^\s*[-*]\s+\S", s):
                tek["punkty"].append((n + 1, s))
        nomera = [k["n"] for k in kartochki]
        if not kartochki:
            o.err(put, "в «Основные направления» нет карточек «### N. Название»")
        if len(set(nomera)) != len(nomera):
            o.err(put, f"номера карточек повторяются: {nomera}")
        if nomera != sorted(nomera):
            o.warn(put, f"карточки не по порядку номеров: {nomera}")
        for k in kartochki:
            mesto = f"{put}:{k['stroka']}"
            if not 1 <= k["n"] <= 8:
                o.err(mesto, f"номер карточки {k['n']}: в шаблоне места 1-8")
            if not 1 <= len(k["punkty"]) <= 3:
                o.err(mesto, f"в карточке {len(k['punkty'])} пунктов, нужно 1-3 (AGENTS.md, п. 3)")
            if not k["uch"]:
                o.warn(mesto, "в карточке нет строки «Участники: …» (AGENTS.md, п. 4.11)")
            if nazv.get(k["n"]) and norm(nazv[k["n"]]) != norm(k["zag"]):
                o.warn(mesto, f"название карточки отличается от nastroyki/rabochie-gruppy.md: «{nazv[k['n']]}»")
            for n, s in k["punkty"]:
                if slov(s) > SLOV_PUNKT:
                    o.warn(f"{put}:{n}", f"пункт из {slov(s)} слов, рекомендуется до 25 (AGENTS.md, п. 3)")
        punktov = sum(len(k["punkty"]) for k in kartochki)
        nenaydeno = sorted(set(range(1, 9)) - set(nomera))
        if nenaydeno:
            o.warn(put, f"нет карточек {nenaydeno}: у каждой активной РГ своя карточка; причина должна быть в «Открытых вопросах» (AGENTS.md, п. 7, шаг 7)")

    # Сноски.
    ssylki = set(re.findall(r"\[\^(\w+)\](?!:)", chist))
    snoski = set(re.findall(r"^\[\^(\w+)\]:", chist, re.M))
    for s_ in ssylki - snoski:
        o.err(put, f"ссылка [^{s_}] без текста сноски «[^{s_}]: …»")
    for s_ in snoski - ssylki:
        o.warn(put, f"сноска [^{s_}] без ссылки в тексте")
    if len(snoski) > 1:
        o.err(put, "в шаблоне место для одной сноски")

    # Служебная часть: обязательные разделы.
    sl = razdely(sluzh)
    imena = [z[2].lower() for z in sl]

    def nomer(nachalo):
        return next((i for i, x in enumerate(imena) if x.startswith(nachalo)), None)

    for nazv_r in ("параметры", "трассировка", "сравнение с прошлым", "не вошло", "открытые вопросы", "источники", "проверка"):
        if nomer(nazv_r) is None:
            o.err(put, f"в служебной части нет раздела «{nazv_r}» (templates/daydzhest.md)")
    if nomer("карта направлений") is None:
        o.warn(put, "в служебной части нет «Карты направлений»")
    i_ov, i_ist = nomer("открытые вопросы"), nomer("источники")
    if i_ov is not None and i_ist is not None and i_ov > i_ist:
        o.err(put, "«Открытые вопросы» должны стоять перед «Источники»")
    if "понятно менеджеру" not in sluzh.lower():
        o.warn(put, "в «Проверке» нет 9 проверок «Понятно менеджеру» (AGENTS.md, раздел 5)")

    # Пометка «нет данных» -> открытый вопрос (AGENTS.md, п. 2.9).
    if MARKER_NET in chist:
        if i_ov is None or MARKER_NET not in "\n".join(telo_razdela(sluzh, sl[i_ov][0])):
            o.err(put, f"в тексте есть «{MARKER_NET}», а в «Открытых вопросах» нет пункта о ней (AGENTS.md, п. 2.9)")
        else:
            o.warn(put, f"в тексте {chist.count(MARKER_NET)} пометок «{MARKER_NET}»: перед отправкой замените данными или уберите строку")

    # Трассировка.
    i_tr = nomer("трассировка")
    trass = []
    if i_tr is not None:
        zag_t, ryady = tablitsa_posle(sluzh, sl[i_tr][0])
        if not zag_t:
            o.err(put, "в разделе «Трассировка» нет таблицы")
        else:
            def kol(nachalo):
                return next((k for k, z in enumerate(zag_t) if z.lower().startswith(nachalo)), None)

            k_str, k_metka, k_status, k_ist = kol("строка"), kol("метка"), kol("статус"), kol("источник")
            k_data = next((k for k, z in enumerate(zag_t) if "период" in z.lower()), None)
            if k_status is None:
                o.warn(put, "в трассировке нет колонки «Статус» (AGENTS.md, п. 2.3)")
            for n, ya in ryady:
                if not any(y for y in ya[1:]):
                    continue
                mesto = f"{put}:{n + smeshchenie}"
                yach = lambda k: ya[k] if k is not None and k < len(ya) else ""  # noqa: E731
                trass.append((norm(yach(k_str)), yach(k_ist)))
                if "[ФАКТ]" not in yach(k_metka):
                    o.err(mesto, "в трассировке строка без метки [ФАКТ] (AGENTS.md, п. 2.2)")
                if not re.search(DATA, yach(k_ist)):
                    o.err(mesto, "у источника нет даты [документ, место, ДД.ММ.ГГГГ] (AGENTS.md, п. 2.1)")
                if yach(k_data).lower().startswith("нет"):
                    o.err(mesto, "событие вне периода в читательской части (AGENTS.md, п. 4.2)")
                st = yach(k_status).lower()
                if k_status is not None and not any(x in st for x in STATUSY):
                    o.warn(mesto, f"статус «{yach(k_status)}» не из перечня AGENTS.md, п. 2.3")
            if len(trass) < punktov:
                o.warn(put, f"строк трассировки {len(trass)}, пунктов в карточках {punktov}: у каждой строки дайджеста должна быть строка трассировки")

    # Действия регулятора сверяются с официальным сайтом (AGENTS.md, п. 2.5).
    for n, s in enumerate(stroki, 1):
        if not DEYSTVIE_REGULYATORA.search(s):
            continue
        ns = norm(s)
        nashli = [ist for nachalo, ist in trass if nachalo and len(nachalo) >= 10 and nachalo[:40] in ns]
        if not nashli:
            o.warn(f"{put}:{n}", "действие регулятора: не нашёл строку трассировки; источник должен быть с официального сайта (AGENTS.md, п. 2.5)")
        elif not any(OFITSIALNYE.search(ist) for ist in nashli):
            o.err(f"{put}:{n}", "действие регулятора без источника с официального сайта (cbr.ru и др.) в трассировке (AGENTS.md, п. 2.5)")

    # Регуляторные документы: «актуально на» хотя бы в одном упоминании.
    pervoe, proverennye = {}, set()
    for n, s in enumerate(tekst.splitlines(), 1):
        doki = {re.sub(r"\s+", " ", x.group(0)) for x in REGULYATORNYE.finditer(s)}
        if AKTUALNO.search(s):
            proverennye |= doki
        for d in doki:
            pervoe.setdefault(d, n)
    for d, n in sorted(pervoe.items(), key=lambda x: x[1]):
        if d not in proverennye:
            o.err(f"{put}:{n}", f"«{d}»: нет «актуально на ДД.ММ.ГГГГ» ни в одном упоминании (AGENTS.md, п. 2.5)")


def fayly(argumenty):
    for a in argumenty:
        p = Path(a)
        if p.is_dir():
            yield from sorted(
                x for x in p.rglob("*.md")
                if x.name != "README.md" and not x.name.startswith(("kontakty", "rabochie-gruppy"))
            )
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
    print(f"Проверено файлов: {proshlo}. Ошибок: {len(o.oshibki)}. Предупреждений: {len(o.preduprezhdeniya)}.")
    return 1 if o.oshibki else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
