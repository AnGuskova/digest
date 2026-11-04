#!/usr/bin/env python3
"""Сборка итогового дайджеста PPTX по шаблону templates/daydzhest.pptx.

На входе черновик: Markdown с читательской и служебной частью (AGENTS.md, п. 3).
На выходе итоговый файл на 2 страницах: только читательская часть, без меток,
ссылок и служебной части, оформленная по шаблону АФТ.

Что куда попадает:
    «# Открытые API: дайджест за <период>»  -> плашка периода, строка над карточками стр. 2
    «## Ключевые итоги», 1-3 пункта
        «- **Заголовок.** Текст»            -> блок «Ключевые итоги» (стр. 1)
    «## Основные направления»
        «### N. Название»                   -> карточка N (1-4 на стр. 1, 5-8 на стр. 2)
        «- пункт» (1-3 пункта)              -> пункты карточки
        «Участники: …»                      -> строка «Участники:» внизу карточки
    «[^1]» в тексте и «[^1]: …»             -> надстрочный номер и сноска внизу стр. 1
    «В предоставленных данных информации нет» -> выделяется жёлтым
    workspace/kontakty.md или nastroyki/kontakty.md -> таблица «Как подключиться» (стр. 2)

Номер карточки N совпадает с номером места в шаблоне: у каждого места свой
значок (nastroyki/rabochie-gruppy.md, колонка «№»). Отсутствующая карточка
убирается, остальные на той же странице сдвигаются.

Высота блоков считается по длине текста. Если текст не помещается, скрипт
уменьшает шрифт карточек с 9,5 до 9 pt, а если и это не помогает, завершается
с ошибкой: дайджест сокращается отбором (AGENTS.md, п. 7, шаг 7).

Запуск:
    python3 scripts/sborka.py "workspace/outputs/Дайджест Открытые API - сентябрь-октябрь 2026_черновик.md"
        [--vyhod файл.pptx] [--shablon templates/daydzhest.pptx]
        [--kontakty файл.md] [--bez-proverki]

Перед сборкой запускается scripts/proverka.py; при ошибках сборки нет.
Нужна библиотека python-pptx: pip install python-pptx
"""

import argparse
import copy
import datetime
import re
import subprocess
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
except ImportError:  # pragma: no cover
    sys.exit("Нужна библиотека python-pptx: pip install python-pptx")

KOREN = Path(__file__).resolve().parent.parent
SHABLON = KOREN / "templates" / "daydzhest.pptx"
MARKER_NET = "В предоставленных данных информации нет"
E = 914400  # EMU в дюйме
PT = 12700  # EMU в пункте

# ---------------------------------------------------------------- разбор черновика


def bez_razmetki(s):
    s = re.sub(r"`([^`]*)`", r"\1", s)
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"(?<!\w)[*_](.+?)[*_](?!\w)", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


def chitatelskaya_chast(tekst):
    m = re.search(r"^---\s*$", tekst, re.M)
    if not m:
        sys.exit("Нет разделителя «---»: это не черновик дайджеста (AGENTS.md, п. 3)")
    return re.sub(r"<!--.*?-->", "", tekst[: m.start()], flags=re.S), tekst[m.end():]


def razobrat(tekst):
    """Читательская часть черновика -> словарь для шаблона."""
    chitatel, sluzh = chitatelskaya_chast(tekst)

    m = re.search(r"^#\s+Открытые API: дайджест за\s+(.+?)\s*$", chitatel, re.M)
    if not m:
        sys.exit("Нет заголовка «# Открытые API: дайджест за …»")
    period = bez_razmetki(m.group(1))
    m = re.search(r"Данные на\s+(\d{2}\.\d{2}\.\d{4})", chitatel)
    data_sreza = m.group(1) if m else None
    ne_zavershen = bool(re.search(r"Период завершён:\s*нет", sluzh, re.I))

    snoski = {}
    for m in re.finditer(r"^\[\^(\w+)\]:\s*(.+?)\s*$", chitatel, re.M):
        snoski[m.group(1)] = bez_razmetki(m.group(2))

    itogi, kartochki = [], []
    razdel, tek = None, None
    for s in chitatel.splitlines():
        if re.match(r"^\[\^\w+\]:", s):
            continue
        z2 = re.match(r"^##\s+(.*?)\s*$", s)
        if z2:
            razdel = bez_razmetki(z2.group(1)).lower()
            tek = None
            continue
        if razdel is None:
            continue
        p = re.match(r"^\s*[-*]\s+(.*\S)\s*$", s)
        if razdel.startswith("ключевые итоги"):
            if p:
                mz = re.match(r"^\*\*(.+?)\*\*\s*(.*)$", p.group(1))
                if mz:
                    zag = bez_razmetki(mz.group(1))
                    zag = zag[:-1] if zag.endswith(".") else zag
                    itogi.append((zag, bez_razmetki(mz.group(2))))
                else:
                    itogi.append(("", bez_razmetki(p.group(1))))
            continue
        if razdel.startswith("основные направления"):
            z3 = re.match(r"^###\s+(\d+)\.\s*(.+?)\s*$", s)
            if z3:
                tek = {"n": int(z3.group(1)), "zagolovok": bez_razmetki(z3.group(2)),
                       "punkty": [], "uchastniki": ""}
                kartochki.append(tek)
                continue
            if tek is None:
                continue
            mu = re.match(r"^\s*\**Участники:?\**:?\s*(.*\S)\s*$", s)
            if mu:
                tek["uchastniki"] = bez_razmetki(mu.group(1))
            elif p:
                tek["punkty"].append(bez_razmetki(p.group(1)))

    oshibki = []
    if not 1 <= len(itogi) <= 3:
        oshibki.append(f"в «Ключевые итоги» {len(itogi)} пунктов, в шаблоне места для 1-3")
    nomera = [k["n"] for k in kartochki]
    if len(set(nomera)) != len(nomera):
        oshibki.append(f"номера карточек повторяются: {nomera}")
    for k in kartochki:
        if not 1 <= k["n"] <= 8:
            oshibki.append(f"карточка «{k['zagolovok']}»: номер {k['n']}, в шаблоне места 1-8")
        if not 1 <= len(k["punkty"]) <= 3:
            oshibki.append(f"карточка {k['n']} «{k['zagolovok']}»: пунктов {len(k['punkty'])}, нужно 1-3")
    if not kartochki:
        oshibki.append("в «Основные направления» нет карточек «### N. Название»")
    ssylki = set(re.findall(r"\[\^(\w+)\](?!:)", chitatel))
    for s_ in ssylki - set(snoski):
        oshibki.append(f"ссылка [^{s_}] без текста сноски «[^{s_}]: …»")
    if len(snoski) > 1:
        oshibki.append("в шаблоне место для одной сноски")
    if oshibki:
        sys.exit("Черновик не подходит для шаблона:\n- " + "\n- ".join(oshibki))
    return {
        "period": period,
        "data_sreza": data_sreza if ne_zavershen else None,
        "itogi": itogi,
        "kartochki": sorted(kartochki, key=lambda k: k["n"]),
        "snoski": snoski,
    }


def kontakty(put):
    """Таблица контактов: | Рабочие группы (через ;) | ФИО | Почта |."""
    if not put or not put.exists():
        return []
    rez = []
    for s in put.read_text(encoding="utf-8").splitlines():
        s = s.strip()
        if not (s.startswith("|") and s.endswith("|")):
            continue
        ya = [y.strip() for y in s.strip("|").split("|")]
        if len(ya) < 3 or all(re.fullmatch(r":?-{3,}:?", y) for y in ya if y):
            continue
        if ya[0].lower().startswith("рабоч") or "<заполнить>" in s or not ya[1]:
            continue
        rg = [bez_razmetki(x) for x in ya[0].split(";") if x.strip()]
        rez.append((rg, bez_razmetki(ya[1]), bez_razmetki(ya[2])))
    return rez


# ---------------------------------------------------------------- измерение текста


class Izmeritel:
    """Ширина строки в пунктах по метрике Arial (или Liberation Sans)."""

    KANDIDATY = {
        False: [
            "C:/Windows/Fonts/arial.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ],
        True: [
            "C:/Windows/Fonts/arialbd.ttf",
            "/Library/Fonts/Arial Bold.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ],
    }

    def __init__(self):
        self.shrifty = {}
        try:
            from PIL import ImageFont

            for zhirnyy, spisok in self.KANDIDATY.items():
                for f in spisok:
                    if Path(f).exists():
                        self.shrifty[zhirnyy] = ImageFont.truetype(f, 100)
                        break
        except ImportError:
            pass

    def shirina(self, s, pt, zhirnyy=False):
        f = self.shrifty.get(zhirnyy)
        if f is None:  # грубая оценка: средний знак 0,56 кегля
            return len(s) * pt * (0.6 if zhirnyy else 0.56)
        return f.getlength(s) * pt / 100

    def strok(self, tekst, pt, shirina_dyuym, zhirnyy=False, zapas=0.96):
        """Число строк при переносе по словам, с запасом на разницу программ."""
        if not tekst:
            return 0
        predel = shirina_dyuym * 72 * zapas
        n, tek = 1, ""
        for slovo in tekst.split():
            proba = f"{tek} {slovo}".strip()
            if self.shirina(proba, pt, zhirnyy) <= predel or not tek:
                tek = proba
            else:
                n += 1
                tek = slovo
        return n


def vys(strok, pt, interval=1.0):
    """Высота текста в дюймах."""
    return strok * pt * 1.2 * interval / 72


# ---------------------------------------------------------------- работа с XML


def figury(slayd):
    return {sh.name: sh for sh in slayd.shapes}


def postavit(sh, x=None, y=None, w=None, h=None):
    """Положение и размер в дюймах."""
    if x is not None:
        sh.left = int(x * E)
    if y is not None:
        sh.top = int(y * E)
    if w is not None:
        sh.width = int(w * E)
    if h is not None:
        sh.height = int(h * E)


def udalit(sh):
    sh._element.getparent().remove(sh._element)


def telo(el):
    t = el.find(".//" + qn("p:txBody"))
    return t if t is not None else el.find(".//" + qn("a:txBody"))


def segmenty(tekst):
    """Текст -> [(кусок, вид)]: вид '' обычный, 'sup' номер сноски, 'hl' пометка."""
    rez = []
    for chast in re.split(r"(\[\^\w+\]|" + re.escape(MARKER_NET) + ")", tekst):
        if not chast:
            continue
        m = re.fullmatch(r"\[\^(\w+)\]", chast)
        if m:
            rez.append((m.group(1), "sup"))
        elif chast == MARKER_NET:
            rez.append((chast, "hl"))
        else:
            rez.append((chast, ""))
    return rez


def progon(obrazets_rPr, tekst, vid="", razmer=None):
    r = copy.deepcopy(obrazets_rPr.getparent())
    for lishnee in r.findall(qn("a:t")):
        r.remove(lishnee)
    rPr = r.find(qn("a:rPr"))
    for a in ("err", "dirty", "baseline"):
        rPr.attrib.pop(a, None)
    if razmer:
        rPr.set("sz", str(int(razmer * 100)))
    if vid == "sup":
        rPr.set("baseline", "30000")
    for h in rPr.findall(qn("a:highlight")):
        rPr.remove(h)
    if vid == "hl":
        h = rPr.makeelement(qn("a:highlight"), {})
        c = h.makeelement(qn("a:srgbClr"), {"val": "FFFF00"})
        h.append(c)
        fill = rPr.find(qn("a:solidFill"))
        if fill is not None:
            fill.addnext(h)
        else:
            rPr.insert(0, h)
    t = r.makeelement(qn("a:t"), {})
    t.text = tekst
    r.append(t)
    return r


def zapolnit_abzats(p, tekst, razmer=None, obrazets_rPr=None):
    """Абзац p: все прогоны заменяются текстом с учётом сносок и пометок."""
    rs = p.findall(qn("a:r"))
    rPr = obrazets_rPr if obrazets_rPr is not None else rs[0].find(qn("a:rPr"))
    rPr = copy.deepcopy(rPr)
    hold = p.makeelement(qn("a:r"), {})
    hold.append(rPr)
    for r in rs:
        p.remove(r)
    end = p.find(qn("a:endParaRPr"))
    for kusok, vid in segmenty(tekst):
        r = progon(rPr, kusok, vid, razmer)
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)


def tekst_v_figuru(sh, tekst, razmer=None):
    tb = telo(sh._element)
    ps = tb.findall(qn("a:p"))
    for p in ps[1:]:
        tb.remove(p)
    zapolnit_abzats(ps[0], tekst, razmer)


def punkty_v_figuru(el, punkty, razmer=None):
    tb = telo(el)
    ps = tb.findall(qn("a:p"))
    obrazets = copy.deepcopy(ps[0])
    for p in ps:
        tb.remove(p)
    for t in punkty:
        p = copy.deepcopy(obrazets)
        zapolnit_abzats(p, t, razmer)
        tb.append(p)


def uchastniki_v_figuru(sh, spisok):
    p = telo(sh._element).find(qn("a:p"))
    rs = p.findall(qn("a:r"))
    rs[0].find(qn("a:t")).text = "Участники: "
    for r in rs[2:]:
        p.remove(r)
    rs[1].find(qn("a:t")).text = spisok


def zamenit_period(slayd, period):
    for sh in slayd.shapes:
        if not sh.has_text_frame:
            continue
        for a in sh.text_frame.paragraphs:
            for r in a.runs:
                if "ПЕРИОД" in r.text:
                    r.text = r.text.replace("ПЕРИОД", period.upper())


# ---------------------------------------------------------------- геометрия шаблона (дюймы)

SHIRINA_ITOGA = 6.45  # summary-title/text
KARTOCHKA_W = 3.54
KOLONKI_X = (0.50, 4.22)
TEKST_W = 3.22 - 0.14  # wg-text минус отступ маркера
ZAG_W = 2.38  # wg-title
ZAZOR_RYADOV = 0.10
PODVAL_1 = 11.12  # низ места на стр. 1 без сноски
SNOSKA_Y = 11.18
PODVAL_2 = 11.10  # низ места на стр. 2
MAKS_DOBAVKA = 0.6  # на сколько можно растянуть ряд карточек, чтобы заполнить страницу


class Mesto:
    """Нужная высота карточки (дюймы) и её части."""

    def __init__(self, k, izm, kegl):
        stroki_zag = izm.strok(k["zagolovok"], 11, ZAG_W, True)
        self.h_zag = vys(stroki_zag, 11)
        self.h_tekst = sum(vys(izm.strok(re.sub(r"\[\^\w+\]", "1", p), kegl, TEKST_W), kegl) for p in k["punkty"])
        self.h_uch = vys(max(1, izm.strok("Участники: " + k["uchastniki"], 8, 3.22)), 8) if k["uchastniki"] else 0
        self.nuzhno = 0.60 + self.h_tekst + 0.12 + self.h_uch + 0.08
        self.stroki_zag = stroki_zag


def kartochka_na_mesto(f, k, mesto, x, y, h, kegl):
    n = k["n"]
    postavit(f[f"wg-card-{n}"], x=x, y=y, h=h)
    postavit(f[f"wg-badge-{n}"], x=x + 0.16, y=y + 0.16)
    postavit(f[f"wg-icon-{n}"], x=x + 3.09, y=y + 0.18)
    zag = f[f"wg-title-{n}"]
    tekst_v_figuru(zag, k["zagolovok"])
    postavit(zag, x=x + 0.61, y=y + 0.33 - mesto.h_zag / 2, h=mesto.h_zag)
    tekst = f[f"wg-text-{n}"]
    punkty_v_figuru(tekst._element, k["punkty"], kegl)
    postavit(tekst, x=x + 0.16, y=y + 0.60, h=mesto.h_tekst + 0.04)
    uch = f[f"wg-members-{n}"]
    if k["uchastniki"]:
        uchastniki_v_figuru(uch, k["uchastniki"])
        postavit(uch, x=x + 0.16, y=y + h - 0.08 - mesto.h_uch, h=mesto.h_uch)
    else:
        udalit(uch)


def ubrat_kartochku(f, n):
    for chast in ("card", "badge", "icon", "title", "text", "members"):
        imya = f"wg-{chast}-{n}"
        if imya in f:
            udalit(f[imya])


def rasstavit_kartochki(f, spisok, verh, niz, izm, kegl):
    """Карточки двумя колонками от verh; свободное место делится между рядами
    (не больше MAKS_DOBAVKA на ряд). Возвращает (нехватка места, низ карточек) в дюймах."""
    mesta = [Mesto(k, izm, kegl) for k in spisok]
    ryady = [list(range(i, min(i + 2, len(spisok)))) for i in range(0, len(spisok), 2)]
    vysoty = [max(mesta[i].nuzhno for i in r) for r in ryady]
    est = niz - verh - ZAZOR_RYADOV * (len(ryady) - 1)
    lishnee = est - sum(vysoty)
    if lishnee < 0:
        return -lishnee, verh
    dobavka = min(MAKS_DOBAVKA, lishnee / len(ryady)) if ryady else 0
    y = verh
    for r, h in zip(ryady, vysoty):
        h += dobavka
        for j, i in enumerate(r):
            kartochka_na_mesto(f, spisok[i], mesta[i], KOLONKI_X[j], y, h, kegl)
        y += h + ZAZOR_RYADOV
    return 0.0, y - ZAZOR_RYADOV


def stranitsa_1(slayd, d, izm, kegl):
    f = figury(slayd)
    # Ключевые итоги.
    y = 1.70 + 0.14
    for i in range(1, 4):
        if i > len(d["itogi"]):
            for imya in (f"summary-title-{i}", f"summary-text-{i}", f"summary-icon-{i}", f"summary-divider-{i - 1}"):
                if imya in f:
                    udalit(f[imya])
            continue
        zag, tekst = d["itogi"][i - 1]
        if i > 1:
            postavit(f[f"summary-divider-{i - 1}"], y=y + 0.08)
            y += 0.18
        h_zag = vys(max(1, izm.strok(zag, 12, SHIRINA_ITOGA, True, 0.99)), 12)
        tekst_v_figuru(f[f"summary-title-{i}"], zag or " ")
        postavit(f[f"summary-title-{i}"], y=y, h=h_zag)
        postavit(f[f"summary-icon-{i}"], y=y + min(h_zag, 0.2) / 2 - 0.125)
        y += h_zag + 0.05
        h_t = vys(izm.strok(re.sub(r"\[\^\w+\]", "1", tekst), 11, SHIRINA_ITOGA), 11, 1.05)
        tekst_v_figuru(f[f"summary-text-{i}"], tekst)
        postavit(f[f"summary-text-{i}"], y=y, h=h_t)
        y += h_t
    niz_paneli = y + 0.09
    postavit(f["summary-panel"], h=niz_paneli - 1.70)
    postavit(f["directions-heading"], y=niz_paneli + 0.11)
    postavit(f["directions-marker"], y=niz_paneli + 0.21)
    verh = niz_paneli + 0.11 + 0.48

    # Сноска.
    if d["snoski"]:
        (klyuch, tekst), = d["snoski"].items()
        sn = f["summary-footnote-1"]
        p = telo(sn._element).find(qn("a:p"))
        rs = p.findall(qn("a:r"))
        rs[0].find(qn("a:t")).text = klyuch
        rs[1].find(qn("a:t")).text = " " + tekst
        h_sn = vys(izm.strok(klyuch + " " + tekst, 8, 6.2), 8)
        postavit(sn, y=SNOSKA_Y, w=6.2, h=h_sn)
        niz = SNOSKA_Y - 0.10
    else:
        udalit(f["summary-footnote-1"])
        niz = PODVAL_1

    spisok = [k for k in d["kartochki"] if k["n"] <= 4]
    for n in set(range(1, 5)) - {k["n"] for k in spisok}:
        ubrat_kartochku(f, n)
    return rasstavit_kartochki(f, spisok, verh, niz, izm, kegl)[0]


def tablitsa_kontaktov(f, spisok, izm):
    """Заполняет таблицу, возвращает её высоту в дюймах."""
    tbl = f["contacts-table"]._element.find(".//" + qn("a:tbl"))
    ryady = tbl.findall(qn("a:tr"))
    obraztsy = [copy.deepcopy(ryady[1]), copy.deepcopy(ryady[2])]
    for r in ryady[1:]:
        tbl.remove(r)
    obshchaya = int(ryady[0].get("h")) / E
    for i, (rg, fio, pochta) in enumerate(spisok):
        tr = copy.deepcopy(obraztsy[i % 2])
        yach = tr.findall(qn("a:tc"))
        punkty_v_figuru(yach[0], rg)
        tb = yach[1].find(qn("a:txBody"))
        a_fio, a_pochta = tb.findall(qn("a:p"))[:2]
        zapolnit_abzats(a_fio, fio)
        if pochta:
            zapolnit_abzats(a_pochta, pochta)
        else:
            tb.remove(a_pochta)
        strok = sum(izm.strok(x, 10, 5.12 - 0.28 - 0.16) for x in rg)
        h = max(0.45, vys(max(strok, 2), 10) + 0.14)
        tr.set("h", str(int(h * E)))
        tbl.append(tr)
        obshchaya += h
    postavit(f["contacts-table"], h=obshchaya)
    return obshchaya


def stranitsa_2(slayd, d, kontakty_spisok, izm, kegl):
    f = figury(slayd)
    spisok = [k for k in d["kartochki"] if k["n"] >= 5]
    for n in set(range(5, 9)) - {k["n"] for k in spisok}:
        ubrat_kartochku(f, n)

    # Высота блоков под карточками: «Полезная информация» и «Как подключиться».
    info = [n for n in f if n.startswith("resource-")] + ["info-heading", "info-marker"]
    verh_info = f["info-heading"].top / E
    sdvig_info = {n: f[n].top / E - verh_info for n in info}
    h_info = 0.44 + 1.44 + 0.18
    if kontakty_spisok:
        h_tabl = tablitsa_kontaktov(f, kontakty_spisok, izm)
        h_kont = 0.44 + h_tabl
    else:
        for imya in ("contacts-marker", "contacts-heading", "contacts-table"):
            udalit(f[imya])
        h_kont = 0

    verh = 0.82
    mesta = [Mesto(k, izm, kegl) for k in spisok]
    ryady = [list(range(i, min(i + 2, len(spisok)))) for i in range(0, len(spisok), 2)]
    nuzhno = sum(max(mesta[i].nuzhno for i in r) for r in ryady) + ZAZOR_RYADOV * max(0, len(ryady) - 1)
    niz_kartochek = PODVAL_2 - h_kont - h_info - 0.18
    if verh + nuzhno > niz_kartochek:
        return verh + nuzhno - niz_kartochek
    if spisok:
        _, niz_y = rasstavit_kartochki(f, spisok, verh, niz_kartochek, izm, kegl)
        y_info = niz_y + 0.18
    else:
        y_info = verh
    for n in info:
        postavit(f[n], y=y_info + sdvig_info[n])
    if kontakty_spisok:
        y_k = y_info + h_info
        postavit(f["contacts-heading"], y=y_k)
        postavit(f["contacts-marker"], y=y_k + 0.10)
        postavit(f["contacts-table"], y=y_k + 0.44)
    return 0.0


def period_v_plashku(slayd, d):
    if not d["data_sreza"]:
        return
    tb = telo(figury(slayd)["period"]._element)
    a = copy.deepcopy(tb.find(qn("a:p")))
    r = a.find(qn("a:r"))
    r.find(qn("a:t")).text = f"данные на {d['data_sreza']}"
    r.find(qn("a:rPr")).set("sz", "850")
    r.find(qn("a:rPr")).set("b", "0")
    tb.append(a)


# ---------------------------------------------------------------- сборка


def ubrat_sluzhebnoe(prez):
    """Комментарии, авторы, заметки и сведения о версиях в итоговый файл не идут."""
    for s in prez.slides:
        for rId, rel in list(s.part.rels.items()):
            if "comment" in rel.reltype or rel.reltype.endswith("/notesSlide"):
                s.part.drop_rel(rId)
    for rId, rel in list(prez.part.rels.items()):
        if any(k in rel.reltype for k in ("authors", "revisionInfo", "commentAuthors")):
            prez.part.drop_rel(rId)


def sobrat(d, shablon, spisok_kontaktov, vyhod):
    izm = Izmeritel()
    for kegl in (9.5, 9.0):
        prez = Presentation(str(shablon))
        s1, s2 = prez.slides[0], prez.slides[1]
        nehvatka = max(stranitsa_1(s1, d, izm, kegl), stranitsa_2(s2, d, spisok_kontaktov, izm, kegl))
        if nehvatka == 0:
            break
    else:
        sys.exit(
            f"Текст не помещается на 2 страницы даже шрифтом 9 pt: не хватает {nehvatka * 2.54:.1f} см. "
            "Сократите отбором: меньше пунктов в карточках или короче ключевые итоги (AGENTS.md, п. 7, шаг 7)."
        )
    for s in (s1, s2):
        zamenit_period(s, d["period"])
    period_v_plashku(s1, d)
    ubrat_sluzhebnoe(prez)

    cp = prez.core_properties
    cp.title = f"Дайджест «Открытые API» за {d['period']}"
    cp.author = "Ассоциация ФинТех"
    cp.last_modified_by = ""
    cp.revision = 1
    seychas = datetime.datetime.now().replace(microsecond=0)
    cp.created = cp.modified = seychas

    vyhod.parent.mkdir(parents=True, exist_ok=True)
    prez.save(str(vyhod))
    return kegl


def main(argv=None):
    p = argparse.ArgumentParser(description="Итоговый дайджест PPTX по шаблону")
    p.add_argument("chernovik", help="черновик дайджеста .md (читательская и служебная часть)")
    p.add_argument("--vyhod", help="итоговый файл .pptx; по умолчанию рядом с черновиком")
    p.add_argument("--shablon", default=str(SHABLON))
    p.add_argument("--kontakty", help="таблица контактов; по умолчанию workspace/kontakty.md или nastroyki/kontakty.md")
    p.add_argument("--bez-proverki", action="store_true", help="не запускать proverka.py")
    a = p.parse_args(argv)

    chernovik = Path(a.chernovik)
    if not a.bez_proverki:
        rez = subprocess.run([sys.executable, str(KOREN / "scripts" / "proverka.py"), str(chernovik)])
        if rez.returncode:
            sys.exit("Проверка нашла ошибки: исправьте черновик, затем повторите сборку.")

    d = razobrat(chernovik.read_text(encoding="utf-8"))

    if a.kontakty:
        put_k = Path(a.kontakty)
    else:
        put_k = next(
            (x for x in (KOREN / "workspace" / "kontakty.md", KOREN / "nastroyki" / "kontakty.md") if x.exists()),
            None,
        )
    spisok = kontakty(put_k)

    if a.vyhod:
        vyhod = Path(a.vyhod)
    else:
        imya = re.sub(r"_черновик$", "", chernovik.stem)
        vyhod = chernovik.with_name(imya + ".pptx")
    if vyhod.resolve() == chernovik.resolve():
        sys.exit("Итоговый файл совпадает с черновиком")

    kegl = sobrat(d, Path(a.shablon), spisok, vyhod)
    try:
        pokaz = vyhod.resolve().relative_to(KOREN)
    except ValueError:
        pokaz = vyhod
    print(f"Итоговый файл: {pokaz}")
    nomera = [k["n"] for k in d["kartochki"]]
    print(f"Ключевых итогов: {len(d['itogi'])}, карточек: {len(nomera)} ({', '.join(map(str, nomera))}).")
    net = sorted(set(range(1, 9)) - set(nomera))
    if net:
        print(f"ВНИМАНИЕ: нет карточек {', '.join(map(str, net))}; их места убраны. Проверьте вид страниц.")
    if kegl < 9.5:
        print(f"ВНИМАНИЕ: чтобы текст поместился, шрифт карточек уменьшен до {kegl} pt. Лучше сократить отбором.")
    if not spisok:
        print("ВНИМАНИЕ: контакты не заполнены (workspace/kontakty.md или nastroyki/kontakty.md), "
              "блок «Как подключиться» убран.")
    print("Откройте файл и проверьте вид перед отправкой.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
