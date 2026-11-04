#!/usr/bin/env python3
"""Сборка итогового дайджеста в PPTX по шаблону templates/daydzhest.pptx.

На входе черновик: файл Markdown с читательской и служебной частью
(AGENTS.md, п. 3). На выходе итоговый файл: только читательская часть,
без меток, ссылок и служебной части, оформленная по шаблону АФТ.

Что куда попадает:
    «## Главное за период»          -> «Ключевые итоги» (2-3 пункта
                                       вида «- **Заголовок.** Текст»)
    «**Инициатива**» и её пункты    -> карточки «Основные направления»
    «## События периода»,
    «## В следующем …»              -> отдельные карточки с этим заголовком
    nastroyki/kontakty.md           -> таблица «Контакты» на 2-й странице
    (сначала ищется workspace/kontakty.md)

Высота блоков считается по длине текста. Если текст не помещается
на первую страницу, скрипт уменьшает шрифт карточек до 9 pt, а если
и это не помогает, завершается с ошибкой: дайджест сокращается отбором
(AGENTS.md, п. 6, шаги 6 и 8).

Запуск:
    python3 scripts/sborka.py workspace/outputs/2026-10-05_дайджест_3кв-2026_черновик.md
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
EMU_PT = 12700
MAKS_KARTOCHEK = 8

# ---------------------------------------------------------------- разбор Markdown


def bez_razmetki(s):
    s = re.sub(r"`([^`]*)`", r"\1", s)
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"(?<!\w)[*_](.+?)[*_](?!\w)", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


def punkt(s):
    m = re.match(r"^\s*[-*]\s+(.*\S)\s*$", s)
    return m.group(1) if m else None


def razobrat(tekst):
    """Читательская часть черновика -> словарь для шаблона."""
    m = re.search(r"^---\s*$", tekst, re.M)
    if not m:
        sys.exit("Нет разделителя «---»: это не черновик дайджеста (AGENTS.md, п. 3)")
    chitatel = re.sub(r"<!--.*?-->", "", tekst[: m.start()], flags=re.S)
    sluzh = tekst[m.end() :]

    m = re.search(r"^#\s+Открытые API: дайджест за\s+(.+?)\s*$", chitatel, re.M)
    if not m:
        sys.exit("Нет заголовка «# Открытые API: дайджест за …»")
    period = bez_razmetki(m.group(1))
    m = re.search(r"Данные на\s+(\d{2}\.\d{2}\.\d{4})", chitatel)
    data_sreza = m.group(1) if m else None
    ne_zavershen = bool(re.search(r"Период завершён:\s*нет", sluzh, re.I))

    itogi, kartochki = [], []
    razdel, tekushchaya = None, None
    for s in chitatel.splitlines():
        z = re.match(r"^##\s+(.*?)\s*$", s)
        if z:
            razdel = bez_razmetki(z.group(1))
            tekushchaya = None
            if not razdel.startswith(("Главное", "По инициативам")):
                tekushchaya = {"zagolovok": razdel, "punkty": []}
                kartochki.append(tekushchaya)
            continue
        if razdel is None:
            continue
        p = punkt(s)
        if razdel.startswith("Главное"):
            if p:
                mz = re.match(r"^\*\*(.+?)\*\*[\s.:—-]*(.*)$", p)
                if mz:
                    itogi.append(
                        (bez_razmetki(mz.group(1)).rstrip(".:"), bez_razmetki(mz.group(2)))
                    )
                else:
                    itogi.append(("", bez_razmetki(p)))
            continue
        if razdel.startswith("По инициативам") and not p:
            mb = re.match(r"^\s*\*\*(.+?)\*\*(.*)$", s)
            if mb:
                tekushchaya = {
                    "zagolovok": bez_razmetki(mb.group(1) + mb.group(2)),
                    "punkty": [],
                }
                kartochki.append(tekushchaya)
            continue
        if p and tekushchaya is not None:
            tekushchaya["punkty"].append(bez_razmetki(p))

    kartochki = [k for k in kartochki if k["punkty"]]
    if not itogi:
        sys.exit("В «Главное за период» нет пунктов")
    if len(itogi) > 3:
        sys.exit(f"В «Главное за период» {len(itogi)} пунктов, в шаблоне места для 3")
    if len(kartochki) > MAKS_KARTOCHEK:
        sys.exit(
            f"Блоков {len(kartochki)}, в шаблоне места для {MAKS_KARTOCHEK} карточек. "
            "Сократите отбором (AGENTS.md, п. 6, шаг 6)."
        )
    return {
        "period": period,
        "data_sreza": data_sreza if ne_zavershen else None,
        "itogi": itogi,
        "kartochki": kartochki,
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
    """Ширина строки в пунктах. Шрифт с метрикой Arial, если найден."""

    KANDIDATY = {
        False: [
            "C:/Windows/Fonts/arial.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
        True: [
            "C:/Windows/Fonts/arialbd.ttf",
            "/Library/Fonts/Arial Bold.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
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

    def strok(self, tekst, pt, shirina_emu, zhirnyy=False):
        """Число строк при переносе по словам, с запасом 4 % на разницу программ."""
        if not tekst:
            return 0
        predel = shirina_emu / EMU_PT * 0.96
        n, tek = 1, ""
        for slovo in tekst.split():
            proba = f"{tek} {slovo}".strip()
            if self.shirina(proba, pt, zhirnyy) <= predel or not tek:
                tek = proba
            else:
                n += 1
                tek = slovo
        return n


def vysota(strok, pt, interval=1.0, posle_pt=0.0):
    return int(strok * pt * 1.2 * interval * EMU_PT + posle_pt * EMU_PT)


# ---------------------------------------------------------------- работа с XML


def figury(slayd):
    return {sh.name: sh for sh in slayd.shapes}


def polozhit(el, x=None, y=None, w=None, h=None):
    xfrm = el.find(".//" + qn("a:xfrm"))
    if xfrm is None:  # таблица: p:xfrm
        xfrm = el.find(qn("p:xfrm"))
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if x is not None:
        off.set("x", str(int(x)))
    if y is not None:
        off.set("y", str(int(y)))
    if w is not None:
        ext.set("cx", str(int(w)))
    if h is not None:
        ext.set("cy", str(int(h)))


def telo_teksta(el):
    """txBody фигуры (p:txBody) или ячейки таблицы (a:txBody)."""
    t = el.find(".//" + qn("p:txBody"))
    return t if t is not None else el.find(".//" + qn("a:txBody"))


def tekst_v_figuru(el, tekst, razmer_pt=None):
    """Текст в первый абзац первого прогона; остальные абзацы удаляются."""
    txBody = telo_teksta(el)
    abzatsy = txBody.findall(qn("a:p"))
    for a in abzatsy[1:]:
        txBody.remove(a)
    progony = abzatsy[0].findall(qn("a:r"))
    for r in progony[1:]:
        abzatsy[0].remove(r)
    progony[0].find(qn("a:t")).text = tekst
    if razmer_pt:
        progony[0].find(qn("a:rPr")).set("sz", str(int(razmer_pt * 100)))


def punkty_v_figuru(el, punkty, razmer_pt):
    txBody = telo_teksta(el)
    abzatsy = txBody.findall(qn("a:p"))
    obrazets = copy.deepcopy(abzatsy[0])
    for a in abzatsy:
        txBody.remove(a)
    for p in punkty:
        a = copy.deepcopy(obrazets)
        r = a.find(qn("a:r"))
        r.find(qn("a:t")).text = p
        r.find(qn("a:rPr")).set("sz", str(int(razmer_pt * 100)))
        txBody.append(a)


class Nomera:
    def __init__(self, slayd):
        self.n = max(int(sh.shape_id) for sh in slayd.shapes) + 1

    def __call__(self, el, imya):
        c = el.find(".//" + qn("p:cNvPr"))
        c.set("id", str(self.n))
        c.set("name", imya)
        self.n += 1
        return el


def zamenit_period(slayd, period):
    for sh in slayd.shapes:
        if not sh.has_text_frame:
            continue
        for a in sh.text_frame.paragraphs:
            for r in a.runs:
                if "ПЕРИОД" in r.text:
                    r.text = r.text.replace("ПЕРИОД", period)


# ---------------------------------------------------------------- страница 1

# Координаты шаблона, EMU.
NIZ_STRANITSY = 10124885 - 91440  # верх подвала минус зазор
SHIRINA_TEKSTA_ITOGA = 5895466
SHIRINA_KARTOCHKI = 3240341
SHAG_KOLONKI = 3404725  # 3862133 - 457200
ZAZOR_RYADOV = 164592


def stranitsa_1(slayd, d, izm, kegl_kartochek, plotno=False):
    zazor_ryadov = 91440 if plotno else ZAZOR_RYADOV
    otstup_tela = 54864 if plotno else 91440
    niz_kartochki = 91440 if plotno else 137160
    f = figury(slayd)
    nomer = Nomera(slayd)
    der = slayd.shapes._spTree
    obr = {k: copy.deepcopy(f[k]._element) for k in (
        "summary-dot-1", "summary-title-1", "summary-text-1", "summary-divider-1",
        "wg-card-1", "wg-badge-1", "wg-title-1", "wg-body-1",
    )}
    for imya, sh in f.items():
        if re.match(r"^(summary-(dot|title|text|divider)|wg-(card|badge|title|body))-\d+$", imya):
            der.remove(sh._element)

    # Ключевые итоги.
    panel = f["summary-panel"]._element
    verh_paneli = 1737360
    y = verh_paneli + 164592
    kegl_itoga = 11
    for i, (zag, tekst) in enumerate(d["itogi"], 1):
        if i > 1:
            el = nomer(copy.deepcopy(obr["summary-divider-1"]), f"summary-divider-{i}")
            polozhit(el, y=y)
            der.append(el)
            y += 73152
        verh = y
        if zag:
            h = max(237744, vysota(izm.strok(zag, 12, SHIRINA_TEKSTA_ITOGA, True), 12))
            el = nomer(copy.deepcopy(obr["summary-title-1"]), f"summary-title-{i}")
            tekst_v_figuru(el, zag)
            polozhit(el, y=y, h=h)
            der.append(el)
            y += h + 18288
        if tekst:
            h = vysota(izm.strok(tekst, kegl_itoga, SHIRINA_TEKSTA_ITOGA), kegl_itoga, 1.05)
            el = nomer(copy.deepcopy(obr["summary-text-1"]), f"summary-text-{i}")
            tekst_v_figuru(el, tekst, kegl_itoga)
            polozhit(el, y=y, h=h)
            der.append(el)
            y += h
        el = nomer(copy.deepcopy(obr["summary-dot-1"]), f"summary-dot-{i}")
        polozhit(el, y=verh + 68580)
        der.append(el)
        y += 109728
    niz_paneli = y + 54864
    polozhit(panel, h=niz_paneli - verh_paneli)

    # Заголовок «Основные направления».
    zag_napr = niz_paneli + 347472
    polozhit(f["directions-heading"]._element, y=zag_napr)
    polozhit(f["directions-marker"]._element, y=zag_napr + 91440)

    # Карточки.
    vysoty = []
    for k in d["kartochki"]:
        h_zag = max(310896, vysota(izm.strok(k["zagolovok"], 12, 2490533, True), 12))
        h_tela = sum(
            vysota(izm.strok(p, kegl_kartochek, 2490533 - 146304), kegl_kartochek, 1.0, 2)
            for p in k["punkty"]
        )
        vysoty.append((h_zag, h_tela, 146304 + h_zag + otstup_tela + h_tela + niz_kartochki))
    y = zag_napr + 438913
    for i in range(0, len(vysoty), 2):
        ryad = vysoty[i : i + 2]
        h_ryada = max(v[2] for v in ryad)
        for j, (h_zag, h_tela, _) in enumerate(ryad):
            n = i + j + 1
            k = d["kartochki"][n - 1]
            x = 457200 + j * SHAG_KOLONKI
            el = nomer(copy.deepcopy(obr["wg-card-1"]), f"wg-card-{n}")
            polozhit(el, x=x, y=y, h=h_ryada)
            der.append(el)
            el = nomer(copy.deepcopy(obr["wg-badge-1"]), f"wg-badge-{n}")
            tekst_v_figuru(el, str(n))
            polozhit(el, x=x + 164592, y=y + 146304)
            der.append(el)
            el = nomer(copy.deepcopy(obr["wg-title-1"]), f"wg-title-{n}")
            tekst_v_figuru(el, k["zagolovok"])
            polozhit(el, x=x + 585216, y=y + 146304, h=h_zag)
            der.append(el)
            el = nomer(copy.deepcopy(obr["wg-body-1"]), f"wg-body-{n}")
            punkty_v_figuru(el, k["punkty"], kegl_kartochek)
            polozhit(el, x=x + 585216, y=y + 146304 + h_zag + otstup_tela, h=h_tela)
            der.append(el)
        y += h_ryada + zazor_ryadov
    return y - zazor_ryadov  # низ последнего ряда


def period_v_plashku(slayd, d):
    if not d["data_sreza"]:
        return
    txBody = telo_teksta(figury(slayd)["period"]._element)
    a = copy.deepcopy(txBody.find(qn("a:p")))
    r = a.find(qn("a:r"))
    r.find(qn("a:t")).text = f"данные на {d['data_sreza']}"
    r.find(qn("a:rPr")).set("sz", "850")
    r.find(qn("a:rPr")).set("b", "0")
    txBody.append(a)


# ---------------------------------------------------------------- страница 2


def stranitsa_2(slayd, spisok, izm):
    f = figury(slayd)
    if not spisok:
        der = slayd.shapes._spTree
        for imya in ("contacts-marker", "contacts-heading", "contacts-table"):
            der.remove(f[imya]._element)
        return False
    tbl = f["contacts-table"]._element.find(".//" + qn("a:tbl"))
    ryady = tbl.findall(qn("a:tr"))
    obraztsy = [copy.deepcopy(ryady[1]), copy.deepcopy(ryady[2])]
    for r in ryady[1:]:
        tbl.remove(r)
    obshchaya = int(ryady[0].get("h"))
    for i, (rg, fio, pochta) in enumerate(spisok):
        tr = copy.deepcopy(obraztsy[i % 2])
        yach = tr.findall(qn("a:tc"))
        punkty_v_figuru(yach[0], rg, 11)
        tb = yach[1].find(qn("a:txBody"))
        a_fio, a_pochta = tb.findall(qn("a:p"))[:2]
        a_fio.find(".//" + qn("a:t")).text = fio
        if pochta:
            a_pochta.find(".//" + qn("a:t")).text = pochta
        else:
            tb.remove(a_pochta)
        strok = sum(izm.strok(x, 11, 3654901 - 365760 - 146304) for x in rg)
        h = max(548640, vysota(max(strok, 2), 11) + 182880 + 45720)
        tr.set("h", str(h))
        tbl.append(tr)
        obshchaya += h
    polozhit(f["contacts-table"]._element, h=obshchaya)
    return True


# ---------------------------------------------------------------- сборка


def ubrat_kommentarii(prez):
    for s in prez.slides:
        for rId, rel in list(s.part.rels.items()):
            if "comments" in rel.reltype:
                s.part.rels.pop(rId)
    for rId, rel in list(prez.part.rels.items()):
        if "authors" in rel.reltype or "revisionInfo" in rel.reltype:
            prez.part.rels.pop(rId)


def sobrat(d, shablon, spisok_kontaktov, vyhod):
    izm = Izmeritel()
    # Сначала вид шаблона, затем плотнее отступы, затем меньше шрифт карточек.
    for kegl, plotno in ((10.5, False), (10.5, True), (10, True), (9.5, True), (9, True)):
        prez = Presentation(str(shablon))
        s1, s2 = prez.slides[0], prez.slides[1]
        niz = stranitsa_1(s1, d, izm, kegl, plotno)
        if niz <= NIZ_STRANITSY:
            break
    else:
        lishnee = (niz - NIZ_STRANITSY) / 914400 * 2.54
        sys.exit(
            f"Текст не помещается на первую страницу даже шрифтом 9 pt: лишние {lishnee:.1f} см. "
            "Сократите отбором: меньше строк в карточках или меньше карточек (AGENTS.md, п. 6, шаг 8)."
        )
    for s in (s1, s2):
        zamenit_period(s, d["period"])
    period_v_plashku(s1, d)
    est_kontakty = stranitsa_2(s2, spisok_kontaktov, izm)
    ubrat_kommentarii(prez)

    cp = prez.core_properties
    cp.title = f"Открытые API: дайджест за {d['period']}"
    cp.author = "Ассоциация ФинТех"
    cp.last_modified_by = ""
    cp.revision = 1
    seychas = datetime.datetime.now().replace(microsecond=0)
    cp.created = cp.modified = seychas

    vyhod.parent.mkdir(parents=True, exist_ok=True)
    prez.save(str(vyhod))
    return kegl, plotno, est_kontakty


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

    kegl, plotno, est_kontakty = sobrat(d, Path(a.shablon), spisok, vyhod)
    try:
        pokaz = vyhod.resolve().relative_to(KOREN)
    except ValueError:
        pokaz = vyhod
    print(f"Итоговый файл: {pokaz}")
    print(f"Ключевых итогов: {len(d['itogi'])}, карточек: {len(d['kartochki'])}.")
    if plotno:
        print(
            "ВНИМАНИЕ: чтобы текст поместился, отступы уплотнены"
            + (f" и шрифт карточек уменьшен до {kegl} pt" if kegl < 10.5 else "")
            + ". Лучше сократить отбором."
        )
    if not est_kontakty:
        print("ВНИМАНИЕ: контакты не заполнены (nastroyki/kontakty.md), блок «Контакты» убран со 2-й страницы.")
    print("Откройте файл и проверьте вид перед отправкой.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
