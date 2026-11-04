#!/usr/bin/env python3
"""Комментарии рецензента в PPTX: список с текстом, к которому они относятся.

Нужен для правок по комментариям (AGENTS.md, раздел 8): агент видит каждое
замечание рядом с фрагментом слайда, вносит правку в черновик, собирает
итоговый файл заново и, если нужно, отмечает комментарии решёнными в копии
исходного файла.

Понимает современные комментарии PowerPoint (ppt/comments/modernComment_*.xml)
и старые (ppt/comments/comment*.xml).

Запуск:
    python3 scripts/kommentarii.py файл.pptx                 список комментариев
    python3 scripts/kommentarii.py файл.pptx --md отчёт.md   то же в файл Markdown
    python3 scripts/kommentarii.py файл.pptx --reshit копия.pptx
                                                             копия, где все современные
                                                             комментарии отмечены решёнными

Сторонние библиотеки не нужны.
"""

import argparse
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def local(tag):
    return tag.rsplit("}", 1)[-1]


def tekst(el):
    if el is None:
        return ""
    abz = []
    for p in el.iter(f"{{{NS['a']}}}p"):
        abz.append("".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t")))
    if not abz:
        abz = ["".join(t.text or "" for t in el.iter() if local(t.tag) == "t")]
    return "\n".join(x for x in abz if x).strip()


def poryadok_slaydov(z):
    """[(номер, путь к слайду)] по порядку в презентации."""
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    tseli = {r.get("Id"): r.get("Target") for r in rels}
    rez = []
    for i, s in enumerate(pres.iter(f"{{{NS['p']}}}sldId"), 1):
        rid = s.get(f"{{{NS['r']}}}id")
        rez.append((i, posixpath.normpath(posixpath.join("ppt", tseli[rid]))))
    return rez


def figury_slayda(xml):
    """{id фигуры: (имя, текст)}."""
    rez = {}
    for sp in xml.iter():
        if local(sp.tag) not in ("sp", "graphicFrame", "pic"):
            continue
        c = next((x for x in sp.iter() if local(x.tag) == "cNvPr"), None)
        if c is None:
            continue
        rez[c.get("id")] = (c.get("name", ""), tekst(sp))
    return rez


def avtory(z):
    rez = {}
    for imya in ("ppt/authors.xml", "ppt/commentAuthors.xml"):
        if imya in z.namelist():
            for a in ET.fromstring(z.read(imya)):
                rez[a.get("id")] = a.get("name", "")
    return rez


def kommentarii(put):
    rez = []
    with zipfile.ZipFile(put) as z:
        names = set(z.namelist())
        avt = avtory(z)
        for nomer, slayd in poryadok_slaydov(z):
            xml = ET.fromstring(z.read(slayd))
            fig = figury_slayda(xml)
            rels_put = posixpath.join(posixpath.dirname(slayd), "_rels", posixpath.basename(slayd) + ".rels")
            if rels_put not in names:
                continue
            for r in ET.fromstring(z.read(rels_put)):
                if "comment" not in r.get("Type", "").lower() or "author" in r.get("Type", "").lower():
                    continue
                f = posixpath.normpath(posixpath.join(posixpath.dirname(slayd), r.get("Target")))
                if f not in names:
                    continue
                koren = ET.fromstring(z.read(f))
                for cm in koren:
                    if local(cm.tag) != "cm":
                        continue
                    sp = next((x for x in cm.iter() if local(x.tag) == "spMk"), None)
                    tx = next((x for x in cm.iter() if local(x.tag) == "txMk"), None)
                    imya, fragment = fig.get(sp.get("id"), ("", "")) if sp is not None else ("", "")
                    if tx is not None and fragment:
                        try:
                            nach, dl = int(tx.get("cp", 0)), int(tx.get("len", 0))
                            # cp считается по тексту фигуры без переводов строк между абзацами
                            ploskiy = fragment.replace("\n", "")
                            if dl:
                                fragment = ploskiy[nach: nach + dl]
                        except ValueError:
                            pass
                    body = next((x for x in cm if local(x.tag) in ("txBody", "text")), None)
                    otvety = [
                        (avt.get(x.get("authorId"), ""), tekst(next((y for y in x if local(y.tag) == "txBody"), None)))
                        for x in cm.iter() if local(x.tag) == "reply"
                    ]
                    rez.append({
                        "slayd": nomer,
                        "fayl": f,
                        "id": cm.get("id") or cm.get("idx"),
                        "avtor": avt.get(cm.get("authorId"), ""),
                        "status": {"resolved": "решён", "closed": "закрыт", "active": "активен"}.get(cm.get("status", "active"), cm.get("status")),
                        "figura": imya,
                        "fragment": fragment,
                        "tekst": tekst(body) if local(body.tag) == "txBody" else (body.text or "") if body is not None else "",
                        "otvety": otvety,
                    })
    return rez


def otchet(spisok):
    stroki = []
    for i, k in enumerate(spisok, 1):
        stroki.append(f"## {i}. Слайд {k['slayd']}, {k['figura'] or 'фигура не указана'} ({k['status']})")
        if k["fragment"]:
            stroki.append(f"Фрагмент: «{k['fragment'][:300]}»")
        stroki.append(f"Комментарий ({k['avtor'] or 'автор не указан'}): {k['tekst']}")
        for a, t in k["otvety"]:
            stroki.append(f"- ответ ({a}): {t}")
        stroki.append("")
    stroki.append(f"Всего комментариев: {len(spisok)}.")
    return "\n".join(stroki)


def reshit(put, vyhod):
    n = 0
    with zipfile.ZipFile(put) as zin, zipfile.ZipFile(vyhod, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if re.match(r"ppt/comments/modernComment_.*\.xml$", it.filename):
                s = data.decode("utf-8")
                s, k = re.subn(r'(<p188:cm\b[^>]*?)\sstatus="[^"]*"', r"\1", s)
                s, k = re.subn(r"<p188:cm\b", '<p188:cm status="resolved"', s)
                n += k
                data = s.encode("utf-8")
            zout.writestr(it, data)
    return n


def main(argv=None):
    p = argparse.ArgumentParser(description="Комментарии в PPTX")
    p.add_argument("fayl", help="файл .pptx с комментариями")
    p.add_argument("--md", help="сохранить список в файл Markdown")
    p.add_argument("--reshit", metavar="КОПИЯ.pptx", help="сохранить копию, где комментарии отмечены решёнными")
    a = p.parse_args(argv)

    if a.reshit:
        if a.reshit == a.fayl:
            sys.exit("Укажите другое имя для копии: исходный файл не меняется")
        n = reshit(a.fayl, a.reshit)
        print(f"Отмечено решёнными: {n}. Копия: {a.reshit}")
        return 0
    spisok = kommentarii(a.fayl)
    tekst_otcheta = otchet(spisok)
    if a.md:
        with open(a.md, "w", encoding="utf-8") as f:
            f.write("# Комментарии в «" + a.fayl + "»\n\n" + tekst_otcheta + "\n")
        print(f"Сохранено: {a.md}. Комментариев: {len(spisok)}.")
    else:
        print(tekst_otcheta)
    return 0


if __name__ == "__main__":
    sys.exit(main())
