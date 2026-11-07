"""
Kelime Atölyesi - Flask sunucusu
Görevi: Word/PDF dosyalarını okuyup içindeki kelimeleri (İngilizce - Türkçe) ayıklamak.
Çalışma ekranı (index.html) tarayıcıda çalışır, ilerleme tarayıcıda saklanır.
"""
import io
import os
import re
from collections import Counter

import pdfplumber
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from flask import Flask, jsonify, request, send_from_directory

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # en fazla 25 MB yükleme

SECTION_RE = re.compile(r"^(P\s?\d{1,3}|\d{1,2}\s?[A-Za-z]?|unit\s?\d+\w*|theme\s?\d+\w*)$", re.I)
BULLETS = "-•·\uf0b7\u2022*–—▪●"


def clean(s):
    s = (s or "").replace("\u00a0", " ").replace("*", "")
    s = s.strip()
    return re.sub(r"\s+", " ", s)


def split_pair(line):
    """Bir satırı (ingilizce, türkçe) olarak ikiye böler. Olmazsa None döner."""
    raw = line.strip().lstrip(BULLETS).strip()
    if "\t" in raw:
        parts = [p.strip() for p in re.split(r"\t+", raw) if p.strip()]
        if len(parts) >= 2:
            return clean(parts[0]), clean(" ".join(parts[1:]))
    m = re.split(r"\s{3,}", raw)
    if len(m) >= 2:
        return clean(m[0]), clean(" ".join(m[1:]))
    if ":" in raw:
        a, b = raw.split(":", 1)
        if a.strip() and b.strip() and len(a.split()) <= 6:
            return clean(a), clean(b)
    if " = " in raw:
        a, b = raw.split(" = ", 1)
        return clean(a), clean(b)
    m = re.split(r"\s[-–—]\s", raw, maxsplit=1)
    if len(m) == 2 and m[0].strip() and m[1].strip():
        return clean(m[0]), clean(m[1])
    return None


# ----------------------------- WORD -----------------------------
def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def parse_docx(fp):
    doc = Document(fp)
    entries, section = [], ""
    for block in iter_blocks(doc):
        lines = []
        if isinstance(block, Paragraph):
            lines = block.text.split("\n")
        else:
            for row in block.rows:
                cells = []
                for c in row.cells:
                    t = c.text.strip()
                    if t and (not cells or cells[-1] != t):
                        cells.append(t)
                if len(cells) >= 2:
                    lines.append("\t".join(cells))
        for ln in lines:
            t = clean(ln.strip().lstrip(BULLETS))
            if not t:
                continue
            if SECTION_RE.match(t) and "\t" not in ln:
                section = t.upper() if re.match(r"^P", t, re.I) else t
                continue
            pair = split_pair(ln)
            if pair and pair[0] and pair[1]:
                entries.append({"en": pair[0], "tr": pair[1], "ex": "", "sec": section})
            elif entries and len(t.split()) >= 3:
                e = entries[-1]
                e["ex"] = (e["ex"] + " " + t).strip()
    return entries


# ------------------------------ PDF ------------------------------
def group_lines(ws):
    ws = sorted(ws, key=lambda w: (w["top"], w["x0"]))
    lines, cur, top = [], [], None
    for w in ws:
        if top is None or abs(w["top"] - top) <= 3:
            cur.append(w)
            top = w["top"] if top is None else top
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur, top = [w], w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))
    return lines


def parse_pdf(fp):
    entries, section = [], ""
    with pdfplumber.open(fp) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1.5, y_tolerance=2)
            xs = sorted(w["x0"] for w in words if w["text"] == "-")
            clusters = []
            for x in xs:
                if clusters and x - clusters[-1][-1] <= 6:
                    clusters[-1].append(x)
                else:
                    clusters.append([x])
            starts = [sum(c) / len(c) for c in clusters if len(c) >= 3]
            if not starts:
                continue
            cols = [[] for _ in starts]
            for w in words:
                idx = 0
                for k, s in enumerate(starts):
                    if w["x0"] >= s - 4:
                        idx = k
                cols[idx].append(w)
            for cw in cols:
                lines = group_lines(cw)
                cands = []
                for ln in lines:
                    if ln[0]["text"] == "-":
                        rest = ln[1:]
                        best = None
                        for a, b in zip(rest, rest[1:]):
                            g = b["x0"] - a["x1"]
                            if g > 7 and (best is None or g > best[0]):
                                best = (g, b["x0"])
                        if best:
                            cands.append(round(best[1] / 4))
                tr_x = Counter(cands).most_common(1)[0][0] * 4 if cands else None
                prev, last_ex = None, False
                for ln in lines:
                    first = ln[0]
                    toks = [w for w in ln if w["text"] not in ("\uf0b7", "•", "·")]
                    text = " ".join(w["text"] for w in toks).strip()
                    if not text:
                        continue
                    if first["text"] == "-":
                        rest = ln[1:]
                        if not rest:
                            continue
                        split_i = None
                        if tr_x is not None:
                            for i, w in enumerate(rest):
                                if i > 0 and abs(w["x0"] - tr_x) <= 3:
                                    split_i = i
                                    break
                        if split_i is None:
                            best = None
                            for i in range(1, len(rest)):
                                g = rest[i]["x0"] - rest[i - 1]["x1"]
                                if g > 7 and (best is None or g > best[0]):
                                    best = (g, i)
                            split_i = best[1] if best else None
                        en = " ".join(w["text"] for w in (rest[:split_i] if split_i else rest))
                        tr = " ".join(w["text"] for w in rest[split_i:]) if split_i else ""
                        prev = {"en": clean(en), "tr": clean(tr), "ex": "", "sec": section}
                        entries.append(prev)
                        last_ex = False
                        continue
                    if SECTION_RE.match(text):
                        section = text.upper() if text[0] in "Pp" else text
                        prev, last_ex = None, False
                        continue
                    if prev is None:
                        continue
                    aligned = tr_x is not None and abs(first["x0"] - tr_x) <= 3
                    nwords = len(text.split())
                    short_tr = (
                        not last_ex
                        and not prev["ex"]
                        and text[0].islower()
                        and not re.search(r"[.?!]", text)
                        and (nwords <= 3 or (re.search("[çğıöşüÇĞİÖŞÜ]", text) and nwords <= 6))
                    )
                    if (aligned and not last_ex) or short_tr or (not prev["tr"] and first["text"] not in ("\uf0b7",)):
                        prev["tr"] = clean(prev["tr"] + " " + text)
                    else:
                        prev["ex"] = (prev["ex"] + " " + text).strip()
                        last_ex = True
    return [e for e in entries if e["en"]]


def dedupe(entries):
    seen, out = set(), []
    for e in entries:
        k = (e["en"].lower(), e["tr"].lower())
        if k in seen:
            continue
        seen.add(k)
        out.append(e)
    return out


@app.route("/")
def index():
    return send_from_directory(BASE, "index.html")


@app.route("/api/parse", methods=["POST"])
def api_parse():
    result = []
    for f in request.files.getlist("files"):
        name = f.filename or "dosya"
        ext = os.path.splitext(name)[1].lower()
        try:
            data = io.BytesIO(f.read())
            if ext == ".docx":
                words = parse_docx(data)
            elif ext == ".pdf":
                words = parse_pdf(data)
            else:
                result.append({"name": name, "error": "Sadece .docx ve .pdf destekleniyor", "words": []})
                continue
            result.append({"name": os.path.splitext(name)[0], "words": dedupe(words)})
        except Exception as ex:  # noqa
            result.append({"name": name, "error": "Dosya okunamadı: %s" % ex, "words": []})
    return jsonify(result)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
