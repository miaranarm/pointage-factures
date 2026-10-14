#!/usr/bin/env python3
"""Pointage prudent de lignes Excel vers des pièces justificatives locales.

Entrées attendues :
  input/P.MBS.xlsx
  input/P.DISTRIMAX.xlsx
  documents/** (PDF, JPG, JPEG, PNG, TIFF)
Sorties dans output/. Aucun document n'est envoyé à un service externe.
"""
from __future__ import annotations
import argparse, csv, re, sys, unicodedata, os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import openpyxl
import fitz
from PIL import Image
import pytesseract
from rapidfuzz import fuzz
from dateutil import parser as dateparser

EXTS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff"}
TARGET_SHEETS = {"P.MBS.xlsx": "Sheet1 (3)", "P.DISTRIMAX.xlsx": "Sheet1 (2)"}
ALIASES = {
 "numero": ["Numéro", "Numero", "N°", "No"],
 "reference": ["Référence", "Reference"],
 "reference_interne": ["Référence Interne", "Reference Interne"],
 "partner": ["Nom d'affichage du partenaire de la facture"],
 "date": ["Date de facturation"],
 "ht": ["Montant hors taxes signé dans la devise"],
 "tax": ["Taxe signée"],
 "ttc": ["Total signé en devises"],
 "file": ["Case à cocher"],
}
TEXT_EXTENSIONS = EXTS

def norm(value):
    if value is None: return ""
    s = unicodedata.normalize("NFKD", str(value))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    return re.sub(r"[^A-Z0-9]+", "", s)

def amount(value):
    if value is None or value == "": return None
    if isinstance(value, (int, float, Decimal)):
        try: return Decimal(str(value))
        except InvalidOperation: return None
    s = str(value).replace("\u00a0", "").replace(" ", "").replace(" MGA", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    else: s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-+]", "", s)
    try: return Decimal(s)
    except InvalidOperation: return None

def parse_date(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if not value: return None
    try: return dateparser.parse(str(value), dayfirst=True, fuzzy=True).date()
    except Exception: return None

def extract_pdf(path):
    pages = []
    with fitz.open(path) as doc:
        for page in doc:
            txt = page.get_text("text") or ""
            if len(re.sub(r"\s", "", txt)) < 25:
                try:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
                    import io
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    txt += "\n" + pytesseract.image_to_string(img, lang="fra+eng")
                except Exception as e:
                    txt += f"\n[OCR_ERROR:{type(e).__name__}]"
            pages.append(txt)
    return "\n".join(pages)

def extract_image(path):
    return pytesseract.image_to_string(Image.open(path), lang="fra+eng")

def get_text(path):
    try:
        return extract_pdf(path) if path.suffix.lower() == ".pdf" else extract_image(path)
    except Exception as e:
        return f"[READ_ERROR:{type(e).__name__}:{e}]"

def find_col(headers, aliases):
    normalized = {norm(v): i for i,v in enumerate(headers) if v is not None}
    for a in aliases:
        if norm(a) in normalized: return normalized[norm(a)]
    return None

def row_fields(headers, row):
    result = {}
    for key, aliases in ALIASES.items():
        idx = find_col(headers, aliases)
        result[key] = row[idx] if idx is not None and idx < len(row) else None
    return result

def token_hits(value, text_norm):
    n = norm(value)
    if len(n) < 4: return False
    return n in text_norm

GENERIC_PARTNERS = {
    "FOURNISSEUR DIVERS", "TRANSPORTEUR DIVERS", "FOURNISSEUR", "DIVERS",
    "TRANSPORTEUR", "CLIENT DIVERS", "PARTENAIRE DIVERS", "INCONNU",
}
GENERIC_PARTNER_TOKENS = {"FOURNISSEUR", "TRANSPORTEUR", "DIVERS", "SOCIETE", "SARL", "SA", "EURL"}

def score_row(fields, doc):
    """Score explicable et conservateur; les indices génériques ne doivent pas suffire."""
    t = doc["norm"]
    score, reasons = 0, []
    identifiers = [("numero", 45), ("reference", 24), ("reference_interne", 18)]
    for key, weight in identifiers:
        v = fields.get(key)
        if token_hits(v, t):
            score += weight
            reasons.append(key)

    # Une date comptable peut différer de la date de facture : elle est un indice,
    # jamais une preuve à elle seule. On ne lui attribue donc pas de pénalité ici.
    d = parse_date(fields.get("date"))
    if d and re.search(rf"(?<!\d){d.day:02d}[./-]{d.month:02d}[./-]{d.year}(?!\d)", doc["text"]):
        score += 8
        reasons.append("date")

    # Les achats peuvent être négatifs dans le grand livre alors que la facture
    # présente un total positif : la comparaison textuelle ignore donc le signe.
    for key, weight in (("ht", 12), ("tax", 8), ("ttc", 15)):
        val = amount(fields.get(key))
        if val is None:
            continue
        absolute = abs(val)
        candidates = {norm(absolute), norm(f"{absolute:,.2f}"), norm(f"{absolute:.2f}")}
        if any(len(x) >= 4 and x in t for x in candidates):
            score += weight
            reasons.append(key)

    raw_partner = str(fields.get("partner") or "").strip()
    partner = norm(raw_partner)
    if len(partner) >= 5 and partner not in {norm(x) for x in GENERIC_PARTNERS}:
        # Exclure les mots génériques : « FOURNISSEUR DIVERS » ne distingue
        # aucune facture et ne doit pas apporter de points.
        partner_tokens = {
            norm(tok) for tok in re.findall(r"\w+", raw_partner, flags=re.UNICODE)
            if len(norm(tok)) >= 3 and norm(tok) not in GENERIC_PARTNER_TOKENS
        }
        doc_tokens = doc.get("tokens", set())
        if partner and partner in t:
            score += 14
            reasons.append("partenaire")
        elif partner_tokens:
            overlap = partner_tokens & doc_tokens
            # Une seule petite occurrence ne suffit pas. Pour un partenaire
            # composé, demander au moins deux mots discriminants concordants.
            if len(overlap) >= 2:
                score += 7
                reasons.append("partenaire_approximatif")
            elif len(partner_tokens) == 1 and len(next(iter(partner_tokens))) >= 7 and overlap:
                score += 5
                reasons.append("partenaire_approximatif")
    return min(score, 100), reasons

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="input")
    ap.add_argument("--documents", default="documents")
    ap.add_argument("--output", default="output")
    ap.add_argument("--min-score", type=int, default=65)
    ap.add_argument("--margin", type=int, default=12)
    args = ap.parse_args()
    root, docs_dir, out = Path(args.input), Path(args.documents), Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    paths = sorted(p for p in docs_dir.rglob("*") if p.is_file() and p.suffix.lower() in EXTS)
    if not paths:
        raise SystemExit(f"Aucun PDF/image trouvé dans {docs_dir}. Arrêt sans modifier les classeurs.")
    # OCR en parallèle : les fichiers sont indépendants et Tesseract est lancé
    # comme processus externe. Limiter à 4 tâches pour éviter de saturer le runner.
    documents = []
    max_workers = min(4, max(1, os.cpu_count() or 2))
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(get_text, path): path for path in paths}
        for i, future in enumerate(as_completed(futures), 1):
            path = futures[future]
            text = future.result()
            documents.append({"path": path, "name": path.name, "text": text,
                              "norm": norm(text), "lines": text.splitlines(),
                              "tokens": {norm(tok) for tok in re.findall(r"\\w+", text, flags=re.UNICODE)
                                         if len(norm(tok)) >= 3}})
            print(f"Indexation {i}/{len(paths)} : {path.name}", flush=True)
    # L'ordre de traitement des tâches étant parallèle, rétablir un ordre stable.
    documents.sort(key=lambda doc: str(doc["path"]).casefold())
    with (out / "rapport_pointage.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["classeur","feuille","ligne_excel","numero","partenaire","date","total_ttc","statut","fichier_retenu","score","candidats","motifs"])
        writer.writeheader()
        for filename, sheetname in TARGET_SHEETS.items():
            src = root / filename
            if not src.exists():
                print(f"Classeur absent : {src}", file=sys.stderr); continue
            wb = openpyxl.load_workbook(src)
            if sheetname not in wb.sheetnames:
                raise SystemExit(f"Feuille cible {sheetname!r} absente de {filename}; feuilles présentes: {wb.sheetnames}")
            ws = wb[sheetname]
            headers = [c.value for c in ws[1]]
            checkbox_col = find_col(headers, ALIASES["file"])
            if checkbox_col is None:
                raise SystemExit(f"Colonne 'Case à cocher' absente de {filename}/{sheetname}; aucun changement effectué.")
            for rownum in range(2, ws.max_row + 1):
                values = [ws.cell(rownum, c).value for c in range(1, ws.max_column + 1)]
                fields = row_fields(headers, values)
                scored = []
                for doc in documents:
                    score, reasons = score_row(fields, doc)
                    if score > 0: scored.append((score, doc, reasons))
                scored.sort(key=lambda x: x[0], reverse=True)
                chosen, status, score, reasons = "", "NON_POINTE", 0, []
                candidates = scored[:3]
                if candidates:
                    score, doc, reasons = candidates[0]
                    gap = score - candidates[1][0] if len(candidates) > 1 else score
                    if score >= args.min_score and gap >= args.margin:
                        chosen, status = doc["name"], "POINTE_A_CONTROLER"
                    elif score >= args.min_score:
                        status = "AMBIGU"
                    else:
                        status = "CANDIDAT_FAIBLE"
                # Case à cocher reçoit exclusivement le nom de fichier, jamais une coche.
                ws.cell(rownum, checkbox_col + 1).value = chosen
                writer.writerow({
                    "classeur": filename, "feuille": sheetname, "ligne_excel": rownum,
                    "numero": fields.get("numero") or "", "partenaire": fields.get("partner") or "",
                    "date": str(fields.get("date") or ""), "total_ttc": fields.get("ttc") if fields.get("ttc") is not None else "",
                    "statut": status, "fichier_retenu": chosen, "score": score,
                    "candidats": " | ".join(f"{d['name']}:{s}" for s,d,_ in candidates),
                    "motifs": ",".join(reasons),
                })
            outpath = out / f"pointe_{filename}"
            wb.save(outpath)
            print(f"Classeur écrit : {outpath}", flush=True)
    # Contrôle transversal : un même justificatif ne doit pas être affecté
    # silencieusement à des références comptables distinctes.
    report_path = out / "rapport_pointage.csv"
    with report_path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        report_fields = reader.fieldnames
        report_rows = list(reader)

    groups = {}
    for row in report_rows:
        name = row.get("fichier_retenu", "").strip()
        if name and row.get("statut") == "POINTE_A_CONTROLER":
            groups.setdefault(name, []).append(row)

    flagged = [rows for rows in groups.values()
               if len({r.get("numero", "").strip() for r in rows
                       if r.get("numero", "").strip()}) > 1]
    affected = {}
    for rows in flagged:
        for row in rows:
            row["statut"] = "DOUBLON_A_CONTROLER"
            row["fichier_retenu"] = ""
            reasons = [x for x in row.get("motifs", "").split(",") if x]
            if "justificatif_reutilise" not in reasons:
                reasons.append("justificatif_reutilise")
            row["motifs"] = ",".join(reasons)
            affected.setdefault(row["classeur"], set()).add(int(row["ligne_excel"]))

    # Refléter les alertes dans les classeurs : laisser la cellule de pointage
    # vide pour qu'une proposition réutilisée ne soit pas prise pour un pointage.
    for filename, rownums in affected.items():
        workbook_path = out / f"pointe_{filename}"
        wb = openpyxl.load_workbook(workbook_path)
        sheetname = TARGET_SHEETS[filename]
        ws = wb[sheetname]
        headers = [c.value for c in ws[1]]
        checkbox_col = find_col(headers, ALIASES["file"])
        for rownum in rownums:
            ws.cell(rownum, checkbox_col + 1).value = ""
        wb.save(workbook_path)
        print(f"Doublons suspects à contrôler dans {filename} : {len(rownums)} ligne(s)", flush=True)

    if flagged:
        with report_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=report_fields)
            writer.writeheader()
            writer.writerows(report_rows)

    print(f"Terminé. Documents indexés : {len(documents)}. Examiner le rapport avant validation.")

if __name__ == "__main__":
    main()
