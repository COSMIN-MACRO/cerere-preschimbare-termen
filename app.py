# -*- coding: utf-8 -*-
"""
Generator Cerere Preschimbare Termen Soluționare (Formular nr. 5 - model nou ONRC)
Aplicație desktop Python + customtkinter.

Rescris de la zero pentru formatul REAL de text (frază narativă), de tipul:
  "DEHELEAN SILVIA , cetatean roman, nascuta la data de 22,05,70 in Simeria jud HD,
   domiciliata in sat Banpotoc com Hărău nr 93 jud Hunedoara identificata cu Carte de
   Identitate seria XD nr 061488 /SPCLEP Deva/13,05,22-03,08,2031 CNP 2700522205503
   UNITED BETON S.R.L. J2024026945008, cui 50626150 Sat Dumbrăviţa, ... Județ Timiş"

Flux:
  1. Lipești textul (frază de identificare + eventual print screen ONRC) sau tastezi.
  2. "EXTRAGE DATE AUTOMAT" -> completează câmpurile (fundal galben = auto).
  3. Verifici/corectezi, bifezi ce dorești, scrii motivul.
  4. "GENEREAZĂ CEREREA" -> completează formularul PDF oficial (editabil) și îl salvează.
"""

import os
import re
import threading
import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter

# ---- drag & drop (opțional) -------------------------------------------------
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    _HAS_DND = False

# ---- OCR (opțional) ---------------------------------------------------------
try:
    import pytesseract
    from PIL import Image
    _HAS_OCR = True
except Exception:
    _HAS_OCR = False

# ---- extragere text din documente ------------------------------------------
try:
    import docx as _docxmod  # python-docx
    _HAS_DOCX = True
except Exception:
    _HAS_DOCX = False

# =============================================================================
# CONSTANTE
# =============================================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PDF = os.path.join(APP_DIR, "Formular5_nou_blank.pdf")
OUTPUT_BASE = os.path.join(APP_DIR, "OUTPUT")

COL_HEADER = "#1a3a5c"
COL_BLUE = "#2563eb"
COL_LIGHT = "#dbeafe"
COL_BG = "#f0f4f8"
COL_WHITE = "#ffffff"
COL_BORDER = "#cbd5e1"
COL_YELLOW = "#fef9c3"

# cheie logică -> nume câmp AcroForm în PDF (verificat vizual pe șablon)
PDF_TEXT = {
    "tribunal": "1",
    "cerere_nr_data": "1_2",
    "nume": "2", "prenume": "3", "cnp": "4",
    "domiciliu": "10",
    "judet_sector": "11", "tara": "15", "cetatenie": "12",
    "nasc_loc": "14", "nasc_judet": "16",
    "data_nasterii": "18", "act_ident": "19", "seria": "20",
    "nr_act": "27", "emis_de": "21", "data_emitere": "22",
    "valabil": "24", "calitate": "30",
    "reprez_firma": "25", "pt_firma": "32",
    "inscris1": "45", "inscris2": "46", "inscris3": "47", "inscris4": "48",
    "inscris5": "49", "inscris6": "50", "inscris7": "51", "inscris8": "52",
    "data_amanare": "22052026-56", "termen": "22052026-57",
    "motiv1": "22052026-53", "motiv2": "22052026-54", "motiv3": "22052026-55",
    "semnat_nume": "59", "semnat_data": "60", "gdpr_data": "60_2",
}

PDF_CHECK = {
    "cb_completare_II": "Box1",
    "cb_preschimbare_II": "Box2",
    "cb_audienta_II": "Box3",
    "cb_renuntare_II": "Box3_2",
    "cb_completare_V": "Box4",
    "cb_preschimbare_VI": "Box5",
}

LUNI = ["", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]


# =============================================================================
# NORMALIZARE + PARSER (format narativ real)
# =============================================================================
def norm_date(d):
    """'22,05,70' / '13.05.22' / '03-08-2031' -> 'DD.MM.YYYY'."""
    parts = [p for p in re.split(r'[.,/\s-]+', d.strip()) if p]
    if len(parts) != 3:
        return d.strip()
    dd, mm, yy = parts
    dd, mm = dd.zfill(2), mm.zfill(2)
    if len(yy) == 2:
        yy = ("19" if int(yy) > 40 else "20") + yy
    return f"{dd}.{mm}.{yy}"

def data_din_cnp(cnp):
    if not cnp or len(cnp) < 7:
        return ""
    s = cnp[0]
    yy, mm, dd = cnp[1:3], cnp[3:5], cnp[5:7]
    sec = {"1": "19", "2": "19", "3": "18", "4": "18", "5": "20", "6": "20"}.get(s, "19")
    return f"{dd}.{mm}.{sec}{yy}"

def extrage(text):
    """Extrage toate datele din textul narativ. Returnează dict cu chei = câmpuri GUI."""
    d = {}
    raw = text or ""
    tt = " " + re.sub(r'\s+', ' ', raw).strip() + " "

    # --- Nume + prenume: cuvinte MAJUSCULE de la început, până la prima virgulă
    m = re.match(r'\s*([A-ZĂÂÎȘȚŞŢ][A-ZĂÂÎȘȚŞŢ\'\- ]+?)\s*,', raw)
    if m:
        full = re.sub(r'\s+', ' ', m.group(1).strip())
        toks = full.split(" ")
        d["nume"] = toks[0]
        d["prenume"] = " ".join(toks[1:])
        d["semnat_nume"] = full

    # --- Cetățenie
    if re.search(r'cet[a-zăâîșțşţ]+ean\w*\s+rom[âaî]n', tt, re.I) or re.search(r'\bROU\b', tt):
        d["cetatenie"] = "Română"
    d["tara"] = "România"

    # --- CNP (+ data nașterii ca rezervă)
    m = re.search(r'CNP\s*[:\-]?\s*([1-8]\d{12})', tt) or re.search(r'\b([1-8]\d{12})\b', tt)
    if m:
        d["cnp"] = m.group(1)

    # --- Data nașterii + loc + județ ("nascuta la data de 22,05,70 in Simeria jud HD")
    m = re.search(r'n[ăa]scut[ăa]?\s+la data de\s+([\d.,/ -]{6,12}?)\s+[îi]n\s+'
                  r'([A-ZĂÂÎȘȚŞŢ][A-Za-zĂÂÎȘȚăâîșț\.\-]+)\s+jud\.?\s*'
                  r'([A-ZĂÂÎȘȚ]{2}|[A-ZĂÂÎȘȚ][a-zăâîșț]+)', tt, re.I)
    if m:
        d["data_nasterii"] = norm_date(m.group(1))
        d["nasc_loc"] = m.group(2).strip(" .")
        d["nasc_judet"] = m.group(3).strip()
    else:
        m = re.search(r'la data de\s+([\d.,/ -]{6,12})', tt)
        if m:
            d["data_nasterii"] = norm_date(m.group(1))
    if not d.get("data_nasterii") and d.get("cnp"):
        d["data_nasterii"] = data_din_cnp(d["cnp"])

    # --- Domiciliu ("domiciliata in <...> identificata/CNP")
    m = re.search(r'domiciliat[ăa]?\s+[îi]n\s+(.*?)\s+'
                  r'(?:identificat|legitimat|posesor|posesoar|cu\s+[Cc]arte|cu\s+CI|CNP)', tt, re.I)
    if m:
        d["domiciliu"] = re.sub(r'\s+', ' ', m.group(1).strip(" ,."))

    # --- Act de identitate
    if re.search(r'[Cc]arte de [Ii]dentitate|\bC\.?I\.?\b|buletin', tt):
        d["act_ident"] = "CI"
    m = re.search(r'seria\s+([A-Z]{1,2})\s*,?\s*nr\.?\s*([0-9]{5,7})', tt, re.I)
    if m:
        d["seria"] = m.group(1).upper()
        d["nr_act"] = m.group(2)
    # emitent + date:  /SPCLEP Deva/13,05,22-03,08,2031
    m = re.search(r'/\s*([^/]+?)\s*/\s*([\d.,/ ]{6,12})\s*-\s*([\d.,/ ]{6,12})', tt)
    if m:
        d["emis_de"] = re.sub(r'\s+', ' ', m.group(1).strip())
        d["data_emitere"] = norm_date(m.group(2))
        d["valabil"] = norm_date(m.group(3))
    else:
        m = re.search(r'(SPCLEP\s+[A-Za-zĂÂÎȘȚăâîșț]+)', tt)
        if m:
            d["emis_de"] = m.group(1)
        m = re.search(r'valabil[ăa]?\s+p[âa]n[ăa]?\s+la\s+([\d.,/ -]{6,12})', tt, re.I)
        if m:
            d["valabil"] = norm_date(m.group(1))

    # --- Calitate: implicit "administrator" (cel care depune este mereu administratorul)
    d["calitate"] = "administrator"

    # --- Firmă
    m = re.search(r'((?:[A-ZĂÂÎȘȚ][A-Za-zĂÂÎȘȚăâîșț&\.\-]+\s+){1,5}(?:S\.?R\.?L\.?|S\.?A\.?))\b', tt)
    if m:
        den = re.sub(r'\s+', ' ', m.group(1).strip())
        den = re.sub(r'S\.?R\.?L\.?$', 'S.R.L.', den)
        den = re.sub(r'S\.?A\.?$', 'S.A.', den)
        d["denumire"] = den
    m = re.search(r'\b(J\s?[0-9]{2,4}[0-9/]{3,})\b', tt)
    if m:
        d["J"] = m.group(1).replace(" ", "")
    m = re.search(r'(?:cui|C\.?U\.?I\.?)\s*:?\s*(RO)?\s*(\d{5,10})', tt, re.I)
    if m:
        d["CUI"] = m.group(2)
    m = re.search(r'(ROONRC\.[A-Z0-9\.]+)', tt)
    if m:
        d["EUID"] = m.group(1)
    js = re.findall(r'Jude[țţt]\s*:?\s*([A-ZĂÂÎȘȚŞŢ][A-Za-zĂÂÎȘȚăâîșțŞşŢţ\-]+)', tt)
    if js:
        d["judet"] = js[-1]

    # --- Dosar ONRC (dacă textul conține și print screen-ul / OCR-ul paginii)
    # Rândul din tabel: "1  2885801  21.07.2026  Mentiune" -> nr. cerere + data
    m = re.search(r'\b(\d{6,8})\s+(\d{1,2}[.,/]\d{1,2}[.,/]\d{2,4})\b', tt)
    if m:
        d["nr_cerere"] = m.group(1)
        d["data_cerere"] = norm_date(m.group(2))
    # rezerve pe etichete (tolerante la OCR: "Numi/Numar/Num... de inregistrare")
    if not d.get("nr_cerere"):
        m = re.search(r'(?:cererea nr|Num\w*\s+de\s+[îi]nregistrare|Nr\.?\s*de\s*[îi]nregistrare)\s*:?\.?\s*(\d{6,8})', tt, re.I)
        if m:
            d["nr_cerere"] = m.group(1)
    if not d.get("data_cerere"):
        m = re.search(r'Data\s+[îi]nregistr\w*\s*:?\s*(\d{1,2}[.,/-]\d{1,2}[.,/-]\d{2,4})', tt, re.I)
        if m:
            d["data_cerere"] = norm_date(m.group(1))
    m = re.search(r'Termen\s+de\s+solu[țţtc]?ionare\s*:?\s*(\d{1,2}[.,/-]\d{1,2}[.,/-]\d{2,4})', tt, re.I)
    if m:
        d["termen"] = norm_date(m.group(1))
    if d.get("data_cerere") and not d.get("data_amanare"):
        d["data_amanare"] = d["data_cerere"]

    return d


# =============================================================================
# EXTRAGERE TEXT DIN FIȘIERE (opțional)
# =============================================================================
def text_from_pdf(path):
    try:
        r = PdfReader(path)
        return "\n".join((p.extract_text() or "") for p in r.pages)
    except Exception:
        return ""

def text_from_docx(path):
    if not _HAS_DOCX:
        return ""
    try:
        doc = _docxmod.Document(path)
        parts = [p.text for p in doc.paragraphs]
        for tb in doc.tables:
            for row in tb.rows:
                for c in row.cells:
                    parts.append(c.text)
        return "\n".join(parts)
    except Exception:
        return ""

def text_from_image(path):
    if not _HAS_OCR:
        return "__NO_OCR__"
    try:
        img = Image.open(path)
        try:
            return pytesseract.image_to_string(img, lang="ron")
        except Exception:
            return pytesseract.image_to_string(img)
    except Exception:
        return ""

def text_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return text_from_pdf(path)
    if ext in (".docx", ".doc"):
        return text_from_docx(path)
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"):
        return text_from_image(path)
    if ext == ".txt":
        try:
            return open(path, encoding="utf-8", errors="ignore").read()
        except Exception:
            return ""
    return ""


# =============================================================================
# COMPLETARE PDF (formular oficial, editabil)
# =============================================================================
def genereaza_pdf(values, checks, out_path):
    reader = PdfReader(TEMPLATE_PDF)
    writer = PdfWriter()
    writer.append(reader)
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        pass
    text_fields = {PDF_TEXT[k]: v for k, v in values.items() if k in PDF_TEXT and v}
    if text_fields:
        writer.update_page_form_field_values(writer.pages[0], text_fields, auto_regenerate=False)
    for key, fld in PDF_CHECK.items():
        state = "/Yes" if checks.get(key) else "/Off"
        try:
            writer.update_page_form_field_values(writer.pages[0], {fld: state}, auto_regenerate=False)
        except Exception:
            pass
    with open(out_path, "wb") as fh:
        writer.write(fh)


# =============================================================================
# GUI
# =============================================================================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

if _HAS_DND:
    class _Root(ctk.CTk, TkinterDnD.DnDWrapper):
        pass
else:
    class _Root(ctk.CTk):
        pass


class App(_Root):
    def __init__(self):
        super().__init__()
        if _HAS_DND:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:
                pass
        self.title("Cerere Preschimbare Termen Soluționare — Formular nr. 5")
        self.geometry("1160x860")
        self.minsize(1000, 720)
        self.configure(fg_color=COL_BG)
        ic = os.path.join(APP_DIR, "icon.ico")
        if os.path.exists(ic):
            try:
                self.iconbitmap(ic)
            except Exception:
                pass

        self.entries = {}
        self.check_vars = {}

        self._build_header()
        self._build_body()
        self._build_footer()

        if _HAS_DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    # ---- header -----------------------------------------------------------
    def _build_header(self):
        h = ctk.CTkFrame(self, height=74, fg_color=COL_HEADER, corner_radius=0)
        h.pack(fill="x")
        left = ctk.CTkFrame(h, fg_color="transparent")
        left.pack(side="left", padx=18, pady=10)
        ctk.CTkLabel(left, text="📄  Cerere Preschimbare Termen Soluționare",
                     font=("Segoe UI", 21, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(left, text="Formular nr. 5 ONRC — model nou · completare automată din textul de identificare",
                     font=("Segoe UI", 12), text_color="#a9c2dd").pack(anchor="w")
        ctk.CTkLabel(h, text="ONRC  ", font=("Segoe UI", 15, "bold"),
                     text_color="#dbeafe").pack(side="right", padx=18)

    # ---- body -------------------------------------------------------------
    def _build_body(self):
        tabs = ctk.CTkTabview(self, fg_color=COL_WHITE,
                              segmented_button_selected_color=COL_BLUE)
        tabs.pack(fill="both", expand=True, padx=10, pady=8)
        self.tab_paste = tabs.add("1. Paste rapid / Import")
        self.tab_form = tabs.add("2. Date cerere")
        self._build_tab_paste()
        self._build_tab_form()

    def _build_tab_paste(self):
        f = self.tab_paste
        ctk.CTkLabel(f, text="Lipește aici textul de identificare (fraza cu numele, CNP, seria/nr., firma) "
                             "și, dacă vrei, textul din print screen-ul ONRC:",
                     font=("Segoe UI", 13), text_color=COL_HEADER, wraplength=1050,
                     justify="left").pack(anchor="w", padx=12, pady=(12, 2))
        ctk.CTkLabel(f, text="Exemplu: DEHELEAN SILVIA , cetatean roman, nascuta la data de 22,05,70 in Simeria "
                             "jud HD, domiciliata in ... seria XD nr 061488 /SPCLEP Deva/13,05,22-03,08,2031 "
                             "CNP ... UNITED BETON S.R.L. J..., cui ... Județ Timiș",
                     font=("Segoe UI", 10), text_color="#64748b", wraplength=1050,
                     justify="left").pack(anchor="w", padx=12, pady=(0, 6))
        self.txt = ctk.CTkTextbox(f, height=340, fg_color=COL_WHITE,
                                  border_color=COL_BORDER, border_width=1, font=("Segoe UI", 12))
        self.txt.pack(fill="both", expand=True, padx=12, pady=6)

        bar = ctk.CTkFrame(f, fg_color="transparent")
        bar.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bar, text="📂  Încarcă fișier (PDF / Word / imagine)", command=self._browse_file,
                      fg_color=COL_HEADER, hover_color="#25506f").pack(side="left")
        ctk.CTkButton(bar, text="⚡  EXTRAGE DATE AUTOMAT", command=self._extract,
                      fg_color=COL_BLUE, hover_color="#1d4fd8",
                      font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkButton(bar, text="🧹 Golește", command=lambda: self.txt.delete("1.0", "end"),
                      fg_color="#64748b", hover_color="#475569", width=90).pack(side="left")

        note = "Drag & drop: " + ("ACTIV" if _HAS_DND else "indisponibil") \
               + "   |   OCR imagine: " + ("ACTIV" if _HAS_OCR else "indisponibil (instalează Tesseract)")
        ctk.CTkLabel(f, text=note, font=("Segoe UI", 11), text_color="#64748b").pack(anchor="w", padx=12)

    def _build_tab_form(self):
        f = self.tab_form
        legend = ctk.CTkFrame(f, fg_color=COL_YELLOW, corner_radius=6)
        legend.pack(fill="x", padx=8, pady=(8, 0))
        ctk.CTkLabel(legend, text="🟡  Câmpurile pe fundal galben au fost completate automat — verifică-le.",
                     font=("Segoe UI", 12), text_color=COL_HEADER, anchor="w").pack(fill="x", padx=10, pady=6)

        sf = ctk.CTkScrollableFrame(f, fg_color=COL_WHITE)
        sf.pack(fill="both", expand=True, padx=6, pady=6)

        # checkbox-uri
        self._section(sf, "Tip cerere (bifează / debifează după caz)")
        cbf = ctk.CTkFrame(sf, fg_color="transparent")
        cbf.pack(fill="x", padx=14, pady=4)
        defaults = {"cb_completare_II": True, "cb_preschimbare_II": True,
                    "cb_audienta_II": False, "cb_renuntare_II": False,
                    "cb_completare_V": True, "cb_preschimbare_VI": True}
        labels = {"cb_completare_II": "Completare înscrisuri (secț. II)",
                  "cb_preschimbare_II": "Preschimbare termen (secț. II)",
                  "cb_audienta_II": "Audiență publică (secț. II)",
                  "cb_renuntare_II": "Renunțare la soluționare (secț. II)",
                  "cb_completare_V": "Completare înscrisuri (secț. V)",
                  "cb_preschimbare_VI": "Preschimbare termen (secț. VI)"}
        for i, key in enumerate(labels):
            var = ctk.BooleanVar(value=defaults[key])
            self.check_vars[key] = var
            ctk.CTkCheckBox(cbf, text=labels[key], variable=var, onvalue=True, offvalue=False,
                            fg_color=COL_BLUE).grid(row=i // 2, column=i % 2, sticky="w", padx=8, pady=4)

        self._section(sf, "Societate")
        self._row(sf, [("denumire", "Denumire firmă", 2), ("J", "Nr. ordine RC (J...)", 1)])
        self._row(sf, [("CUI", "CUI", 1), ("EUID", "EUID", 1), ("judet", "Județ sediu → Tribunal", 1)])

        self._section(sf, "Persoană (asociat / administrator)")
        self._row(sf, [("nume", "Nume", 1), ("prenume", "Prenume", 1), ("cnp", "CNP", 1)])
        self._row(sf, [("domiciliu", "Domiciliu (complet)", 3)])
        self._row(sf, [("calitate", "În calitate de", 1), ("cetatenie", "Cetățenia", 1), ("tara", "Țara", 1)])

        self._section(sf, "Act de identitate")
        self._row(sf, [("act_ident", "Act", 1), ("seria", "Seria", 1), ("nr_act", "Nr.", 1)])
        self._row(sf, [("emis_de", "Emis de", 1), ("data_emitere", "Data emiterii", 1), ("valabil", "Valabil până la", 1)])
        self._row(sf, [("data_nasterii", "Data nașterii", 1), ("nasc_loc", "Loc naștere", 1), ("nasc_judet", "Sector/județ naștere", 1)])

        self._section(sf, "Dosar ONRC (din print screen)")
        self._row(sf, [("nr_cerere", "Nr. cerere", 1), ("data_cerere", "Data cererii", 1)])
        self._row(sf, [("data_amanare", "Data încheierii de amânare", 1), ("termen", "Cu termen la (soluționare)", 1)])

        self._section(sf, "Completare înscrisuri (secț. V) — max. 8 rânduri")
        for i in range(1, 9):
            self._row(sf, [(f"inscris{i}", f"{i}.", 3)])

        self._section(sf, "Motivul preschimbării (se completează MANUAL)")
        for i in range(1, 4):
            self._row(sf, [(f"motiv{i}", f"Motiv rând {i}", 3)])

        self._section(sf, "Semnare")
        self._row(sf, [("semnat_nume", "Nume și prenume (semnatar)", 2), ("semnat_data", "Data", 1)])
        self.entries["semnat_data"].insert(0, self._today())
        self.entries["calitate"].insert(0, "administrator")

        ctk.CTkButton(f, text="✅  GENEREAZĂ CEREREA (PDF)", command=self._generate,
                      fg_color="#16a34a", hover_color="#15803d", height=48,
                      font=("Segoe UI", 15, "bold")).pack(fill="x", padx=10, pady=(4, 10))

    def _section(self, parent, title):
        ctk.CTkLabel(parent, text=title, font=("Segoe UI", 13, "bold"),
                     text_color=COL_HEADER, fg_color=COL_LIGHT, corner_radius=6,
                     anchor="w").pack(fill="x", padx=8, pady=(12, 2))

    def _row(self, parent, cols):
        r = ctk.CTkFrame(parent, fg_color="transparent")
        r.pack(fill="x", padx=10, pady=2)
        total = sum(w for _, _, w in cols)
        for c in range(total):
            r.grid_columnconfigure(c, weight=1)
        pos = 0
        for key, label, w in cols:
            cell = ctk.CTkFrame(r, fg_color="transparent")
            cell.grid(row=0, column=pos, columnspan=w, sticky="ew", padx=4)
            ctk.CTkLabel(cell, text=label, font=("Segoe UI", 11), text_color="#475569", anchor="w").pack(fill="x")
            e = ctk.CTkEntry(cell, fg_color=COL_WHITE, border_color=COL_BORDER, border_width=1)
            e.pack(fill="x")
            self.entries[key] = e
            pos += w

    # ---- footer -----------------------------------------------------------
    def _build_footer(self):
        footer = ctk.CTkFrame(self, height=30, fg_color=COL_BORDER, corner_radius=0)
        footer.pack(fill="x")
        self.status = ctk.CTkLabel(footer, text="Gata.", font=("Segoe UI", 11), text_color="#334155")
        self.status.pack(side="left", padx=12)

    def _set_status(self, msg):
        self.status.configure(text=msg)
        self.update_idletasks()

    # ---- acțiuni ----------------------------------------------------------
    def _today(self):
        n = datetime.date.today()
        return f"{n.day:02d}.{LUNI[n.month]}.{n.year}"

    def _set_entry(self, key, val, highlight=True):
        if not val:
            return
        e = self.entries.get(key)
        if e is None:
            return
        e.delete(0, "end")
        e.insert(0, val)
        if highlight:
            e.configure(fg_color=COL_YELLOW)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Alege document",
            filetypes=[("Documente", "*.pdf *.docx *.doc *.png *.jpg *.jpeg *.tif *.tiff *.txt"),
                       ("Toate", "*.*")])
        if path:
            self._ingest_file(path)

    def _on_drop(self, event):
        for p in re.findall(r'\{[^}]*\}|\S+', event.data):
            p = p.strip("{}")
            if os.path.isfile(p):
                self._ingest_file(p)

    def _ingest_file(self, path):
        self._set_status(f"Citesc {os.path.basename(path)} ...")
        txt = text_from_file(path)
        if txt == "__NO_OCR__":
            messagebox.showwarning("OCR indisponibil",
                "Pentru imagini este nevoie de Tesseract OCR. Instalează-l sau lipește textul manual.")
            self._set_status("OCR indisponibil.")
            return
        if not txt.strip():
            messagebox.showwarning("Fără text",
                "Nu am putut extrage text (posibil PDF scanat). Lipește textul manual în casetă.")
            self._set_status("Fără text extras.")
            return
        cur = self.txt.get("1.0", "end").strip()
        self.txt.insert("end", ("\n\n" if cur else "") + txt)
        self._set_status(f"Am adăugat textul din {os.path.basename(path)}. Apasă EXTRAGE DATE AUTOMAT.")

    def _extract(self):
        t = self.txt.get("1.0", "end")
        if not t.strip():
            messagebox.showinfo("Gol", "Lipește mai întâi text sau încarcă un fișier.")
            return
        data = extrage(t)
        for k, v in data.items():
            self._set_entry(k, v)
        # semnatar din nume+prenume dacă există
        nume = self.entries["nume"].get().strip()
        pren = self.entries["prenume"].get().strip()
        if nume or pren:
            self._set_entry("semnat_nume", (nume + " " + pren).strip())
        self._set_status("Date extrase. Verifică câmpurile galbene, apoi generează.")

    def _collect(self):
        v = {k: self.entries[k].get().strip() for k in self.entries}
        firma_txt = ""
        if v.get("denumire"):
            firma_txt = v["denumire"]
            extra = []
            if v.get("J"):
                extra.append(v["J"])
            if v.get("CUI"):
                extra.append("cui " + v["CUI"])
            if extra:
                firma_txt += " " + ", ".join(extra)

        out = dict(v)
        out["cerere_nr_data"] = (f'{v.get("nr_cerere","")}-{v.get("data_cerere","")}').strip("-")
        out["tribunal"] = v.get("judet", "")
        out["reprez_firma"] = firma_txt
        out["pt_firma"] = firma_txt
        out["gdpr_data"] = v.get("semnat_data", "")

        checks = {k: self.check_vars[k].get() for k in self.check_vars}
        denum = v.get("denumire") or "Firma"
        return out, checks, denum

    def _generate(self):
        if not os.path.exists(TEMPLATE_PDF):
            messagebox.showerror("Lipsă șablon", f"Nu găsesc șablonul:\n{TEMPLATE_PDF}")
            return
        values, checks, denum = self._collect()
        self._set_status("Generez cererea ...")
        threading.Thread(target=self._do_generate, args=(values, checks, denum), daemon=True).start()

    def _do_generate(self, values, checks, denum):
        try:
            slug = re.sub(r'[^\w]', '_', denum)[:40].strip("_") or "Firma"
            folder = os.path.join(OUTPUT_BASE, slug)
            os.makedirs(folder, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(folder, f"Cerere_preschimbare_{slug}_{stamp}.pdf")
            genereaza_pdf(values, checks, out_path)
            self._set_status(f"✅ Generat: {out_path}")
            for p in (folder, out_path):
                try:
                    os.startfile(p)
                except Exception:
                    pass
            self.after(0, lambda: messagebox.showinfo(
                "Gata",
                "Cererea a fost generată.\n\nFormularul rămâne EDITABIL: poți ajusta bifele și scrie "
                "motivul și în cititorul PDF, apoi aplici semnătura digitală.\n\n" + out_path))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Eroare", str(e)))
            self._set_status("Eroare la generare.")


if __name__ == "__main__":
    App().mainloop()
