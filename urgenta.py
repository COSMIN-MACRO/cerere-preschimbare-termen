# -*- coding: utf-8 -*-
"""Set urgenta ORC: Formular 5 (preschimbare termen) + cerere urgenta (DOCX+PDF) + raport client.
Utilizare: python3 urgenta.py config.json   (vezi SKILL.md pentru campuri)"""
import sys, os, json, re, datetime, unicodedata, subprocess, shutil
from pypdf import PdfReader, PdfWriter
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL

EMAIL = "transmiterecereri@tm.onrc.ro"
JUD = {"01":"ALBA","02":"ARAD","03":"ARGEȘ","04":"BACĂU","05":"BIHOR","06":"BISTRIȚA-NĂSĂUD","07":"BOTOȘANI",
"08":"BRAȘOV","09":"BRĂILA","10":"BUZĂU","11":"CARAȘ-SEVERIN","12":"CLUJ","13":"CONSTANȚA","14":"COVASNA",
"15":"DÂMBOVIȚA","16":"DOLJ","17":"GALAȚI","18":"GORJ","19":"HARGHITA","20":"HUNEDOARA","21":"IALOMIȚA",
"22":"IAȘI","23":"ILFOV","24":"MARAMUREȘ","25":"MEHEDINȚI","26":"MUREȘ","27":"NEAMȚ","28":"OLT","29":"PRAHOVA",
"30":"SATU MARE","31":"SĂLAJ","32":"SIBIU","33":"SUCEAVA","34":"TELEORMAN","35":"TIMIȘ","36":"TULCEA",
"37":"VASLUI","38":"VÂLCEA","39":"VRANCEA","40":"BUCUREȘTI","51":"CĂLĂRAȘI","52":"GIURGIU"}
F5 = {"tribunal":"1","cerere_nr_data":"1_2","nume":"2","prenume":"3","cnp":"4","domiciliu":"10","judet_dom":"11",
"cetatenie":"12","nasc_loc":"14","tara":"15","nasc_judet":"16","data_nasterii":"18","act":"19","seria":"20",
"emis_de":"21","data_emitere":"22","valabil":"24","reprez_firma":"25","nr_act":"27","calitate":"30","pt_firma":"32",
"termen":"22052026-57","data_amanare":"22052026-56","motiv1":"22052026-53","motiv2":"22052026-54","motiv3":"22052026-55",
"semnat_nume":"59","semnat_data":"60","gdpr_data":"60_2"}

def nd(s): return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
def pct(x): return x if x.endswith(".") else x + "."
def d2(s): return datetime.datetime.strptime(s, "%d.%m.%Y").date()

def judet_din_J(J):
    J = (J or "").replace(" ", "").upper()
    m = re.match(r"^[JFC](\d{4})(\d{6})(\d{2})\d$", J)          # format nou J2024012345357
    if m: return JUD.get(m.group(3))
    m = re.match(r"^[JFC](\d{2})/", J)                         # format vechi J35/123/2020
    if m: return JUD.get(m.group(1))

def verificari(c):
    w = []; p = c["p"]
    cnp = p.get("cnp", "")
    if len(cnp) == 13:
        sec = {"1":"19","2":"19","5":"20","6":"20"}.get(cnp[0], "19")
        dn = f"{cnp[5:7]}.{cnp[3:5]}.{sec}{cnp[1:3]}"
        if p.get("data_nasterii") and p["data_nasterii"] != dn:
            w.append(f"Data nașterii {p['data_nasterii']} nu corespunde CNP-ului ({dn}).")
        sx = "M" if cnp[0] in "1357" else "F"
        if p.get("sex") and p["sex"] != sx: w.append("Sexul nu corespunde CNP-ului.")
    elif cnp: w.append(f"CNP-ul {cnp} nu are 13 cifre.")
    jj = judet_din_J(c.get("J"))
    if jj and jj != c["orc"].upper(): w.append(f"Numărul J indică județul {jj}, dar cererea e adresată ORC {c['orc']}.")
    if not c.get("sediu"): w.append("Lipsește sediul — am lăsat puncte de completat.")
    try:
        if d2(c["termen"]) < d2(c["data_doc"]) and not c.get("amanat"):
            w.append(f"Termenul {c['termen']} a trecut. Dacă dosarul e amânat, completează data încheierii de amânare în Formularul 5.")
    except Exception: pass
    if p.get("valabil"):
        try:
            if d2(p["valabil"]) < d2(c["data_doc"]): w.append(f"Actul de identitate a expirat la {p['valabil']}.")
        except Exception: pass
    return w

# ---------------- Formular 5 ----------------
def formular5(c, tpl, out):
    p = c["p"]; pfa = c["forma"] == "pfa"; inm = c["tip"] == "inmatriculare"
    firma_id = c["firma"] + (" (in curs de inregistrare)" if inm and pfa else " (in curs de inmatriculare)" if inm else
               "".join(f", {x}" for x in [c.get("J"), ("CUI " + c["CUI"]) if c.get("CUI") else None] if x))
    ptf = c["firma"] + (f" - cerere ref. nr. {c['ref']}" if c.get("ref") else "")
    m = c["motiv_f5"] + ["", "", ""]
    v = {"tribunal": c["orc"].title(), "cerere_nr_data": f"{c['nr']}-{c['data']}", "nume": p["nume"], "prenume": p["prenume"],
         "cnp": p.get("cnp",""), "domiciliu": p.get("domiciliu_scurt") or p.get("domiciliu",""), "judet_dom": p.get("judet_dom",""),
         "cetatenie": p.get("cetatenie","Romana"), "tara": "Romania", "nasc_loc": p.get("nasc_loc",""), "nasc_judet": p.get("nasc_judet",""),
         "data_nasterii": p.get("data_nasterii",""), "act": p.get("act","CI"), "seria": p.get("seria",""), "nr_act": p.get("nr_act",""),
         "emis_de": p.get("emis_de",""), "data_emitere": p.get("data_emitere",""), "valabil": p.get("valabil",""),
         "reprez_firma": firma_id, "pt_firma": ptf, "calitate": c["calitate"] + (" PFA" if pfa and c["calitate"]=="titular" else ""),
         "termen": c["termen"] if not c.get("amanat") else "", "data_amanare": c.get("data_amanare",""),
         "motiv1": m[0], "motiv2": m[1], "motiv3": m[2], "semnat_nume": f"{p['nume']} {p['prenume']}",
         "semnat_data": c["data_doc"], "gdpr_data": c["data_doc"]}
    tf = {F5[k]: nd(x) for k, x in v.items() if x}
    w = PdfWriter(); w.append(PdfReader(tpl)); w.set_need_appearances_writer(True)
    w.update_page_form_field_values(w.pages[0], tf, auto_regenerate=False)
    comp = bool(c.get("inscrisuri"))
    for i, t in enumerate(c.get("inscrisuri", [])[:8]): w.update_page_form_field_values(w.pages[0], {str(45+i): nd(t)}, auto_regenerate=False)
    for f, on in {"Box1": comp, "Box2": True, "Box3": False, "Box3_2": False, "Box4": comp, "Box5": True}.items():
        w.update_page_form_field_values(w.pages[0], {f: "/Yes" if on else "/Off"}, auto_regenerate=False)
    with open(out, "wb") as fh: w.write(fh)

# ---------------- Cerere urgenta ----------------
def para(d, segs, al=AL.JUSTIFY, after=6, indent=True, bold=False):
    q = d.add_paragraph(); q.alignment = al; f = q.paragraph_format
    f.space_after = Pt(after); f.space_before = Pt(0)
    if indent: f.first_line_indent = Cm(1.25)
    for t, b in (segs if isinstance(segs, list) else [(segs, bold)]):
        r = q.add_run(t); r.bold = b
    return q

def cerere(c, out):
    p = c["p"]; pfa = c["forma"] == "pfa"; inm = c["tip"] == "inmatriculare"; F = (p.get("sex") or ("F" if p.get("cnp","1")[0] in "2468" else "M")) == "F"
    obj = {"inmatriculare": "cererii de înregistrare (înmatriculare)", "radiere": "cererii de radiere"}.get(c["tip"], "cererii de înscriere mențiuni")
    d = docx.Document(); s = d.sections[0]
    s.left_margin = s.right_margin = Cm(2.5); s.top_margin = Cm(2); s.bottom_margin = Cm(2.54)
    st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    for t, a in [("CĂTRE,", 0), (f"OFICIUL REGISTRULUI COMERȚULUI DE PE LÂNGĂ TRIBUNALUL {c['orc'].upper()}", 0),
                 ("În atenția Domnului / Doamnei Director", 24), ("CERERE", 0),
                 (f"privind soluționarea în regim de urgență a {obj}", 24)]:
        para(d, t, AL.CENTER, a, False, True)
    act = (f"carte electronică de identitate nr. {p['nr_act']}, emisă" if p.get("act") == "CEI"
           else f"C.I. seria {p.get('seria','')} nr. {p.get('nr_act','')}, eliberată")
    ent = (f"titular al " if pfa else f"{c['calitate']} al societății " + ("în curs de înmatriculare " if inm else ""))
    tail = "" if inm else "".join(f", {x}" for x in [c.get("J"), ("CUI " + c["CUI"]) if c.get("CUI") else None] if x)
    firma_txt = c["firma"] + (" (în curs de înregistrare)" if inm and pfa else "")
    sediu = c.get("sediu") or "……………………………………………………"
    para(d, [(f"Subsemnat{'a' if F else 'ul'} ", False), (f"{p['nume']} {p['prenume']}", True),
             (f", cetățean român, născut{'ă' if F else ''} la data de {p['data_nasterii']} în {p['nascut_in']}, "
              f"domiciliat{'ă' if F else ''} în {p['domiciliu']}, identificat{'ă' if F else ''} cu {act} de {p['emis_de']} "
              f"la data de {p['data_emitere']}, CNP {p['cnp']}, în calitate de {ent}", False),
             (firma_txt, True), (f"{tail}, cu sediul {'profesional' if pfa else 'social'} în {sediu},", False)])
    ref = f" (cerere de referință nr. {c['ref']})" if c.get("ref") else ""
    para(d, [("vă rog să binevoiți a dispune soluționarea în regim de urgență a ", False),
             ({"inmatriculare": "cererii de înregistrare (autorizare și înmatriculare)", "radiere": "cererii de radiere"}
              .get(c["tip"], "cererii de înscriere mențiuni"), inm),
             (" înregistrată sub nr. ", False), (f"{c['nr']} din data de {c['data']}", True),
             (f"{ref}" + (f", {c['obiect']}" if c.get("obiect") else "") + ".", False)])
    para(d, "Solicitarea este motivată de faptul că " + c["motiv_cerere"])
    para(d, "Menționez că dosarul a fost depus complet, cu toate actele doveditoare prevăzute de lege.")
    para(d, "Vă mulțumesc pentru înțelegere și sprijin.", after=30)
    t = d.add_table(rows=1, cols=2); a, b = t.rows[0].cells
    a.paragraphs[0].add_run(f"Data: {c['data_doc']}")
    q = b.paragraphs[0]; q.alignment = AL.CENTER
    q.add_run("Semnătura,").bold = True
    q.add_run("\n" + c["firma"] + "\n" + ("Titular," if pfa else f"Prin {c['calitate']},") + f"\n{p['nume']} {p['prenume']}\n\n……………………………")
    d.save(out)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(out), out],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)

# ---------------- Raport client ----------------
def raport(c):
    pfa = c["forma"] == "pfa"; inm = c["tip"] == "inmatriculare"; p = c["p"]; nume = f"{p['nume']} {p['prenume']}"
    obj = {"inmatriculare": "cererii de înregistrare", "radiere": "cererii de radiere"}.get(c["tip"], "cererii de înscriere mențiuni")
    ref = c["firma"] if pfa else "societatea " + c["firma"]
    subj = f"Cerere de soluționare în regim de urgență – {c['firma']}" + (" (în curs de înmatriculare)" if inm else "") + f", cererea nr. {c['nr']}/{c['data']}"
    return "\n".join([
        f"Am pregătit cererea de urgență și cererea de preschimbare a termenului pentru {pct(c['firma'])} Ambele sunt pe numele {nume}, {c['calitate']}, cu data de {c['data_doc']}.", "",
        "Vă rugăm să semnați documentele ori cu semnătură electronică, ori chiar și fizic (PDF-uri în telefon sau listate, semnate, iar apoi le scanați înapoi în PDF). Când sunt semnate, le trimiteți ca atașament în e-mail, după instrucțiunile de mai jos.", "",
        " " * 12 + "**Instrucțiuni**", "",
        "Vă rog să trimiteți e-mailul de pe adresa personală sau de pe adresa de e-mail a firmei", "",
        f"**către adresa**: {EMAIL}", "",
        f"**cu subiectul**: {subj}", "", "**Text pentru corpul e-mailului:**", "", "Stimată Doamnă Director,", "",
        f"Vă transmitem atașat cererea privind soluționarea în regim de urgență a {obj} nr. {c['nr']} din {c['data']} referitoare la {pct(ref)} Atașăm și cererea de preschimbare a termenului de soluționare.", "",
        f"Urgența este motivată de faptul că {c['motiv_mail']}", "", "Vă mulțumim pentru sprijin.", "", "Cu stimă,  ", nume])

def raport_docx(txt, out):
    d = docx.Document(); st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
    st.element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Calibri")
    for bloc in txt.split("\n\n"):
        q = d.add_paragraph(); q.paragraph_format.space_after = Pt(10); q.paragraph_format.line_spacing = 1.15
        if bloc.startswith("\u00a0"): q.paragraph_format.left_indent = Cm(1.25); bloc = bloc.strip("\u00a0")
        for i, linie in enumerate(bloc.replace("  \n", "\n").split("\n")):
            if i: q.add_run().add_break()
            for j, part in enumerate(linie.split("**")):
                if part: q.add_run(part).bold = bool(j % 2)
    d.save(out)

def main(cfgp):
    c = json.load(open(cfgp, encoding="utf-8"))
    c.setdefault("data_doc", datetime.date.today().strftime("%d.%m.%Y")); c.setdefault("orc", "TIMIȘ")
    out = c.get("out", "."); os.makedirs(out, exist_ok=True)
    tag = re.sub(r"[^A-Za-z0-9]+", "_", nd(c["firma"]).replace("S.R.L.", "SRL")).strip("_")
    tpl = c.get("template") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "Formular5_nou_blank.pdf")
    f5 = f"{out}/Cerere_preschimbare_termen_{tag}.pdf"; cu = f"{out}/Cerere_urgenta_ORC_{tag}.docx"; rp = f"{out}/Raport_client_{tag}.docx"
    formular5(c, tpl, f5); cerere(c, cu)
    txt = raport(c); raport_docx(txt, rp)
    pages = len(PdfReader(cu[:-5] + ".pdf").pages)
    w = verificari(c) + ([f"Cererea de urgență are {pages} pagini (de regulă 1)."] if pages > 1 else [])
    print("FISIERE:", f5, cu, cu[:-5] + ".pdf", rp, sep="\n  ")
    print("ATENTIONARI:" if w else "ATENTIONARI: niciuna", *w, sep="\n  - ")
    print("\n===== RAPORT CLIENT =====\n" + txt)

if __name__ == "__main__": main(sys.argv[1])
