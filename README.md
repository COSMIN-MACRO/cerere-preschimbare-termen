# Cerere Preschimbare Termen Soluționare — Formular nr. 5 ONRC

Aplicație desktop (Windows, Python + customtkinter) care completează automat **formularul oficial ONRC „Formular nr. 5" (model nou)** — cererea de preschimbare a termenului de soluționare / completare înscrisuri — pornind de la textul de identificare al persoanei și de la print screen-ul dosarului de pe portalul ONRC.

Datele se extrag local (fără internet), se completează în formularul PDF oficial (care rămâne **editabil**), iar tu doar verifici, bifezi ce dorești și scrii motivul. Ideal pentru consultanți / firme care depun frecvent astfel de cereri.

---

## Caracteristici

- **Extragere automată** din text lipit (frază de identificare de tipul „NUME PRENUME, cetatean roman, nascut... CNP... FIRMA S.R.L. J..., cui...").
- **Import fișiere**: PDF, Word (.docx) și imagini; **drag & drop** peste fereastră.
- **OCR local** (opțional, prin Tesseract) pentru citirea numărului de dosar direct dintr-un **print screen** ONRC.
- Completează **formularul PDF oficial**, păstrându-l editabil (poți ajusta bifele / motivul și în cititorul PDF, apoi aplici semnătura digitală).
- Checkbox-uri pentru tipul cererii (completare / preschimbare / audiență / renunțare), cu valori implicite rezonabile.
- Câmpurile completate automat sunt evidențiate cu **galben** pentru verificare rapidă.
- Calitatea semnatarului este implicit **administrator**.
- Salvează cererea generată în `OUTPUT/<Firma>/` și deschide folderul automat.

---

## Cerințe

- **Windows** cu **Python 3.10+** (la instalare bifează „Add Python to PATH").
- Module Python (vezi `requirements.txt`): `customtkinter`, `python-docx`, `pypdf`, `pillow`, `pywin32`, `tkinterdnd2`, `pytesseract`.
- **Opțional**, doar pentru OCR pe imagini: [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki). Fără el aplicația funcționează normal — poți lipi textul în locul imaginii.

---

## Instalare

```bash
git clone https://github.com/COSMIN-MACRO/cerere-preschimbare-termen.git
cd cerere-preschimbare-termen
pip install -r requirements.txt
python app.py
```

Pentru utilizatorii Windows fără experiență există și varianta cu dublu-click:

1. `INSTALARE_module.bat` — instalează modulele (o singură dată).
2. `Porneste_Aplicatia.bat` — pornește aplicația.
3. `Creeaza_Shortcut_Desktop.bat` — pune o scurtătură pe Desktop.

Ghidul complet pas cu pas: **`INSTRUCTIUNI_INSTALARE.pdf`**.

---

## Utilizare

1. Tab **„1. Paste rapid / Import"** — lipește fraza de identificare (și, opțional, textul/print screen-ul dosarului ONRC), sau trage fișierele peste fereastră.
2. Apasă **„EXTRAGE DATE AUTOMAT"**.
3. Tab **„2. Date cerere"** — verifică câmpurile galbene, bifează/debifează tipul cererii, scrie **motivul**.
4. Apasă **„GENEREAZĂ CEREREA"**. PDF-ul apare în `OUTPUT/<Firma>/` și se deschide automat.

---

## Structura proiectului

```
cerere-preschimbare-termen/
├── app.py                       # aplicația (GUI + extragere + completare PDF)
├── Formular5_nou_blank.pdf      # șablonul oficial ONRC (AcroForm) — necesar
├── icon.ico                     # iconița aplicației
├── make_shortcut.py             # creare scurtătură pe Desktop
├── INSTALARE_module.bat         # instalare module (Windows)
├── Porneste_Aplicatia.bat       # lansare (Windows)
├── Creeaza_Shortcut_Desktop.bat # scurtătură Desktop (Windows)
├── INSTRUCTIUNI_INSTALARE.pdf   # ghid pas cu pas
├── CITESTE_MA.txt               # rezumat instalare/utilizare
├── requirements.txt
├── .gitignore
└── LICENSE
```

---

## Note

- **Diacritice**: formularul folosește `NeedAppearances`, deci diacriticele (ă, â, î, ș, ț) apar corect în Adobe Reader / Foxit / Edge / browser. Unele cititoare foarte vechi le pot afișa greșit — în acest caz deschide PDF-ul în Adobe Reader.
- Cererea generată rămâne **editabilă**: poți completa/modifica direct în PDF înainte de semnare.
- Aplicația rulează **direct cu Python** (`python app.py`), fără PyInstaller.

## Confidențialitate

Toată prelucrarea are loc **local**, pe calculatorul tău. Aplicația nu trimite date către niciun server.

## Disclaimer

Proiect neoficial, fără legătură cu ONRC / Ministerul Justiției. Șablonul „Formular nr. 5" aparține ONRC. Verifică întotdeauna corectitudinea datelor înainte de depunere.

## Licență

Distribuit sub licența [MIT](LICENSE).
