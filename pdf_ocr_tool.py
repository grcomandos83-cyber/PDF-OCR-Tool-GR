"""
PDF OCR Tool GR - Επαγγελματικό Εργαλείο OCR για Ελληνικά & Αρχαία/Πολυτονικά Ελληνικά
Βασισμένο στο Tesseract OCR Engine με υποστήριξη:
- Πολυτονικών & Αρχαίων Ελληνικών (grc), Νέων Ελληνικών (ell), Αγγλικών (eng)
- High-res PDF Rendering (PyMuPDF), Searchable PDF, DOCX, TXT
- Προηγμένη προεπεξεργασία εικόνας & εξειδικευμένο φίλτρο καθαρισμού ελληνικού κειμένου
"""

import os
import sys
import shutil
import threading
import queue
import time
import subprocess
import re
import unicodedata

# Αυτόματος έλεγχος και εγκατάσταση εξαρτήσεων κατά την εκκίνηση (μόνο σε Python script mode)
def ensure_dependencies():
    if getattr(sys, 'frozen', False):
        return  # Όταν τρέχει ως συσκευασμένο EXE, όλες οι βιβλιοθήκες είναι ήδη ενσωματωμένες

    required_modules = {
        "pymupdf": "pymupdf",
        "PIL": "Pillow",
        "docx": "python-docx",
        "pytesseract": "pytesseract"
    }

    missing_packages = []
    for mod_name, pkg_name in required_modules.items():
        try:
            __import__(mod_name)
        except ImportError:
            missing_packages.append(pkg_name)

    if missing_packages:
        print(f"Εντοπίστηκαν βιβλιοθήκες που λείπουν: {', '.join(missing_packages)}")
        print("Αυτόματη εγκατάσταση μέσω διαχειριστή πακέτων...")
        
        installed = False
        if shutil.which("uv"):
            try:
                cmd = ["uv", "pip", "install", "--python", sys.executable] + missing_packages
                subprocess.check_call(cmd)
                installed = True
            except Exception:
                pass

        if not installed:
            try:
                cmd = [sys.executable, "-m", "pip", "install", "--break-system-packages"] + missing_packages
                subprocess.check_call(cmd)
                installed = True
            except Exception:
                try:
                    cmd = [sys.executable, "-m", "pip", "install"] + missing_packages
                    subprocess.check_call(cmd)
                    installed = True
                except Exception as e:
                    print(f"Σφάλμα κατά την εγκατάσταση: {e}")

try:
    ensure_dependencies()
except Exception as e:
    print(f"Dependency check error: {e}")

# Εισαγωγή βιβλιοθηκών
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
except ImportError:
    Image = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    import docx
except ImportError:
    docx = None


def clean_greek_ocr_text(text: str) -> str:
    """
    Εξειδικευμένο φίλτρο μετα-επεξεργασίας (Post-processing) για διόρθωση
    τυπικών σφαλμάτων OCR σε ελληνικά, μονοτονικά και πολυτονικά κείμενα.
    """
    if not text:
        return ""

    # 1. Κανονικοποίηση Unicode (NFC)
    text = unicodedata.normalize("NFC", text)

    # 2. Διόρθωση συμβόλου micro (µ -> μ) και άλλων ειδικών συμβόλων
    text = text.replace('\u00B5', 'μ')
    text = text.replace('入', 'λ')  # Διόρθωση τυχόν παρασιτικών χαρακτήρων
    text = text.replace('△', 'Δ')

    # 3. Διόρθωση κλασικών σφαλμάτων σε συχνές λέξεις / μόρια
    text = re.sub(r'\b(?:eis|eig|sig|vig|rig)\b', 'εἰς', text, flags=re.IGNORECASE)
    text = re.sub(r'\bzal\b', 'καὶ', text)
    text = re.sub(r'\bzai\b', 'καὶ', text)
    text = re.sub(r'\b[kK]ui\b', 'Καὶ', text)
    text = re.sub(r'\bot\b', 'οἱ', text)
    text = re.sub(r"‘Oc\b", "Ὡς", text)
    text = re.sub(r"‘Yad\b", "Ὑπὸ", text)
    text = re.sub(r"\*Ey\b", "Ἐν", text)
    text = re.sub(r"\bOa\b", "θὰ", text)
    text = re.sub(r"\bTa\b", "τὰ", text)

    # 4. Διόρθωση αριθμού 6 που παρερμηνεύτηκε αντί για πνεύμα/άρθρο (Ὁ / ὁ)
    text = re.sub(r'(^|\s)6(?=[A-ZΑ-Ω])', r'\1Ὁ ', text)
    text = re.sub(r'(^|\s)6(?=\s+[Α-ΩA-Z])', r'\1Ὁ', text)
    text = re.sub(r'(^|\s)6(?=\s+[α-ωa-z])', r'\1ὁ', text)

    # 5. Διόρθωση τελικού » ή « ή > αντί για τελικό 'ν'
    greek_vowels = r'[αεηιουωάέήίόύώὰὲὴὶὸὺὼᾶῆῖῦῶἄἔἤἴὄὔὤἂἒἢἲὂὒὢἆἦἶὖὦᾄᾔᾤᾂᾒᾢᾆᾖᾦᾀᾐᾠᾁᾑᾡᾳῃῳ]'
    text = re.sub(rf'({greek_vowels})[»«>]+(?=[,\.;:!?\s\n\r]|$)', r'\1ν', text)

    # 6. Αντικατάσταση λατινικών ομογλύφων που παρεισφρέουν εντός ελληνικών λέξεων
    latin_to_greek_lower = {
        'a': 'α', 'b': 'β', 'c': 'σ', 'd': 'δ', 'e': 'ε', 'i': 'ι',
        'k': 'κ', 'n': 'η', 'o': 'ο', 'p': 'ρ', 't': 'τ', 'u': 'υ',
        'v': 'ν', 'w': 'ω', 'x': 'χ', 'y': 'υ'
    }
    latin_to_greek_upper = {
        'A': 'Α', 'B': 'Β', 'E': 'Ε', 'H': 'Η', 'I': 'Ι', 'K': 'Κ',
        'M': 'Μ', 'N': 'Ν', 'O': 'Ο', 'P': 'Ρ', 'T': 'Τ', 'X': 'Χ', 'Y': 'Υ', 'Z': 'Ζ'
    }

    def replace_latin_in_greek_word(match):
        word = match.group(0)
        has_greek = any('\u0370' <= c <= '\u03FF' or '\u1F00' <= c <= '\u1FFF' for c in word)
        if not has_greek:
            return word

        res = []
        for c in word:
            if c in latin_to_greek_lower:
                res.append(latin_to_greek_lower[c])
            elif c in latin_to_greek_upper:
                res.append(latin_to_greek_upper[c])
            else:
                res.append(c)
        return "".join(res)

    text = re.sub(r'[\w\u0370-\u03FF\u1F00-\u1FFF]+', replace_latin_in_greek_word, text)

    # 7. Διόρθωση εισαγωγικών
    text = text.replace("“", "«").replace("”", "»")

    return text


def find_tesseract_binary():
    """Αυτόματος εντοπισμός της εγκατάστασης του Tesseract OCR."""
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Tesseract-OCR\tesseract.exe"),
        shutil.which("tesseract"),
        shutil.which("tesseract.exe"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    return None


def get_available_tesseract_languages(tesseract_path):
    """Επιστρέφει τη λίστα των διαθέσιμων traineddata γλωσσών του Tesseract."""
    if not tesseract_path or not os.path.isfile(tesseract_path):
        return []
    
    tessdata_dir = os.path.join(os.path.dirname(tesseract_path), "tessdata")
    prefix_env = os.environ.get("TESSDATA_PREFIX")
    
    dirs_to_check = []
    if prefix_env and os.path.isdir(prefix_env):
        dirs_to_check.append(prefix_env)
    if os.path.isdir(tessdata_dir):
        dirs_to_check.append(tessdata_dir)
        
    langs = set()
    for d in dirs_to_check:
        try:
            for f in os.listdir(d):
                if f.endswith(".traineddata"):
                    langs.add(f[:-12])
        except Exception:
            pass
    return sorted(list(langs))


def parse_page_range(range_str, total_pages):
    """Μετατροπή συμβολοσειράς π.χ. '1-3, 5, 8-10' σε λίστα μηδενικής βάσης [0, 1, 2, 4, 7, 8, 9]."""
    if not range_str or range_str.strip().lower() in ("όλες", "ολες", "all", "*", ""):
        return list(range(total_pages))
    
    pages = set()
    parts = range_str.split(",")
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            sub = part.split("-")
            if len(sub) == 2:
                try:
                    start = int(sub[0].strip())
                    end = int(sub[1].strip())
                    for p in range(max(1, start), min(total_pages, end) + 1):
                        pages.add(p - 1)
                except ValueError:
                    pass
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    pages.add(p - 1)
            except ValueError:
                pass
    res = sorted(list(pages))
    return res if res else list(range(total_pages))


class OcrWorker(threading.Thread):
    def __init__(self, tesseract_cmd, input_path, output_base, lang, export_pdf, export_txt, 
                 export_docx, dpi, enhance_image, clean_text, page_range_str, msg_queue, cancel_event):
        super().__init__()
        self.tesseract_cmd = tesseract_cmd
        self.input_path = input_path
        self.output_base = output_base
        self.lang = lang
        self.export_pdf = export_pdf
        self.export_txt = export_txt
        self.export_docx = export_docx
        self.dpi = dpi
        self.enhance_image = enhance_image
        self.clean_text = clean_text
        self.page_range_str = page_range_str
        self.msg_queue = msg_queue
        self.cancel_event = cancel_event

    def preprocess_image(self, pil_img):
        """Προαιρετική βελτιστοποίηση εικόνας για μέγιστη ακρίβεια OCR."""
        if not self.enhance_image:
            return pil_img
        try:
            gray = pil_img.convert("L")
            if hasattr(ImageOps, "autocontrast"):
                gray = ImageOps.autocontrast(gray, cutoff=0.5)
            enhancer = ImageEnhance.Contrast(gray)
            enhanced = enhancer.enhance(1.25)
            filtered = enhanced.filter(ImageFilter.UnsharpMask(radius=1.2, percent=110, threshold=3))
            return filtered
        except Exception:
            return pil_img

    def run(self):
        try:
            if not pymupdf or not Image:
                raise Exception("Απαιτούνται οι βιβλιοθήκες pymupdf και Pillow.")

            if not pytesseract or not self.tesseract_cmd or not os.path.isfile(self.tesseract_cmd):
                raise Exception("Το Tesseract OCR δεν είναι διαθέσιμο ή δεν έχει οριστεί έγκυρη διαδρομή.")
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

            is_pdf = self.input_path.lower().endswith(".pdf")
            created_files = []
            extracted_pages_text = []

            tess_config = "--oem 1 --psm 3"

            if is_pdf:
                doc = pymupdf.open(self.input_path)
                total_doc_pages = len(doc)
                pages_to_process = parse_page_range(self.page_range_str, total_doc_pages)
                total_to_do = len(pages_to_process)

                if total_to_do == 0:
                    raise Exception("Δεν επιλέχθηκε καμία έγκυρη σελίδα προς επεξεργασία.")

                self.msg_queue.put(("STATUS", f"Επεξεργασία {total_to_do} σελίδων (Tesseract OCR)..."))

                pdf_writer = pymupdf.open() if self.export_pdf else None
                zoom = self.dpi / 72.0
                matrix = pymupdf.Matrix(zoom, zoom)

                for idx, page_num in enumerate(pages_to_process, 1):
                    if self.cancel_event.is_set():
                        self.msg_queue.put(("STATUS", "Η διαδικασία διακόπηκε από τον χρήστη."))
                        return

                    self.msg_queue.put(("PROGRESS", idx, total_to_do, f"Σελίδα {page_num + 1}/{total_doc_pages} ({idx}/{total_to_do})"))
                    
                    page = doc[page_num]
                    pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                    img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                    processed_img = self.preprocess_image(img)

                    # Tesseract OCR εκτέλεση
                    raw_text = pytesseract.image_to_string(processed_img, lang=self.lang, config=tess_config)
                    page_text = clean_greek_ocr_text(raw_text) if self.clean_text else raw_text

                    if self.export_pdf and pdf_writer is not None:
                        pdf_bytes = pytesseract.image_to_pdf_or_hocr(processed_img, extension='pdf', lang=self.lang, config=tess_config)
                        page_pdf_doc = pymupdf.open("pdf", pdf_bytes)
                        pdf_writer.insert_pdf(page_pdf_doc)
                        page_pdf_doc.close()

                    extracted_pages_text.append((page_num + 1, page_text))
                    self.msg_queue.put(("PAGE_TEXT", page_num + 1, page_text))

                doc.close()

                # Αποθήκευση Searchable PDF
                if self.export_pdf and pdf_writer is not None:
                    out_pdf = self.output_base + ".pdf"
                    pdf_writer.save(out_pdf)
                    pdf_writer.close()
                    created_files.append(out_pdf)

            else:
                # Επεξεργασία μεμονωμένης εικόνας (PNG, JPG, TIFF κ.λπ.)
                self.msg_queue.put(("STATUS", "Επεξεργασία εικόνας (Tesseract OCR)..."))
                self.msg_queue.put(("PROGRESS", 1, 1, "Ανάγνωση εικόνας (1/1)"))
                
                img = Image.open(self.input_path).convert("RGB")
                processed_img = self.preprocess_image(img)

                raw_text = pytesseract.image_to_string(processed_img, lang=self.lang, config=tess_config)
                page_text = clean_greek_ocr_text(raw_text) if self.clean_text else raw_text

                if self.export_pdf:
                    pdf_bytes = pytesseract.image_to_pdf_or_hocr(processed_img, extension='pdf', lang=self.lang, config=tess_config)
                    out_pdf = self.output_base + ".pdf"
                    with open(out_pdf, "wb") as fp:
                        fp.write(pdf_bytes)
                    created_files.append(out_pdf)

                extracted_pages_text.append((1, page_text))
                self.msg_queue.put(("PAGE_TEXT", 1, page_text))

            # Αποθήκευση TXT
            if self.export_txt and extracted_pages_text:
                out_txt = self.output_base + ".txt"
                with open(out_txt, "w", encoding="utf-8") as f:
                    for p_num, p_text in extracted_pages_text:
                        if len(extracted_pages_text) > 1:
                            f.write(f"\n--- [ Σελίδα {p_num} ] ---\n\n")
                        f.write(p_text.strip() + "\n")
                created_files.append(out_txt)

            # Αποθήκευση Word DOCX
            if self.export_docx and extracted_pages_text and docx is not None:
                out_docx = self.output_base + ".docx"
                doc_obj = docx.Document()
                doc_obj.add_heading("Αποτέλεσμα OCR", level=0)
                
                for p_num, p_text in extracted_pages_text:
                    if len(extracted_pages_text) > 1:
                        doc_obj.add_heading(f"Σελίδα {p_num}", level=2)
                    for paragraph in p_text.split("\n\n"):
                        clean_p = paragraph.strip()
                        if clean_p:
                            doc_obj.add_paragraph(clean_p)
                doc_obj.save(out_docx)
                created_files.append(out_docx)

            self.msg_queue.put(("SUCCESS", created_files))

        except Exception as e:
            self.msg_queue.put(("ERROR", str(e)))


class ModernOcrApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF OCR Tool GR - Επαγγελματικό Εργαλείο OCR (Tesseract Engine)")
        self.root.geometry("880x730")
        self.root.minsize(820, 650)
        
        # State variables
        self.input_file = tk.StringVar()
        self.output_base = tk.StringVar()
        
        self.tesseract_path = tk.StringVar(value=find_tesseract_binary() or "")
        self.available_tess_langs = get_available_tesseract_languages(self.tesseract_path.get())
        
        # Γλώσσες
        default_tess = "grc+ell+eng" if "grc" in self.available_tess_langs else "ell+eng"
        self.tess_lang_mode = tk.StringVar(value=default_tess)
        self.tess_custom_lang = tk.StringVar(value="")
        
        self.opt_pdf = tk.BooleanVar(value=True)
        self.opt_txt = tk.BooleanVar(value=True)
        self.opt_docx = tk.BooleanVar(value=False)
        self.opt_enhance = tk.BooleanVar(value=True)
        self.opt_clean_text = tk.BooleanVar(value=True)
        
        self.page_range_str = tk.StringVar(value="Όλες")
        self.dpi_val = tk.IntVar(value=300)
        
        self.msg_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.last_created_files = []

        self.apply_styles()
        self.build_ui()
        self.check_queue()

    def apply_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # Primary colors & fonts
        style.configure(".", font=("Segoe UI", 9))
        style.configure("Title.TLabel", font=("Segoe UI", 11, "bold"), foreground="#1a365d")
        style.configure("BannerOk.TLabel", font=("Segoe UI", 9, "bold"), foreground="#22543d", background="#c6f6d5")
        style.configure("BannerWarn.TLabel", font=("Segoe UI", 9, "bold"), foreground="#744210", background="#feebc8")
        style.configure("BannerErr.TLabel", font=("Segoe UI", 9, "bold"), foreground="#742a2a", background="#fed7d7")
        style.configure("Header.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#2b6cb0")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background="#2b6cb0", foreground="white")
        style.map("Accent.TButton", background=[("active", "#2c5282"), ("disabled", "#cbd5e0")])
        style.configure("Stop.TButton", font=("Segoe UI", 10, "bold"), background="#c53030", foreground="white")
        style.map("Stop.TButton", background=[("active", "#9b2c2c"), ("disabled", "#cbd5e0")])

    def build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        tab_main = ttk.Frame(notebook, padding=10)
        tab_preview = ttk.Frame(notebook, padding=10)
        tab_settings = ttk.Frame(notebook, padding=10)

        notebook.add(tab_main, text=" 📄 Κύρια Επεξεργασία ")
        notebook.add(tab_preview, text=" 👁️ Προεπισκόπηση & Κείμενο ")
        notebook.add(tab_settings, text=" ⚙️ Ρυθμίσεις Tesseract ")

        # ================= TAB 1: Κύρια Επεξεργασία =================
        # Banner κατάστασης
        self.banner_frame = ttk.Frame(tab_main)
        self.banner_frame.pack(fill="x", pady=(0, 8))
        self.lbl_main_banner = ttk.Label(self.banner_frame, text="", padding=(10, 5))
        self.lbl_main_banner.pack(fill="x")

        # 1. Αρχεία
        grp_files = ttk.LabelFrame(tab_main, text=" 📁 Επιλογή Αρχείων ", style="Header.TLabelframe")
        grp_files.pack(fill="x", pady=(0, 10))

        ttk.Label(grp_files, text="Αρχείο Εισόδου:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ent_in = ttk.Entry(grp_files, textvariable=self.input_file)
        ent_in.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        btn_in = ttk.Button(grp_files, text="Αναζήτηση...", command=self.browse_input)
        btn_in.grid(row=0, column=2, padx=8, pady=6)

        ttk.Label(grp_files, text="Βασικό Όνομα Εξόδου:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ent_out = ttk.Entry(grp_files, textvariable=self.output_base)
        ent_out.grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        btn_out = ttk.Button(grp_files, text="Αλλαγή...", command=self.browse_output)
        btn_out.grid(row=1, column=2, padx=8, pady=6)

        grp_files.columnconfigure(1, weight=1)

        # 2. Γλώσσα & Μορφές Εξαγωγής
        mid_frame = ttk.Frame(tab_main)
        mid_frame.pack(fill="x", pady=(0, 10))

        # Πλαίσιο Γλωσσών
        grp_lang = ttk.LabelFrame(mid_frame, text=" 🌐 Γλώσσα Αναγνώρισης (OCR Language) ", style="Header.TLabelframe")
        grp_lang.pack(side="left", fill="both", expand=True, padx=(0, 5))

        tess_opts = [
            ("Πολυτονικά & Νέα Ελληνικά + Αγγλικά (grc+ell+eng) ⭐", "grc+ell+eng"),
            ("Ελληνικά & Αγγλικά (ell+eng) - Μονοτονικό", "ell+eng"),
            ("Αρχαία / Πολυτονικά Ελληνικά (grc)", "grc"),
            ("Νέα Ελληνικά (ell)", "ell"),
            ("Αγγλικά (eng)", "eng"),
            ("Προσαρμοσμένο (Custom)", "custom"),
        ]
        for text, val in tess_opts:
            ttk.Radiobutton(grp_lang, text=text, variable=self.tess_lang_mode, value=val, command=self.on_tess_lang_change).pack(anchor="w", padx=8, pady=3)
        
        custom_box = ttk.Frame(grp_lang)
        custom_box.pack(fill="x", padx=8, pady=3)
        ttk.Label(custom_box, text="Κωδικοί:").pack(side="left")
        self.ent_custom_lang = ttk.Entry(custom_box, textvariable=self.tess_custom_lang, width=15, state="disabled")
        self.ent_custom_lang.pack(side="left", padx=5)

        # Μορφές Εξαγωγής
        grp_export = ttk.LabelFrame(mid_frame, text=" 💾 Μορφές Εξαγωγής & Φίλτρα ", style="Header.TLabelframe")
        grp_export.pack(side="right", fill="both", expand=True, padx=(5, 0))

        ttk.Checkbutton(grp_export, text="Searchable PDF (.pdf) - Αναζητήσιμο PDF", variable=self.opt_pdf).pack(anchor="w", padx=10, pady=4)
        ttk.Checkbutton(grp_export, text="Αρχείο Κειμένου (.txt) - Καθαρό κείμενο", variable=self.opt_txt).pack(anchor="w", padx=10, pady=4)
        ttk.Checkbutton(grp_export, text="Έγγραφο Word (.docx) - Επεξεργάσιμο", variable=self.opt_docx).pack(anchor="w", padx=10, pady=4)
        ttk.Checkbutton(grp_export, text="Βελτιστοποίηση Εικόνας (Autocontrast / Denoise)", variable=self.opt_enhance).pack(anchor="w", padx=10, pady=4)
        ttk.Checkbutton(grp_export, text="✨ Αυτόματος Καθαρισμός Ελληνικών (Post-process)", variable=self.opt_clean_text).pack(anchor="w", padx=10, pady=4)

        # 3. Ρυθμίσεις Σελίδων & Ποιότητας
        grp_opts = ttk.LabelFrame(tab_main, text=" ⚙️ Επιλογές Σελίδων & Ποιότητας ", style="Header.TLabelframe")
        grp_opts.pack(fill="x", pady=(0, 10))

        row_opts = ttk.Frame(grp_opts)
        row_opts.pack(fill="x", padx=10, pady=6)

        ttk.Label(row_opts, text="Σελίδες:").pack(side="left", padx=(0, 5))
        ttk.Entry(row_opts, textvariable=self.page_range_str, width=18).pack(side="left", padx=(0, 20))
        ttk.Label(row_opts, text="(π.χ. 'Όλες' ή '1-3, 5, 8')", foreground="#718096").pack(side="left", padx=(0, 30))

        ttk.Label(row_opts, text="Ανάλυση DPI:").pack(side="left", padx=(0, 5))
        dpi_combo = ttk.Combobox(row_opts, textvariable=self.dpi_val, values=[150, 200, 300, 400], width=8, state="readonly")
        dpi_combo.pack(side="left")
        ttk.Label(row_opts, text="(300 DPI βέλτιστο)", foreground="#718096").pack(side="left", padx=(5, 0))

        # 4. Κουμπιά Ενέργειας
        act_frame = ttk.Frame(tab_main)
        act_frame.pack(fill="x", pady=8)

        self.btn_run = ttk.Button(act_frame, text="🚀 Έναρξη OCR (Tesseract)", style="Accent.TButton", command=self.start_ocr)
        self.btn_run.pack(side="left", expand=True, fill="x", padx=(0, 5), ipady=6)

        self.btn_stop = ttk.Button(act_frame, text="⏹️ Διακοπή", style="Stop.TButton", state="disabled", command=self.stop_ocr)
        self.btn_stop.pack(side="right", expand=True, fill="x", padx=(5, 0), ipady=6)

        # ================= TAB 2: Προεπισκόπηση & Κείμενο =================
        top_preview_bar = ttk.Frame(tab_preview)
        top_preview_bar.pack(fill="x", pady=(0, 5))

        btn_copy = ttk.Button(top_preview_bar, text="📋 Αντιγραφή στο Πρόχειρο", command=self.copy_preview_text)
        btn_copy.pack(side="left", padx=5)

        btn_clear = ttk.Button(top_preview_bar, text="🗑️ Καθαρισμός", command=self.clear_preview_text)
        btn_clear.pack(side="left", padx=5)

        self.btn_open_file = ttk.Button(top_preview_bar, text="📂 Άνοιγμα Εγγράφου", state="disabled", command=self.open_output_file)
        self.btn_open_file.pack(side="right", padx=5)

        self.btn_open_folder = ttk.Button(top_preview_bar, text="📁 Άνοιγμα Φακέλου", state="disabled", command=self.open_output_folder)
        self.btn_open_folder.pack(side="right", padx=5)

        self.txt_preview = ScrolledText(tab_preview, wrap="word", font=("Consolas", 10))
        self.txt_preview.pack(fill="both", expand=True, pady=5)

        # ================= TAB 3: Ρυθμίσεις Tesseract =================
        grp_tess = ttk.LabelFrame(tab_settings, text=" ⚙️ Διαμόρφωση Tesseract OCR Engine ", style="Header.TLabelframe")
        grp_tess.pack(fill="x", pady=5)

        ttk.Label(grp_tess, text="Διαδρομή Εκτελέσιμου (tesseract.exe):").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(grp_tess, textvariable=self.tesseract_path).grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(grp_tess, text="Αναζήτηση...", command=self.browse_tesseract).grid(row=0, column=2, padx=8, pady=8)
        grp_tess.columnconfigure(1, weight=1)

        self.lbl_tess_status = ttk.Label(grp_tess, text="", font=("Segoe UI", 9, "italic"))
        self.lbl_tess_status.grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=4)

        self.lbl_tess_langs = ttk.Label(grp_tess, text="", font=("Segoe UI", 9))
        self.lbl_tess_langs.grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=(2, 8))

        # ================= Bottom Progress & Status Bar =================
        bot_frame = ttk.Frame(self.root, padding=(10, 5))
        bot_frame.pack(fill="x", side="bottom")

        self.progress_bar = ttk.Progressbar(bot_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 4))

        self.lbl_status = ttk.Label(bot_frame, text="Έτοιμο", relief="sunken", anchor="w", padding=(6, 2))
        self.lbl_status.pack(fill="x")

        # Αρχικοποίηση κατάστασης
        self.update_tesseract_info()

    def on_tess_lang_change(self):
        if self.tess_lang_mode.get() == "custom":
            self.ent_custom_lang.config(state="normal")
        else:
            self.ent_custom_lang.config(state="disabled")

    def update_tesseract_info(self):
        path = self.tesseract_path.get()
        if path and os.path.isfile(path):
            self.available_tess_langs = get_available_tesseract_languages(path)
            langs_str = ", ".join(self.available_tess_langs) if self.available_tess_langs else "Κανένα"
            self.lbl_tess_status.config(text=f"✅ Εντοπίστηκε έγκυρο Tesseract OCR: {path}", foreground="#276749")
            self.lbl_tess_langs.config(text=f"📚 Διαθέσιμα πακέτα γλωσσών: {langs_str}")
            self.lbl_main_banner.config(
                text="🟢 Tesseract OCR: Έτοιμο για χρήση (Υποστηρίζει Νέα & Αρχαία/Πολυτονικά Ελληνικά)", 
                style="BannerOk.TLabel"
            )
        else:
            self.available_tess_langs = []
            self.lbl_tess_status.config(text="❌ Το Tesseract OCR δεν εντοπίστηκε αυτόματα.", foreground="#c53030")
            self.lbl_tess_langs.config(text="")
            self.lbl_main_banner.config(
                text="🟡 Tesseract OCR: Δεν βρέθηκε εκτελέσιμο (Ορίστε το αρχείο στις Ρυθμίσεις)", 
                style="BannerWarn.TLabel"
            )

    def browse_tesseract(self):
        filename = filedialog.askopenfilename(
            title="Επιλογή tesseract.exe",
            filetypes=[("Εκτελέσιμο Tesseract", "tesseract.exe"), ("Όλα τα αρχεία", "*.*")]
        )
        if filename:
            self.tesseract_path.set(filename)
            self.update_tesseract_info()

    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Επιλέξτε αρχείο PDF ή Εικόνα",
            filetypes=[("Όλα τα υποστηριζόμενα", "*.pdf *.png *.jpg *.jpeg *.tiff *.bmp *.webp"),
                       ("Αρχεία PDF", "*.pdf"),
                       ("Αρχεία Εικόνων", "*.png *.jpg *.jpeg *.tiff *.bmp *.webp"),
                       ("Όλα τα αρχεία", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            base, _ = os.path.splitext(filename)
            self.output_base.set(base + "_OCR")

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Επιλέξτε βασικό όνομα αρχείων εξόδου",
            initialfile=os.path.basename(self.output_base.get()) if self.output_base.get() else "output_OCR"
        )
        if filename:
            base, _ = os.path.splitext(filename)
            self.output_base.set(base)

    def copy_preview_text(self):
        text = self.txt_preview.get("1.0", tk.END).strip()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("Αντιγραφή", "Το κείμενο αντιγράφηκε επιτυχώς στο πρόχειρο!")

    def clear_preview_text(self):
        self.txt_preview.delete("1.0", tk.END)

    def open_output_file(self):
        if self.last_created_files:
            target = self.last_created_files[0]
            if os.path.exists(target):
                os.startfile(target)

    def open_output_folder(self):
        if self.last_created_files:
            folder = os.path.dirname(self.last_created_files[0])
            if os.path.exists(folder):
                subprocess.Popen(f'explorer "{folder}"')
        elif self.output_base.get():
            folder = os.path.dirname(self.output_base.get())
            if os.path.exists(folder):
                subprocess.Popen(f'explorer "{folder}"')

    def get_selected_lang(self):
        mode = self.tess_lang_mode.get()
        if mode == "custom":
            c = self.tess_custom_lang.get().strip()
            return c if c else "ell+eng"
        return mode

    def start_ocr(self):
        infile = self.input_file.get().strip()
        outbase = self.output_base.get().strip()
        tess_cmd = self.tesseract_path.get().strip()

        if not infile or not os.path.exists(infile):
            messagebox.showwarning("Προσοχή", "Παρακαλώ επιλέξτε ένα έγκυρο αρχείο εισόδου.")
            return

        if not outbase:
            messagebox.showwarning("Προσοχή", "Παρακαλώ καθορίστε το βασικό όνομα εξόδου.")
            return

        if not tess_cmd or not os.path.exists(tess_cmd):
            messagebox.showerror("Σφάλμα", "Το Tesseract OCR δεν βρέθηκε. Μεταβείτε στις Ρυθμίσεις για να ορίσετε το tesseract.exe.")
            return

        if not (self.opt_pdf.get() or self.opt_txt.get() or self.opt_docx.get()):
            messagebox.showwarning("Προσοχή", "Παρακαλώ επιλέξτε τουλάχιστον μία μορφή εξαγωγής (PDF, TXT ή Word).")
            return

        lang = self.get_selected_lang()

        # UI State update
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_open_file.config(state="disabled")
        self.btn_open_folder.config(state="disabled")
        self.progress_bar["value"] = 0
        self.lbl_status.config(text="Εκκίνηση διαδικασίας Tesseract OCR...")
        self.clear_preview_text()

        self.cancel_event.clear()
        self.last_created_files = []

        self.worker_thread = OcrWorker(
            tesseract_cmd=tess_cmd,
            input_path=infile,
            output_base=outbase,
            lang=lang,
            export_pdf=self.opt_pdf.get(),
            export_txt=self.opt_txt.get(),
            export_docx=self.opt_docx.get(),
            dpi=self.dpi_val.get(),
            enhance_image=self.opt_enhance.get(),
            clean_text=self.opt_clean_text.get(),
            page_range_str=self.page_range_str.get(),
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event
        )
        self.worker_thread.start()

    def stop_ocr(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.cancel_event.set()
            self.lbl_status.config(text="Αίτημα διακοπής... Παρακαλώ περιμένετε...")
            self.btn_stop.config(state="disabled")

    def check_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                msg_type = msg[0]

                if msg_type == "STATUS":
                    self.lbl_status.config(text=msg[1])

                elif msg_type == "PROGRESS":
                    idx, total, text = msg[1], msg[2], msg[3]
                    percent = int((idx / total) * 100) if total > 0 else 0
                    self.progress_bar["value"] = percent
                    self.lbl_status.config(text=f"Επεξεργασία: {text} ({percent}%)")

                elif msg_type == "PAGE_TEXT":
                    page_num, text = msg[1], msg[2]
                    self.txt_preview.insert(tk.END, f"\n=== Σελίδα {page_num} ===\n\n{text}\n")
                    self.txt_preview.see(tk.END)

                elif msg_type == "SUCCESS":
                    created_files = msg[1]
                    self.last_created_files = created_files
                    self.progress_bar["value"] = 100
                    self.lbl_status.config(text="✅ Η μετατροπή ολοκληρώθηκε επιτυχώς!")
                    self.btn_run.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    self.btn_open_file.config(state="normal")
                    self.btn_open_folder.config(state="normal")
                    
                    files_str = "\n".join([f"• {os.path.basename(f)}" for f in created_files])
                    messagebox.showinfo("Επιτυχία OCR", f"Η διαδικασία ολοκληρώθηκε επιτυχώς!\n\nΔημιουργήθηκαν τα αρχεία:\n{files_str}")

                elif msg_type == "ERROR":
                    err_msg = msg[1]
                    self.lbl_status.config(text="❌ Προέκυψε σφάλμα κατά την επεξεργασία.")
                    self.btn_run.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    messagebox.showerror("Σφάλμα OCR", f"Προέκυψε σφάλμα:\n{err_msg}")

        except queue.Empty:
            pass

        self.root.after(100, self.check_queue)


if __name__ == "__main__":
    root = tk.Tk()
    app = ModernOcrApp(root)
    root.mainloop()
