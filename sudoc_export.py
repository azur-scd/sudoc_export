"""
Script SudocExport
Outil de récupération de notices MARCXML depuis le Sudoc (ABES)
à partir d'une liste de PPN — Export Excel
Récupération via : https://www.sudoc.fr/{ppn}.xml
Auteur : généré avec Claude (Anthropic)
Usage  : python sudoc_export.py  (ou double-clic sur le .exe compilé)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import requests
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import re
import time
import os
import sys

# ─────────────────────────────────────────────
#  CONSTANTES MARC
# ─────────────────────────────────────────────

SUDOC_XML_URL = "https://www.sudoc.fr/{ppn}.xml"

# Champs à extraire avec leurs libellés pour l'en-tête Excel
COLUMNS = [
    ("PPN",                  "ppn"),
    ("Type de document",     "type_doc"),
    ("Titre",                "titre"),
    ("Mention de resp.",     "resp"),
    ("Auteur(s)",            "auteurs"),
    ("Éditeur",              "editeur"),
    ("Lieu d'édition",       "lieu"),
    ("Date de publication",  "date"),
    ("Date codée",           "date_codee"),
    ("Pays d'édition",       "pays"),
    ("Description physique", "description"),
    ("Collection",           "collection"),
    ("Langue",               "langue"),
    ("ISBN",                 "isbn"),
    ("EAN",                  "ean"),
    ("ISSN",                 "issn"),
    ("Lien Sudoc",           "url"),
    ("Statut",               "statut"),
]


# ─────────────────────────────────────────────
#  HELPERS MARCXML
# ─────────────────────────────────────────────

def _clean(val):
    """Retire la ponctuation ISBD finale souvent présente dans les données Sudoc."""
    return val.rstrip(" .,;:/")


def get_subfields(record, tag, *codes):
    """Retourne les sous-champs demandés avec ponctuation.
    - Plusieurs sous-champs d'une même occurrence : joints par ", "
    - Plusieurs occurrences du même champ       : séparées par " ; "
    - Ponctuation finale héritée du Sudoc nettoyée sur chaque sous-champ.
    """
    values = []
    for field in record.findall(f"datafield[@tag='{tag}']"):
        parts = []
        for code in codes:
            for sf in field.findall(f"subfield[@code='{code}']"):
                if sf.text:
                    parts.append(_clean(sf.text.strip()))
        if parts:
            values.append(", ".join(parts))
    return " ; ".join(values) if values else ""


# Ponctuation ISBD précédant chaque sous-champ du 200
_TITRE_PUNCT = {"a": "", "e": " : ", "h": ". ", "i": ", "}

def get_titre_200(record):
    """Construit le titre avec ponctuation ISBD à partir du champ 200.
    $a Titre propre
    $e Sous-titre                    ->  précédé de ' : '
    $h Numéro de partie              ->  précédé de '. '
    $i Titre de partie               ->  précédé de ', '
    Plusieurs occurrences du 200 (rare) sont séparées par ' ; '.
    Les titres parallèles ($d) ne sont pas exploités.
    """
    titres = []
    for field in record.findall("datafield[@tag='200']"):
        result = ""
        for sf in field.findall("subfield"):
            code = sf.get("code", "")
            val  = (sf.text or "").strip()
            if not val or code not in _TITRE_PUNCT:
                continue
            punct = _TITRE_PUNCT[code]
            if result and punct:
                # Supprimer la ponctuation finale du Sudoc si elle existe déjà
                tail = result.rstrip()
                if tail and tail[-1] in (":", ".", ","):
                    result = tail[:-1].rstrip()
                result += punct + val
            else:
                result += punct + val
        if result:
            titres.append(result.strip())
    return " ; ".join(titres)


def get_controlfield(record, tag):
    """Retourne le contenu d'un champ de contrôle."""
    cf = record.find(f"controlfield[@tag='{tag}']")
    return cf.text.strip() if cf is not None and cf.text else ""


def get_leader_type(record):
    """Détermine le type de document à partir des positions 6 et 7 du Leader.

    Position 6 : type de record
    Position 7 : niveau bibliographique
    La combinaison des deux permet de distinguer monographie, périodique,
    document électronique, etc.
    """
    leader = record.find("leader")
    ldr = leader.text if leader is not None and leader.text else ""
    # Leader UNIMARC standard 24 caractères :
    # position 6 = type de record, position 7 = niveau bibliographique
    if len(ldr) < 8:
        return "Inconnu"

    p6 = ldr[6]   # type de record
    p7 = ldr[7]   # niveau bibliographique

    # Table de décodage combinée pos.6 + pos.7
    # pos.6 : type de record   pos.7 : niveau bibliographique
    combined = {
        # Ressources textuelles (a=imprimé, b=manuscrit)
        ("a", "m"): "Monographie",
        ("a", "s"): "Périodique",
        ("a", "a"): "Article / partie composante",
        ("a", "i"): "Ressource textuelle intégratrice",
        ("a", "c"): "Collection",
        ("b", "m"): "Manuscrit (monographie)",
        ("b", "s"): "Manuscrit (périodique)",
        ("b", "a"): "Manuscrit (partie composante)",
        # Musique notée (c=imprimée, d=manuscrite)
        ("c", "m"): "Musique notée (monographie)",
        ("c", "s"): "Musique notée (périodique)",
        ("c", "a"): "Musique notée (partie composante)",
        ("d", "m"): "Musique notée manuscrite (monographie)",
        ("d", "a"): "Musique notée manuscrite (partie composante)",
        # Ressources cartographiques (e=imprimée, f=manuscrite)
        ("e", "m"): "Ressource cartographique (monographie)",
        ("e", "s"): "Ressource cartographique (périodique)",
        ("e", "a"): "Ressource cartographique (partie composante)",
        ("f", "m"): "Ressource cartographique manuscrite",
        # Ressource projetée ou vidéo
        ("g", "m"): "Vidéo / image projetée (monographie)",
        ("g", "s"): "Vidéo / image projetée (périodique)",
        ("g", "a"): "Vidéo / image projetée (partie composante)",
        # Enregistrements sonores
        ("i", "m"): "Enregistrement sonore non musical (monographie)",
        ("i", "s"): "Enregistrement sonore non musical (périodique)",
        ("i", "a"): "Enregistrement sonore non musical (partie composante)",
        ("j", "m"): "Enregistrement sonore musical (monographie)",
        ("j", "s"): "Enregistrement sonore musical (périodique)",
        ("j", "a"): "Enregistrement sonore musical (partie composante)",
        # Ressource graphique à deux dimensions
        ("k", "m"): "Image / ressource graphique (monographie)",
        ("k", "s"): "Image / ressource graphique (périodique)",
        ("k", "a"): "Image / ressource graphique (partie composante)",
        # Ressource électronique (l = L minuscule)
        ("l", "m"): "Ressource électronique (monographie)",
        ("l", "s"): "Ressource électronique (périodique)",
        ("l", "a"): "Ressource électronique (partie composante)",
        ("l", "i"): "Ressource électronique intégratrice",
        # Ressource multimédia
        ("m", "m"): "Ressource multimédia (monographie)",
        ("m", "s"): "Ressource multimédia (périodique)",
        ("m", "a"): "Ressource multimédia (partie composante)",
        # Objet tridimensionnel
        ("r", "m"): "Objet tridimensionnel",
        ("r", "a"): "Objet tridimensionnel (partie composante)",
    }

    if (p6, p7) in combined:
        return combined[(p6, p7)]
    # Repli sur pos.6 seule si combinaison inconnue
    fallback = {
        "a": "Ressource textuelle",
        "b": "Ressource textuelle manuscrite",
        "c": "Musique notée",
        "d": "Musique notée manuscrite",
        "e": "Ressource cartographique",
        "f": "Ressource cartographique manuscrite",
        "g": "Vidéo / image projetée",
        "i": "Enregistrement sonore non musical",
        "j": "Enregistrement sonore musical",
        "k": "Image / ressource graphique",
        "l": "Ressource électronique",
        "m": "Ressource multimédia",
        "r": "Objet tridimensionnel",
    }
    label = fallback.get(p6, f"Inconnu (pos6={p6!r})")
    niv = {"a": "partie composante", "i": "intégrateur",
           "m": "monographie", "s": "périodique", "c": "collection"}
    return f"{label} — {niv.get(p7, p7)}"


def parse_record(record):
    """Extrait les données bibliographiques d'un record MARCXML."""
    data = {}

    # PPN (001)
    data["ppn"] = get_controlfield(record, "001")

    # Type de document — décodage leader puis affinage via 105$a pos.4
    data["type_doc"] = get_leader_type(record)
    _sf105 = record.find("datafield[@tag='105']/subfield[@code='a']")
    if _sf105 is not None and _sf105.text and len(_sf105.text) >= 5:
        _pos4 = _sf105.text[4]
        if _pos4 == "m":
            data["type_doc"] = "Thèse (version de soutenance)"
        elif _pos4 == "7":
            data["type_doc"] = "Mémoire (version de soutenance)"

    # Titre (200 $a : $e. $h, $i) avec ponctuation ISBD
    data["titre"] = get_titre_200(record)

    # Mention de responsabilité (200 $f $g)
    data["resp"] = get_subfields(record, "200", "f", "g")

    # Auteurs : 700 (auteur principal personne), 701 (co-auteur), 702 (autre resp.)
    auteurs = []
    for tag in ("700", "701", "702"):
        v = get_subfields(record, tag, "a", "b")
        if v:
            auteurs.append(v)
    # Collectivités : 710, 711
    for tag in ("710", "711"):
        v = get_subfields(record, tag, "a", "b")
        if v:
            auteurs.append(v)
    data["auteurs"] = " ; ".join(auteurs)

    # Éditeur / lieu / date — préférence 214, sinon 210
    if get_subfields(record, "214", "c"):
        data["editeur"] = get_subfields(record, "214", "c")
        data["lieu"]    = get_subfields(record, "214", "a")
        data["date"]    = get_subfields(record, "214", "d")
    else:
        data["editeur"] = get_subfields(record, "210", "c")
        data["lieu"]    = get_subfields(record, "210", "a")
        data["date"]    = get_subfields(record, "210", "d")

    # Date codée : positions 9-12 de la zone 100 $a (lecture brute, sans nettoyage)
    _sf100 = record.find("datafield[@tag='100']/subfield[@code='a']")
    _z100  = (_sf100.text or "") if _sf100 is not None else ""
    data["date_codee"] = _z100[9:13].strip() if len(_z100) >= 13 else ""

    # Pays d'édition (102 $a)
    data["pays"] = get_subfields(record, "102", "a")

    # Description physique (215 $a uniquement)
    data["description"] = get_subfields(record, "215", "a")

    # Collection (225 $a)
    data["collection"] = get_subfields(record, "225", "a")

    # Langue (101 $a)
    data["langue"] = get_subfields(record, "101", "a")

    # ISBN (010 $a)
    data["isbn"] = get_subfields(record, "010", "a")

    # EAN (073 $a)
    data["ean"] = get_subfields(record, "073", "a")

    # ISSN (011 $a)
    data["issn"] = get_subfields(record, "011", "a")

    # Lien vers la notice Sudoc
    ppn = data["ppn"]
    data["url"] = f"https://www.sudoc.fr/{ppn}" if ppn else ""

    data["statut"] = "OK"
    return data


# ─────────────────────────────────────────────
#  REQUÊTE SUDOC (URL directe {ppn}.xml)
# ─────────────────────────────────────────────

def _empty_row(ppn, statut):
    """Retourne une ligne vide avec le PPN et le statut d'erreur."""
    row = {k: "" for _, k in COLUMNS}
    row["ppn"]    = ppn
    row["statut"] = statut
    return row


def fetch_ppn(ppn):
    """Récupère https://www.sudoc.fr/{ppn}.xml et retourne un dict bibliographique."""
    ppn = ppn.strip()
    if not ppn:
        return None

    url = SUDOC_XML_URL.format(ppn=ppn)

    try:
        resp = requests.get(url, timeout=15)
    except requests.exceptions.Timeout:
        return _empty_row(ppn, "Erreur : timeout")
    except requests.exceptions.RequestException as e:
        return _empty_row(ppn, f"Erreur réseau : {e}")

    if resp.status_code == 404:
        return _empty_row(ppn, "PPN introuvable (404)")
    if resp.status_code != 200:
        return _empty_row(ppn, f"Erreur HTTP {resp.status_code}")

    try:
        root = ET.fromstring(resp.content)
    except ET.ParseError:
        return _empty_row(ppn, "Erreur : XML invalide")

    # Le document retourné est directement un <record> MARCXML
    # (espace de noms http://www.loc.gov/MARC21/slim)
    record_el = root  # la racine EST le record
    # Vérification : si la racine n'est pas un record MARC, chercher en profondeur
    local = root.tag.split("}")[-1] if "}" in root.tag else root.tag
    if local != "record":
        record_el = root.find(".//{http://www.loc.gov/MARC21/slim}record")
        if record_el is None:
            record_el = root.find(".//record")
    if record_el is None:
        return _empty_row(ppn, "Erreur : structure MARCXML inattendue")

    return parse_record(record_el)


# ─────────────────────────────────────────────
#  EXPORT EXCEL
# ─────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", fgColor="1F3864")
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)
ERROR_FILL    = PatternFill("solid", fgColor="FFDDD5")
ALTROW_FILL   = PatternFill("solid", fgColor="EEF2F8")
LINK_FONT     = Font(color="1155CC", underline="single")
THIN          = Side(style="thin", color="CCCCCC")
CELL_BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def export_xlsx(rows, filepath):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Notices Sudoc"

    # En-tête
    headers = [label for label, _ in COLUMNS]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font        = HEADER_FONT
        cell.fill        = HEADER_FILL
        cell.alignment   = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border      = CELL_BORDER
    ws.row_dimensions[1].height = 28

    # Données
    for row_idx, data in enumerate(rows, 2):
        is_error = data.get("statut", "OK") != "OK"
        is_alt   = (row_idx % 2 == 0) and not is_error

        for col_idx, (_, key) in enumerate(COLUMNS, 1):
            value = data.get(key, "")
            cell  = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border    = CELL_BORDER

            if is_error:
                cell.fill = ERROR_FILL
            elif is_alt:
                cell.fill = ALTROW_FILL

            # Hyperlien sur la colonne URL
            if key == "url" and value.startswith("http"):
                cell.hyperlink = value
                cell.font      = LINK_FONT
                cell.value     = value

    # Largeurs de colonnes
    col_widths = {
        "PPN": 14, "Type de document": 22, "Titre": 55,
        "Mention de resp.": 30, "Auteur(s)": 30, "Éditeur": 25,
        "Lieu d'édition": 20, "Date de publication": 12, "Date codée": 12, "Pays d'édition": 16,
        "Description physique": 20, "Collection": 22, "Langue": 10,
        "ISBN": 18, "EAN": 16, "ISSN": 16, "Lien Sudoc": 32, "Statut": 18,
    }
    for col_idx, (label, _) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(label, 18)

    # Figer la première ligne
    ws.freeze_panes = "A2"

    # Filtre automatique
    ws.auto_filter.ref = ws.dimensions

    wb.save(filepath)


# ─────────────────────────────────────────────
#  INTERFACE GRAPHIQUE
# ─────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sudoc — Export notices MARCXML")
        self.resizable(True, True)
        self.minsize(680, 560)
        self._build_ui()
        self._results = []
        # Centrer la fenêtre
        self.update_idletasks()
        w, h = 760, 620
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # ── Couleurs ──────────────────────────────
        BG      = "#F5F7FA"
        ACCENT  = "#1F3864"
        BTN_FG  = "#FFFFFF"
        BTN_ACT = "#2E4F8C"

        self.configure(bg=BG)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame",       background=BG)
        style.configure("TLabel",       background=BG, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=BG, font=("Segoe UI", 14, "bold"), foreground=ACCENT)
        style.configure("Sub.TLabel",   background=BG, font=("Segoe UI", 9), foreground="#555")
        style.configure("TProgressbar", troughcolor="#DDE3EE", background=ACCENT, thickness=10)

        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        # ── Titre ─────────────────────────────────
        ttk.Label(main, text="Sudoc — Export de notices", style="Title.TLabel").pack(anchor="w")
        ttk.Label(main, text="Collez vos PPN (un par ligne) ou importez un fichier texte, puis lancez la récupération.", style="Sub.TLabel").pack(anchor="w", pady=(2, 12))

        # ── Zone de saisie des PPN ────────────────
        ppn_header = ttk.Frame(main)
        ppn_header.pack(fill="x")
        ttk.Label(ppn_header, text="Liste de PPN :").pack(side="left")
        tk.Button(
            ppn_header, text="📂  Importer un fichier .txt",
            bg="#5C6BC0", fg=BTN_FG, activebackground="#3949AB", activeforeground=BTN_FG,
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            padx=10, pady=3, bd=0,
            command=self._import_file,
        ).pack(side="right")
        frame_text = ttk.Frame(main)
        frame_text.pack(fill="both", expand=True, pady=(4, 0))

        self._ppn_text = tk.Text(
            frame_text, height=10, font=("Consolas", 10),
            relief="flat", bd=1, highlightthickness=1,
            highlightbackground="#CDD5E0", highlightcolor=ACCENT,
            wrap="none",
        )
        sb_y = ttk.Scrollbar(frame_text, orient="vertical",   command=self._ppn_text.yview)
        sb_x = ttk.Scrollbar(frame_text, orient="horizontal", command=self._ppn_text.xview)
        self._ppn_text.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self._ppn_text.pack(side="left", fill="both", expand=True)

        # ── Barre de progression ───────────────────
        self._progress_var = tk.DoubleVar()
        self._progress = ttk.Progressbar(main, variable=self._progress_var, maximum=100, style="TProgressbar")
        self._progress.pack(fill="x", pady=(12, 0))

        self._status_var = tk.StringVar(value="Prêt.")
        ttk.Label(main, textvariable=self._status_var, style="Sub.TLabel").pack(anchor="w", pady=(3, 0))

        # ── Délai entre requêtes ───────────────────
        delay_frame = ttk.Frame(main)
        delay_frame.pack(anchor="w", pady=(8, 0))
        ttk.Label(delay_frame, text="Délai entre requêtes (ms) :").pack(side="left")
        self._delay_var = tk.StringVar(value="500")
        vcmd = (self.register(lambda s: s.isdigit() or s == ""), "%P")
        delay_entry = tk.Entry(
            delay_frame, textvariable=self._delay_var, width=6,
            font=("Segoe UI", 10), relief="flat", bd=1,
            highlightthickness=1, highlightbackground="#CDD5E0",
            highlightcolor=ACCENT, justify="right",
            validate="key", validatecommand=vcmd,
        )
        delay_entry.pack(side="left", padx=(6, 4))
        ttk.Label(delay_frame, text="(0 = pas de délai)", style="Sub.TLabel").pack(side="left")

        # ── Sauvegarde automatique ─────────────────
        save_frame = ttk.Frame(main)
        save_frame.pack(anchor="w", pady=(4, 0))
        ttk.Label(save_frame, text="Sauvegarde auto toutes les").pack(side="left")
        self._autosave_var = tk.StringVar(value="100")
        vcmd2 = (self.register(lambda s: s.isdigit() or s == ""), "%P")
        tk.Entry(
            save_frame, textvariable=self._autosave_var, width=6,
            font=("Segoe UI", 10), relief="flat", bd=1,
            highlightthickness=1, highlightbackground="#CDD5E0",
            highlightcolor=ACCENT, justify="right",
            validate="key", validatecommand=vcmd2,
        ).pack(side="left", padx=(6, 4))
        ttk.Label(save_frame, text="notices  (0 = désactivée)", style="Sub.TLabel").pack(side="left")

        # ── Boutons ────────────────────────────────
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(14, 0))

        btn_cfg = dict(
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            padx=18, pady=8, bd=0,
        )

        self._btn_run = tk.Button(
            btn_frame, text="▶  Lancer la récupération",
            bg=ACCENT, fg=BTN_FG, activebackground=BTN_ACT, activeforeground=BTN_FG,
            command=self._start,
            **btn_cfg,
        )
        self._btn_run.pack(side="left", padx=(0, 10))

        self._btn_export = tk.Button(
            btn_frame, text="💾  Exporter en Excel",
            bg="#2E7D32", fg=BTN_FG, activebackground="#1B5E20", activeforeground=BTN_FG,
            command=self._export,
            state="disabled",
            **btn_cfg,
        )
        self._btn_export.pack(side="left")

        self._btn_clear = tk.Button(
            btn_frame, text="Effacer",
            bg="#888", fg=BTN_FG, activebackground="#555", activeforeground=BTN_FG,
            command=self._clear,
            **btn_cfg,
        )
        self._btn_clear.pack(side="right")

        # ── Log ────────────────────────────────────
        ttk.Label(main, text="Journal :").pack(anchor="w", pady=(16, 2))
        log_frame = ttk.Frame(main)
        log_frame.pack(fill="both", expand=False)

        self._log = tk.Text(
            log_frame, height=7, font=("Consolas", 9),
            relief="flat", bd=0, state="disabled",
            bg="#1E2A3A", fg="#A8C7E8",
            insertbackground="white",
        )
        log_sb = ttk.Scrollbar(log_frame, orient="vertical", command=self._log.yview)
        self._log.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

    # ── Méthodes ─────────────────────────────────

    def _log_write(self, msg, tag=None):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _import_file(self):
        """Ouvre un fichier texte et charge les PPN dans la zone de saisie."""
        filepath = filedialog.askopenfilename(
            title="Importer un fichier de PPN",
            filetypes=[
                ("Fichiers texte", "*.txt"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if not filepath:
            return
        try:
            # Détection automatique de l'encodage : UTF-8 puis fallback latin-1
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    raw = f.read()
            except UnicodeDecodeError:
                with open(filepath, "r", encoding="latin-1") as f:
                    raw = f.read()
            # Extraire les PPN : on garde uniquement les tokens non vides
            ppns = [p.strip() for p in re.split(r"[\n\r,;\t]+", raw) if p.strip()]
            if not ppns:
                messagebox.showwarning("Fichier vide", "Aucun PPN trouvé dans ce fichier.")
                return
            # Insérer dans la zone de texte (remplace le contenu existant)
            self._ppn_text.delete("1.0", "end")
            self._ppn_text.insert("end", "\n".join(ppns))
            self._log_write(f"  📂 {len(ppns)} PPN importé(s) depuis : {filepath}")
        except Exception as e:
            messagebox.showerror("Erreur de lecture", f"Impossible de lire le fichier :\n{e}")

    def _clear(self):
        self._ppn_text.delete("1.0", "end")
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._status_var.set("Prêt.")
        self._progress_var.set(0)
        self._results = []
        self._btn_export.configure(state="disabled")

    def _start(self):
        raw = self._ppn_text.get("1.0", "end").strip()
        ppns = [p.strip() for p in re.split(r"[\n\r,;]+", raw) if p.strip()]
        if not ppns:
            messagebox.showwarning("Attention", "Aucun PPN saisi.")
            return
        self._btn_run.configure(state="disabled")
        self._btn_export.configure(state="disabled")
        self._results = []
        self._progress_var.set(0)
        # Nom du fichier de sauvegarde automatique (fixe pour toute la session)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self._autosave_path = os.path.join(
            os.path.expanduser("~"), f"sudoc_autosave_{timestamp}.xlsx"
        )
        threading.Thread(target=self._run, args=(ppns,), daemon=True).start()

    def _run(self, ppns):
        total = len(ppns)
        try:
            delay_ms = int(self._delay_var.get())
        except ValueError:
            delay_ms = 500
        delay_s = max(0, delay_ms) / 1000.0
        self._log_write(f"Démarrage — {total} PPN à traiter (délai : {delay_ms} ms).")
        ok_count = err_count = 0

        try:
            autosave_every = max(0, int(self._autosave_var.get()))
        except ValueError:
            autosave_every = 100
        if autosave_every > 0:
            self._log_write(f"Sauvegarde automatique activée toutes les {autosave_every} notices.")
            self._log_write(f"  → {self._autosave_path}")

        for i, ppn in enumerate(ppns, 1):
            self._status_var.set(f"Traitement {i}/{total} — PPN {ppn} …")
            data = fetch_ppn(ppn)
            if data:
                self._results.append(data)
                statut = data.get("statut", "?")
                if statut == "OK":
                    ok_count += 1
                    titre = data.get("titre") or "(sans titre)"
                    self._log_write(f"  ✓ {ppn}  {titre[:60]}")
                else:
                    err_count += 1
                    self._log_write(f"  ✗ {ppn}  [{statut}]")
            pct = i / total * 100
            self._progress_var.set(pct)
            # Sauvegarde automatique intermédiaire
            if autosave_every > 0 and i % autosave_every == 0 and self._results:
                try:
                    export_xlsx(self._results, self._autosave_path)
                    self._log_write(f"  💾 Sauvegarde automatique — {i}/{total} notices ({self._autosave_path})")
                except Exception as e:
                    self._log_write(f"  ⚠ Sauvegarde automatique échouée : {e}")
            if i < total and delay_s > 0:
                time.sleep(delay_s)

        summary = f"Terminé — {ok_count} notice(s) récupérée(s), {err_count} erreur(s)."
        self._status_var.set(summary)
        self._log_write(summary)

        if self._results:
            self._btn_export.configure(state="normal")
        self._btn_run.configure(state="normal")

    def _export(self):
        if not self._results:
            messagebox.showinfo("Info", "Aucune donnée à exporter.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"notices_sudoc_{timestamp}.xlsx"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")],
            initialfile=default_name,
            title="Enregistrer le fichier Excel",
        )
        if not filepath:
            return
        try:
            export_xlsx(self._results, filepath)
            self._log_write(f"  💾 Fichier exporté : {filepath}")
            messagebox.showinfo("Export réussi", f"Fichier enregistré :\n{filepath}")
        except Exception as e:
            messagebox.showerror("Erreur d'export", str(e))


# ─────────────────────────────────────────────
#  POINT D'ENTRÉE
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
