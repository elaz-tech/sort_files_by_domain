#!/usr/bin/env python3
"""
app.py  —  File Domain Sorter  (PyQt6 GUI)
────────────────────────────────────────────────────────
Classifies files by content domain using a live-editable
keyword config (config/domains_config.json) — no restart needed.

Requirements:
    pip install PyQt6 pypdf pdfplumber python-docx anthropic --break-system-packages

Run:
    python app.py
"""

import json
import os
import shutil
import sys
import textwrap
from pathlib import Path

from PyQt6.QtCore import (
    QObject, QRunnable, QSettings, Qt, QThreadPool,
    pyqtSignal, pyqtSlot,
)
from PyQt6.QtGui import QColor, QIcon, QPalette, QTextCharFormat, QTextCursor
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QDialog, QFileDialog, QFrame,
    QGroupBox, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QScrollArea,
    QSizePolicy, QSplitter, QStatusBar, QToolButton,
    QVBoxLayout, QWidget,
)

# ── optional heavy deps ────────────────────────────────────────────────────────
try:    import pdfplumber;               HAS_PDFPLUMBER = True
except ImportError:                      HAS_PDFPLUMBER = False
try:    from pypdf import PdfReader;     HAS_PYPDF = True
except ImportError:                      HAS_PYPDF = False
try:    from docx import Document as DD; HAS_DOCX = True
except ImportError:                      HAS_DOCX = False
try:    import anthropic;                HAS_ANTHROPIC = True
except ImportError:                      HAS_ANTHROPIC = False

# ── palette ────────────────────────────────────────────────────────────────────
DARK   = "#0d1117"
PANEL  = "#161b22"
PANEL2 = "#1c2128"
BORDER = "#30363d"
ACCENT = "#58a6ff"
GREEN  = "#3fb950"
YELLOW = "#d29922"
RED    = "#f85149"
MUTED  = "#8b949e"
FG     = "#e6edf3"

STYLE = f"""
QMainWindow, QWidget {{
    background:{DARK}; color:{FG};
    font-family:'JetBrains Mono','Consolas','Courier New',monospace;
    font-size:13px;
}}
QGroupBox {{
    border:1px solid {BORDER}; border-radius:8px;
    margin-top:14px; padding:10px 14px 14px 14px;
    font-size:11px; color:{MUTED}; font-weight:bold;
    letter-spacing:1.5px;
}}
QGroupBox::title {{ subcontrol-origin:margin; left:12px; top:2px; }}
QLineEdit {{
    background:{PANEL}; border:1px solid {BORDER}; border-radius:6px;
    padding:8px 10px; color:{FG}; selection-background-color:{ACCENT};
}}
QLineEdit:focus {{ border-color:{ACCENT}; }}
QPushButton {{
    background:{PANEL}; border:1px solid {BORDER}; border-radius:6px;
    padding:7px 16px; color:{FG}; font-weight:bold;
}}
QPushButton:hover {{ border-color:{ACCENT}; color:{ACCENT}; }}
QPushButton#runBtn {{
    background:{ACCENT}; border:none; color:{DARK};
    font-size:14px; padding:10px 30px; border-radius:8px;
    font-weight:bold; letter-spacing:0.5px;
}}
QPushButton#runBtn:hover {{ background:#79c0ff; }}
QPushButton#runBtn:disabled {{ background:{BORDER}; color:{MUTED}; }}
QPushButton#stopBtn {{
    background:transparent; border:1px solid {RED}; color:{RED};
    border-radius:8px; font-weight:bold; padding:10px 20px;
}}
QPushButton#stopBtn:hover {{ background:{RED}; color:white; }}
QPushButton#dangerBtn {{
    background:transparent; border:1px solid {RED}; color:{RED};
    border-radius:6px; padding:5px 12px;
}}
QPushButton#dangerBtn:hover {{ background:{RED}; color:white; }}
QPushButton#addBtn {{
    background:transparent; border:1px solid {GREEN}; color:{GREEN};
    border-radius:6px; padding:5px 12px; font-weight:bold;
}}
QPushButton#addBtn:hover {{ background:{GREEN}; color:{DARK}; }}
QCheckBox {{ spacing:8px; color:{FG}; }}
QCheckBox::indicator {{
    width:16px; height:16px; border:1px solid {BORDER};
    border-radius:4px; background:{PANEL};
}}
QCheckBox::indicator:checked {{ background:{ACCENT}; border-color:{ACCENT}; }}
QPlainTextEdit {{
    background:#010409; border:1px solid {BORDER}; border-radius:8px;
    color:#b3c3d4;
    font-family:'JetBrains Mono','Fira Code','Courier New',monospace;
    font-size:12px; padding:8px; selection-background-color:{ACCENT};
}}
QProgressBar {{
    background:{PANEL}; border:1px solid {BORDER}; border-radius:4px;
    height:6px; color:transparent;
}}
QProgressBar::chunk {{ background:{ACCENT}; border-radius:4px; }}
QScrollBar:vertical {{
    background:{PANEL}; width:8px; border-radius:4px;
}}
QScrollBar::handle:vertical {{
    background:{BORDER}; border-radius:4px; min-height:30px;
}}
QScrollBar::handle:vertical:hover {{ background:{MUTED}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0px; }}
QScrollBar:horizontal {{
    background:{PANEL}; height:8px; border-radius:4px;
}}
QScrollBar::handle:horizontal {{
    background:{BORDER}; border-radius:4px; min-width:30px;
}}
QScrollBar::handle:horizontal:hover {{ background:{MUTED}; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width:0px; }}
QStatusBar {{
    background:{PANEL}; border-top:1px solid {BORDER};
    color:{MUTED}; font-size:11px;
}}
QListWidget {{
    background:{PANEL2}; border:1px solid {BORDER}; border-radius:6px;
    color:{FG}; padding:4px; font-size:12px;
}}
QListWidget::item {{ padding:5px 8px; border-radius:4px; }}
QListWidget::item:selected {{ background:{ACCENT}; color:{DARK}; }}
QListWidget::item:hover:!selected {{ background:{BORDER}; }}
QSplitter::handle {{ background:{BORDER}; }}
"""

# ── paths ──────────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DEFAULT_CONFIG_FILE = APP_DIR / "config" / "domains_config.json"
APP_ICON = APP_DIR / "assets" / "icon.png"
USER_CONFIG_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "DomainSorter"
CONFIG_FILE = USER_CONFIG_DIR / "domains_config.json"

FALLBACK_CONFIG: dict = {
    "Finance":    ["invoice","balance sheet","revenue","profit","tax","budget","financial","accounting","payment","bank","expense","ledger","equity"],
    "Legal":      ["contract","agreement","clause","liability","court","law","attorney","jurisdiction","plaintiff","defendant","statute","compliance"],
    "Medical":    ["patient","diagnosis","treatment","clinical","hospital","physician","prescription","symptom","disease","health","surgery","therapy"],
    "Technology": ["software","hardware","api","cloud","server","database","algorithm","code","network","python","javascript","machine learning"],
    "Marketing":  ["campaign","brand","market","customer","advertising","social media","seo","conversion","audience","engagement","funnel"],
    "HR":         ["employee","recruitment","onboarding","performance review","payroll","benefits","workforce","salary","hiring","talent"],
    "Science":    ["experiment","hypothesis","data analysis","research","methodology","laboratory","findings","peer review","statistical"],
    "Education":  ["course","curriculum","student","learning","university","school","teacher","exam","lecture","syllabus","degree"],
    "Government": ["policy","regulation","government","ministry","official","public sector","legislation","federal","agency","mandate"],
    "Other":      [],
}

# ── config I/O ─────────────────────────────────────────────────────────────────
def _sanitize_config(raw: dict) -> dict:
    return {
        k: (v if isinstance(v, list) else [])
        for k, v in raw.items()
        if isinstance(k, str) and not k.startswith("_")
    }

def _load_default_config() -> dict:
    if DEFAULT_CONFIG_FILE.exists():
        try:
            raw = json.loads(DEFAULT_CONFIG_FILE.read_text("utf-8"))
            data = _sanitize_config(raw)
            if data:
                return data
        except Exception:
            pass
    return dict(FALLBACK_CONFIG)

def _ensure_user_config() -> None:
    if CONFIG_FILE.exists():
        return
    save_config(_load_default_config())

def load_config() -> dict:
    _ensure_user_config()
    if CONFIG_FILE.exists():
        try:
            raw = json.loads(CONFIG_FILE.read_text("utf-8"))
            data = _sanitize_config(raw)
            if data:
                return data
        except Exception:
            pass
    data = _load_default_config()
    save_config(data)
    return data

def save_config(data: dict) -> None:
    payload = {
        "_comment": "Domain Sorter — keyword configuration file.",
        "_instructions": [
            "Each key is a domain name; the value is a list of keywords.",
            "Add/remove domains or keywords freely — changes take effect on next Run.",
            "Keywords are case-insensitive and matched as substrings.",
            "Keep 'Other' as a catch-all fallback."
        ],
    }
    payload.update(data)
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(payload, indent=2, ensure_ascii=False), "utf-8")

# ── text extraction ────────────────────────────────────────────────────────────
MAX_CHARS = 3000
SUPPORTED  = {".pdf", ".txt", ".md", ".docx", ".doc", ".csv", ".json", ".xml"}

def extract_text(path: Path) -> str:
    sfx  = path.suffix.lower()
    text = ""
    if sfx == ".pdf":
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(path) as pdf:
                    for pg in pdf.pages[:10]:
                        t = pg.extract_text()
                        if t: text += t + "\n"
                        if len(text) >= MAX_CHARS: break
            except Exception: pass
        if not text.strip() and HAS_PYPDF:
            try:
                r = PdfReader(str(path))
                for pg in r.pages[:10]:
                    t = pg.extract_text()
                    if t: text += t + "\n"
                    if len(text) >= MAX_CHARS: break
            except Exception: pass
    elif sfx in (".docx", ".doc"):
        if HAS_DOCX:
            try:
                doc = DD(str(path))
                text = "\n".join(p.text for p in doc.paragraphs)
            except Exception: pass
    else:
        for enc in ("utf-8", "latin-1", "cp1252"):
            try: text = path.read_text(encoding=enc); break
            except Exception: continue
    return text.strip()[:MAX_CHARS]

# ── classifiers ────────────────────────────────────────────────────────────────
def classify_keywords(text: str, domains: list, kw_map: dict) -> dict:
    lower  = text.lower()
    scores = {d: sum(lower.count(k.lower()) for k in kw_map.get(d, []))
              for d in domains if d != "Other"}
    if not scores or max(scores.values()) == 0:
        return {"domain": "Other", "confidence": "low", "reason": "No keyword matches."}
    best = max(scores, key=lambda d: scores[d])
    sc   = scores[best]
    return {"domain": best,
            "confidence": "high" if sc >= 5 else ("medium" if sc >= 2 else "low"),
            "reason": f"Keyword score = {sc}"}

def classify_ai(text: str, filename: str, domains: list, client) -> dict:
    prompt = textwrap.dedent(f"""
        Classify this document into ONE domain from: {', '.join(domains)}
        Reply ONLY with JSON: {{"domain":"...","confidence":"high|medium|low","reason":"one sentence"}}
        Filename: {filename}
        Text:
        ---
        {text or '(no text)'}
        ---
    """).strip()
    resp = client.messages.create(
        model="claude-sonnet-4-20250514", max_tokens=200,
        messages=[{"role": "user", "content": prompt}])
    raw = resp.content[0].text.strip()
    raw = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
    result = json.loads(raw)
    if result.get("domain") not in domains:
        result["domain"] = "Other"
    return result

# ── worker ─────────────────────────────────────────────────────────────────────
class Signals(QObject):
    log      = pyqtSignal(str, str)
    progress = pyqtSignal(int, int)
    done     = pyqtSignal(dict)

class SortWorker(QRunnable):
    def __init__(self, cfg: dict):
        super().__init__()
        self.cfg     = cfg
        self.signals = Signals()
        self._stop   = False

    def stop(self): self._stop = True

    @pyqtSlot()
    def run(self):
        cfg     = self.cfg
        inp     = Path(cfg["input"])
        out     = Path(cfg["output"])
        domains = cfg["domains"]
        dry     = cfg["dry_run"]
        use_ai  = cfg["use_ai"]
        api_key = cfg["api_key"]
        recurse = cfg["recursive"]
        kw_map  = cfg["kw_map"]

        def log(msg, lvl="info"): self.signals.log.emit(msg, lvl)

        client = None
        if use_ai:
            if not HAS_ANTHROPIC:
                log("'anthropic' not installed — using keywords.", "warn"); use_ai = False
            elif not api_key:
                log("No API key — using keywords.", "warn"); use_ai = False
            else:
                try:
                    client = anthropic.Anthropic(api_key=api_key)
                    log("Claude AI client ready.", "ok")
                except Exception as e:
                    log(f"AI init failed: {e} — using keywords.", "warn"); use_ai = False

        glob_fn = inp.rglob if recurse else inp.glob
        files   = [f for f in glob_fn("*")
                   if f.is_file() and f.suffix.lower() in SUPPORTED]

        if not files:
            log("No supported files found in input folder.", "warn")
            self.signals.done.emit({}); return

        log(f"Found {len(files)} file(s)  ·  mode: {'AI (Claude)' if use_ai else 'keywords'}", "info")
        if dry: log("DRY RUN — no files will be copied.", "warn")
        log("─" * 52, "info")

        summary: dict = {d: [] for d in domains}
        errors = []

        for i, fp in enumerate(files, 1):
            if self._stop:
                log("⏹  Stopped by user.", "warn"); break
            self.signals.progress.emit(i, len(files))
            log(f"[{i}/{len(files)}]  {fp.name}", "info")
            try:
                text   = extract_text(fp)
                result = (classify_ai(text, fp.name, domains, client)
                          if (use_ai and client)
                          else classify_keywords(text, domains, kw_map))

                domain     = result["domain"]
                confidence = result.get("confidence", "?")
                reason     = result.get("reason", "")
                lvl = "ok" if confidence == "high" else ("warn" if confidence == "medium" else "info")
                log(f"  ↳  {domain}  [{confidence}]  — {reason}", lvl)

                if not dry:
                    dest_dir = out / domain
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest = dest_dir / fp.name
                    c = 2
                    while dest.exists():
                        dest = dest_dir / f"{fp.stem}_{c}{fp.suffix}"; c += 1
                    shutil.copy2(fp, dest)
                    log(f"  ✓  Copied → {domain}/{dest.name}", "ok")
                else:
                    log(f"  ~  Would copy → {domain}/{fp.name}", "info")

                summary[domain].append(fp.name)
            except Exception as e:
                log(f"  ✗  Error: {e}", "error")
                errors.append(f"{fp.name}: {e}")

        log("─" * 52, "info")
        log("DONE  —  Summary:", "ok")
        for d, fns in summary.items():
            if fns: log(f"  {d:<15}  {len(fns):>3} file(s)", "ok")
        if errors: log(f"  Errors: {len(errors)}", "error")
        self.signals.done.emit(summary)

# ── folder-picker row ──────────────────────────────────────────────────────────
class FolderRow(QWidget):
    def __init__(self, label: str, placeholder: str, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self); lay.setContentsMargins(0,0,0,0); lay.setSpacing(8)
        lbl = QLabel(label); lbl.setFixedWidth(60)
        lbl.setStyleSheet(f"color:{MUTED}; font-size:11px; letter-spacing:1px; font-weight:bold;")
        self.edit = QLineEdit(); self.edit.setPlaceholderText(placeholder)
        btn = QToolButton(); btn.setText("…"); btn.setFixedWidth(34)
        btn.setStyleSheet(f"""
            QToolButton{{background:{PANEL};border:1px solid {BORDER};border-radius:6px;
                color:{FG};font-weight:bold;font-size:15px;}}
            QToolButton:hover{{border-color:{ACCENT};color:{ACCENT};}}
        """)
        btn.clicked.connect(self._browse)
        lay.addWidget(lbl); lay.addWidget(self.edit); lay.addWidget(btn)

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "Select Folder",
                                              self.edit.text() or str(Path.home()))
        if d: self.edit.setText(d)

    @property
    def path(self) -> str: return self.edit.text().strip()

# ── log widget ─────────────────────────────────────────────────────────────────
class LogWidget(QPlainTextEdit):
    COLORS = {"info":"#b3c3d4", "ok":GREEN, "warn":YELLOW, "error":RED}

    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)

    def append_line(self, text: str, level: str = "info"):
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(self.COLORS.get(level, self.COLORS["info"])))
        cur = self.textCursor()
        cur.movePosition(QTextCursor.MoveOperation.End)
        if self.toPlainText(): cur.insertText("\n")
        cur.insertText(text, fmt)
        self.setTextCursor(cur)
        self.ensureCursorVisible()

# ── keyword editor dialog ──────────────────────────────────────────────────────
class KeywordEditor(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Domain & Keyword Editor")
        self.setMinimumSize(740, 540)
        self.resize(820, 580)
        self._data: dict = {}
        self._current_domain = None
        self._build_ui()
        self._load()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 14)
        root.setSpacing(12)

        # title
        hdr = QHBoxLayout()
        title = QLabel("⚙  DOMAIN & KEYWORD EDITOR")
        title.setStyleSheet(f"color:{ACCENT}; font-size:15px; font-weight:bold; letter-spacing:2px;")
        cfg_path = QLabel(f"{CONFIG_FILE}")
        cfg_path.setStyleSheet(f"color:{MUTED}; font-size:10px;")
        hdr.addWidget(title); hdr.addStretch(); hdr.addWidget(cfg_path)
        root.addLayout(hdr)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{BORDER};"); root.addWidget(sep)

        # splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(2)
        root.addWidget(splitter, 1)

        # ── LEFT: domain list ──────────────────────────────────────────────────
        lw = QWidget(); lw.setMaximumWidth(220)
        ll = QVBoxLayout(lw); ll.setContentsMargins(0,0,8,0); ll.setSpacing(6)

        lbl = QLabel("DOMAINS")
        lbl.setStyleSheet(f"color:{MUTED}; font-size:11px; letter-spacing:1.5px; font-weight:bold;")
        ll.addWidget(lbl)

        self.domain_list = QListWidget()
        self.domain_list.currentRowChanged.connect(self._on_domain_select)
        ll.addWidget(self.domain_list, 1)

        db = QHBoxLayout(); db.setSpacing(6)
        add_d = QPushButton("+ Domain"); add_d.setObjectName("addBtn")
        del_d = QPushButton("Delete");   del_d.setObjectName("dangerBtn")
        add_d.clicked.connect(self._add_domain)
        del_d.clicked.connect(self._del_domain)
        db.addWidget(add_d); db.addWidget(del_d)
        ll.addLayout(db)
        splitter.addWidget(lw)

        # ── RIGHT: keyword list ────────────────────────────────────────────────
        rw = QWidget()
        rl = QVBoxLayout(rw); rl.setContentsMargins(8,0,0,0); rl.setSpacing(6)

        kh = QHBoxLayout()
        self.kw_label = QLabel("KEYWORDS  —  select a domain")
        self.kw_label.setStyleSheet(f"color:{MUTED}; font-size:11px; letter-spacing:1.5px; font-weight:bold;")
        self.kw_count = QLabel("")
        self.kw_count.setStyleSheet(f"color:{MUTED}; font-size:10px;")
        kh.addWidget(self.kw_label); kh.addStretch(); kh.addWidget(self.kw_count)
        rl.addLayout(kh)

        self.kw_list = QListWidget()
        rl.addWidget(self.kw_list, 1)

        # add-keyword row
        ar = QHBoxLayout(); ar.setSpacing(6)
        self.kw_input = QLineEdit()
        self.kw_input.setPlaceholderText("Type keyword and press Add or Enter…")
        self.kw_input.returnPressed.connect(self._add_keyword)
        add_kw = QPushButton("+ Add"); add_kw.setObjectName("addBtn"); add_kw.setFixedWidth(70)
        add_kw.clicked.connect(self._add_keyword)
        ar.addWidget(self.kw_input); ar.addWidget(add_kw)
        rl.addLayout(ar)

        kb = QHBoxLayout(); kb.setSpacing(6)
        del_kw = QPushButton("Delete Selected"); del_kw.setObjectName("dangerBtn")
        clr_kw = QPushButton("Clear All");       clr_kw.setObjectName("dangerBtn")
        del_kw.clicked.connect(self._del_keyword)
        clr_kw.clicked.connect(self._clear_keywords)
        kb.addWidget(del_kw); kb.addWidget(clr_kw); kb.addStretch()
        rl.addLayout(kb)
        splitter.addWidget(rw)
        splitter.setStretchFactor(0, 0); splitter.setStretchFactor(1, 1)

        # bottom
        bot = QHBoxLayout(); bot.setSpacing(10)
        rel = QPushButton("↺  Reload from file"); rel.clicked.connect(self._load)
        sav = QPushButton("💾  Save to file")
        sav.setStyleSheet(f"""
            QPushButton{{background:{ACCENT};border:none;color:{DARK};
                border-radius:6px;padding:8px 20px;font-weight:bold;}}
            QPushButton:hover{{background:#79c0ff;}}
        """)
        sav.clicked.connect(self._save)
        cls = QPushButton("Close"); cls.clicked.connect(self.accept)
        bot.addWidget(rel); bot.addStretch(); bot.addWidget(sav); bot.addWidget(cls)
        root.addLayout(bot)

    # ── helpers ────────────────────────────────────────────────────────────────
    def _load(self):
        self._data = load_config()
        self._refresh_domain_list()

    def _save(self):
        self._sync_current()
        try:
            save_config(self._data)
            QMessageBox.information(self, "Saved", f"Config saved to:\n{CONFIG_FILE}")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))

    def _sync_current(self):
        if self._current_domain is not None:
            self._data[self._current_domain] = [
                self.kw_list.item(i).text() for i in range(self.kw_list.count())
            ]

    def _refresh_domain_list(self):
        prev = self._current_domain
        self.domain_list.blockSignals(True)
        self.domain_list.clear()
        for d in self._data:
            self.domain_list.addItem(QListWidgetItem(d))
        self.domain_list.blockSignals(False)
        for i in range(self.domain_list.count()):
            if self.domain_list.item(i).text() == prev:
                self.domain_list.setCurrentRow(i); return
        if self.domain_list.count():
            self.domain_list.setCurrentRow(0)

    def _refresh_kw_list(self, domain: str):
        self.kw_list.clear()
        for kw in self._data.get(domain, []):
            self.kw_list.addItem(kw)
        self.kw_label.setText(f"KEYWORDS  —  {domain}")
        n = self.kw_list.count()
        self.kw_count.setText(f"{n} keyword{'s' if n != 1 else ''}")

    # ── domain actions ─────────────────────────────────────────────────────────
    def _on_domain_select(self, row: int):
        if row < 0: return
        self._sync_current()
        self._current_domain = self.domain_list.item(row).text()
        self._refresh_kw_list(self._current_domain)

    def _add_domain(self):
        name, ok = QInputDialog.getText(self, "New Domain", "Domain name:")
        name = name.strip()
        if not ok or not name: return
        if name in self._data:
            QMessageBox.warning(self, "Duplicate", f"'{name}' already exists."); return
        self._sync_current()
        self._data[name] = []
        self._refresh_domain_list()
        for i in range(self.domain_list.count()):
            if self.domain_list.item(i).text() == name:
                self.domain_list.setCurrentRow(i); break

    def _del_domain(self):
        row = self.domain_list.currentRow()
        if row < 0: return
        domain = self.domain_list.item(row).text()
        if domain == "Other":
            QMessageBox.warning(self, "Protected",
                "'Other' is the catch-all fallback and cannot be deleted."); return
        ans = QMessageBox.question(self, "Delete Domain",
            f"Delete domain '{domain}' and all its keywords?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if ans != QMessageBox.StandardButton.Yes: return
        self._data.pop(domain, None)
        if self._current_domain == domain:
            self._current_domain = None
            self.kw_list.clear()
            self.kw_label.setText("KEYWORDS  —  select a domain")
            self.kw_count.setText("")
        self._refresh_domain_list()

    # ── keyword actions ────────────────────────────────────────────────────────
    def _add_keyword(self):
        if self._current_domain is None:
            QMessageBox.warning(self, "No Domain", "Select a domain first."); return
        kw = self.kw_input.text().strip()
        if not kw: return
        existing = [self.kw_list.item(i).text().lower() for i in range(self.kw_list.count())]
        if kw.lower() in existing:
            QMessageBox.information(self, "Duplicate", f"'{kw}' already in list."); return
        self.kw_list.addItem(kw)
        self.kw_input.clear()
        n = self.kw_list.count()
        self.kw_count.setText(f"{n} keyword{'s' if n != 1 else ''}")

    def _del_keyword(self):
        for item in self.kw_list.selectedItems():
            self.kw_list.takeItem(self.kw_list.row(item))
        n = self.kw_list.count()
        self.kw_count.setText(f"{n} keyword{'s' if n != 1 else ''}")

    def _clear_keywords(self):
        if self._current_domain is None: return
        ans = QMessageBox.question(self, "Clear All",
            f"Clear all keywords from '{self._current_domain}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if ans == QMessageBox.StandardButton.Yes:
            self.kw_list.clear()
            self.kw_count.setText("0 keywords")

# ── main window ────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Domain Sorter")
        if APP_ICON.exists():
            self.setWindowIcon(QIcon(str(APP_ICON)))
        self.setMinimumSize(920, 660)
        self.resize(1080, 740)
        self._worker   = None
        self._pool     = QThreadPool.globalInstance()
        self._settings = QSettings("DomainSorter", "App")
        self._domain_checks: dict = {}
        self._build_ui()
        self._restore_settings()
        self._reload_domains()

    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        rl = QVBoxLayout(root)
        rl.setContentsMargins(20, 16, 20, 12); rl.setSpacing(14)

        # header
        hdr_row = QHBoxLayout()
        hdr = QLabel("⬡  FILE DOMAIN SORTER")
        hdr.setStyleSheet(f"color:{ACCENT}; font-size:18px; font-weight:bold; letter-spacing:3px;")
        about_btn = QPushButton("About Me")
        about_btn.setStyleSheet(f"""
            QPushButton{{background:transparent; border:1px solid {BORDER};
                border-radius:6px; padding:6px 12px; color:{MUTED}; font-size:11px;}}
            QPushButton:hover{{border-color:{ACCENT}; color:{ACCENT};}}
        """)
        about_btn.clicked.connect(self._show_about)
        edit_btn = QPushButton("⚙  Edit Domains & Keywords")
        edit_btn.setStyleSheet(f"""
            QPushButton{{background:transparent; border:1px solid {BORDER};
                border-radius:6px; padding:6px 14px; color:{MUTED}; font-size:11px;}}
            QPushButton:hover{{border-color:{ACCENT}; color:{ACCENT};}}
        """)
        edit_btn.clicked.connect(self._open_editor)
        hdr_row.addWidget(hdr); hdr_row.addStretch(); hdr_row.addWidget(about_btn); hdr_row.addWidget(edit_btn)
        rl.addLayout(hdr_row)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{BORDER};"); rl.addWidget(sep)

        # splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(2); rl.addWidget(splitter, 1)

        # ── LEFT ──────────────────────────────────────────────────────────────
        left = QWidget(); left.setMaximumWidth(400)
        ll = QVBoxLayout(left); ll.setContentsMargins(0,0,8,0); ll.setSpacing(14)
        splitter.addWidget(left)

        # folders
        fg = QGroupBox("FOLDERS"); fg_l = QVBoxLayout(fg); fg_l.setSpacing(10)
        self.input_row  = FolderRow("INPUT",  "Folder with files to sort…")
        self.output_row = FolderRow("OUTPUT", "Destination folder…")
        fg_l.addWidget(self.input_row); fg_l.addWidget(self.output_row)
        ll.addWidget(fg)

        # domains
        dg = QGroupBox("ACTIVE DOMAINS"); dg_l = QVBoxLayout(dg); dg_l.setSpacing(6)
        note = QLabel("Check domains to sort into. Use ⚙ to add/remove/edit keywords.")
        note.setStyleSheet(f"color:{MUTED}; font-size:11px;"); note.setWordWrap(True)
        dg_l.addWidget(note)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setMaximumHeight(160)
        self._domains_container = QWidget()
        self._domains_layout    = QVBoxLayout(self._domains_container)
        self._domains_layout.setSpacing(4); self._domains_layout.setContentsMargins(0,0,0,0)
        scroll.setWidget(self._domains_container)
        dg_l.addWidget(scroll)
        ll.addWidget(dg)

        # options
        og = QGroupBox("OPTIONS"); og_l = QVBoxLayout(og); og_l.setSpacing(10)
        self.ai_check  = QCheckBox("Use AI (Claude API) for classification")
        self.dry_check = QCheckBox("Dry run (preview only — no files copied)")
        self.rec_check = QCheckBox("Recursive (include sub-folders)")
        self.dry_check.setChecked(True)
        og_l.addWidget(self.ai_check)
        og_l.addWidget(self.dry_check)
        og_l.addWidget(self.rec_check)

        ar = QHBoxLayout()
        al = QLabel("API KEY"); al.setFixedWidth(60)
        al.setStyleSheet(f"color:{MUTED}; font-size:11px; letter-spacing:1px; font-weight:bold;")
        self.api_edit = QLineEdit()
        self.api_edit.setPlaceholderText("sk-ant-… (or set ANTHROPIC_API_KEY env var)")
        self.api_edit.setEchoMode(QLineEdit.EchoMode.Password)
        ar.addWidget(al); ar.addWidget(self.api_edit)
        og_l.addLayout(ar)
        ll.addWidget(og)
        ll.addStretch()

        # ── RIGHT: log ─────────────────────────────────────────────────────────
        right = QWidget()
        rl2 = QVBoxLayout(right); rl2.setContentsMargins(8,0,0,0); rl2.setSpacing(8)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0); splitter.setStretchFactor(1, 1)

        log_hdr = QHBoxLayout()
        ll_lbl = QLabel("ACTIVITY LOG")
        ll_lbl.setStyleSheet(f"color:{MUTED}; font-size:11px; letter-spacing:1.5px; font-weight:bold;")
        clr = QPushButton("Clear"); clr.setFixedWidth(60)
        clr.setStyleSheet(f"""
            QPushButton{{background:transparent;border:1px solid {BORDER};
                border-radius:5px;color:{MUTED};font-size:11px;padding:3px 8px;}}
            QPushButton:hover{{border-color:{MUTED};color:{FG};}}
        """)
        clr.clicked.connect(lambda: self.log.clear())
        log_hdr.addWidget(ll_lbl); log_hdr.addStretch(); log_hdr.addWidget(clr)
        rl2.addLayout(log_hdr)

        self.log = LogWidget(); rl2.addWidget(self.log, 1)

        self.progress = QProgressBar(); self.progress.setFixedHeight(6)
        self.progress.setRange(0, 100); self.progress.setValue(0)
        rl2.addWidget(self.progress)

        # bottom buttons
        bot = QHBoxLayout(); bot.setSpacing(10)
        self.run_btn = QPushButton("▶  Run Sorter")
        self.run_btn.setObjectName("runBtn"); self.run_btn.setFixedHeight(44)
        self.run_btn.clicked.connect(self._run)
        self.stop_btn = QPushButton("■  Stop")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setFixedHeight(44); self.stop_btn.setFixedWidth(100)
        self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self._stop)
        bot.addWidget(self.run_btn, 1); bot.addWidget(self.stop_btn)
        rl.addLayout(bot)

        self.status = QStatusBar(); self.setStatusBar(self.status)
        self.status.showMessage("Ready  —  select folders and press Run")

    # ── domain checkbox panel ──────────────────────────────────────────────────
    def _reload_domains(self):
        cfg = load_config()
        prev = {d for d, cb in self._domain_checks.items() if cb.isChecked()}
        # clear old
        while self._domains_layout.count():
            item = self._domains_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._domain_checks.clear()
        for domain in cfg:
            cb = QCheckBox(domain)
            cb.setChecked(True if not prev else domain in prev)
            self._domain_checks[domain] = cb
            self._domains_layout.addWidget(cb)
        self._domains_layout.addStretch()

    # ── editor ─────────────────────────────────────────────────────────────────
    def _open_editor(self):
        dlg = KeywordEditor(self)
        dlg.exec()
        self._reload_domains()
        self.log.append_line("Config reloaded from file.", "ok")
        self.status.showMessage("Domain config reloaded.")

    def _show_about(self):
        QMessageBox.information(
            self,
            "About Me",
            "File Domain Sorter\n\nDevelopper: Fouad El Azbi.\n\nEmail: EAF.microservice@gmail.com\n\nVersion: 1.0.0\n\n &copy; 2026 EAF microservice | All right reserved."
        )

    # ── run ────────────────────────────────────────────────────────────────────
    def _run(self):
        inp = self.input_row.path; out = self.output_row.path
        if not inp or not Path(inp).exists():
            QMessageBox.warning(self, "Input Missing", "Please select a valid input folder."); return
        if not out:
            QMessageBox.warning(self, "Output Missing", "Please select an output folder."); return
        domains = [d for d, cb in self._domain_checks.items() if cb.isChecked()]
        if not domains:
            QMessageBox.warning(self, "No Domains", "Select at least one domain."); return

        kw_map = load_config()   # always fresh at run time
        self.log.append_line(f"Config: {CONFIG_FILE}", "info")

        self._save_settings()
        self.log.clear()
        self.progress.setValue(0)
        self.run_btn.setEnabled(False); self.stop_btn.setEnabled(True)
        self.status.showMessage("Running…")

        cfg = {
            "input": inp, "output": out, "domains": domains,
            "dry_run":   self.dry_check.isChecked(),
            "use_ai":    self.ai_check.isChecked(),
            "api_key":   self.api_edit.text().strip() or os.environ.get("ANTHROPIC_API_KEY", ""),
            "recursive": self.rec_check.isChecked(),
            "kw_map":    kw_map,
        }
        self._worker = SortWorker(cfg)
        self._worker.signals.log.connect(self._on_log)
        self._worker.signals.progress.connect(self._on_progress)
        self._worker.signals.done.connect(self._on_done)
        self._pool.start(self._worker)

    def _stop(self):
        if self._worker: self._worker.stop()
        self.stop_btn.setEnabled(False)

    @pyqtSlot(str, str)
    def _on_log(self, msg, level): self.log.append_line(msg, level)

    @pyqtSlot(int, int)
    def _on_progress(self, cur, total):
        self.progress.setValue(int(cur / total * 100) if total else 0)
        self.status.showMessage(f"Processing {cur} / {total}…")

    @pyqtSlot(dict)
    def _on_done(self, summary):
        self.run_btn.setEnabled(True); self.stop_btn.setEnabled(False)
        self.progress.setValue(100)
        total = sum(len(v) for v in summary.values())
        self.status.showMessage(f"Done — {total} file(s) processed.")

    # ── persistence ────────────────────────────────────────────────────────────
    def _save_settings(self):
        s = self._settings
        s.setValue("input",     self.input_row.edit.text())
        s.setValue("output",    self.output_row.edit.text())
        s.setValue("dry_run",   self.dry_check.isChecked())
        s.setValue("recursive", self.rec_check.isChecked())
        s.setValue("use_ai",    self.ai_check.isChecked())

    def _restore_settings(self):
        s = self._settings
        self.input_row.edit.setText(s.value("input",  ""))
        self.output_row.edit.setText(s.value("output", ""))
        self.dry_check.setChecked(s.value("dry_run",   True,  type=bool))
        self.rec_check.setChecked(s.value("recursive", False, type=bool))
        self.ai_check.setChecked(s.value("use_ai",     False, type=bool))

# ── entry point ────────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE)
    if APP_ICON.exists():
        app.setWindowIcon(QIcon(str(APP_ICON)))
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window,     QColor(DARK))
    pal.setColor(QPalette.ColorRole.WindowText, QColor(FG))
    pal.setColor(QPalette.ColorRole.Base,       QColor(PANEL))
    pal.setColor(QPalette.ColorRole.Text,       QColor(FG))
    app.setPalette(pal)

    win = MainWindow()
    win.show()
    win.log.append_line("File Domain Sorter  ready.", "ok")
    win.log.append_line(f"Config: {CONFIG_FILE}", "info")
    win.log.append_line("Select folders, configure options, then press  ▶ Run.", "info")
    win.log.append_line("─" * 52, "info")
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
