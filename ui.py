# -*- coding: utf-8 -*-
"""CustomTkinter arayüzü — kompakt, tek ekranda her şey görünür."""

from __future__ import annotations

import datetime
import os
import platform
import webbrowser
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog

from converter import ConversionResult, run_batch_in_thread
from utils import (
    APP_VERSION,
    LIBREOFFICE_DOWNLOAD_URL,
    SUPPORTED_EXTENSIONS,
    is_supported_file,
    parse_dnd_file_list,
)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _TKDND = True
except ImportError:
    _TKDND = False

# ── Renkler ──────────────────────────────────────────────────────────────────
_ACCENT = "#3b82f6"
_ACCENT_HOVER = "#2563eb"
_GREEN = "#22c55e"
_RED = "#ef4444"
_ORANGE = "#f59e0b"

_BG = ("gray96", "#111111")
_SURFACE = ("white", "#1a1a1a")
_SURFACE2 = ("gray93", "#222222")
_BORDER = ("gray82", "#2e2e2e")
_TEXT = ("gray10", "gray90")
_TEXT2 = ("gray45", "gray55")
_R = 12


def _mono(size: int = 12) -> ctk.CTkFont:
    fam = "SF Mono" if platform.system() == "Darwin" else "Consolas"
    try:
        return ctk.CTkFont(family=fam, size=size)
    except Exception:
        return ctk.CTkFont(size=size)


# ── LibreOffice uyarı penceresi ──────────────────────────────────────────────

class LibreOfficeRequiredWindow(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SwiftPDF — LibreOffice gerekli")
        self.geometry("480x300")
        self.resizable(False, False)
        self.configure(fg_color=_BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        f = ctk.CTkFrame(self, fg_color=_SURFACE, corner_radius=_R, border_width=1, border_color=_BORDER)
        f.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)
        f.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(f, text="⚠️  LibreOffice bulunamadı", font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=24, pady=(24, 8))
        ctk.CTkLabel(f, text=(
            "Bu uygulama Office → PDF dönüşümü için LibreOffice gerektirir.\n"
            "Kurun ve uygulamayı yeniden başlatın."),
            font=ctk.CTkFont(size=14), text_color=_TEXT2, justify="left", wraplength=400,
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 20))

        btns = ctk.CTkFrame(f, fg_color="transparent")
        btns.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 24))
        ctk.CTkButton(btns, text="LibreOffice İndir", height=40, corner_radius=10,
                       fg_color=_ACCENT, hover_color=_ACCENT_HOVER, font=ctk.CTkFont(size=14, weight="bold"),
                       command=lambda: webbrowser.open(LIBREOFFICE_DOWNLOAD_URL)).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btns, text="Kapat", height=40, corner_radius=10,
                       fg_color=("gray80", "gray30"), hover_color=("gray70", "gray38"),
                       command=self.destroy).pack(side="left")
        self.protocol("WM_DELETE_WINDOW", self.destroy)


def show_libreoffice_required_and_wait() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    LibreOfficeRequiredWindow().mainloop()


# ── Ana arayüz mixin ────────────────────────────────────────────────────────

class SwiftPDFUIMixin:
    _soffice_path: str
    _file_order: list[str]
    _file_rows: dict[str, dict]
    _converting: bool

    # ── Arayüz oluşturma ────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.title(f"SwiftPDF v{APP_VERSION}")
        self.geometry("860x700")
        self.minsize(700, 560)
        self.configure(fg_color=_BG)

        self.grid_columnconfigure(0, weight=1)
        # Satır 0: başlık        ~44 px
        # Satır 1: bırakma       ~56 px
        # Satır 2: dosya listesi ~geri kalan (weight 3)
        # Satır 3: çıktı klasörü ~44 px
        # Satır 4: ilerleme      ~36 px
        # Satır 5: günlük        ~geri kalan (weight 1)
        # Satır 6: buton         ~52 px
        self.grid_rowconfigure(2, weight=3)
        self.grid_rowconfigure(5, weight=1)

        self._build_header()      # row 0
        self._build_drop_strip()  # row 1
        self._build_file_list()   # row 2
        self._build_output_row()  # row 3
        self._build_progress()    # row 4
        self._build_log()         # row 5
        self._build_convert_btn() # row 6

    def _build_header(self) -> None:
        hdr = ctk.CTkFrame(self, fg_color="transparent", height=44)
        hdr.grid(row=0, column=0, sticky="ew", padx=20, pady=(14, 6))
        hdr.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(hdr, text=f"SwiftPDF", font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=0, column=0, sticky="w")

        ctk.CTkLabel(hdr, text=f"v{APP_VERSION}", font=ctk.CTkFont(size=11, weight="bold"),
                      text_color="white", fg_color=_ACCENT, corner_radius=6, padx=8, pady=2,
                      ).grid(row=0, column=1, sticky="w", padx=(10, 0))

        exts = "  ·  ".join(sorted(SUPPORTED_EXTENSIONS))
        ctk.CTkLabel(hdr, text=exts, font=ctk.CTkFont(size=11), text_color=_TEXT2).grid(
            row=0, column=2, sticky="e", padx=(0, 12))

        self._theme_var = ctk.StringVar(value="Koyu")
        ctk.CTkSegmentedButton(
            hdr, values=["Koyu", "Açık"], variable=self._theme_var,
            command=self._on_theme, width=140, height=32, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=3, sticky="e")
        ctk.set_appearance_mode("dark")

    def _build_drop_strip(self) -> None:
        strip = ctk.CTkFrame(self, fg_color=_SURFACE, corner_radius=_R, border_width=1, border_color=_BORDER, height=56)
        strip.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 6))
        strip.grid_columnconfigure(1, weight=1)
        strip.grid_propagate(False)

        ctk.CTkLabel(strip, text="📂", font=ctk.CTkFont(size=22)).grid(row=0, column=0, padx=(16, 8), pady=10)

        dnd_hint = "Dosyaları buraya sürükleyin veya →" if _TKDND else "Dosya eklemek için →"
        ctk.CTkLabel(strip, text=dnd_hint, font=ctk.CTkFont(size=14), text_color=_TEXT2,
                      anchor="w").grid(row=0, column=1, sticky="ew")

        self._add_btn = ctk.CTkButton(
            strip, text="+ Dosya ekle", width=130, height=38, corner_radius=10,
            fg_color=_ACCENT, hover_color=_ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"), command=self._browse_files,
        )
        self._add_btn.grid(row=0, column=2, padx=(8, 8), pady=8)

        self._drop_widget = strip

    def _build_file_list(self) -> None:
        wrap = ctk.CTkFrame(self, fg_color=_SURFACE, corner_radius=_R, border_width=1, border_color=_BORDER)
        wrap.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 6))
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(wrap, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))
        top.grid_columnconfigure(0, weight=1)

        self._list_title = ctk.CTkLabel(top, text="Dosya kuyruğu", font=ctk.CTkFont(size=14, weight="bold"))
        self._list_title.grid(row=0, column=0, sticky="w")

        self._file_count = ctk.CTkLabel(top, text="0 dosya", font=ctk.CTkFont(size=12), text_color=_TEXT2)
        self._file_count.grid(row=0, column=1, sticky="e", padx=(0, 8))

        ctk.CTkButton(top, text="Temizle", width=80, height=28, corner_radius=8,
                       fg_color=("gray82", "#2a2a2a"), hover_color=("gray72", "#333333"),
                       font=ctk.CTkFont(size=12), command=self._clear_list).grid(row=0, column=2, sticky="e")

        self._scroll = ctk.CTkScrollableFrame(wrap, fg_color="transparent", corner_radius=0)
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self._scroll.grid_columnconfigure(0, weight=1)

        self._empty_msg = ctk.CTkLabel(
            self._scroll, text="Henüz dosya eklenmedi.\nYukarıdan dosya sürükleyin veya «+ Dosya ekle» ile seçin.",
            font=ctk.CTkFont(size=14), text_color=_TEXT2, justify="center",
        )
        self._empty_msg.pack(pady=40)

    def _build_output_row(self) -> None:
        row = ctk.CTkFrame(self, fg_color=_SURFACE, corner_radius=_R, border_width=1, border_color=_BORDER, height=50)
        row.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 6))
        row.grid_columnconfigure(1, weight=1)
        row.grid_propagate(False)

        ctk.CTkLabel(row, text="Kayıt yeri:", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=(16, 8), pady=10)

        home = str(Path.home() / "Desktop")
        self._out_var = ctk.StringVar(value=home)
        self._out_entry = ctk.CTkEntry(row, textvariable=self._out_var, height=34, corner_radius=8,
                                        font=ctk.CTkFont(size=13), border_width=1)
        self._out_entry.grid(row=0, column=1, sticky="ew", pady=8)

        ctk.CTkButton(row, text="Seç…", width=80, height=34, corner_radius=8,
                       fg_color=("gray82", "#2a2a2a"), hover_color=("gray72", "#333333"),
                       font=ctk.CTkFont(size=12, weight="bold"), command=self._browse_out
                       ).grid(row=0, column=2, padx=(8, 12), pady=8)

    def _build_progress(self) -> None:
        row = ctk.CTkFrame(self, fg_color="transparent", height=36)
        row.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 4))
        row.grid_columnconfigure(0, weight=1)

        self._progress_label = ctk.CTkLabel(row, text="Hazır", font=ctk.CTkFont(size=12, weight="bold"),
                                             text_color=_TEXT2)
        self._progress_label.grid(row=0, column=0, sticky="w")

        self._pct_label = ctk.CTkLabel(row, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color=_TEXT2)
        self._pct_label.grid(row=0, column=1, sticky="e")

        self._progress = ctk.CTkProgressBar(self, height=8, corner_radius=4, progress_color=_ACCENT)
        self._progress.grid(row=4, column=0, sticky="sew", padx=20, pady=(22, 0))
        self._progress.set(0)

    def _build_log(self) -> None:
        self._log = ctk.CTkTextbox(self, height=100, corner_radius=_R, font=_mono(12),
                                    border_width=1, border_color=_BORDER, fg_color=_SURFACE,
                                    text_color=_TEXT)
        self._log.grid(row=5, column=0, sticky="nsew", padx=20, pady=(6, 6))
        self._log_line(f"SwiftPDF v{APP_VERSION} — Uygulama hazır.")

    def _build_convert_btn(self) -> None:
        self._convert_btn = ctk.CTkButton(
            self, text="PDF'e dönüştür", height=48, corner_radius=_R,
            fg_color=_ACCENT, hover_color=_ACCENT_HOVER,
            font=ctk.CTkFont(size=15, weight="bold"), command=self._start_convert,
        )
        self._convert_btn.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 16))

    # ── Sürükle-bırak ───────────────────────────────────────────────────────

    def _setup_dnd(self) -> None:
        if not _TKDND:
            return
        try:
            for w in (self, self._drop_widget, self._scroll):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            self._log_line("Sürükle-bırak kaydı başarısız.")

    def _on_drop(self, event) -> None:
        self._add_paths(parse_dnd_file_list(event.data))

    # ── Tema ─────────────────────────────────────────────────────────────────

    def _on_theme(self, choice: str) -> None:
        ctk.set_appearance_mode("light" if choice == "Açık" else "dark")

    # ── Günlük ───────────────────────────────────────────────────────────────

    def _log_line(self, text: str) -> None:
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}]  {text}\n")
        self._log.see("end")

    # ── Dosya ekleme / kaldırma ──────────────────────────────────────────────

    def _browse_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Office dosyaları seçin",
            filetypes=[("Office", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt"), ("Tüm dosyalar", "*.*")],
        )
        if files:
            self._add_paths(list(files))

    def _browse_out(self) -> None:
        d = filedialog.askdirectory(title="Çıktı klasörü")
        if d:
            self._out_var.set(d)

    def _add_paths(self, paths: list[str]) -> None:
        for p in paths:
            p = os.path.normpath(p)
            if not os.path.isfile(p):
                self._log_line(f"Atlandı (dosya yok): {Path(p).name}")
                continue
            if not is_supported_file(p):
                self._log_line(f"Desteklenmeyen: {Path(p).name}")
                continue
            if p in self._file_rows:
                continue
            if not self._file_order:
                self._empty_msg.pack_forget()
            self._file_order.append(p)
            self._create_file_card(p)
        self._update_count()

    def _create_file_card(self, p: str) -> None:
        ext = Path(p).suffix.lower()
        tag, color = self._badge_info(ext)

        card = ctk.CTkFrame(self._scroll, fg_color=_SURFACE2, corner_radius=10,
                             border_width=1, border_color=_BORDER)
        card.pack(fill="x", pady=3, padx=4)
        card.grid_columnconfigure(1, weight=1)

        # Sol renk şeridi
        ctk.CTkFrame(card, width=4, fg_color=color, corner_radius=2).grid(
            row=0, column=0, rowspan=2, sticky="ns", padx=(8, 0), pady=8)

        # Dosya adı + klasör yolu
        info_frame = ctk.CTkFrame(card, fg_color="transparent")
        info_frame.grid(row=0, column=1, rowspan=2, sticky="ew", padx=(10, 4), pady=8)
        info_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(info_frame, text=Path(p).name, anchor="w",
                      font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="ew")

        parent = str(Path(p).parent)
        if len(parent) > 60:
            parent = "…" + parent[-57:]
        ctk.CTkLabel(info_frame, text=parent, anchor="w",
                      font=ctk.CTkFont(size=11), text_color=_TEXT2).grid(row=1, column=0, sticky="ew")

        # Tür rozeti
        badge = ctk.CTkLabel(card, text=f" {tag} ", font=ctk.CTkFont(size=11, weight="bold"),
                              fg_color=color, text_color="white", corner_radius=6)
        badge.grid(row=0, column=2, padx=(4, 4), pady=(8, 2), sticky="ne")

        # Durum etiketi
        status = ctk.CTkLabel(card, text="Bekliyor", font=ctk.CTkFont(size=11, weight="bold"),
                               text_color=_TEXT2)
        status.grid(row=1, column=2, padx=(4, 4), pady=(0, 8), sticky="se")

        # Kaldır düğmesi
        rm = ctk.CTkButton(card, text="✕", width=32, height=32, corner_radius=8,
                            fg_color="transparent", hover_color=("gray78", "#333333"),
                            font=ctk.CTkFont(size=16), text_color=_TEXT2,
                            command=lambda fp=p: self._remove_one(fp))
        rm.grid(row=0, column=3, rowspan=2, padx=(0, 8), pady=8)

        self._file_rows[p] = {"card": card, "frame": card, "status": status, "badge": badge, "rm": rm}

    @staticmethod
    def _badge_info(ext: str) -> tuple[str, str]:
        if ext in (".doc", ".docx"):
            return ("DOC", "#2563eb")
        if ext in (".xls", ".xlsx"):
            return ("XLS", "#16a34a")
        if ext in (".ppt", ".pptx"):
            return ("PPT", "#ea580c")
        return ("?", "gray50")

    def _remove_one(self, path: str) -> None:
        if self._converting:
            return
        info = self._file_rows.pop(path, None)
        if info:
            info["frame"].destroy()
        if path in self._file_order:
            self._file_order.remove(path)
        self._update_count()
        if not self._file_order:
            self._empty_msg.pack(pady=40)

    def _clear_list(self) -> None:
        if self._converting:
            return
        for p in list(self._file_order):
            info = self._file_rows.pop(p, None)
            if info:
                info["frame"].destroy()
        self._file_order.clear()
        self._update_count()
        self._empty_msg.pack(pady=40)

    def _update_count(self) -> None:
        n = len(self._file_order)
        self._file_count.configure(text=f"{n} dosya")
        if not self._converting:
            if n:
                self._progress_label.configure(text=f"{n} dosya kuyrukta")
            else:
                self._progress_label.configure(text="Hazır")
                self._pct_label.configure(text="")

    def _set_busy(self, busy: bool) -> None:
        self._converting = busy
        st = "disabled" if busy else "normal"
        self._convert_btn.configure(state=st)
        self._add_btn.configure(state=st)
        for info in self._file_rows.values():
            info["rm"].configure(state=st)

    # ── Dönüştürme ───────────────────────────────────────────────────────────

    def _start_convert(self) -> None:
        if self._converting:
            return
        files = list(self._file_order)
        if not files:
            self._log_line("Önce dosya ekleyin.")
            return
        out_dir = self._out_var.get().strip()
        if not out_dir:
            self._log_line("Çıktı klasörü seçin.")
            return
        try:
            Path(out_dir).mkdir(parents=True, exist_ok=True)
        except OSError as e:
            self._log_line(f"Klasör hatası: {e}")
            return

        total = len(files)
        # Tüm kartları sıfırla
        for p in files:
            if p in self._file_rows:
                fi = self._file_rows[p]
                fi["card"].configure(border_color=_BORDER, border_width=1)
                fi["status"].configure(text="Bekliyor", text_color=_TEXT2)

        # İlk dosyayı aktif göster
        if files[0] in self._file_rows:
            fi = self._file_rows[files[0]]
            fi["card"].configure(border_color=_ACCENT, border_width=2)
            fi["status"].configure(text="Dönüştürülüyor…", text_color=_ACCENT)

        self._set_busy(True)
        self._progress.set(0)
        self._progress_label.configure(text=f"0 / {total}")
        self._pct_label.configure(text="0%")
        self._log_line(f"Dönüşüm başladı — {total} dosya")

        def on_progress(i: int, tot: int, res: ConversionResult) -> None:
            def ui() -> None:
                pct = int(100 * i / tot) if tot else 0
                self._progress.set(i / tot if tot else 0)
                self._progress_label.configure(text=f"{i} / {tot}")
                self._pct_label.configure(text=f"{pct}%")

                ip = res.input_path
                short = Path(ip).name
                if ip in self._file_rows:
                    fi = self._file_rows[ip]
                    if res.ok:
                        fi["card"].configure(border_color=_GREEN, border_width=2)
                        fi["status"].configure(text="PDF hazır ✓", text_color=_GREEN)
                    else:
                        fi["card"].configure(border_color=_RED, border_width=2)
                        fi["status"].configure(text="Hata ✗", text_color=_RED)

                if res.ok:
                    self._log_line(f"✓  {short}")
                else:
                    self._log_line(f"✗  {short} — {res.message}")

                # Sıradaki dosyayı aktif yap
                if i < len(files):
                    nxt = files[i]
                    if nxt in self._file_rows:
                        nfi = self._file_rows[nxt]
                        nfi["card"].configure(border_color=_ACCENT, border_width=2)
                        nfi["status"].configure(text="Dönüştürülüyor…", text_color=_ACCENT)

            self.after(0, ui)

        def on_done(_results: list[ConversionResult]) -> None:
            def finish() -> None:
                self._progress.set(1.0)
                self._pct_label.configure(text="100%")
                ok = sum(1 for r in _results if r.ok)
                fail = len(_results) - ok
                msg = f"Tamamlandı — {ok} başarılı"
                if fail:
                    msg += f", {fail} hata"
                self._progress_label.configure(text=msg)
                self._set_busy(False)
                self._log_line(msg)
            self.after(0, finish)

        run_batch_in_thread(self._soffice_path, files, out_dir, on_progress=on_progress, on_done=on_done)


# ── Pencere sınıfları ────────────────────────────────────────────────────────

if _TKDND:
    class SwiftPDFApp(ctk.CTk, TkinterDnD.DnDWrapper, SwiftPDFUIMixin):
        def __init__(self, soffice_path: str) -> None:
            super().__init__()
            self.TkdndVersion = TkinterDnD._require(self)
            self._soffice_path = soffice_path
            self._file_order = []
            self._file_rows = {}
            self._converting = False
            self._build_ui()
            self._setup_dnd()
else:
    class SwiftPDFApp(ctk.CTk, SwiftPDFUIMixin):
        def __init__(self, soffice_path: str) -> None:
            super().__init__()
            self._soffice_path = soffice_path
            self._file_order = []
            self._file_rows = {}
            self._converting = False
            self._build_ui()


def run_app(soffice_path: str) -> None:
    ctk.set_default_color_theme("blue")
    SwiftPDFApp(soffice_path=soffice_path).mainloop()
