# -*- coding: utf-8 -*-
"""Yardımcı fonksiyonlar: desteklenen uzantılar, LibreOffice yolu, DnD yolu ayrıştırma."""

from __future__ import annotations

import os
import platform
import shutil
from pathlib import Path
from urllib.parse import unquote, urlparse

# Desteklenen Office uzantıları (küçük harf)
SUPPORTED_EXTENSIONS = frozenset(
    {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"}
)

APP_VERSION = "1.1.1"

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download"


def is_supported_file(path: str | Path) -> bool:
    """Dosya yolu desteklenen bir formata sahip mi kontrol eder."""
    suffix = Path(path).suffix.lower()
    return suffix in SUPPORTED_EXTENSIONS


def find_libreoffice_executable() -> str | None:
    """
    LibreOffice 'soffice' yürütülebilir dosyasının tam yolunu döndürür.
    Windows ve macOS için bilinen konumlar + PATH üzerinde arama.
    """
    system = platform.system()

    if system == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.isfile(p):
                return p
    elif system == "Darwin":
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.isfile(mac_path):
            return mac_path

    # PATH'te soffice / soffice.exe
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found and os.path.isfile(found):
            return found

    return None


def parse_dnd_file_list(data: str) -> list[str]:
    """
    Sürükle-bırak ile gelen ham metni dosya yolları listesine çevirir.
    Windows: boşlukla ayrılmış yollar; boşluk içerenler { } ile sarılı olabilir.
    macOS: file:// URL'leri olabilir.
    """
    if not data or not str(data).strip():
        return []

    raw = str(data).strip()
    paths: list[str] = []

    # file:// satırları (bazen birden fazla)
    for part in raw.replace("\r", "\n").split("\n"):
        part = part.strip()
        if not part:
            continue
        if part.startswith("file://"):
            p = urlparse(part)
            path = unquote(p.path)
            # Windows: file:///C:/... -> path bazen /C:/... olur
            if platform.system() == "Windows":
                if len(path) >= 3 and path[0] == "/" and path[2] == ":":
                    path = path[1:]
                elif path.startswith("/"):
                    path = path.lstrip("/")
            if path:
                paths.append(path)
        else:
            paths.extend(_parse_windows_style_dnd(part))

    # Tek satırda birden fazla file:// olmayan Windows listesi
    if not paths and raw:
        paths = _parse_windows_style_dnd(raw)

    seen: set[str] = set()
    unique: list[str] = []
    for p in paths:
        p = os.path.normpath(p)
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique


def _parse_windows_style_dnd(s: str) -> list[str]:
    """Boşluk / süslü parantez kurallarına göre yol listesi."""
    result: list[str] = []
    i = 0
    n = len(s)
    while i < n:
        while i < n and s[i].isspace():
            i += 1
        if i >= n:
            break
        if s[i] == "{":
            j = s.find("}", i + 1)
            if j == -1:
                result.append(s[i:].strip("{}").strip())
                break
            result.append(s[i + 1 : j].strip())
            i = j + 1
        else:
            j = i
            while j < n and not s[j].isspace():
                j += 1
            result.append(s[i:j].strip())
            i = j
    return [p for p in result if p]
