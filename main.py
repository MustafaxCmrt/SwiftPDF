# -*- coding: utf-8 -*-
"""
SwiftPDF — Office dosyalarını LibreOffice ile PDF'e dönüştürür.
Başlangıç noktası: LibreOffice kontrolü ve ana arayüz.
"""

from __future__ import annotations

import os
import sys

# PyInstaller _MEI geçici dizin uyarısını bastır
os.environ["PYINSTALLER_SUPPRESS_SPLASH_SCREEN"] = "1"

from utils import find_libreoffice_executable
from ui import run_app, show_libreoffice_required_and_wait


def main() -> None:
    # Kurulum yolu yoksa uyarı göster ve çık; ana uygulama açılmaz
    soffice = find_libreoffice_executable()
    if not soffice:
        show_libreoffice_required_and_wait()
        sys.exit(1)
    run_app(soffice_path=soffice)


if __name__ == "__main__":
    main()
