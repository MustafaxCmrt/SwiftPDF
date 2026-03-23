# -*- coding: utf-8 -*-
"""LibreOffice headless ile Office dosyalarını PDF'e dönüştürme (subprocess)."""

from __future__ import annotations

import atexit
import os
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import Callable

from utils import find_libreoffice_executable, is_supported_file


# Varsayılan dönüşüm zaman aşımı (saniye)
DEFAULT_TIMEOUT_SEC = 300

# LibreOffice için izole geçici kullanıcı profili — profil kilidi sorununu önler
_lo_profile_dir: str | None = None


def _get_lo_profile_dir() -> str:
    """LibreOffice için tek seferlik geçici profil dizini oluşturur."""
    global _lo_profile_dir
    if _lo_profile_dir is None:
        _lo_profile_dir = tempfile.mkdtemp(prefix="swiftpdf_lo_")
        atexit.register(_cleanup_lo_profile)
    return _lo_profile_dir


def _cleanup_lo_profile() -> None:
    global _lo_profile_dir
    if _lo_profile_dir and os.path.isdir(_lo_profile_dir):
        try:
            shutil.rmtree(_lo_profile_dir, ignore_errors=True)
        except Exception:
            pass
        _lo_profile_dir = None


class ConversionResult:
    """Tek bir dosya için dönüşüm sonucu."""

    def __init__(
        self,
        input_path: str,
        ok: bool,
        message: str,
        output_pdf: str | None = None,
    ) -> None:
        self.input_path = input_path
        self.ok = ok
        self.message = message
        self.output_pdf = output_pdf


def convert_file_to_pdf(
    soffice_path: str,
    input_path: str,
    output_dir: str,
    timeout_sec: int = DEFAULT_TIMEOUT_SEC,
) -> ConversionResult:
    """
    Tek dosyayı PDF'e çevirir. Çıktı dosyası genelde aynı taban adla output_dir içinde oluşur.
    """
    inp = Path(input_path).resolve()
    out = Path(output_dir).resolve()

    if not inp.is_file():
        return ConversionResult(str(inp), False, "Dosya bulunamadı.")

    if not is_supported_file(inp):
        return ConversionResult(
            str(inp),
            False,
            f"Desteklenmeyen format: {inp.suffix}",
        )

    if not out.is_dir():
        return ConversionResult(str(inp), False, "Çıktı klasörü geçersiz veya yok.")

    # İzole profil dizini — açık LibreOffice oturumuyla çakışmayı önler
    profile_uri = Path(_get_lo_profile_dir()).as_uri()

    cmd = [
        soffice_path,
        f"-env:UserInstallation={profile_uri}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nologo",
        "--nodefault",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out),
        str(inp),
    ]

    # Windows'ta konsol penceresi açılmasın (3.7+)
    run_kw: dict = {}
    if os.name == "nt":
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        if flags:
            run_kw["creationflags"] = flags
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            **run_kw,
        )
    except subprocess.TimeoutExpired:
        return ConversionResult(
            str(inp),
            False,
            f"Dönüşüm zaman aşımına uğradı ({timeout_sec} sn).",
        )
    except OSError as e:
        return ConversionResult(str(inp), False, f"İşlem başlatılamadı: {e}")

    expected_pdf = out / f"{inp.stem}.pdf"
    if proc.returncode == 0 and expected_pdf.is_file():
        return ConversionResult(
            str(inp),
            True,
            "Tamamlandı.",
            output_pdf=str(expected_pdf),
        )

    err_hint = (proc.stderr or proc.stdout or "").strip()
    if err_hint:
        msg = f"LibreOffice hata kodu {proc.returncode}: {err_hint[:500]}"
    else:
        msg = (
            f"Dönüşüm başarısız (kod {proc.returncode}). "
            f"Beklenen PDF: {expected_pdf}"
        )
    return ConversionResult(str(inp), False, msg)


def convert_batch(
    soffice_path: str | None,
    file_paths: list[str],
    output_dir: str,
    on_progress: Callable[[int, int, ConversionResult], None] | None = None,
    timeout_sec: int = DEFAULT_TIMEOUT_SEC,
) -> list[ConversionResult]:
    """
    Dosyaları sırayla dönüştürür. on_progress(i, total, result) her dosyadan sonra çağrılır.
    Thread içinden güvenli şekilde UI güncellemek için on_progress'te after() kullanın.
    """
    if not soffice_path:
        soffice_path = find_libreoffice_executable()
    if not soffice_path or not os.path.isfile(soffice_path):
        return [
            ConversionResult(p, False, "LibreOffice bulunamadı.") for p in file_paths
        ]

    total = len(file_paths)
    results: list[ConversionResult] = []
    for i, p in enumerate(file_paths, start=1):
        res = convert_file_to_pdf(soffice_path, p, output_dir, timeout_sec)
        results.append(res)
        if on_progress:
            on_progress(i, total, res)
    return results


def run_batch_in_thread(
    soffice_path: str,
    file_paths: list[str],
    output_dir: str,
    on_progress: Callable[[int, int, ConversionResult], None],
    on_done: Callable[[list[ConversionResult]], None],
    timeout_sec: int = DEFAULT_TIMEOUT_SEC,
) -> None:
    """Dönüşümü arka planda çalıştırır; UI kilidi önlenir."""

    def worker() -> None:
        results = convert_batch(
            soffice_path,
            file_paths,
            output_dir,
            on_progress=on_progress,
            timeout_sec=timeout_sec,
        )
        on_done(results)

    t = threading.Thread(target=worker, daemon=True)
    t.start()
