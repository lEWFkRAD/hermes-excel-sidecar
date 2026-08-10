"""Local attachment extractor for the Hermes Excel bridge.

Runs on SHATNER using the same machinery as bearden-writer's ocr_worker
(pypdf text layer -> pypdfium2 raster -> Tesseract), but standalone: no DB,
no client matching, no network. Invoked per-file by broker/server.mjs when
HERMES_EXCEL_DOCLING_API=local.

Usage:
    python local_extract.py <file>       -> JSON {ok, text, method} on stdout
    python local_extract.py --health     -> JSON {ok, tesseract, ...}

Intended interpreter: C:\\Users\\Administrator\\bearden-writer\\venv\\Scripts\\python.exe
(has pypdf, pypdfium2, pillow_heif, PIL).
"""

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

TESSERACT_CMD = os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
TESSERACT_PSM = os.getenv("TESSERACT_PSM", "6")
MAX_OCR_PAGES = int(os.getenv("LOCAL_EXTRACT_MAX_PAGES", "40"))
# OCR engine: "grm" routes scanned pages through the GRM vision-OCR pipeline
# (scripts/grm_ocr_cli.py under the hermes-agent venv, same stack anydoc uses);
# anything else — and every GRM failure — lands on Tesseract.
OCR_ENGINE = os.getenv("HERMES_EXCEL_OCR_ENGINE", "tesseract").strip().lower()
GRM_PYTHON = os.getenv("HERMES_EXCEL_GRM_PYTHON", "")
GRM_CLI = os.getenv("HERMES_EXCEL_GRM_CLI", "")
GRM_PYTHONPATH = os.getenv("HERMES_EXCEL_GRM_PYTHONPATH", "")
GRM_TIMEOUT = int(os.getenv("HERMES_EXCEL_GRM_TIMEOUT", "420"))
# GRM streams ~10-30s per page; past this many pages one CLI call would blow
# the timeout, so large documents go straight to Tesseract.
GRM_MAX_PAGES = int(os.getenv("HERMES_EXCEL_GRM_MAX_PAGES", "12"))
# Same threshold ocr_worker uses to accept the PDF text layer.
MIN_TEXTLAYER_CHARS = 50

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif"}

# Office/document formats parsed locally by anydoc (same lesson as the
# HermesOffice xlsx-sidecar: parse in-process with a zero-network library
# instead of shipping the file to a heavy parser service). PDFs stay on the
# pypdf/Tesseract/GRM path above — anydoc reports "OCR required" for scanned
# PDFs, which is exactly what our ocr_pages() dispatch already handles.
OFFICE_EXTS = {".docx", ".doc", ".odt", ".rtf", ".epub", ".pptx", ".ppt",
               ".xlsx", ".xls", ".ods"}


def office_text(path):
    """Parse an office document to Markdown via anydoc. Raises on failure."""
    try:
        import anydoc
    except ImportError as exc:
        raise RuntimeError(f"anydoc not importable: {exc}")
    data = Path(path).read_bytes()
    fmt = anydoc.format_from_path(path) or anydoc.format_from_bytes(data)
    if fmt is None:
        raise RuntimeError("unrecognized document format")
    md = anydoc.to_markdown_bytes(data)
    return md if isinstance(md, str) else md.decode("utf-8")


def fail(error):
    print(json.dumps({"ok": False, "error": str(error)[:400]}))
    sys.exit(0)


def health():
    out = {"ok": True, "service": "hermes-excel-local-extract"}
    if not os.path.isfile(TESSERACT_CMD):
        out = {"ok": False, "error": f"tesseract not found at {TESSERACT_CMD}"}
    else:
        try:
            import pypdf  # noqa: F401
            import pypdfium2  # noqa: F401
            from PIL import Image  # noqa: F401
            out["tesseract"] = TESSERACT_CMD
        except Exception as exc:
            out = {"ok": False, "error": f"venv import failed: {exc}"}
    # anydoc (office-format local parsing) is informational like GRM: if it's
    # missing, office attachments fall through to the configured parser API.
    if out.get("ok"):
        try:
            import anydoc  # noqa: F401
            out["anydoc"] = {"ok": True}
        except Exception as exc:
            out["anydoc"] = {"ok": False, "error": str(exc)[:200]}
    # GRM engine status is informational: a dead GRM never fails health because
    # every OCR call falls back to Tesseract on its own.
    if out.get("ok") and OCR_ENGINE == "grm":
        try:
            env = dict(os.environ)
            if GRM_PYTHONPATH:
                env["PYTHONPATH"] = GRM_PYTHONPATH
            proc = subprocess.run([GRM_PYTHON, GRM_CLI, "--health"], capture_output=True,
                                  text=True, timeout=30, check=False, env=env)
            out["grm"] = json.loads((proc.stdout or "").strip() or "{}")
        except Exception as exc:
            out["grm"] = {"ok": False, "error": str(exc)[:200]}
    print(json.dumps(out))
    sys.exit(0)


def pdf_textlayer(path):
    import pypdf
    try:
        reader = pypdf.PdfReader(path)
        return "\f".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def rasterize_pdf(path):
    import pypdfium2
    pdf = pypdfium2.PdfDocument(path)
    total = len(pdf)
    rendered = min(total, MAX_OCR_PAGES)
    images = [pdf[i].render(scale=2.0).to_pil() for i in range(rendered)]
    return images, rendered, total


def load_image(path):
    from PIL import Image, ImageOps
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except Exception:
        pass
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    return img


def tesseract_pages(images):
    if not os.path.isfile(TESSERACT_CMD):
        raise RuntimeError(f"tesseract not found at {TESSERACT_CMD}")
    pages = []
    with tempfile.TemporaryDirectory(prefix="hermes-excel-ocr-") as tmpdir:
        tmp = Path(tmpdir)
        for idx, img in enumerate(images, start=1):
            image_path = tmp / f"page-{idx:03d}.png"
            output_base = tmp / f"page-{idx:03d}"
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(image_path, format="PNG")
            proc = subprocess.run(
                [TESSERACT_CMD, str(image_path), str(output_base), "--psm", str(TESSERACT_PSM)],
                capture_output=True, text=True, timeout=180, check=False,
            )
            if proc.returncode != 0:
                err = (proc.stderr or proc.stdout or f"exit {proc.returncode}").strip()
                raise RuntimeError(f"tesseract failed: {err[:200]}")
            pages.append(output_base.with_suffix(".txt").read_text(encoding="utf-8", errors="replace"))
    return "\n\n--- page break ---\n\n".join(pages)


def grm_pages(images):
    if not (GRM_PYTHON and os.path.isfile(GRM_PYTHON) and GRM_CLI and os.path.isfile(GRM_CLI)):
        raise RuntimeError("GRM OCR CLI not configured")
    if len(images) > GRM_MAX_PAGES:
        raise RuntimeError(f"{len(images)} pages exceeds GRM cap of {GRM_MAX_PAGES}")
    env = dict(os.environ)
    if GRM_PYTHONPATH:
        env["PYTHONPATH"] = GRM_PYTHONPATH
    with tempfile.TemporaryDirectory(prefix="hermes-excel-grm-") as tmpdir:
        tmp = Path(tmpdir)
        paths = []
        for idx, img in enumerate(images, start=1):
            page_path = tmp / f"page-{idx:03d}.png"
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(page_path, format="PNG")
            paths.append(str(page_path))
        proc = subprocess.run([GRM_PYTHON, GRM_CLI, *paths], capture_output=True, text=True,
                              timeout=GRM_TIMEOUT, check=False, env=env)
        try:
            data = json.loads((proc.stdout or "").strip() or "{}")
        except ValueError:
            raise RuntimeError(f"grm cli returned non-JSON (exit {proc.returncode})")
        if not data.get("ok"):
            raise RuntimeError(str(data.get("error") or f"grm cli exit {proc.returncode}")[:200])
        pages = data.get("pages")
        if not isinstance(pages, list) or len(pages) != len(paths):
            raise RuntimeError("grm cli returned wrong page count")
        return "\n\n--- page break ---\n\n".join(str(page) for page in pages)


def ocr_pages(images):
    """Engine dispatch. Returns (text, method); GRM failures fall back to Tesseract."""
    if OCR_ENGINE == "grm":
        try:
            return grm_pages(images), "grm-vision"
        except Exception as exc:
            note = f"\n\n[GRM OCR unavailable ({str(exc)[:120]}); used Tesseract fallback]"
            return tesseract_pages(images) + note, f"tesseract:psm{TESSERACT_PSM}:grm-fallback"
    return tesseract_pages(images), f"tesseract:psm{TESSERACT_PSM}"


def main():
    if len(sys.argv) < 2:
        fail("usage: local_extract.py <file> | --health")
    if sys.argv[1] == "--health":
        health()

    path = sys.argv[1]
    if not os.path.isfile(path):
        fail(f"file not found: {path}")
    suffix = Path(path).suffix.lower()

    try:
        if suffix == ".pdf":
            text = pdf_textlayer(path)
            if len(text.strip()) >= MIN_TEXTLAYER_CHARS:
                print(json.dumps({"ok": True, "text": text, "method": "pypdf:text_layer"}))
                return
            images, rendered, total = rasterize_pdf(path)
            text, method = ocr_pages(images)
            if total > rendered:
                text += f"\n\n[truncated after {rendered} of {total} pages]"
            print(json.dumps({"ok": True, "text": text, "method": method}))
        elif suffix in IMAGE_EXTS:
            text, method = ocr_pages([load_image(path)])
            print(json.dumps({"ok": True, "text": text, "method": method}))
        elif suffix in OFFICE_EXTS:
            text = office_text(path)
            print(json.dumps({"ok": True, "text": text, "method": "anydoc:local"}))
        else:
            fail(f"unsupported type {suffix}")
    except Exception as exc:
        fail(f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
