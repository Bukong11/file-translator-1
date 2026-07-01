from __future__ import annotations

import json
import os
import shutil
import urllib.error
import urllib.request
import uuid
from pathlib import Path

import fitz
from docx import Document
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from pypdf import PdfReader


BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "storage"
WORK_DIR.mkdir(exist_ok=True)

SUPPORTED_SUFFIXES = {".docx", ".xlsx", ".pdf"}
OPENAI_API_URL = "https://api.openai.com/v1/responses"
DEFAULT_OPENAI_MODEL = "gpt-5.5"
PDF_FONT_FILE = next(
    (
        path
        for path in (
            Path(r"C:\Windows\Fonts\msyh.ttc"),
            Path(r"C:\Windows\Fonts\simhei.ttf"),
            Path(r"C:\Windows\Fonts\arial.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        )
        if path.exists()
    ),
    None,
)

app = FastAPI(title="File Translator Backend")

allowed_origins = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "*").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


def translate_with_openai(text: str, target_language: str, domain: str, terms: str) -> str:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    model = os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL).strip()

    if not api_key:
        return ""

    language_names = {
        "zh": "Chinese",
        "en": "English",
        "ms": "Malay",
        "ja": "Japanese",
        "ko": "Korean",
        "de": "German",
        "fr": "French",
    }
    prompt = (
        "Translate the text below. Keep the meaning accurate and professional. "
        "Return only the translated text. Do not add explanations, prefixes, quotes, or markdown.\n\n"
        f"Target language: {language_names.get(target_language, target_language)}\n"
        f"Professional domain: {domain}\n"
        f"Terminology requirements: {terms or 'None'}\n\n"
        f"Text:\n{text}"
    )
    payload = {
        "model": model,
        "input": prompt,
    }
    request = urllib.request.Request(
        OPENAI_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise HTTPException(status_code=502, detail=f"OpenAI API error: {detail}") from exc
    except urllib.error.URLError as exc:
        raise HTTPException(status_code=502, detail=f"Could not reach OpenAI API: {exc.reason}") from exc

    if data.get("output_text"):
        return data["output_text"].strip()

    pieces: list[str] = []
    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and content.get("text"):
                pieces.append(content["text"])

    return "\n".join(pieces).strip()


def translate_text(text: str, target_language: str, domain: str, terms: str) -> str:
    clean = text.strip()
    if not clean:
        return text

    translated = translate_with_openai(text, target_language, domain, terms)
    if translated:
        return translated

    language_names = {
        "zh": "Chinese",
        "en": "English",
        "ms": "Malay",
        "ja": "Japanese",
        "ko": "Korean",
        "de": "German",
        "fr": "French",
    }
    label = language_names.get(target_language, target_language)
    return f"[{label} translation placeholder] {text}"


def translate_docx(source: Path, target: Path, target_language: str, domain: str, terms: str) -> None:
    document = Document(source)

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            for run in paragraph.runs:
                if run.text.strip():
                    run.text = translate_text(run.text, target_language, domain, terms)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip():
                        for run in paragraph.runs:
                            if run.text.strip():
                                run.text = translate_text(run.text, target_language, domain, terms)

    document.save(target)


def translate_xlsx(source: Path, target: Path, target_language: str, domain: str, terms: str) -> None:
    workbook = load_workbook(source)

    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    cell.value = translate_text(cell.value, target_language, domain, terms)

    workbook.save(target)


def translate_pdf_to_docx(source: Path, target: Path, target_language: str, domain: str, terms: str) -> None:
    reader = PdfReader(str(source))
    document = Document()
    document.add_heading(source.stem, level=1)

    extracted_any_text = False
    for page_number, page in enumerate(reader.pages, start=1):
      text = page.extract_text() or ""
      clean = text.strip()
      if not clean:
          continue

      extracted_any_text = True
      document.add_heading(f"Page {page_number}", level=2)
      translated = translate_text(clean, target_language, domain, terms)
      for paragraph in translated.splitlines():
          if paragraph.strip():
              document.add_paragraph(paragraph.strip())

    if not extracted_any_text:
        document.add_paragraph(
            "No selectable text was found in this PDF. It may be a scanned image PDF and needs OCR before translation."
        )

    document.save(target)


def should_translate_pdf_text(text: str) -> bool:
    clean = " ".join(text.split())
    if len(clean) < 3:
        return False

    letters = sum(ch.isalpha() for ch in clean)
    cjk = sum("\u4e00" <= ch <= "\u9fff" for ch in clean)
    digits = sum(ch.isdigit() for ch in clean)

    if letters + cjk == 0:
        return False

    if digits and digits / max(len(clean), 1) > 0.65:
        return False

    protected_words = {
        "SUS430",
        "A3",
        "M5",
        "R3",
        "GB/T",
        "Alpha ESS",
    }
    if clean in protected_words:
        return False

    return True


def fitted_font_size(rect: fitz.Rect, text: str) -> float:
    height_based = max(4.0, min(8.0, rect.height * 0.52))
    if len(text) > 120:
        return max(3.8, min(height_based, 5.2))
    if len(text) > 60:
        return max(4.0, min(height_based, 6.0))
    return height_based


def translate_pdf_overlay(source: Path, target: Path, target_language: str, domain: str, terms: str) -> None:
    document = fitz.open(source)

    for page in document:
        blocks = page.get_text("blocks")
        for block in blocks:
            x0, y0, x1, y1, text = block[:5]
            clean = " ".join(str(text).split())
            if not should_translate_pdf_text(clean):
                continue

            rect = fitz.Rect(x0, y0, x1, y1)
            if rect.width < 8 or rect.height < 5:
                continue

            translated = translate_text(clean, target_language, domain, terms)
            if not translated or translated == clean:
                continue

            cover = rect + (-0.5, -0.4, 0.5, 0.4)
            page.draw_rect(cover, color=None, fill=(1, 1, 1), overlay=True)
            page.insert_textbox(
                rect,
                translated,
                fontsize=fitted_font_size(rect, translated),
                fontname="msyh" if PDF_FONT_FILE else "helv",
                fontfile=str(PDF_FONT_FILE) if PDF_FONT_FILE else None,
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT,
                overlay=True,
            )

    document.save(target, garbage=4, deflate=True)
    document.close()


@app.get("/health")
def health() -> dict[str, str]:
    mode = "openai" if os.getenv("OPENAI_API_KEY", "").strip() else "placeholder"
    return {"status": "ok", "translation_mode": mode}


@app.get("/")
def root() -> dict[str, str]:
    return {
        "name": "File Translator Backend",
        "status": "running",
        "health": "/health",
    }


@app.post("/translate")
async def translate_file(
    file: UploadFile = File(...),
    target_language: str = Form("zh"),
    domain: str = Form("business"),
    terms: str = Form(""),
    pdf_mode: str = Form("overlay"),
) -> FileResponse:
    original_name = Path(file.filename or "uploaded").name
    suffix = Path(original_name).suffix.lower()

    if suffix not in SUPPORTED_SUFFIXES:
        raise HTTPException(status_code=400, detail="This demo backend currently supports .docx, .xlsx, and text-based .pdf files.")

    task_id = uuid.uuid4().hex
    task_dir = WORK_DIR / task_id
    task_dir.mkdir()

    source = task_dir / original_name
    output_suffix = ".pdf" if suffix == ".pdf" and pdf_mode == "overlay" else ".docx" if suffix == ".pdf" else suffix
    output_name = f"{Path(original_name).stem}-translated{output_suffix}"
    target = task_dir / output_name

    with source.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    if suffix == ".docx":
        translate_docx(source, target, target_language, domain, terms)
    elif suffix == ".xlsx":
        translate_xlsx(source, target, target_language, domain, terms)
    elif suffix == ".pdf":
        if pdf_mode == "overlay":
            translate_pdf_overlay(source, target, target_language, domain, terms)
        else:
            translate_pdf_to_docx(source, target, target_language, domain, terms)

    return FileResponse(
        target,
        filename=output_name,
        media_type="application/octet-stream",
    )
