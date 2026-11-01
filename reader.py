"""
reader.py — reads PDF, DOCX, and PPTX draft files and returns plain text.

For image-based PDFs (scanned or Arabic fonts rendered as images), falls back
to Gemini vision which reads the pages directly as images.
"""

import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_MIN_CHARS_PER_PAGE = 200  # below this → PDF is image-based, use vision


def read_pdf(path: str, api_key: str = '') -> str:
    import pymupdf
    doc = pymupdf.open(path)
    num_pages = len(doc)
    pages = []
    for page in doc:
        text = page.get_text()
        if text.strip():
            pages.append(text)
    doc.close()

    full_text = "\n\n".join(pages)
    avg_chars = len(full_text) / max(num_pages, 1)

    if avg_chars < _MIN_CHARS_PER_PAGE and api_key:
        logger.info(f'PDF text thin ({avg_chars:.0f} chars/page avg) — using Gemini vision OCR')
        return _read_pdf_with_gemini(path, api_key)

    return full_text


_VISION_BATCH = 6   # pages per Gemini call — keeps payload manageable
_VISION_DPI   = 96  # enough for text reading, ~200KB per page


def _read_pdf_with_gemini(path: str, api_key: str) -> str:
    """Render PDF pages as images and read all text via Gemini vision, batched."""
    import pymupdf
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)

    doc = pymupdf.open(path)
    page_images = []
    for page in doc:
        pix = page.get_pixmap(dpi=_VISION_DPI)
        page_images.append(pix.tobytes('png'))
    doc.close()
    logger.info(f'Rendered {len(page_images)} pages at {_VISION_DPI}dpi for Gemini vision')

    prompt_text = (
        'اقرأ هذه الصفحات واستخرج جميع النصوص العربية والإنجليزية بالترتيب. '
        'اكتب الجداول والأرقام والنقاط والعناوين كما هي. لا تلخص.'
    )

    all_text = []
    for batch_start in range(0, len(page_images), _VISION_BATCH):
        batch = page_images[batch_start: batch_start + _VISION_BATCH]
        parts = [types.Part.from_bytes(data=img, mime_type='image/png') for img in batch]
        parts.append(prompt_text)

        for model in ('gemini-3.8-flash', 'gemini-3.5-flash-lite'):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=parts,
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        max_output_tokens=4096,
                        http_options={'timeout': 60000},
                    )
                )
                chunk = response.text.strip()
                all_text.append(chunk)
                logger.info(f'Batch {batch_start//6+1}: {len(chunk)} chars via {model}')
                break
            except Exception as exc:
                logger.warning(f'Vision [{model}] batch {batch_start//6+1} failed: {exc}')

    result = '\n\n'.join(all_text)
    logger.info(f'Gemini vision total: {len(result)} chars from {len(page_images)}-page PDF')
    return result


def extract_client_logo(pdf_path: str, save_dir: str = '') -> str:
    """
    Extract the client's logo from a tender PDF cover page.

    Strategy:
    1. Look for a natural logo image (not full-page, not tiny dots).
    2. If only full-page background images are found, crop the top-center
       portion where logos typically appear.

    Returns absolute path to saved PNG, or '' if nothing usable found.
    """
    import pymupdf
    import tempfile

    out_dir = Path(save_dir) if save_dir else Path(tempfile.gettempdir()) / 'zawad_logos'
    out_dir.mkdir(parents=True, exist_ok=True)

    safe_stem = ''.join(c for c in Path(pdf_path).stem if c.isalnum() or c in '_-')[:30]
    logo_path = out_dir / f'{safe_stem}_client_logo.png'

    if logo_path.exists():
        return str(logo_path)

    doc = pymupdf.open(pdf_path)
    best_natural  = None   # standalone logo image
    best_fullpage = None   # full-page cover image (fallback — will crop)

    seen_xrefs = set()

    for page_idx in range(min(3, len(doc))):
        page = doc[page_idx]
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            try:
                img = doc.extract_image(xref)
            except Exception:
                continue

            w, h = img.get('width', 0), img.get('height', 0)
            if w < 40 or h < 40:
                continue    # invisible dots or color fills
            aspect = w / h if h else 0

            is_fullpage = (h > 500 and 0.5 < aspect < 1.5)
            is_banner   = (aspect > 3.5)  # wide strips — not logos
            is_tiny     = (w < 80 or h < 60)

            if is_fullpage:
                if best_fullpage is None:
                    best_fullpage = img['image']
            elif not is_banner and not is_tiny:
                # Proportions match a logo (roughly square or moderate width)
                if best_natural is None:
                    best_natural = img['image']

    doc.close()

    if best_natural:
        logo_path.write_bytes(best_natural)
        logger.info(f'Extracted standalone logo from PDF → {logo_path}')
        return str(logo_path)

    if best_fullpage:
        # Crop top-center portion of cover page — logos sit there in Omani tenders
        cropped = _crop_logo_from_cover(best_fullpage)
        if cropped:
            logo_path.write_bytes(cropped)
            logger.info(f'Cropped logo from cover page → {logo_path}')
            return str(logo_path)

    logger.info('No client logo image found in PDF')
    return ''


def _crop_logo_from_cover(img_bytes: bytes) -> bytes:
    """Crop the top-center 20% of a cover page image to isolate the logo."""
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(img_bytes)).convert('RGBA')
        w, h = img.size
        # Crop: top 22%, horizontal center 60%
        x0 = int(w * 0.20)
        x1 = int(w * 0.80)
        y0 = 0
        y1 = int(h * 0.22)
        cropped = img.crop((x0, y0, x1, y1))
        buf = io.BytesIO()
        cropped.save(buf, format='PNG')
        return buf.getvalue()
    except Exception as e:
        logger.warning(f'Logo crop failed: {e}')
        return b''


def read_docx(path: str) -> str:
    from docx import Document
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    # Also extract table text
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                paragraphs.append(" | ".join(cells))
    return "\n\n".join(paragraphs)


def read_pptx(path: str) -> str:
    from pptx import Presentation
    prs = Presentation(path)
    slides_text = []
    for i, slide in enumerate(prs.slides, 1):
        parts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        parts.append(text)
        if parts:
            slides_text.append(f"[Slide {i}]\n" + "\n".join(parts))
    return "\n\n".join(slides_text)


def read_txt(path: str) -> str:
    encodings = ['utf-8', 'utf-8-sig', 'cp1256', 'latin-1']
    for enc in encodings:
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(path, 'rb') as f:
        return f.read().decode('utf-8', errors='replace')


def read_file(path, api_key: str = '') -> str:
    """Auto-detect format and return extracted text."""
    path = str(path)
    ext = Path(path).suffix.lower()

    if ext == '.pdf':
        return read_pdf(path, api_key=api_key)
    elif ext in ('.docx', '.doc'):
        return read_docx(path)
    elif ext in ('.pptx', '.ppt'):
        return read_pptx(path)
    elif ext == '.txt':
        return read_txt(path)
    else:
        try:
            return read_txt(path)
        except Exception:
            raise ValueError(f"Unsupported file format: {ext}")
