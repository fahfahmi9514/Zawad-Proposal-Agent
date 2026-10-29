"""
generator.py — Single-call proposal generator for Zawad.

One focused Gemini call produces the complete proposal JSON.
Works reliably on the free tier (5 RPM limit).

Pipeline:
  1. RAG: retrieve 2 similar past proposals (word-overlap, no API)
  2. Checklist: injected from extractor.py output
  3. Single Gemini call → full JSON
  4. Python validates and fixes all pricing math
"""

import json
import re
import logging
import time
from google import genai
from google.genai import types

logger = logging.getLogger(__name__)

MODELS      = ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-1.5-flash']
MAX_RETRIES = 3
RETRY_WAITS = [30, 60, 90]

SYSTEM_PROMPT = """أنت خبير في كتابة عروض المناقصات الحكومية العُمانية للشركات الاستشارية.
مهمتك: إنشاء عرض فني ومالي احترافي كامل بالعربية لشركة زواد للتطوير والاستشارات.

=== معلومات زواد ===
الاسم: زواد للتطوير والاستشارات | البريد: ghasan@zawad.om | الهاتف: +968 9777 3444
مدير المشروع ورئيس الفريق: د. غسان المعمري — استخدم هذا الاسم دائماً في قسم الفريق

=== بطاقة الأسعار ===
يوم تنفيذ: 4,000 ريال | يوم تصميم: 2,500 ريال | Coaching/ساعة: 120 ريال
تيسير جماعي/ساعة: 150 ريال | تقرير: 1,000 ريال | مرشح مركز تقييم: 481 ريال

=== الضريبة والضمان والصلاحية — اقرأ المناقصة أولاً ===
- ضريبة القيمة المضافة: استخدم النسبة التي تذكرها المناقصة. إن لم تُذكر، استخدم 5% للجهات الحكومية العُمانية، وأسعار بدون ضريبة للجهات الخاصة.
- ضمان حسن التنفيذ: استخدم النسبة التي تذكرها المناقصة بالضبط واكتبها في حقل bond_rate (مثال 0.05 للـ5%). إن لم تُذكر ولم يُطلب، اكتب bond_rate: null ولا تذكر الضمان.
- صلاحية العرض: استخدم المدة التي تذكرها المناقصة. إن لم تُذكر، استخدم 90 يوماً.
- رسوم حماية المنافسة 1.2% (تعميم 13/2024): اذكرها فقط إن كانت الجهة حكومية عُمانية.

=== سجل الإنجازات ===
برنامج تطوير القيادات (فريق 5) — الأكاديمية السلطانية — 119,000 ريال — 25 مشاركاً
برنامج إتقان للتقييم — الأكاديمية السلطانية — 109,000 ريال — 200 مشارك
البرنامج التنفيذي للإدارة المحلية — الأكاديمية السلطانية — 83,000 ريال
برنامج إمكان 3 — وزارة الخدمة المدنية — 81,900 ريال — 120 مشاركاً
تمكين القيادة — شركة عُمان للألمنيوم — 43,000 ريال — 32 مشاركاً

=== قواعد الكتابة — إلزامية ===
- ابدأ الجمل بالفعل: يعتمد / نتميز / نتطلع / نحرص / يضمن / نقدم
- خاطب الجهة بصيغة الجمع: احتياجاتكم / أهدافكم / معكم / شريككم
- كل قسم نصي لا يقل عن فقرتين كاملتين
- اذكر رؤية عُمان 2040 في السياق المناسب
- company_overview: 4 أسطر+ | requirements_understanding: 4 أسطر+ | methodology: 5 أسطر+
- cover_letter: 3 فقرات كاملة تبدأ بالمخاطبة الرسمية
- innovation: 4 أسطر+ تربط البرنامج بتخصص زواد

=== innovation_points — استخدم هذه العناوين الستة بالضبط ===
خبرة حكومية واسعة | محتوى محلي مخصص | سجل نجاح حافل
منهجية تفاعلية | التزام بالجودة | تقييمات متميزة

=== التسعير — قواعد صارمة ===
- احسب: total_omr = quantity × rate_omr لكل بند
- subtotal_omr = مجموع كل total_omr
- حدد vat_rate: استخدم نسبة الضريبة من المناقصة (مثال: 0.05 أو 0.03). إن طُلبت أسعار بدون ضريبة اكتب 0. إن لم تُذكر، اكتب 0.05 للجهات الحكومية.
- vat_omr = subtotal_omr × vat_rate
- grand_total_omr = subtotal_omr + vat_omr
- لا تتجاوز سقف الميزانية إن ذُكر

=== هيكل JSON المطلوب — أعد هذا الهيكل كاملاً فقط ===
{
  "program_title": "عنوان البرنامج بالعربية",
  "program_title_en": "Program Title in English",
  "client_name": "اسم الجهة",
  "client_name_en": "Client Name",
  "client_website": "www.client.gov.om",
  "tender_number": "رقم المناقصة أو ''",
  "program_category": "البرنامج التدريبي",
  "presenter": "زواد للتطوير والاستشارات",
  "submitted_to": "مقدَّم إلى: اسم الجهة",
  "format": "A",
  "bond_rate": null,
  "stats": {
    "participants": 20,
    "days": 5,
    "location": "مقر الجهة",
    "month": "يتم التنسيق مع الجهة",
    "duration": "5 أيام تدريبية",
    "language": "العربية",
    "currency": "ريال عماني",
    "validity": "90 يوماً",
    "components": "ورش عمل وحالات دراسية",
    "prep_period": "3 أسابيع من الإسناد",
    "deliverables": "تقرير نهائي شامل"
  },
  "key_stats": [
    {"number": "20", "label": "مشاركاً"},
    {"number": "5", "label": "أيام تدريبية"},
    {"number": "3", "label": "وحدات تعليمية"},
    {"number": "90", "label": "يوماً صلاحية العرض"}
  ],
  "commitments_grid": [
    {"icon": "👥", "title": "الفئة المستهدفة", "value": "20 مشاركاً"},
    {"icon": "📍", "title": "مكان التنفيذ", "value": "مقر الجهة"},
    {"icon": "⏳", "title": "فترة الإعداد", "value": "3 أسابيع"},
    {"icon": "📋", "title": "تسليم المخرجات", "value": "خلال 15 يوم عمل"},
    {"icon": "💰", "title": "عملة العرض", "value": "الريال العُماني"},
    {"icon": "🛡", "title": "صلاحية العرض", "value": "90 يوماً"}
  ],
  "cover_letter": "خطاب تقديم رسمي — 3 فقرات",
  "sections": {
    "key_message": "رسالة القيمة في جملتين",
    "commitments": "نص الالتزامات",
    "commitments_list": ["بند 1", "بند 2", "بند 3", "بند 4", "بند 5"],
    "company_overview": "نبذة عن زواد — 4 أسطر+",
    "requirements_understanding": "فهم المتطلبات — 4 أسطر+",
    "methodology": "المنهجية — 5 أسطر+",
    "quality_commitment": "التزام الجودة — فقرة",
    "workplan": [
      {"day": 1, "title": "اليوم الأول: العنوان", "activities": ["نشاط 1", "نشاط 2", "نشاط 3", "نشاط 4"]}
    ],
    "team": [
      {"name": "الاسم", "role": "المسمى", "experience": "وصف الخبرة"},
      {"name": "الاسم", "role": "المسمى", "experience": "وصف الخبرة"},
      {"name": "الاسم", "role": "المسمى", "experience": "وصف الخبرة"}
    ],
    "expected_outputs": [
      {"number": "01", "title": "عنوان", "description": "وصف"},
      {"number": "02", "title": "عنوان", "description": "وصف"},
      {"number": "03", "title": "عنوان", "description": "وصف"},
      {"number": "04", "title": "عنوان", "description": "وصف"},
      {"number": "05", "title": "عنوان", "description": "وصف"}
    ],
    "track_record_highlight": "سجل الإنجازات — فقرتان",
    "innovation": "التميز والقيمة المضافة — 4 أسطر+",
    "innovation_points": [
      {"icon": "🌐", "title": "خبرة حكومية واسعة", "description": "وصف"},
      {"icon": "🎯", "title": "محتوى محلي مخصص", "description": "وصف"},
      {"icon": "📊", "title": "سجل نجاح حافل", "description": "وصف"},
      {"icon": "💡", "title": "منهجية تفاعلية", "description": "وصف"},
      {"icon": "✅", "title": "التزام بالجودة", "description": "وصف"},
      {"icon": "⭐", "title": "تقييمات متميزة", "description": "وصف"}
    ],
    "risks": [
      {"risk": "وصف المخاطرة", "probability": "متوسطة", "impact": "عالي",
       "preventive": "الإجراء الوقائي", "contingency": "خطة الطوارئ"},
      {"risk": "...", "probability": "...", "impact": "...", "preventive": "...", "contingency": "..."},
      {"risk": "...", "probability": "...", "impact": "...", "preventive": "...", "contingency": "..."},
      {"risk": "...", "probability": "...", "impact": "...", "preventive": "...", "contingency": "..."},
      {"risk": "...", "probability": "...", "impact": "...", "preventive": "...", "contingency": "..."}
    ],
    "pricing": {
      "line_items": [
        {"description": "أيام تصميم البرنامج", "quantity": 5, "unit": "يوم", "rate_omr": 2500, "total_omr": 12500},
        {"description": "أيام تنفيذ البرنامج", "quantity": 5, "unit": "يوم", "rate_omr": 4000, "total_omr": 20000},
        {"description": "تقرير نهائي شامل", "quantity": 1, "unit": "تقرير", "rate_omr": 1000, "total_omr": 1000}
      ],
      "vat_rate": 0.05,
      "subtotal_omr": 33500,
      "vat_omr": 1675,
      "grand_total_omr": 35175
    }
  }
}"""


def _extract_json(text: str) -> dict:
    """Extract JSON from Gemini response, handling markdown fences and extra text."""
    # 1. Direct parse
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        pass

    # 2. Strip markdown fences (greedy — handles very long JSON blocks)
    stripped = re.sub(r'^```(?:json)?\s*', '', text.strip(), flags=re.IGNORECASE)
    stripped = re.sub(r'\s*```\s*$', '', stripped)
    try:
        return json.loads(stripped.strip())
    except json.JSONDecodeError:
        pass

    # 3. Find the outermost { ... } block by brace matching
    start = text.find('{')
    if start != -1:
        depth = 0
        for i, ch in enumerate(text[start:], start):
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i + 1])
                    except json.JSONDecodeError:
                        break

    raise ValueError(f'Cannot parse JSON from response (first 300 chars): {text[:300]}')


def _fix_pricing(proposal: dict) -> dict:
    """Python always recalculates pricing math — never trusts Gemini for totals."""
    pricing = proposal.get('sections', {}).get('pricing', {})
    if not pricing:
        return proposal
    items = pricing.get('line_items', [])
    for item in items:
        try:
            qty  = float(item.get('quantity', 0) or 0)
            rate = float(item.get('rate_omr', 0) or 0)
        except (TypeError, ValueError):
            qty, rate = 0.0, 0.0
        item['quantity']  = qty
        item['rate_omr']  = rate
        item['total_omr'] = round(qty * rate, 3)
    subtotal = round(sum(i['total_omr'] for i in items), 3)
    vat_rate = float(pricing.get('vat_rate', 0.05))
    pricing['subtotal_omr']    = subtotal
    pricing['vat_omr']         = round(subtotal * vat_rate, 3)
    pricing['grand_total_omr'] = round(subtotal * (1 + vat_rate), 3)
    return proposal


def _get_rag(draft_text: str, api_key: str) -> str:
    try:
        from rag import get_rag_context
        return get_rag_context(draft_text, api_key)
    except Exception as exc:
        logger.warning(f'RAG skipped: {exc}')
        return ''


def generate_proposal(draft_text: str, api_key: str,
                      checklist: str = None) -> dict:
    """
    Generate a complete proposal with a single Gemini call.
    checklist: formatted requirements string from extractor.format_checklist().
    Returns the complete proposal dict ready for build_pptx().
    """
    rag = _get_rag(draft_text, api_key)

    checklist_section = ''
    if checklist:
        checklist_section = f'\n\n=== قائمة المتطلبات المستخرجة — يجب تغطية كل بند ===\n{checklist[:800]}'

    rag_section = ''
    if rag:
        rag_section = f'\n\n=== أمثلة من عروض زواد السابقة (للأسلوب فقط) ===\n{rag[:1000]}'

    user_prompt = f"""الوثائق المرفقة (اقرأ كل وثيقة بالكامل وافهم دورها):
---
{draft_text[:12000]}
---
{checklist_section}{rag_section}

تعليمات:
- إذا كانت هناك وثيقة مناقصة: استخرج منها متطلبات الجهة وأهدافها واكتب العرض ليغطيها بالكامل
- إذا كانت هناك مسودة عرض سابقة: استخدمها كمرجع للأسلوب والبنية وحسّن عليها
- إذا كانت هناك وثائق إضافية: اقرأ محتواها واستخدم ما يخدم العرض
- اكتب workplan بعدد الأيام الصحيح المذكور في المناقصة
- احسب التسعير بدقة من بطاقة الأسعار
- غطِّ كل متطلب في قائمة المتطلبات أعلاه
- أعد JSON فقط بدون أي نص خارجه"""

    client = genai.Client(api_key=api_key, http_options={'timeout': 120000})
    cfg = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        temperature=0.4,
        max_output_tokens=16000,
    )

    last_exc = None
    for attempt in range(MAX_RETRIES):
        model = MODELS[min(attempt, len(MODELS) - 1)]
        try:
            if attempt > 0:
                wait = RETRY_WAITS[attempt - 1]
                logger.info(f'  Gemini retry {attempt + 1} in {wait}s (model: {model})...')
                time.sleep(wait)
                client = genai.Client(api_key=api_key, http_options={'timeout': 120000})
            logger.info(f'Generator: calling Gemini attempt {attempt + 1} with {model}...')
            raw = client.models.generate_content(
                model=model, contents=user_prompt, config=cfg
            ).text.strip()
            proposal = _extract_json(raw)
            proposal = _fix_pricing(proposal)
            logger.info(f'Generator: done — {proposal.get("program_title")} (model: {model})')
            return proposal
        except Exception as exc:
            last_exc = exc
            err_str = str(exc).lower()
            logger.warning(f'  Attempt {attempt + 1} ({model}) failed: {exc}')
            retryable = ('429' in err_str or 'resource_exhausted' in err_str or
                         '504' in err_str or 'deadline_exceeded' in err_str or
                         '503' in err_str or 'unavailable' in err_str or
                         'timeout' in err_str)
            if not retryable:
                break

    raise RuntimeError(f'Gemini failed after {MAX_RETRIES} attempts: {last_exc}')


def patch_proposal(current_data: dict, correction: str, api_key: str) -> dict:
    """Apply a user correction to an existing proposal. Returns the updated dict."""
    system = (
        f"{SYSTEM_PROMPT}\n\n"
        "أنت محرر عروض. طبّق التعديلات المطلوبة بدقة. "
        "لا تُغيِّر ما لم يُطلب. أعد الـ JSON الكامل المُعدَّل فقط."
    )
    prompt = f"""العرض الحالي:
---
{json.dumps(current_data, ensure_ascii=False, indent=2)[:12000]}
---

تعليمات التعديل:
{correction}

أعد JSON المُعدَّل كاملاً فقط."""

    client = genai.Client(api_key=api_key, http_options={'timeout': 120000})
    cfg = types.GenerateContentConfig(
        system_instruction=system,
        temperature=0.2,
        max_output_tokens=16000,
    )

    last_exc = None
    for attempt in range(MAX_RETRIES):
        model = MODELS[min(attempt, len(MODELS) - 1)]
        try:
            if attempt > 0:
                wait = RETRY_WAITS[attempt - 1]
                logger.info(f'  Patch retry {attempt + 1} in {wait}s (model: {model})...')
                time.sleep(wait)
                client = genai.Client(api_key=api_key, http_options={'timeout': 120000})
            raw = client.models.generate_content(
                model=model, contents=prompt, config=cfg
            ).text.strip()
            result = _extract_json(raw)
            result = _fix_pricing(result)
            return result
        except Exception as exc:
            last_exc = exc
            err_str = str(exc).lower()
            logger.warning(f'  Patch attempt {attempt + 1} ({model}) failed: {exc}')
            retryable = ('429' in err_str or 'resource_exhausted' in err_str or
                         '504' in err_str or 'deadline_exceeded' in err_str or
                         '503' in err_str or 'unavailable' in err_str or
                         'timeout' in err_str)
            if not retryable:
                break

    raise RuntimeError(f'Patch failed after {MAX_RETRIES} attempts: {last_exc}')
