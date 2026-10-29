"""
extractor.py — Requirements Extraction Agent for Zawad Proposal Agent.

Reads the raw tender text and extracts a structured checklist of:
  - Mandatory requirements (participants, days, budget ceiling, etc.)
  - Compliance items (VAT, bond, validity)
  - Preferred/optional requirements
  - Program type and client name

The checklist is injected into generator.py so Gemini writes the proposal
knowing exactly what it must cover. It is also passed to reviewer.py so the
adversarial reviewer can verify every requirement is addressed.
"""

import json
import logging
from google import genai
from google.genai import types

logger = logging.getLogger(__name__)

MODEL = 'gemini-3.5-flash-lite'

EXTRACTOR_SYSTEM = """أنت محلل مناقصات خبير ومتخصص في تحليل وثائق المناقصات الحكومية العُمانية.

مهمتك الوحيدة: استخرج كل متطلب وقيد وشرط مذكور في نص المناقصة.

أعد JSON فقط بهذا الهيكل — لا تضف أي نص خارج JSON:
{
  "client_name": "اسم الجهة الطالبة كما وردت بالضبط",
  "tender_number": "رقم المناقصة كما ورد أو 'غير مذكور'",
  "program_type": "نوع العمل المطلوب (برنامج تدريبي / مركز تقييم / استشارة / ...)",
  "participants": "عدد المشاركين كرقم أو 'غير محدد'",
  "duration_days": "عدد الأيام كرقم أو 'غير محدد'",
  "budget_ceiling_omr": "سقف الميزانية بالأرقام فقط أو 0 إن لم يُذكر",
  "mandatory_requirements": [
    "كل متطلب إلزامي واضح يجب أن يغطيه العرض"
  ],
  "compliance_items": [
    "ضريبة القيمة المضافة 5%",
    "أي التزام قانوني أو مالي مذكور في وثيقة المناقصة"
  ],
  "preferred_requirements": [
    "أي متطلب مفضل أو إضافي غير إلزامي"
  ],
  "red_flags": [
    "أي شرط استبعادي أو نقطة حرجة قد تؤدي لرفض العرض إن أُهملت"
  ]
}"""


def extract_requirements(draft_text: str, api_key: str) -> dict:
    """
    Extract structured requirements from tender text.
    Returns a dict. Returns {} on failure (caller handles gracefully).
    """
    client = genai.Client(api_key=api_key, http_options={'timeout': 25000})
    cfg = types.GenerateContentConfig(
        system_instruction=EXTRACTOR_SYSTEM,
        temperature=0.0,
        max_output_tokens=512,
    )

    prompt = f"""استخرج جميع المتطلبات من نص المناقصة التالية:

---
{draft_text[:2000]}
---

أعد JSON فقط."""

    try:
        response = client.models.generate_content(
            model=MODEL, contents=prompt, config=cfg
        )
        raw = response.text.strip()
        if raw.startswith('```'):
            raw = raw.split('```')[1]
            if raw.startswith('json'):
                raw = raw[4:]
        return json.loads(raw.strip())
    except Exception as exc:
        logger.warning(f'Extractor skipped (will continue without checklist): {exc}')
        return {}


def format_checklist(req: dict) -> str:
    """
    Convert the extracted requirements dict into a formatted Arabic checklist
    string ready to inject into the generator prompt.
    """
    if not req:
        return ''

    lines = ['=== قائمة متطلبات المناقصة المستخرجة (يجب تغطية كل بند) ===']

    if req.get('client_name'):
        lines.append(f'• الجهة: {req["client_name"]}')
    if req.get('tender_number') and req['tender_number'] != 'غير مذكور':
        lines.append(f'• رقم المناقصة: {req["tender_number"]}')
    if req.get('program_type'):
        lines.append(f'• نوع البرنامج: {req["program_type"]}')
    if req.get('participants') and req['participants'] != 'غير محدد':
        lines.append(f'• عدد المشاركين: {req["participants"]}')
    if req.get('duration_days') and req['duration_days'] != 'غير محدد':
        lines.append(f'• المدة: {req["duration_days"]} يوم')
    budget = req.get('budget_ceiling_omr', 0)
    if budget and int(budget) > 0:
        lines.append(f'• سقف الميزانية: {budget} ريال عماني — لا يجوز تجاوزه')

    if req.get('mandatory_requirements'):
        lines.append('\nالمتطلبات الإلزامية — يجب تغطية كل متطلب بوضوح في العرض:')
        for r in req['mandatory_requirements']:
            lines.append(f'  ✦ {r}')

    if req.get('compliance_items'):
        lines.append('\nبنود الامتثال الإلزامية:')
        for c in req['compliance_items']:
            lines.append(f'  ✦ {c}')

    if req.get('red_flags'):
        lines.append('\n⚠️ نقاط حرجة — تجاهلها يؤدي لرفض العرض:')
        for f in req['red_flags']:
            lines.append(f'  ⚠ {f}')

    if req.get('preferred_requirements'):
        lines.append('\nمتطلبات مفضلة (إن أمكن تضمينها):')
        for p in req['preferred_requirements']:
            lines.append(f'  ◦ {p}')

    return '\n'.join(lines)


def requirements_summary(req: dict) -> str:
    """Short one-line summary for the chatbot UI message."""
    if not req:
        return ''
    n = len(req.get('mandatory_requirements', []))
    flags = len(req.get('red_flags', []))
    budget = req.get('budget_ceiling_omr', 0)
    parts = [f'{n} متطلباً إلزامياً']
    if flags:
        parts.append(f'{flags} نقطة حرجة')
    if budget and int(budget) > 0:
        parts.append(f'سقف الميزانية: {budget} ريال')
    return ' | '.join(parts)
