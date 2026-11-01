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

مهمتك: اقرأ نص المناقصة كاملاً كلمة بكلمة واستخرج كل متطلب وقيد وشرط.
لا تفترض شيئاً — فقط ما هو مذكور صراحةً في النص.

أعد JSON فقط بهذا الهيكل — لا تضف أي نص خارج JSON:
{
  "client_name": "اسم الجهة الطالبة كما وردت بالضبط",
  "tender_number": "رقم المناقصة كما ورد أو 'غير مذكور'",
  "program_type": "نوع العمل المطلوب (برنامج تدريبي / مركز تقييم / استشارة / مشروع تطوير / ...)",
  "project_phase": "مرحلة المشروع المطلوبة (تصميم كامل / تنفيذ / تصميم وتنفيذ / تقييم / غير محدد)",
  "participants": "عدد المشاركين كرقم أو 'غير محدد'",
  "cohort_gender": "female إذا طُلبت مشاركات إناث — male إذا ذكور — mixed أو 'غير محدد'",
  "duration_days": "عدد الأيام كرقم أو 'غير محدد'",
  "delivery_location": "مكان التنفيذ كما ورد في المناقصة أو 'غير محدد'",
  "delivery_mode": "حضوري / عن بُعد / هجين / غير محدد",
  "budget_ceiling_omr": "سقف الميزانية بالأرقام فقط أو 0 إن لم يُذكر",
  "team_requirements": [
    "كل متطلب خاص بالفريق: الجنس، سنوات الخبرة، التخصص، الشهادات، الجنسية — كما ورد بالضبط"
  ],
  "coaching_type": "فردي / جماعي / كلاهما / غير مذكور",
  "curriculum_units": [
    "كل وحدة أو موضوع تعليمي مذكور بالاسم في المناقصة"
  ],
  "submission_documents": [
    "كل وثيقة أو نموذج مطلوب تقديمه مع العرض (نموذج ج، نموذج ب، شهادة وزارة العمل، تقرير الملاءة، ...)"
  ],
  "pricing_format": "وصف هيكل التسعير المطلوب إن كان محدداً في المناقصة — أو 'غير محدد'",
  "mandatory_requirements": [
    "كل متطلب إلزامي واضح يجب أن يغطيه العرض"
  ],
  "compliance_items": [
    "ضريبة القيمة المضافة، الضمان، الصلاحية، وأي التزام قانوني أو مالي مذكور"
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
        max_output_tokens=2048,
    )

    prompt = f"""اقرأ نص المناقصة التالية كاملاً كلمة بكلمة واستخرج جميع المتطلبات:

---
{draft_text[:10000]}
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
    if req.get('project_phase') and req['project_phase'] != 'غير محدد':
        lines.append(f'• مرحلة المشروع: {req["project_phase"]}')
    if req.get('participants') and req['participants'] != 'غير محدد':
        lines.append(f'• عدد المشاركين: {req["participants"]}')
    if req.get('cohort_gender') and req['cohort_gender'] not in ('غير محدد', 'mixed'):
        lines.append(f'• جنس المشاركين: {req["cohort_gender"]}')
    if req.get('duration_days') and req['duration_days'] != 'غير محدد':
        lines.append(f'• المدة: {req["duration_days"]} يوم')
    if req.get('delivery_location') and req['delivery_location'] != 'غير محدد':
        lines.append(f'• مكان التنفيذ: {req["delivery_location"]}')
    if req.get('delivery_mode') and req['delivery_mode'] != 'غير محدد':
        lines.append(f'• طريقة التقديم: {req["delivery_mode"]}')
    budget = req.get('budget_ceiling_omr', 0)
    if budget and int(budget) > 0:
        lines.append(f'• سقف الميزانية: {budget} ريال عماني — لا يجوز تجاوزه')
    if req.get('coaching_type') and req['coaching_type'] != 'غير مذكور':
        lines.append(f'• نوع التدريب/الكوتشينج: {req["coaching_type"]}')

    if req.get('team_requirements'):
        lines.append('\nمتطلبات الفريق — يجب مطابقتها بدقة:')
        for r in req['team_requirements']:
            lines.append(f'  ✦ {r}')

    if req.get('curriculum_units'):
        lines.append('\nوحدات المنهج المطلوبة — لا تخترع غيرها:')
        for u in req['curriculum_units']:
            lines.append(f'  ✦ {u}')

    if req.get('submission_documents'):
        lines.append('\nوثائق التقديم المطلوبة:')
        for d in req['submission_documents']:
            lines.append(f'  ✦ {d}')

    if req.get('pricing_format') and req['pricing_format'] != 'غير محدد':
        lines.append(f'\nهيكل التسعير المطلوب: {req["pricing_format"]}')

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
