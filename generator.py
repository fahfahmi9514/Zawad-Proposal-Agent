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
مهمتك: إنشاء عرض فني ومالي احترافي كامل لشركة زواد للتطوير والاستشارات.

=== اللغة ===
- إذا طلب المستخدم الإنجليزية (English / in English / بالإنجليزي)، اكتب language: "en" واكتب جميع النصوص بالإنجليزية.
- إذا طلب العربية أو لم يحدد، اكتب language: "ar" واكتب جميع النصوص بالعربية.
- لا تخلط بين اللغتين في نفس الحقل.

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

=== الجنس والضمائر — إلزامي ===
- إذا كانت المناقصة تستهدف فئة نسائية أو تطلب مدربات إناث أو قائدات:
  • اكتب cohort_gender: "female"
  • استخدم الصيغة المؤنثة في كل النصوص: مشاركة / مدربة / قائدة / خبيرة
  • أضف ملاحظة في الفريق أن 70%+ من الفريق من الخبيرات الإناث
- إذا كانت الفئة مختلطة أو غير محدد: cohort_gender: "mixed"

=== هيكل الوحدات التعليمية — إلزامي ===
- قسّم workplan إلى وحدات (وحدة 1، وحدة 2...) وليس أياماً مجردة
- يجب تغطية المواضيع المطلوبة حسب سياق المناقصة:
  الذكاء العاطفي | إدارة الأهداف بالنتائج MBO/KPIs | الذكاء الاصطناعي في العمل
  التميز المؤسسي | بناء الفريق وديناميكياته | إنتاج المحتوى | التخطيط الاستراتيجي
- كل وحدة تشمل: العنوان، المفاهيم الرئيسية، الأنشطة التطبيقية، مخرجات التعلم

=== 10 مشاريع تطبيقية ===
- أضف applied_projects: قائمة 10 مشاريع تطبيقية جماعية على مدى شهرين
- كل مشروع: العنوان، الهدف، المدة، طريقة التوجيه والإشراف

=== جلسة الحوار القيادي ===
- أضف leadership_dialogue: اقترح متحدثاً ومحاور الجلسة

=== قياس التعلم (الشريحة 08) ===
- أضف learning_measurement: أدوات وأساليب قياس التعلم (قبلي / بعدي / أثر)
- اذكر: نموذج كيركباتريك، التقييم القبلي والبعدي، متابعة 60-90 يوماً

=== الإقرارات الإلزامية ===
- أضف ip_acknowledgment: جميع المواد المنتجة ملك للجهة بعد إتمام السداد
- أضف party2_obligations: التزامات الجهة (توفير المواد 3 أسابيع مسبقاً، السفر والإقامة، الهوية البصرية)

=== التسعير — قواعد صارمة ===
- إذا كانت المناقصة تحتوي نموذج مالي (نموذج ج / جدول أسعار): اقرأ هيكله وبنوده واتبعه بالضبط
- إذا لم يوجد نموذج محدد: استخدم البنية القياسية (تصميم + تنفيذ + تقرير)
- احسب: total_omr = quantity × rate_omr لكل بند
- subtotal_omr = مجموع كل total_omr
- حدد vat_rate: استخدم نسبة الضريبة من المناقصة (مثال: 0.05 أو 0.03). إن طُلبت أسعار بدون ضريبة اكتب 0. إن لم تُذكر، اكتب 0.05 للجهات الحكومية.
- vat_omr = subtotal_omr × vat_rate
- grand_total_omr = subtotal_omr + vat_omr
- لا تتجاوز سقف الميزانية إن ذُكر

=== الفريق — استخراج السيرة الذاتية ===
- استخرج جميع المعلومات من السيرة الذاتية المرفقة — لا تخترع معلومات.
- اكتب experience: فقرتين كاملتين (5 جمل+) من محتوى السيرة الذاتية.
- اكتب experience_en: نفس المحتوى بالإنجليزية.
- اكتب qualifications: قائمة من 3-5 بنود (مؤهلات، خبرات، شهادات) مستخرجة من السيرة.
- اكتب qualifications_en: نفس القائمة بالإنجليزية.
- اكتب years: عدد سنوات الخبرة كنص مثل "15+" أو "10+" (استخرجه من السيرة).
- إن كان الاسم معروفاً لكن المعلومات ناقصة: أضف البطاقة بالاسم فقط واترك باقي الحقول فارغة ("" أو []).
- د. غسان المعمري يجب أن يكون دائماً الأول في قائمة الفريق.

=== هيكل JSON المطلوب — أعد هذا الهيكل كاملاً فقط ===
{
  "language": "ar",
  "cohort_gender": "mixed",
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
      {
        "name": "د. غسان المعمري",
        "name_en": "Dr. Ghasan Al-Maamari",
        "role": "مدير المشروع ورئيس الفريق",
        "role_en": "Project Director & Team Lead",
        "years": "20+",
        "experience": "فقرتان كاملتان 5 جمل+ من السيرة الذاتية بالعربية",
        "experience_en": "Two full paragraphs 5+ sentences from CV in English",
        "qualifications": ["مؤهل أكاديمي", "خبرة واسعة", "شهادة مهنية"],
        "qualifications_en": ["Academic qualification", "Extensive experience", "Professional cert"]
      }
    ],
    "expected_outputs": [
      {"number": "01", "title": "عنوان", "description": "وصف"},
      {"number": "02", "title": "عنوان", "description": "وصف"},
      {"number": "03", "title": "عنوان", "description": "وصف"},
      {"number": "04", "title": "عنوان", "description": "وصف"},
      {"number": "05", "title": "عنوان", "description": "وصف"}
    ],
    "track_record_highlight": "سجل الإنجازات — فقرتان",
    "applied_projects": [
      {"title": "عنوان المشروع", "objective": "الهدف", "duration": "أسبوعان", "mentoring": "كيفية التوجيه"}
    ],
    "leadership_dialogue": {
      "speaker_proposal": "اقتراح المتحدث وخلفيته",
      "session_axes": ["محور 1", "محور 2", "محور 3"]
    },
    "learning_measurement": {
      "pre_assessment": "التقييم القبلي وأدواته",
      "post_assessment": "التقييم البعدي وأدواته",
      "impact_followup": "متابعة الأثر بعد 60-90 يوماً وفق كيركباتريك",
      "tools": ["أداة 1", "أداة 2", "أداة 3"]
    },
    "ip_acknowledgment": "جميع المواد والمحتوى المنتج في إطار هذا البرنامج ملك حصري للجهة المستفيدة بعد إتمام السداد الكامل، ولا يحق لزواد استخدامها لصالح جهات أخرى.",
    "party2_obligations": [
      "توفير المواد والوثائق المرجعية قبل 3 أسابيع من انطلاق البرنامج",
      "تغطية نفقات السفر والإقامة للمدربين خارج مسقط إن اقتضى الأمر",
      "توفير الهوية البصرية والشعار الرسمي للجهة للاستخدام في مواد البرنامج",
      "تسهيل إجراءات الدخول للمنشأة ومنح الوصول اللازم للمدربين"
    ],
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


def _load_team_cvs() -> str:
    """Load team CV library and return a compact summary for the prompt."""
    try:
        import os
        cv_path = os.path.join(os.path.dirname(__file__), 'config', 'team_cvs.json')
        with open(cv_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        members = data.get('members', [])
        lines = ['=== مكتبة السير الذاتية لفريق زواد — اختر من هذه الأسماء لقسم الفريق ===']
        for m in members:
            name = m.get('name_ar', m.get('name_en', ''))
            role = m.get('current_role', '')
            years = m.get('years_experience', '')
            specs = ' | '.join(m.get('specializations', [])[:4])
            certs = ', '.join(m.get('certifications', [])[:4])
            gender = m.get('gender', '')
            bio = m.get('bio_ar', '')
            suitable = ' | '.join(m.get('suitable_for', [])[:4])
            lines.append(
                f'\n[{name}] ({gender})\n'
                f'  الدور: {role}\n'
                f'  الخبرة: {years} سنة\n'
                f'  التخصصات: {specs}\n'
                f'  الشهادات: {certs or "—"}\n'
                f'  مناسب لـ: {suitable}\n'
                f'  نبذة: {bio[:300]}'
            )
        return '\n'.join(lines)
    except Exception as exc:
        logger.warning(f'CV library load skipped: {exc}')
        return ''


def generate_proposal(draft_text: str, api_key: str,
                      checklist: str = None) -> dict:
    """
    Generate a complete proposal with a single Gemini call.
    checklist: formatted requirements string from extractor.format_checklist().
    Returns the complete proposal dict ready for build_pptx().
    """
    rag = _get_rag(draft_text, api_key)
    cv_library = _load_team_cvs()

    checklist_section = ''
    if checklist:
        checklist_section = f'\n\n=== قائمة المتطلبات المستخرجة — يجب تغطية كل بند ===\n{checklist[:800]}'

    rag_section = ''
    if rag:
        rag_section = f'\n\n=== أمثلة من عروض زواد السابقة (للأسلوب فقط) ===\n{rag[:1000]}'

    cv_section = ''
    if cv_library:
        cv_section = f'\n\n{cv_library[:4000]}\n\nتعليمات الفريق: اختر 2-4 أعضاء من المكتبة أعلاه حسب متطلبات المناقصة. إذا طلبت المناقصة مدربات إناث، أعطِ الأولوية للأعضاء ذوي gender: female. د. غسان المعمري يجب أن يكون دائماً الأول.'

    user_prompt = f"""الوثائق المرفقة (اقرأ كل وثيقة بالكامل وافهم دورها):
---
{draft_text[:12000]}
---
{checklist_section}{rag_section}{cv_section}

تعليمات:
- إذا كانت هناك وثيقة مناقصة: استخرج منها متطلبات الجهة وأهدافها واكتب العرض ليغطيها بالكامل
- إذا كانت هناك مسودة عرض سابقة: استخدمها كمرجع للأسلوب والبنية وحسّن عليها
- إذا كانت هناك وثائق إضافية: اقرأ محتواها واستخدم ما يخدم العرض
- اكتب workplan بعدد الأيام الصحيح المذكور في المناقصة
- احسب التسعير بدقة من بطاقة الأسعار
- غطِّ كل متطلب في قائمة المتطلبات أعلاه
- استخدم السير الذاتية من مكتبة الفريق لملء قسم team بمعلومات حقيقية
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
