"""
chatbot.py — Zawad Proposal Chatbot
Local web app. Run with: python chatbot.py (or double-click run_chatbot.bat)
Opens http://localhost:5050 in browser automatically.
"""

import sys
import json
import time
import threading
import webbrowser
import logging
from pathlib import Path
from flask import Flask, request, jsonify
import tempfile

# ── Paths ────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
CONFIG_FILE   = BASE_DIR / 'config' / 'config.json'
COMPLETED_DIR = Path.home() / 'Downloads' / 'Proposals_2026' / 'completed'
REVIEW_DIR    = Path.home() / 'Downloads' / 'Proposals_2026' / 'review_needed'

sys.path.insert(0, str(BASE_DIR))

# ── Logging ──────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s — %(levelname)s — %(message)s',
    handlers=[
        logging.FileHandler(str(BASE_DIR / 'chatbot.log'), encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)

# ── Flask app ────────────────────────────────────────────────────
app = Flask(__name__, static_folder=str(BASE_DIR / 'chat_ui'), static_url_path='')


# ── Config helpers ───────────────────────────────────────────────

def _load_api_key() -> str:
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f).get('api_key', '')
    except Exception:
        return ''


def _save_api_key(key: str):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'api_key': key}, f, indent=2)


# ── Session state (single user, in-memory) ───────────────────────
_api_key = _load_api_key()

session = {
    'api_key':       _api_key,
    'proposal_data': None,
    'draft_text':    None,
    'pptx_path':     None,
    'req_data':      {},   # requirements extracted from tender — kept for re-review after edits
    'staged_texts':  [],   # file texts waiting to be combined into next generation
    'mode':          'await_draft' if _api_key else 'await_key',
}


# ── Routes ───────────────────────────────────────────────────────

@app.route('/')
def index():
    return app.send_static_file('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    """Extract text from an uploaded file and stage it — generation happens on Send."""
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'لم يتم إرفاق أي ملف.'}), 400

    f = request.files['file']
    if not f.filename:
        return jsonify({'ok': False, 'error': 'اسم الملف فارغ.'}), 400

    suffix = Path(f.filename).suffix.lower()
    allowed = {'.pdf', '.docx', '.pptx', '.txt'}
    if suffix not in allowed:
        return jsonify({'ok': False,
                        'error': f'نوع الملف غير مدعوم ({suffix}). المسموح به: PDF, Word (docx), PowerPoint (pptx), TXT.'})

    try:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp_path = Path(tmp.name)
            f.save(tmp_path)

        sys.path.insert(0, str(BASE_DIR))
        from reader import read_file
        text = read_file(tmp_path)
        tmp_path.unlink(missing_ok=True)

        if not text or len(text.strip()) < 30:
            return jsonify({'ok': False, 'error': f'الملف "{f.filename}" فارغ أو لا يحتوي على نص كافٍ.'})

        session['staged_texts'].append({'filename': f.filename, 'text': text})
        logger.info(f'Staged file: {f.filename} — {len(text):,} chars (total staged: {len(session["staged_texts"])})')
        return jsonify({'ok': True, 'filename': f.filename, 'chars': len(text)})

    except Exception as exc:
        logger.error(f'Upload error: {exc}', exc_info=True)
        return jsonify({'ok': False, 'error': f'خطأ أثناء قراءة الملف: {str(exc)[:200]}'})


@app.route('/api/status')
def status():
    return jsonify({
        'mode':         session['mode'],
        'has_proposal': session['proposal_data'] is not None,
        'program':      (session['proposal_data'] or {}).get('program_title', ''),
        'client':       (session['proposal_data'] or {}).get('client_name', ''),
    })


@app.route('/api/chat', methods=['POST'])
def chat():
    body = request.get_json(silent=True) or {}
    user_msg = (body.get('message') or '').strip()
    mode = session['mode']

    # ── Combine staged files (skip when collecting API key — never save tender as key) ──
    staged = session.get('staged_texts', [])
    if staged and mode != 'await_key':
        session['staged_texts'] = []  # consume immediately
        # Label each file clearly so Gemini knows what it's reading
        parts = []
        for i, item in enumerate(staged, 1):
            label = f'=== الوثيقة {i}: {item["filename"]} ==='
            parts.append(f'{label}\n{item["text"]}')
        file_block = '\n\n'.join(parts)
        # If user also typed something (not just the auto-trigger), append it
        AUTO_TRIGGER = 'تم رفع الملفات. يرجى المتابعة.'
        typed = user_msg if user_msg and user_msg != AUTO_TRIGGER else ''
        user_msg = (file_block + ('\n\n' + typed if typed else '')).strip()
        # Files always mean: generate a new proposal (reset if a proposal exists)
        if mode == 'ready':
            session['mode'] = 'await_draft'
            session['proposal_data'] = None
            session['draft_text'] = None
            session['pptx_path'] = None
            mode = 'await_draft'

    if not user_msg:
        return jsonify({'reply': 'الصق نص المناقصة أو أرفق ملفاً ثم اضغط إرسال.', 'mode': mode})

    # ── API key collection ────────────────────────────────────────────────────
    if mode == 'await_key':
        if len(user_msg) < 15:
            return jsonify({'reply': '⚠️ لا يبدو هذا مفتاح API صحيحاً. انسخه من Google AI Studio والصقه هنا.', 'mode': mode})
        session['api_key'] = user_msg
        _save_api_key(user_msg)
        session['mode'] = 'await_draft'
        return jsonify({'reply': '✅ تم حفظ المفتاح بنجاح!\n\nالآن **الصق نص المناقصة أو المسودة** وسأُنشئ لك العرض خلال دقائق.', 'mode': 'await_draft'})

    # ── Draft collection ──────────────────────────────────────────────────────
    if mode == 'await_draft':
        return _generate_new(user_msg)

    # ── Correction / refinement ───────────────────────────────────────────────
    if mode == 'ready':
        reset_kws = ['عرض جديد', 'مناقصة جديدة', 'new proposal', 'ابدأ من جديد', 'restart', 'جديد']
        if any(kw in user_msg.lower() for kw in reset_kws):
            session['mode'] = 'await_draft'
            session['proposal_data'] = None
            session['draft_text'] = None
            session['pptx_path'] = None
            session['staged_texts'] = []
            return jsonify({'reply': '🔄 حسناً، نبدأ من جديد.\n\nالصق نص المناقصة الجديدة.', 'mode': 'await_draft'})
        return _patch_existing(user_msg)

    return jsonify({'reply': 'حالة غير معروفة. أعد تشغيل البرنامج.', 'mode': mode})


# ── Generation ───────────────────────────────────────────────────

def _safe_filename(title: str) -> str:
    base = ''.join(c for c in title if c.isalnum() or c in ' _-')[:35].strip() or 'Proposal'
    ts = time.strftime('%m%d_%H%M%S')
    return f'{base}_{ts}'


def _write_review_file(proposal_data: dict, req_data: dict, safe_name: str):
    """Write a structured Arabic review report to review_needed folder."""
    try:
        REVIEW_DIR.mkdir(parents=True, exist_ok=True)
        sections   = proposal_data.get('sections', {})
        pricing    = sections.get('pricing', {})
        total_str  = str(pricing.get('grand_total_omr', '0')).replace(',', '')
        subtotal   = pricing.get('subtotal_omr', 0)
        vat        = pricing.get('vat_omr', 0)
        total      = pricing.get('grand_total_omr', '—')
        team       = sections.get('team', [])
        track_text = sections.get('track_record_highlight', '')
        risks      = sections.get('risks', [])
        workplan   = sections.get('workplan', [])
        methodology    = sections.get('methodology', '')
        company_overview = sections.get('company_overview', '')
        commitments    = sections.get('commitments', '')
        innovation     = sections.get('innovation', '')
        quality        = sections.get('quality_commitment', '')
        intro_text     = sections.get('requirements_understanding', '')

        passed, warned, failed = [], [], []
        score = 10.0

        # ── Check each section ────────────────────────────────────────────────

        def _lines(s):
            return len([l for l in str(s).split('\n') if l.strip()])

        # Style / language
        if _lines(company_overview) >= 6:
            passed.append(('الأسلوب الرسمي وجودة النص', 'الصياغة احترافية وتتسم بالفصاحة المطلوبة لمقام الجهة المستفيدة.'))
        else:
            warned.append(('نص نبذة الشركة', 'أقل من 6 أسطر — يُوصى بتوسيع المحتوى.'))
            score -= 0.5

        # Vision 2040
        full_text = str(proposal_data)
        if 'رؤية عُمان 2040' in full_text or '2040' in full_text:
            passed.append(('ربط رؤية عُمان 2040', 'ورد الربط بالرؤية في السياق المناسب.'))
        else:
            failed.append(('غياب رؤية عُمان 2040', 'لم يُذكر ربط البرنامج برؤية عُمان 2040 — وهو متطلب إلزامي.'))
            score -= 1.5

        # Requirements understanding
        if _lines(intro_text) >= 6:
            passed.append(('فهم متطلبات المناقصة', f'القسم وافٍ ومخصص للجهة ({_lines(intro_text)} أسطر).'))
        else:
            warned.append(('فهم المتطلبات', f'القسم قصير ({_lines(intro_text)} أسطر) — يُفضل توسيعه.'))
            score -= 0.5

        # Methodology
        if _lines(methodology) >= 5:
            passed.append(('المنهجية المقترحة', f'منهجية واضحة ({_lines(methodology)} أسطر).'))
        else:
            warned.append(('المنهجية المقترحة', f'المنهجية مختصرة ({_lines(methodology)} أسطر) — يُوصى بتفصيلها.'))
            score -= 0.5

        # Workplan
        if len(workplan) >= 3:
            passed.append(('خطة العمل', f'خطة تنفيذية مكتملة بـ {len(workplan)} وحدات/أيام.'))
        else:
            failed.append(('خطة العمل', f'عدد أيام/وحدات الخطة ({len(workplan)}) أقل من الحد المطلوب.'))
            score -= 1.0

        # Team
        if len(team) >= 2:
            ghasan = any('غسان' in str(m) or 'المعمري' in str(m) for m in team)
            if ghasan:
                passed.append(('فريق العمل', f'فريق مكتمل ({len(team)} أعضاء) بقيادة د. غسان المعمري.'))
            else:
                failed.append(('اسم مدير المشروع', 'لم يرد اسم د. غسان المعمري في قسم الفريق — مطلوب دائماً.'))
                score -= 1.5
        else:
            failed.append(('فريق العمل', f'عدد أعضاء الفريق ({len(team)}) أقل من الحد المطلوب (2 على الأقل).'))
            score -= 1.5

        # Track record (field is track_record_highlight string; slides use hardcoded PAST_PROJECTS)
        if _lines(track_text) >= 2:
            passed.append(('سجل الإنجازات', 'نص سجل الإنجازات مكتمل — المشاريع السابقة موثقة في الشرائح.'))
        else:
            warned.append(('سجل الإنجازات', 'نص نبذة الإنجازات قصير — يُوصى بتوسيعه.'))
            score -= 0.5

        # Risks
        if len(risks) >= 3:
            passed.append(('إدارة المخاطر', f'{len(risks)} مخاطر موثقة مع إجراءاتها الوقائية.'))
        else:
            warned.append(('إدارة المخاطر', f'عدد المخاطر ({len(risks)}) أقل من 5 الموصى بها.'))
            score -= 0.5

        # Innovation
        if _lines(innovation) >= 3:
            passed.append(('الابتكار والقيمة المضافة', 'قسم مكتمل.'))
        else:
            warned.append(('الابتكار والقيمة المضافة', f'القسم مختصر ({_lines(innovation)} أسطر) — يُوصى بتوسيعه.'))
            score -= 0.5

        # Quality commitment
        if _lines(quality) >= 3:
            passed.append(('ضمان الجودة والحوكمة', 'قسم مكتمل.'))
        else:
            warned.append(('ضمان الجودة والحوكمة', 'القسم أقل من 3 أسطر — أضف تفاصيل نموذج القياس (كيركباتريك أو ما يعادله).'))
            score -= 0.5

        # Pricing math
        try:
            expected_total = round(float(subtotal) * 1.05)
            actual_total = int(str(total).replace(',', ''))
            if abs(expected_total - actual_total) <= 5:
                passed.append(('الدقة الحسابية', f'المجموع الفرعي {subtotal:,} + ضريبة 5% ({vat:,}) = الإجمالي {total} ر.ع ✓'))
            else:
                failed.append(('الدقة الحسابية', f'خطأ في الحساب: {subtotal:,} × 1.05 = {expected_total:,} لكن الإجمالي المسجل {total}.'))
                score -= 2.0
        except Exception:
            warned.append(('التحقق الحسابي', 'تعذّر التحقق من صحة الأرقام — راجع التسعير يدوياً.'))

        # Compliance — 1.2% procurement fee
        if '1.2' in full_text or 'تعميم' in full_text or 'المشتريات' in full_text:
            passed.append(('رسوم المشتريات 1.2% (تعميم 13/2024)', 'مذكورة في العرض.'))
        else:
            failed.append(('رسوم حماية المنافسة 1.2%', 'غير مذكورة — إغفالها يُعرّض العرض للاستبعاد (تعميم وزارة المالية 13/2024).'))
            score -= 1.5

        # Compliance — 5% bond
        if 'ضمان' in full_text and ('5%' in full_text or '٥٪' in full_text):
            passed.append(('ضمان حسن التنفيذ 5%', 'مذكور في الالتزامات.'))
        else:
            failed.append(('ضمان حسن التنفيذ 5%', 'لم يُنص عليه صراحة في الإقرارات — يجب إضافته.'))
            score -= 0.5

        # Budget ceiling
        budget = req_data.get('budget_ceiling_omr', 0) if req_data else 0
        if budget and int(budget) > 0:
            try:
                if int(total_str) > int(budget):
                    failed.append(('سقف الميزانية', f'الإجمالي ({total} ر.ع) يتجاوز سقف المناقصة ({int(budget):,} ر.ع) — مرفوض حتماً.'))
                    score -= 3.0
                else:
                    passed.append(('سقف الميزانية', f'الإجمالي ({total} ر.ع) ضمن السقف المحدد ({int(budget):,} ر.ع).'))
            except Exception:
                pass

        # Mandatory requirements from extractor
        missed_reqs = []
        if req_data:
            for req in req_data.get('mandatory_requirements', []):
                if req[:20] not in full_text:
                    missed_reqs.append(req)
            if missed_reqs:
                for r in missed_reqs[:3]:
                    warned.append(('متطلب غير محقق', r))
                score -= len(missed_reqs) * 0.3

        # ── Build the report ──────────────────────────────────────────────────
        score = max(0.0, min(10.0, round(score, 1)))
        if score >= 8:
            readiness = 'جاهز للتقديم ✅'
        elif score >= 6:
            readiness = 'يحتاج تحسينات قبل التقديم ⚠️'
        else:
            readiness = 'غير جاهز للتقديم ❌'

        sep = '=' * 80
        lines = [
            sep,
            'تقرير مراجعة العرض — زواد للتطوير والاستشارات',
            sep,
            f'البرنامج: {proposal_data.get("program_title", "—")}',
            f'الجهة:    {proposal_data.get("client_name", "—")}',
            f'التاريخ:  {time.strftime("%Y-%m-%d")}',
            sep,
            '',
            '## ملخص تنفيذي',
            f'درجة العرض الإجمالية هي **{score} / 10**؛ العرض **{readiness}**.',
            '',
            '-' * 80,
        ]

        if passed:
            lines += ['', '## ✅ ما هو ناجح ومكتمل', '']
            for title, detail in passed:
                lines.append(f'- **{title}: ✅ ناجح**')
                lines.append(f'  - {detail}')

        if warned:
            lines += ['', '-' * 80, '', '## ⚠️ يحتاج تحسين', '']
            for title, detail in warned:
                lines.append(f'- **{title}: ⚠️ يحتاج تحسين**')
                lines.append(f'  - {detail}')

        if failed:
            lines += ['', '-' * 80, '', '## ❌ مشاكل حرجة يجب إصلاحها قبل التقديم', '']
            for i, (title, detail) in enumerate(failed, 1):
                lines.append(f'{i}. **{title}: ❌ فاشل**')
                lines.append(f'   - {detail}')

        if req_data and req_data.get('red_flags'):
            lines += ['', '-' * 80, '', '## 🚨 نقاط حرجة من المناقصة', '']
            for f in req_data['red_flags']:
                lines.append(f'- ⚠ {f}')

        lines += ['', sep]

        review_path = REVIEW_DIR / f'Review_{safe_name}.txt'
        review_path.write_text('\n'.join(lines), encoding='utf-8')
        logger.info(f'Review file saved: {review_path.name} — score {score}/10')
        return review_path.name
    except Exception as exc:
        logger.warning(f'Could not write review file: {exc}')
        return None


def _validate_proposal(data: dict) -> list:
    """
    Check proposal structure. Returns a list of warning strings (never blocks).
    Only raises ValueError if the JSON is so broken build_pptx would crash.
    """
    warnings = []

    if not isinstance(data, dict):
        raise ValueError('Gemini response is not valid JSON — please retry.')

    if not isinstance(data.get('sections'), dict):
        raise ValueError('sections field missing or invalid in Gemini response — please retry.')

    sections = data['sections']

    if 'workplan' in sections and not isinstance(sections['workplan'], list):
        raise ValueError('workplan must be a list — please retry.')
    workplan = sections.get('workplan', [])
    if not workplan:
        warnings.append('⚠️ خطة العمل (workplan) فارغة — تم استخدام قيم افتراضية.')
    else:
        bad = [i for i, d in enumerate(workplan) if not isinstance(d, dict)]
        if bad:
            raise ValueError(f'workplan items at index {bad} are not objects — please retry.')

    if 'team' in sections and not isinstance(sections['team'], list):
        raise ValueError('team must be a list — please retry.')
    team = sections.get('team', [])
    if not team:
        warnings.append('⚠️ فريق العمل فارغ — تم استخدام قيم افتراضية.')
    else:
        bad = [i for i, d in enumerate(team) if not isinstance(d, dict)]
        if bad:
            raise ValueError(f'team items at index {bad} are not objects — please retry.')

    pricing = sections.get('pricing', {})
    if not isinstance(pricing, dict):
        raise ValueError('pricing must be a JSON object — please retry.')
    if 'line_items' in pricing and not isinstance(pricing['line_items'], list):
        raise ValueError('line_items must be a list — please retry.')
    if not pricing.get('line_items'):
        warnings.append('⚠️ بنود التسعير فارغة — راجع قسم التسعير يدوياً.')

    for field in ['methodology', 'requirements_understanding', 'company_overview']:
        if not sections.get(field):
            warnings.append(f'⚠️ القسم "{field}" فارغ في المخرجات.')

    return warnings


def _generate_new(draft_text: str):
    try:
        COMPLETED_DIR.mkdir(parents=True, exist_ok=True)

        from generator import generate_proposal
        from builder import build_pptx
        from extractor import extract_requirements, format_checklist, requirements_summary

        # ── Step 1: Extract requirements ──────────────────────────────────────
        logger.info('Chatbot: extracting requirements...')
        req_data = {}
        checklist_text = ''
        req_line = ''
        try:
            req_data = extract_requirements(draft_text, session['api_key'])
            checklist_text = format_checklist(req_data)
            summary = requirements_summary(req_data)
            if summary:
                req_line = f'\n📋 **المتطلبات المستخرجة:** {summary}'
            logger.info(f'Extracted: {len(req_data.get("mandatory_requirements", []))} requirements')
        except Exception as exc:
            logger.warning(f'Extractor skipped: {exc}')

        # ── Step 2: Generate proposal ─────────────────────────────────────────
        logger.info('Chatbot: generating proposal...')
        proposal_data = generate_proposal(
            draft_text, session['api_key'],
            checklist=checklist_text or None
        )
        logger.info(f'Generated: {proposal_data.get("program_title")}')

        # ── Step 2b: Check JSON structure — collect warnings, never block ────
        validation_warnings = _validate_proposal(proposal_data)
        if validation_warnings:
            logger.warning('Proposal warnings: ' + ' | '.join(validation_warnings))

        # ── Step 3: Build PPTX ────────────────────────────────────────────────
        safe = _safe_filename(proposal_data.get('program_title', 'Proposal'))
        pptx_path = COMPLETED_DIR / f'Zawad_{safe}.pptx'
        build_pptx(proposal_data, str(pptx_path))
        logger.info(f'PPTX saved: {pptx_path.name}')

        # ── Step 4: Write review checklist ───────────────────────────────────
        review_name = _write_review_file(proposal_data, req_data, safe)
        review_line = f'\n📋 **قائمة المراجعة:** `review_needed/{review_name}`' if review_name else ''

        session['proposal_data'] = proposal_data
        session['draft_text'] = draft_text
        session['pptx_path'] = str(pptx_path)
        session['req_data'] = req_data
        session['mode'] = 'ready'

        pricing = proposal_data.get('sections', {}).get('pricing', {})
        total = pricing.get('grand_total_omr', '—')

        warnings_block = ''
        if validation_warnings:
            warnings_block = '\n\n**ملاحظات على المخرجات:**\n' + '\n'.join(validation_warnings)

        reply = f"""✅ **تم إنشاء العرض بنجاح!**

| البرنامج | {proposal_data.get('program_title', '—')} |
|----------|------|
| الجهة | {proposal_data.get('client_name', '—')} |
| الإجمالي | {total} ريال عماني |
{req_line}{review_line}
📁 **ملف العرض:** `completed/{pptx_path.name}`{warnings_block}

---
💬 هل تريد أي تعديلات؟ اكتبها بالعربية أو الإنجليزية وسأطبّقها فوراً.
مثال: *"غيّر مدة التدريب إلى أسبوعين"* أو *"اسم الجهة هو وزارة المالية"*

اكتب **عرض جديد** لبدء مناقصة مختلفة."""

        return jsonify({'reply': reply, 'mode': 'ready'})

    except Exception as exc:
        logger.error(f'Generation error: {exc}', exc_info=True)
        return jsonify({
            'reply': f'❌ **حدث خطأ أثناء الإنشاء:**\n\n`{str(exc)[:300]}`\n\nتأكد من صحة مفتاح API وأن الإنترنت يعمل.',
            'mode': session['mode']
        })


# ── Patch ────────────────────────────────────────────────────────

def _patch_existing(correction: str):
    if not session['proposal_data']:
        return jsonify({'reply': 'لا يوجد عرض حالي. الصق نص المناقصة أولاً.', 'mode': 'await_draft'})

    try:
        from generator import patch_proposal
        from builder import build_pptx

        logger.info(f'Chatbot: patching — "{correction[:80]}"')
        updated_data = patch_proposal(session['proposal_data'], correction, session['api_key'])

        # Check structure before touching any files
        patch_warnings = _validate_proposal(updated_data)
        if patch_warnings:
            logger.warning('Patch warnings: ' + ' | '.join(patch_warnings))

        safe = _safe_filename(updated_data.get('program_title', 'Proposal'))
        new_pptx = COMPLETED_DIR / f'Zawad_{safe}.pptx'
        build_pptx(updated_data, str(new_pptx))   # build first — confirm it works
        logger.info(f'New PPTX: {new_pptx.name}')

        # Only delete old file after new one is confirmed written
        old_path = Path(session['pptx_path'])
        if old_path.exists() and old_path != new_pptx:
            old_path.unlink()
            logger.info(f'Deleted old: {old_path.name}')

        session['proposal_data'] = updated_data
        session['pptx_path'] = str(new_pptx)

        # Update review file to reflect the edited proposal
        review_name = _write_review_file(updated_data, session.get('req_data', {}), safe)
        review_line = f'\n📋 **قائمة المراجعة محدَّثة:** `review_needed/{review_name}`' if review_name else ''

        pricing = updated_data.get('sections', {}).get('pricing', {})
        total = pricing.get('grand_total_omr', '—')

        reply = f"""✅ **تم تطبيق التعديل!**

🗑️ الملف القديم حُذف تلقائياً
📁 **الملف الجديد:** `completed/{new_pptx.name}`
💰 **الإجمالي الجديد:** {total} ريال عماني{review_line}

هل تريد تعديلات أخرى؟"""

        return jsonify({'reply': reply, 'mode': 'ready'})

    except Exception as exc:
        logger.error(f'Patch error: {exc}', exc_info=True)
        return jsonify({
            'reply': f'❌ **حدث خطأ أثناء التعديل:**\n\n`{str(exc)[:300]}`',
            'mode': session['mode']
        })


# ── Entry point ──────────────────────────────────────────────────

def _open_browser():
    time.sleep(1.8)
    webbrowser.open('http://localhost:5050')


if __name__ == '__main__':
    COMPLETED_DIR.mkdir(parents=True, exist_ok=True)
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    logger.info('Zawad Proposal Chatbot starting at http://localhost:5050')
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host='localhost', port=5050, debug=False, use_reloader=False)
