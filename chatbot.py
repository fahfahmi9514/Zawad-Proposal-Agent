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

def _load_config() -> dict:
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _load_api_key() -> str:
    return _load_config().get('api_key', '')


def _save_api_key(key: str):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    cfg = _load_config()
    cfg['api_key'] = key
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=2)


# ── Session state (single user, in-memory) ───────────────────────
_api_key = _load_api_key()

session = {
    'api_key':       _api_key,
    'proposal_data': None,
    'draft_text':    None,
    'pptx_path':     None,
    'req_data':      {},   # requirements extracted from tender — kept for re-review after edits
    'staged_texts':  [],   # file texts waiting to be combined into next generation
    'pending_draft': None, # tender text held while waiting for user confirmation
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
        # Always go conversational when files are uploaded — analyze first, build only after confirm
        if mode == 'await_draft':
            session['pending_draft'] = user_msg
            session['mode'] = 'await_confirm'
            return _handle_with_question(user_msg, typed)

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
        # If the message is short and looks like a question (not a full tender), answer it
        QUESTION_SIGNALS = ['؟', '?', 'ما هي', 'ما هو', 'كيف', 'هل', 'what', 'how', 'can you',
                            'متطلبات', 'requirements', 'اشرح', 'explain', 'وضح', 'أخبرني']
        is_question = len(user_msg) < 300 and any(kw in user_msg for kw in QUESTION_SIGNALS)
        if is_question and not session.get('pending_draft'):
            return jsonify({
                'reply': '📎 من فضلك أرفق ملف المناقصة أولاً، ثم اكتب سؤالك وأرسل — سأجيبك وأنتظر موافقتك قبل البدء.',
                'mode': mode
            })
        return _generate_new(user_msg)

    # ── Conversational confirmation before generating ─────────────────────────
    if mode == 'await_confirm':
        YES_KWS = ['نعم', 'يلا', 'yes', 'go', 'ابدأ', 'تفضل', 'اكمل', 'أكمل',
                   'ok', 'okay', 'موافق', 'ماشي', 'sure', 'proceed', 'هيا', 'يس',
                   'ابني', 'انشئ', 'أنشئ', 'ابدأ', 'بالتوفيق', 'جهز', 'اعمل']
        NO_KWS  = ['لا', 'no', 'cancel', 'إلغاء', 'stop', 'الغ', 'ألغِ']
        msg_lower = user_msg.lower().strip()

        if any(kw in msg_lower for kw in YES_KWS):
            draft = session.get('pending_draft') or user_msg
            session['pending_draft'] = None
            session['mode'] = 'await_draft'
            return _generate_new(draft)

        if any(kw in msg_lower for kw in NO_KWS):
            session['pending_draft'] = None
            session['mode'] = 'await_draft'
            return jsonify({'reply': '🔄 تم الإلغاء. أرفق المناقصة من جديد عندما تكون جاهزاً.', 'mode': 'await_draft'})

        # User asked another question — answer it and ask again
        return _answer_about_tender(user_msg)

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


def _format_req_summary(req: dict) -> str:
    if not req:
        return '_تعذّر استخراج المتطلبات تلقائياً — سيتم التعامل مع النص مباشرةً._'
    lines = []
    if req.get('client_name'):      lines.append(f'🏢 **الجهة:** {req["client_name"]}')
    if req.get('tender_number'):    lines.append(f'📄 **رقم المناقصة:** {req["tender_number"]}')
    if req.get('program_type'):     lines.append(f'📚 **نوع العمل:** {req["program_type"]}')
    if req.get('participants'):     lines.append(f'👥 **المشاركون:** {req["participants"]}')
    if req.get('duration_days'):    lines.append(f'📅 **المدة:** {req["duration_days"]} يوم')
    budget = req.get('budget_ceiling_omr', 0)
    if budget and str(budget) not in ('0', ''):
        lines.append(f'💰 **سقف الميزانية:** {budget} ريال عماني')
    mandatory = req.get('mandatory_requirements', [])
    if mandatory:
        lines.append('\n✅ **المتطلبات الإلزامية:**')
        for r in mandatory[:6]: lines.append(f'   • {r}')
    compliance = req.get('compliance_items', [])
    if compliance:
        lines.append('\n⚖️ **الالتزامات القانونية:**')
        for c in compliance[:4]: lines.append(f'   • {c}')
    red_flags = req.get('red_flags', [])
    if red_flags:
        lines.append('\n🚩 **نقاط حرجة — لا تُهمل:**')
        for f in red_flags[:3]: lines.append(f'   • {f}')
    return '\n'.join(lines) if lines else '_لا متطلبات محددة._'


def _generate_targeted_questions(tender_text: str) -> str:
    """Ask Gemini to generate 3-5 targeted questions the user must answer before building."""
    try:
        from google import genai
        from google.genai import types as _types
        client = genai.Client(api_key=session['api_key'])
        prompt = (
            'أنت مساعد تحليل مناقصات حكومية عُمانية.\n'
            'اقرأ هذه المناقصة وحدّد 3 إلى 5 أسئلة يجب على صاحب العرض الإجابة عليها '
            'قبل كتابة العرض، بناءً على متطلبات خاصة بهذه المناقصة تحديداً.\n'
            'ركّز على معلومات لا يمكن للنظام افتراضها أو اختراعها مثل:\n'
            '- أسماء المدربين/المدربات المتاحين (خاصة إن اشترطت المناقصة إناثاً)\n'
            '- الشهادات والمؤهلات المحددة المطلوبة\n'
            '- أي اشتراطات خاصة بالجنس أو الخبرة أو الجنسية\n'
            '- تفاصيل المتحدث في جلسة الحوار إن وُجدت\n'
            '- أي قيود أو شروط خاصة بهذه الجهة\n\n'
            'اكتب الأسئلة بصيغة مباشرة وواضحة باللغة العربية، واحدة في كل سطر.\n\n'
            f'=== المناقصة ===\n{tender_text[:6000]}'
        )
        resp = client.models.generate_content(
            model='gemini-3.5-flash-lite',
            contents=prompt,
            config=_types.GenerateContentConfig(temperature=0.3, max_output_tokens=500)
        )
        questions = resp.text.strip()
        return f'❓ **أسئلة يجب الإجابة عليها قبل بناء العرض:**\n\n{questions}'
    except Exception as e:
        logger.warning(f'Targeted questions failed: {e}')
        return '❓ **يرجى تزويدنا بأي معلومات إضافية عن فريق التنفيذ أو المتطلبات الخاصة قبل البدء.**'


def _answer_about_tender(question: str):
    """Answer a question about the pending tender, then ask for confirmation."""
    tender_text = session.get('pending_draft', '')
    try:
        from google import genai
        client = genai.Client(api_key=session['api_key'])
        prompt = (
            'أنت مساعد تحليل مناقصات. أجب على سؤال المستخدم بشكل مباشر ومختصر '
            'بناءً على محتوى الوثيقة فقط. لا تخترع معلومات غير موجودة.\n\n'
            f'=== الوثيقة ===\n{tender_text[:8000]}\n\n'
            f'=== سؤال المستخدم ===\n{question}'
        )
        from google.genai import types as _types
        resp = client.models.generate_content(
            model='gemini-3.5-flash-lite',
            contents=prompt,
            config=_types.GenerateContentConfig(temperature=0.2, max_output_tokens=800)
        )
        answer = resp.text.strip()
    except Exception as e:
        logger.warning(f'Tender Q&A failed: {e}')
        answer = '📋 تم استلام سؤالك. يمكنني الإجابة بشكل أفضل بعد إنشاء العرض.'

    reply = (
        f'{answer}\n\n'
        '---\n'
        '💬 هل أبدأ الآن بإنشاء العرض؟ اكتب **نعم** أو **يلا** للمتابعة، '
        'أو اطرح أي سؤال آخر.'
    )
    return jsonify({'reply': reply, 'mode': 'await_confirm'})


def _handle_with_question(full_content: str, typed_question: str):
    """Extract requirements, generate targeted questions, optionally answer user's question."""
    from extractor import extract_requirements
    req = {}
    try:
        req = extract_requirements(full_content, session['api_key'])
        session['req_data'] = req
    except Exception as e:
        logger.warning(f'Pre-extract failed: {e}')

    summary = _format_req_summary(req)
    targeted_qs = _generate_targeted_questions(full_content)

    answer_block = ''
    if typed_question:
        try:
            from google import genai
            from google.genai import types as _types
            client = genai.Client(api_key=session['api_key'])
            prompt = (
                'أنت مساعد تحليل مناقصات. أجب على سؤال المستخدم بشكل مباشر ومختصر '
                'بناءً على محتوى الوثيقة فقط. لا تخترع معلومات.\n\n'
                f'=== الوثيقة ===\n{full_content[:8000]}\n\n'
                f'=== سؤال المستخدم ===\n{typed_question}'
            )
            resp = client.models.generate_content(
                model='gemini-3.5-flash-lite',
                contents=prompt,
                config=_types.GenerateContentConfig(temperature=0.2, max_output_tokens=800)
            )
            answer_block = resp.text.strip() + '\n\n---\n\n'
        except Exception as e:
            logger.warning(f'Q&A failed: {e}')

    reply = (
        f'{answer_block}'
        f'📋 **ملخص متطلبات المناقصة:**\n{summary}\n\n'
        f'---\n\n'
        f'{targeted_qs}\n\n'
        f'---\n'
        f'💬 أجب على الأسئلة أعلاه ثم اكتب **نعم** أو **يلا** لبدء الإنشاء، '
        f'أو اطرح أي سؤال آخر عن المناقصة.'
    )
    return jsonify({'reply': reply, 'mode': 'await_confirm'})


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

        # ── Step 3b: Upload to shared Drive (optional) ───────────────────────
        from drive_upload import upload_proposal as _drive_upload
        drive_link = _drive_upload(str(pptx_path), _load_config())
        drive_line = f'\n☁️ **Google Drive:** [افتح الملف]({drive_link})' if drive_link else ''

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
📁 **ملف العرض:** `completed/{pptx_path.name}`
📂 **[افتح في PowerPoint](http://localhost:5050/api/open_pptx)** ← اضغط لفتح الملف مباشرة{drive_line}{warnings_block}

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

        # Upload updated file to Drive
        from drive_upload import upload_proposal as _drive_upload
        drive_link = _drive_upload(str(new_pptx), _load_config())
        drive_line = f'\n☁️ **Google Drive:** [افتح الملف]({drive_link})' if drive_link else ''

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
💰 **الإجمالي الجديد:** {total} ريال عماني{review_line}{drive_line}
📂 **[افتح في PowerPoint](http://localhost:5050/api/open_pptx)** ← اضغط لفتح الملف مباشرة

هل تريد تعديلات أخرى؟"""

        return jsonify({'reply': reply, 'mode': 'ready'})

    except Exception as exc:
        logger.error(f'Patch error: {exc}', exc_info=True)
        return jsonify({
            'reply': f'❌ **حدث خطأ أثناء التعديل:**\n\n`{str(exc)[:300]}`',
            'mode': session['mode']
        })


# ── Slide Preview ───────────────────────────────────────────────

def _build_preview_html(data: dict) -> str:
    """Build a complete HTML preview page from proposal data."""
    import html as _html
    sec = data.get('sections', {})
    pricing = sec.get('pricing', {})
    lang = data.get('language', 'ar')

    def esc(s): return _html.escape(str(s)) if s else ''
    def trunc(s, n=350): return esc(str(s)[:n]) + ('…' if len(str(s)) > n else '')

    slides = []
    num = [0]
    def next_num():
        num[0] += 1
        return str(num[0]).zfill(2)

    def card(title, body_html, bg='#fff'):
        n = next_num()
        return f'''<div class="slide-card" id="slide-{n}">
  <div class="slide-header">
    <span class="slide-num">{n}</span>
    <span class="slide-title">{esc(title)}</span>
    <button class="edit-btn" onclick="openEdit('{n}','{esc(title)}')">✏️ تعديل</button>
  </div>
  <div class="slide-body">{body_html}</div>
</div>'''

    # 01 Cover
    slides.append(card('الغلاف', f'''
      <div class="kv"><b>البرنامج:</b> {esc(data.get('program_title',''))}</div>
      <div class="kv"><b>الجهة:</b> {esc(data.get('client_name',''))}</div>
      <div class="kv"><b>النوع:</b> {esc(data.get('program_category',''))}</div>
      <div class="kv"><b>رقم المناقصة:</b> {esc(data.get('tender_number','—'))}</div>'''))

    # 02 Cover Letter
    slides.append(card('خطاب التقديم', f'<p>{trunc(sec.get("cover_letter",""),400)}</p>'))

    # 03 Executive Summary / Stats
    stats_html = '<div class="stat-grid">'
    for s in data.get('key_stats', []):
        stats_html += f'<div class="stat-tile"><div class="stat-num">{esc(s.get("number",""))}</div><div class="stat-lbl">{esc(s.get("label",""))}</div></div>'
    stats_html += '</div>'
    stats_html += f'<p style="margin-top:10px">{trunc(sec.get("key_message",""),250)}</p>'
    slides.append(card('الملخص التنفيذي', stats_html))

    # 04 Commitments grid
    cl = sec.get('commitments_list', [])
    clist_html = '<ul>' + ''.join(f'<li>{esc(c)}</li>' for c in cl) + '</ul>' if cl else f'<p>{trunc(sec.get("commitments",""))}</p>'
    slides.append(card('الالتزامات الرئيسية', clist_html))

    # 05 About Zawad
    slides.append(card('نبذة عن زواد', f'<p>{trunc(sec.get("company_overview",""),300)}</p>'))

    # 06 Requirements Understanding
    slides.append(card('فهمنا لنطاق العمل', f'<p>{trunc(sec.get("requirements_understanding",""),400)}</p>'))

    # 07 Expected Outputs
    outputs = sec.get('expected_outputs', [])
    out_html = '<div class="output-grid">'
    for o in outputs:
        out_html += f'<div class="output-card"><span class="out-num">{esc(o.get("number",""))}</span><b>{esc(o.get("title",""))}</b><p>{esc(o.get("description",""))}</p></div>'
    out_html += '</div>'
    slides.append(card('المخرجات المتوقعة', out_html))

    # 08 Methodology
    slides.append(card('المنهجية المقترحة', f'<p>{trunc(sec.get("methodology",""),300)}</p>'))

    # 09+ Workplan days
    for d in sec.get('workplan', []):
        acts = d.get('activities', [])
        body = f'<b>{esc(d.get("title",""))}</b><ul>' + ''.join(f'<li>{esc(a)}</li>' for a in acts) + '</ul>'
        slides.append(card(f'خطة العمل — {esc(d.get("title","يوم"))}', body))

    # Gantt
    slides.append(card('الإطار الزمني التفصيلي', '<p>الإطار الزمني التفصيلي — انظر الملف للجدول الكامل</p>'))

    # Team slides
    team = sec.get('team', [])
    for m in team:
        name_key = 'name_en' if lang == 'en' else 'name'
        role_key = 'role_en' if lang == 'en' else 'role'
        exp_key = 'experience_en' if lang == 'en' else 'experience'
        quals = m.get('qualifications_en' if lang == 'en' else 'qualifications', [])
        body = f'''<b style="font-size:16px">{esc(m.get(name_key) or m.get('name',''))}</b>
          <div class="role-badge">{esc(m.get(role_key) or m.get('role',''))}</div>
          <p style="margin-top:8px">{trunc(m.get(exp_key) or m.get('experience',''),200)}</p>
          {'<ul>' + ''.join(f"<li>{esc(q)}</li>" for q in quals[:4]) + '</ul>' if quals else ''}'''
        slides.append(card(f'فريق العمل — {esc(m.get(name_key) or m.get("name",""))}', body))

    # Learning Measurement (08 in TOC)
    lm = sec.get('learning_measurement', {})
    if lm:
        tools = lm.get('tools', [])
        lm_body = f'''<div class="kv"><b>التقييم القبلي:</b> {trunc(lm.get('pre_assessment',''),150)}</div>
          <div class="kv"><b>التقييم البعدي:</b> {trunc(lm.get('post_assessment',''),150)}</div>
          <div class="kv"><b>متابعة الأثر:</b> {trunc(lm.get('impact_followup',''),150)}</div>
          {'<ul>' + ''.join(f"<li>{esc(t)}</li>" for t in tools) + '</ul>' if tools else ''}'''
        slides.append(card('أدوات قياس التعلم', lm_body))
    else:
        slides.append(card('أدوات قياس التعلم', '<p>التقييم القبلي والبعدي — كيركباتريك المستوى 1–4 — متابعة 60-90 يوماً</p>'))

    # Governance
    slides.append(card('الحوكمة وضمان الجودة', '<p>الحوكمة وضمان الجودة — كيركباتريك المستوى 1–4 — مراجعة داخلية — اعتماد الجهة قبل كل مرحلة</p>'))

    # Risk Matrix
    risks = sec.get('risks', [])
    if risks:
        risk_html = '<table class="data-table"><thead><tr><th>المخاطرة</th><th>الاحتمالية</th><th>الأثر</th><th>الإجراء الوقائي</th><th>خطة الطوارئ</th></tr></thead><tbody>'
        for r in risks:
            risk_html += f'<tr><td>{esc(r.get("risk",""))}</td><td>{esc(r.get("probability",""))}</td><td>{esc(r.get("impact",""))}</td><td>{esc(r.get("preventive",""))}</td><td>{esc(r.get("contingency",""))}</td></tr>'
        risk_html += '</tbody></table>'
        slides.append(card('إدارة المخاطر', risk_html))

    # Innovation
    pts = sec.get('innovation_points', [])
    if pts:
        inn_html = '<div class="inno-grid">'
        for p in pts:
            inn_html += f'<div class="inno-card"><div class="inno-icon">{esc(p.get("icon","⭐"))}</div><b>{esc(p.get("title",""))}</b><p>{esc(p.get("description",""))}</p></div>'
        inn_html += '</div>'
        slides.append(card('الابتكار والقيمة المضافة', inn_html))

    # Track Record
    slides.append(card('المشاريع المكتملة', '<p>12 مشروعاً مكتملاً — انظر الملف للجدول الكامل</p>'))

    # Financial Summary
    subtotal = pricing.get('subtotal_omr', 0)
    vat = pricing.get('vat_omr', 0)
    total = pricing.get('grand_total_omr', 0)
    vat_pct = int(round(float(pricing.get('vat_rate', 0.05)) * 100))
    validity = data.get('stats', {}).get('validity', '90 يوماً')
    fin_body = f'''<div class="fin-grid">
      <div class="fin-box mint"><div class="fin-val">{int(round(float(total))):,}</div><div class="fin-lbl">الإجمالي الكلي (ر.ع)</div></div>
      <div class="fin-box teal"><div class="fin-val">{int(round(float(vat))):,}</div><div class="fin-lbl">ضريبة {vat_pct}% (ر.ع)</div></div>
      <div class="fin-box navy"><div class="fin-val">{int(round(float(subtotal))):,}</div><div class="fin-lbl">قبل الضريبة (ر.ع)</div></div>
    </div>
    <p style="margin-top:10px;color:#0E7C86"><b>صلاحية العرض:</b> {esc(validity)}</p>'''
    slides.append(card('العرض المالي — ملخص', fin_body))

    # Financial Table
    items = pricing.get('line_items', [])
    if items:
        ft_html = '<table class="data-table"><thead><tr><th>البند</th><th>الكمية</th><th>الوحدة</th><th>السعر</th><th>الإجمالي</th></tr></thead><tbody>'
        for it in items:
            try: rate = f"{int(round(float(it.get('rate_omr',0)))):,}"
            except: rate = str(it.get('rate_omr',''))
            try: tot = f"{int(round(float(it.get('total_omr',0)))):,}"
            except: tot = str(it.get('total_omr',''))
            ft_html += f'<tr><td>{esc(it.get("description",""))}</td><td>{esc(str(it.get("quantity","")))}</td><td>{esc(str(it.get("unit","")))}</td><td>{rate}</td><td>{tot}</td></tr>'
        ft_html += f'</tbody></table><p style="text-align:left;margin-top:8px;font-weight:bold">الإجمالي الكلي: {int(round(float(total))):,} ر.ع</p>'
        slides.append(card('العرض المالي — تفصيلي', ft_html))

    # Commitments page
    slides.append(card('الإقرارات والالتزامات', clist_html))

    # Declarations
    decl_html = '<ul><li>إقرار الاطلاع على شروط المناقصة وقبولها</li><li>السجل التجاري ساري المفعول</li><li>شهادة التسجيل الضريبي</li><li>السيرة الذاتية لفريق العمل</li><li>إقرار عدم التعارض</li></ul>'
    slides.append(card('الوثائق والإقرارات', decl_html))

    # Closing
    slides.append(card('الختام', f'<h2 style="color:#0A2647">شكراً لثقتكم</h2><p>{esc(data.get("client_name",""))} · info@zawad.om · +968 99232775 · www.zawad.om</p>'))

    # Build sidebar entries
    sidebar_items = ''
    for i in range(1, num[0] + 1):
        n = str(i).zfill(2)
        sidebar_items += f'<a class="sidebar-item" href="#slide-{n}" onclick="highlight(\'{n}\')">{n}</a>'

    cards_html = '\n'.join(slides)

    return f'''<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="60">
<title>معاينة الشرائح — زواد</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Segoe UI",Arial,sans-serif;background:#F0F4F8;color:#1B2A3A;display:flex;flex-direction:column;height:100vh;overflow:hidden}}
.top-bar{{background:#0A2647;color:#fff;padding:12px 20px;display:flex;align-items:center;gap:14px;flex-shrink:0}}
.top-bar h1{{font-size:16px;color:#2EC4B6;flex:1}}
.top-bar button{{padding:6px 14px;background:transparent;border:1.5px solid #2EC4B6;color:#2EC4B6;border-radius:8px;cursor:pointer;font-size:13px}}
.top-bar button:hover{{background:#2EC4B6;color:#0A2647}}
.layout{{display:flex;flex:1;overflow:hidden}}
.sidebar{{width:70px;background:#0A2647;overflow-y:auto;padding:10px 0;flex-shrink:0}}
.sidebar-item{{display:block;text-align:center;padding:8px 4px;color:rgba(255,255,255,0.65);font-size:12px;font-weight:700;cursor:pointer;text-decoration:none;border-right:3px solid transparent;transition:.15s}}
.sidebar-item:hover{{color:#2EC4B6;border-right-color:#2EC4B6}}
.sidebar-item.active{{color:#2EC4B6;border-right-color:#2EC4B6;background:rgba(46,196,182,0.1)}}
.main{{flex:1;overflow-y:auto;padding:20px}}
.slide-card{{background:#fff;border-radius:12px;margin-bottom:18px;box-shadow:0 2px 8px rgba(0,0,0,0.08);overflow:hidden;scroll-margin-top:20px}}
.slide-header{{background:#0A2647;padding:12px 16px;display:flex;align-items:center;gap:12px}}
.slide-num{{width:32px;height:32px;border-radius:50%;background:#2EC4B6;color:#0A2647;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:13px;flex-shrink:0}}
.slide-title{{color:#fff;font-size:14px;font-weight:600;flex:1}}
.edit-btn{{padding:4px 12px;background:transparent;border:1.5px solid #2EC4B6;color:#2EC4B6;border-radius:6px;cursor:pointer;font-size:12px}}
.edit-btn:hover{{background:#2EC4B6;color:#0A2647}}
.slide-body{{padding:16px;font-size:13px;line-height:1.7}}
.kv{{margin-bottom:6px}}
.stat-grid{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px}}
.stat-tile{{background:#0A2647;border-radius:8px;padding:12px 16px;min-width:80px;text-align:center}}
.stat-num{{color:#2EC4B6;font-size:22px;font-weight:900}}
.stat-lbl{{color:#fff;font-size:11px;margin-top:4px}}
.role-badge{{display:inline-block;background:#2EC4B6;color:#0A2647;border-radius:6px;padding:3px 10px;font-size:12px;font-weight:700;margin:6px 0}}
.output-grid{{display:flex;flex-wrap:wrap;gap:8px}}
.output-card{{background:#F0F4F8;border-radius:8px;padding:10px;min-width:160px;flex:1}}
.out-num{{display:inline-block;background:#0A2647;color:#2EC4B6;border-radius:50%;width:24px;height:24px;text-align:center;line-height:24px;font-size:11px;font-weight:700;margin-bottom:6px}}
.fin-grid{{display:flex;gap:12px;flex-wrap:wrap}}
.fin-box{{border-radius:10px;padding:14px 18px;flex:1;min-width:120px;text-align:center}}
.fin-box.mint{{background:#2EC4B6}}.fin-box.teal{{background:#0E7C86}}.fin-box.navy{{background:#0A2647}}
.fin-val{{color:#fff;font-size:22px;font-weight:900}}.fin-lbl{{color:rgba(255,255,255,0.8);font-size:12px;margin-top:4px}}
.inno-grid{{display:flex;flex-wrap:wrap;gap:8px}}
.inno-card{{background:#F0F4F8;border-radius:8px;padding:10px;flex:1;min-width:140px;text-align:center}}
.inno-icon{{font-size:22px;margin-bottom:6px}}
.data-table{{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}}
.data-table th{{background:#0A2647;color:#fff;padding:7px 10px;text-align:right}}
.data-table td{{border:1px solid #D0DCE8;padding:6px 10px;text-align:right}}
.data-table tr:nth-child(even) td{{background:#F0F4F8}}
ul{{padding-right:18px;margin:6px 0}}li{{margin-bottom:4px}}
.modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:999;align-items:center;justify-content:center}}
.modal-overlay.open{{display:flex}}
.modal{{background:#fff;border-radius:12px;padding:24px;width:480px;max-width:95vw;direction:rtl}}
.modal h3{{color:#0A2647;margin-bottom:12px}}
.modal textarea{{width:100%;border:1.5px solid #D0DCE8;border-radius:8px;padding:10px;font-size:14px;font-family:inherit;resize:vertical;min-height:80px}}
.modal-actions{{display:flex;gap:10px;margin-top:12px;justify-content:flex-end}}
.btn-copy{{padding:8px 18px;background:#2EC4B6;border:none;border-radius:8px;color:#0A2647;font-weight:700;cursor:pointer}}
.btn-close{{padding:8px 18px;background:transparent;border:1.5px solid #D0DCE8;border-radius:8px;cursor:pointer}}
</style>
</head>
<body>
<div class="top-bar">
  <h1>🖥️ معاينة الشرائح — زواد للتطوير والاستشارات</h1>
  <span style="color:rgba(255,255,255,0.6);font-size:12px">{esc(data.get('program_title',''))} · {esc(data.get('client_name',''))}</span>
  <button onclick="location.reload()">↺ تحديث</button>
</div>
<div class="layout">
  <div class="sidebar">{sidebar_items}</div>
  <div class="main">{cards_html}</div>
</div>
<div class="modal-overlay" id="edit-modal">
  <div class="modal">
    <h3>✏️ تعديل الشريحة</h3>
    <p style="color:#6B7D8E;font-size:12px;margin-bottom:8px">انسخ النص واكتبه في نافذة المحادثة</p>
    <textarea id="edit-text" rows="3"></textarea>
    <div class="modal-actions">
      <button class="btn-close" onclick="closeEdit()">إغلاق</button>
      <button class="btn-copy" onclick="copyEdit()">📋 نسخ التعديل</button>
    </div>
  </div>
</div>
<script>
function openEdit(num, title){{
  document.getElementById('edit-text').value = 'عدّل الشريحة ' + num + ' (' + title + '): ';
  document.getElementById('edit-modal').classList.add('open');
  setTimeout(()=>document.getElementById('edit-text').focus(),50);
}}
function closeEdit(){{ document.getElementById('edit-modal').classList.remove('open'); }}
function copyEdit(){{
  const t = document.getElementById('edit-text');
  if(navigator.clipboard){{ navigator.clipboard.writeText(t.value); }}
  else{{ t.select(); document.execCommand('copy'); }}
  closeEdit();
  alert('تم النسخ! الصقه في نافذة المحادثة.');
}}
function highlight(n){{
  document.querySelectorAll('.sidebar-item').forEach(a=>a.classList.remove('active'));
  const target = document.querySelector('.sidebar-item[href="#slide-'+n+'"]');
  if(target) target.classList.add('active');
}}
document.getElementById('edit-modal').addEventListener('click', function(e){{
  if(e.target===this) closeEdit();
}});
</script>
</body>
</html>'''


@app.route('/preview')
def preview():
    data = session.get('proposal_data')
    if not data:
        return '<html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0A2647;color:#2EC4B6;text-align:center"><h2>لا يوجد عرض حالياً</h2><p style="color:white;margin-top:12px">أنشئ عرضاً أولاً من نافذة المحادثة</p></body></html>'
    return _build_preview_html(data)


@app.route('/api/open_pptx')
def open_pptx():
    """Open the current PPTX file natively in PowerPoint (Windows only)."""
    pptx_path = session.get('pptx_path')
    if not pptx_path or not Path(pptx_path).exists():
        return jsonify({'ok': False, 'error': 'لا يوجد ملف عرض حالياً.'}), 404
    try:
        import os
        os.startfile(pptx_path)
        return jsonify({'ok': True, 'path': pptx_path})
    except Exception as exc:
        logger.error(f'open_pptx error: {exc}')
        return jsonify({'ok': False, 'error': str(exc)}), 500


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
