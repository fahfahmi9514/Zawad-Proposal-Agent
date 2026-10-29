"""
reviewer.py — Adversarial Dual-Reviewer for Zawad Proposal Agent.

Two-pass review:
  Pass A (Attacker): pretends to be the client's rejection committee.
                     Aggressively finds every weakness, gap, and compliance miss.
  Pass B (Defender): reads the attacks, filters false alarms, keeps only real issues.
                     Also verifies every extracted requirement is covered.
                     Produces the final actionable report.

The attacker and defender use different prompts so they don't share blind spots.
"""

import json
import logging
import time
from google import genai
from google.genai import types

logger = logging.getLogger(__name__)
MODEL = 'gemini-3.8-flash'

# ── Attacker prompt ──────────────────────────────────────────────────────────

ATTACKER_SYSTEM = """أنت لجنة تقييم المناقصات في جهة حكومية عُمانية. مهمتك: رفض هذا العرض إن وجدت أي سبب لذلك.

ابحث بعدوانية عن:
1. أقسام غامضة أو فارغة أو مكررة
2. أي متطلب من المناقصة لم يُغطَّ في العرض
3. أخطاء في الأسعار أو الحسابات
4. لغة عربية ضعيفة أو عامية أو غير رسمية
5. أقسام قصيرة جداً (جملتان = فشل)
6. أي ادعاء غير موثق (مثل "خبرة 20 عاماً" بدون دليل)
7. عدم ذكر رؤية عُمان 2040 أو ذكرها بشكل سطحي
8. غياب خطة الطوارئ أو المخاطر الواقعية
9. أي تناقض بين الأرقام (عدد الأيام في stats vs workplan مثلاً)
10. غياب أي التزام قانوني (ضريبة / ضمان / صلاحية)

اكتب كل مشكلة في سطر منفصل يبدأ بـ [مشكلة]:
[مشكلة]: وصف المشكلة بدقة مع ذكر القسم المتأثر
[مشكلة]: ...

إذا كان العرض مثالياً في جانب ما، لا تخترع مشاكل. لكن كن صارماً ودقيقاً في كل ما تجد."""


# ── Defender prompt ───────────────────────────────────────────────────────────

DEFENDER_SYSTEM = """أنت مدير جودة خبير في شركة زواد. لديك تقرير هجوم من لجنة التقييم، وعليك:

1. لكل مشكلة مذكورة: قرر هل هي صحيحة أم مبالغ فيها أم خاطئة
2. تحقق من تغطية كل متطلب في قائمة المتطلبات المستخرجة
3. تحقق من صحة الأرقام المالية باستخدام الأرقام المحسوبة المعطاة لك
4. اكتب تقرير المراجعة النهائي بهذا التنسيق الدقيق:

## ملخص تنفيذي
[جملة واحدة: درجة العرض من 10 وهل هو جاهز للتقديم]

## ✅ ما هو ناجح
[بنود سليمة لا تحتاج تعديل]

## ⚠️ يحتاج تحسين
[مشاكل حقيقية لكن غير حرجة — مع اقتراح جملة بديلة أو رقم صحيح]

## ❌ مشاكل حرجة — يجب إصلاحها قبل التقديم
[فقط المشاكل الحرجة الحقيقية — لا تضع هنا ما هو مقبول]

## 📋 تغطية المتطلبات
[لكل متطلب مستخرج: ✅ مغطى / ❌ مفقود]

## 🎯 أولويات الإصلاح
1. [أول إصلاح — مع الرقم الصحيح أو الجملة البديلة]
2. [ثانياً]
3. [ثالثاً إن وجد]

كن دقيقاً وعملياً. لا تنتقد ما هو مقبول."""


# ── Math helper (identical to original) ─────────────────────────────────────

def _math_note(proposal_data: dict) -> str:
    pricing = proposal_data.get('sections', {}).get('pricing', {})
    subtotal = pricing.get('subtotal_omr', 0)
    vat = pricing.get('vat_omr', 0)
    grand_total = pricing.get('grand_total_omr', 0)
    items = pricing.get('line_items', [])
    c_sub = sum(i.get('total_omr', 0) for i in items)
    c_vat = round(c_sub * 0.05, 3)
    c_grand = round(c_sub + c_vat, 3)
    ok = abs(c_sub - subtotal) < 0.1 and abs(c_vat - vat) < 0.1
    return (
        f'[أرقام الحساب]\n'
        f'مجموع البنود المحسوب: {c_sub:,.3f} | VAT المحسوبة: {c_vat:,.3f} | الإجمالي المحسوب: {c_grand:,.3f}\n'
        f'ما في JSON — subtotal: {subtotal} | vat: {vat} | grand_total: {grand_total}\n'
        f'{"✅ الأرقام صحيحة" if ok else f"❌ خطأ في الحسابات — الصحيح: subtotal={c_sub:,.3f} | vat={c_vat:,.3f} | grand_total={c_grand:,.3f}"}'
    )


def _call_gemini(system: str, prompt: str, api_key: str,
                 max_tokens: int = 2048, temperature: float = 0.2) -> str:
    client = genai.Client(api_key=api_key, http_options={'timeout': 90000})
    cfg = types.GenerateContentConfig(
        system_instruction=system,
        temperature=temperature,
        max_output_tokens=max_tokens,
    )
    for attempt in range(4):
        try:
            if attempt > 0:
                wait = 30 * attempt
                logger.info(f'  Reviewer retry {attempt + 1} in {wait}s...')
                time.sleep(wait)
            r = client.models.generate_content(model=MODEL, contents=prompt, config=cfg)
            return r.text.strip()
        except Exception as exc:
            logger.warning(f'  Reviewer attempt {attempt + 1} failed: {exc}')
            if '503' not in str(exc) and 'UNAVAILABLE' not in str(exc):
                break
    return ''


# ── Public API ───────────────────────────────────────────────────────────────

def review_proposal(proposal_data: dict, draft_text: str,
                    api_key: str, requirements: dict = None) -> str:
    """
    Two-pass adversarial review.
    requirements: the dict returned by extractor.extract_requirements() — optional.
    Returns a full Arabic review report string.
    """
    proposal_str = json.dumps(proposal_data, ensure_ascii=False, indent=2)
    math = _math_note(proposal_data)

    # Format requirements checklist for injection
    req_section = ''
    if requirements:
        from extractor import format_checklist
        req_section = f'\n\nقائمة المتطلبات المستخرجة من المناقصة:\n{format_checklist(requirements)}'

    # ── Pass A: Attacker ──────────────────────────────────────────────────────
    logger.info('  Reviewer Pass A (Attacker)...')
    attack_prompt = f"""المسودة الأصلية (أول 2000 حرف):
---
{draft_text[:2000]}
---
{req_section}

محتوى العرض المُنشأ (JSON):
---
{proposal_str[:6000]}
---

{math}

ابحث عن كل مشكلة واكتبها."""

    attack_findings = _call_gemini(
        ATTACKER_SYSTEM, attack_prompt, api_key,
        max_tokens=1500, temperature=0.3
    )

    if not attack_findings:
        attack_findings = '[تعذّر تشغيل المراجع الأول]'

    logger.info(f'  Attacker found {attack_findings.count("[مشكلة]")} issues')

    # ── Pass B: Defender ──────────────────────────────────────────────────────
    logger.info('  Reviewer Pass B (Defender)...')
    defend_prompt = f"""تقرير المراجع الأول (محاولة رفض العرض):
---
{attack_findings}
---
{req_section}

محتوى العرض (JSON):
---
{proposal_str[:5000]}
---

{math}

قيّم كل مشكلة وأنتج التقرير النهائي."""

    final_report = _call_gemini(
        DEFENDER_SYSTEM, defend_prompt, api_key,
        max_tokens=2500, temperature=0.1
    )

    if not final_report:
        final_report = '[تعذّر توليد التقرير النهائي]'

    # ── Assemble full report ──────────────────────────────────────────────────
    import datetime
    header = (
        f'{"=" * 80}\n'
        f'تقرير المراجعة الثنائية — زواد للتطوير والاستشارات\n'
        f'{"=" * 80}\n'
        f'البرنامج : {proposal_data.get("program_title", "غير محدد")}\n'
        f'الجهة    : {proposal_data.get("client_name", "غير محدد")}\n'
        f'التاريخ  : {datetime.date.today().strftime("%Y-%m-%d")}\n'
        f'{"=" * 80}\n\n'
    )

    attacker_section = (
        f'\n\n{"─" * 40}\n'
        f'[ملف داخلي] نتائج المراجع الأول (الهجوم):\n'
        f'{"─" * 40}\n'
        f'{attack_findings}\n'
    )

    return header + final_report + attacker_section


def extract_critical_issues(review_report: str) -> list:
    """
    Pull out lines from the ❌ critical section for display in the chatbot UI.
    Returns a list of short strings (up to 3 issues).
    """
    issues = []
    in_critical = False
    for line in review_report.splitlines():
        if '❌' in line and 'مشاكل حرجة' in line:
            in_critical = True
            continue
        if in_critical:
            if line.startswith('##'):
                break
            stripped = line.strip('- •*').strip()
            if stripped and len(stripped) > 10:
                issues.append(stripped)
                if len(issues) >= 3:
                    break
    return issues
