"""
rag.py — RAG (Retrieval-Augmented Generation) for Zawad Proposal Agent.

How it works:
  1. Place past proposal files (PPTX, PDF, DOCX) in the past_proposals/ folder
  2. On first use, RAG reads each file, embeds it via Gemini, and saves rag_index.json
  3. New files are auto-indexed; changed files are re-indexed; deleted files are removed
  4. When generating a new proposal, the top 2 most similar past proposals are injected
     into the Gemini prompt as style/content examples

No extra dependencies — uses google-genai (already installed) and pure Python math.
"""

import json
import logging
import hashlib
from pathlib import Path

logger = logging.getLogger(__name__)

BASE_DIR   = Path(__file__).parent
PAST_DIR   = BASE_DIR / 'past_proposals'
INDEX_FILE = BASE_DIR / 'rag_index.json'

STORE_CHARS  = 3000   # chars stored per proposal
INJECT_CHARS = 1200   # chars injected per result into the prompt


def _tokenize(text: str) -> set:
    """Split text into unique word tokens, stripping punctuation."""
    import re
    tokens = re.findall(r'[\w؀-ۿ]+', text.lower())
    return set(t for t in tokens if len(t) > 2)


def _jaccard(a: set, b: set) -> float:
    """Jaccard similarity between two token sets."""
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _file_md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def build_index(api_key: str, force: bool = False) -> int:
    """
    Scan past_proposals/ and embed any new or changed files.
    Saves embeddings to rag_index.json.
    Returns count of newly indexed files (0 = already up to date).
    """
    PAST_DIR.mkdir(exist_ok=True)

    # Load existing index
    index = {}
    if INDEX_FILE.exists() and not force:
        try:
            with open(INDEX_FILE, 'r', encoding='utf-8') as f:
                index = json.load(f)
        except Exception:
            index = {}

    from reader import read_file

    supported = {'.pptx', '.ppt', '.pdf', '.docx', '.doc', '.txt'}
    files = [f for f in PAST_DIR.iterdir()
             if f.is_file() and f.suffix.lower() in supported]

    # Find files that need (re-)indexing
    to_index = []
    for f in files:
        fhash = _file_md5(f)
        if f.name in index and index[f.name].get('hash') == fhash:
            continue
        try:
            text = read_file(str(f))
            if not text or len(text.strip()) < 50:
                logger.warning(f'RAG: skipping empty file {f.name}')
                continue
            to_index.append((f.name, fhash, text))
            logger.info(f'RAG: indexed — {f.name}')
        except Exception as exc:
            logger.warning(f'RAG: could not read {f.name}: {exc}')

    # Remove stale entries for deleted files
    current_names = {f.name for f in files}
    for stale_key in [k for k in index if k not in current_names]:
        del index[stale_key]
        logger.info(f'RAG: removed stale entry — {stale_key}')

    if not to_index:
        return 0

    # Store text + token set (no embedding API needed)
    for name, fhash, text in to_index:
        index[name] = {
            'hash':   fhash,
            'text':   text[:STORE_CHARS],
            'tokens': list(_tokenize(text[:3000])),
        }

    _save_index(index)
    return len(to_index)


def _save_index(index: dict):
    with open(INDEX_FILE, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def retrieve(draft_text: str, top_k: int = 2) -> list:
    """
    Return top_k most similar past proposals using word-overlap (Jaccard similarity).
    No embedding API required.
    Each result: {'name': str, 'text': str, 'score': float}
    """
    if not INDEX_FILE.exists():
        return []

    try:
        with open(INDEX_FILE, 'r', encoding='utf-8') as f:
            index = json.load(f)
    except Exception:
        return []

    if not index:
        return []

    query_tokens = _tokenize(draft_text[:3000])
    scored = []
    for name, entry in index.items():
        stored_tokens = set(entry.get('tokens', []))
        if not stored_tokens:
            continue
        score = _jaccard(query_tokens, stored_tokens)
        scored.append({'name': name, 'text': entry['text'], 'score': score})

    scored.sort(key=lambda x: x['score'], reverse=True)
    return scored[:top_k]


def get_rag_context(draft_text: str, api_key: str) -> str:
    """
    Build a prompt-ready RAG context block from similar past proposals.
    Returns '' if past_proposals/ is empty or embedding fails.
    """
    # Auto-update index with any new/changed files
    try:
        n = build_index(api_key)
        if n:
            logger.info(f'RAG: indexed {n} new file(s)')
    except Exception as exc:
        logger.warning(f'RAG: build_index error: {exc}')

    results = retrieve(draft_text)
    if not results:
        return ''

    lines = ['=== أمثلة من عروض زواد السابقة المشابهة — استخدمها كمرجع للأسلوب والمحتوى ===']
    for i, r in enumerate(results, 1):
        score_pct = int(r['score'] * 100)
        lines.append(f'\n--- مثال {i}: {r["name"]} (تشابه {score_pct}%) ---')
        lines.append(r['text'][:INJECT_CHARS])

    logger.info(f'RAG: injecting {len(results)} example(s) into prompt')
    return '\n'.join(lines)
