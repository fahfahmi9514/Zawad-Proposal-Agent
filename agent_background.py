"""
agent_background.py — Headless Zawad Proposal Agent.
No windows, no taskbar. Runs silently in the background.
Auto-started by Task Scheduler at every Windows login.
"""

import os
import sys
import json
import time
import threading
from pathlib import Path
import logging

# Log file in the agent folder
log_path = Path(__file__).parent / 'agent.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s — %(levelname)s — %(message)s',
    handlers=[
        logging.FileHandler(str(log_path), encoding='utf-8'),
    ]
)
logger = logging.getLogger(__name__)

# ── Paths ───────────────────────────────────────────────────────
DOWNLOADS     = Path.home() / 'Downloads'
WATCH_FOLDER  = DOWNLOADS / 'Proposals_2026'
COMPLETED_DIR = WATCH_FOLDER / 'completed'
REVIEW_DIR    = WATCH_FOLDER / 'review_needed'
CONFIG_FILE   = Path(__file__).parent / 'config' / 'config.json'

SUPPORTED_EXT = {'.pdf', '.docx', '.doc', '.pptx', '.ppt', '.txt'}


def _ensure_folders():
    for folder in [WATCH_FOLDER, COMPLETED_DIR, REVIEW_DIR]:
        folder.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)


def _load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_config(data: dict):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def _ask_api_key_with_window() -> str | None:
    """Show a tkinter window to get the API key (first time only)."""
    import tkinter as tk

    result = {'key': None}

    root = tk.Tk()
    root.title('Zawad Proposal Agent — First Time Setup')
    root.geometry('500x360')
    root.resizable(False, False)
    root.configure(bg='#0A2647')

    hdr = tk.Frame(root, bg='#0A2647')
    hdr.pack(fill='x', pady=(20, 10))
    tk.Label(hdr, text='زواد للتطوير والاستشارات',
             font=('Arial', 15, 'bold'), fg='#2EC4B6', bg='#0A2647').pack()
    tk.Label(hdr, text='One-time Setup — Enter your Gemini API Key',
             font=('Arial', 10), fg='#FFFFFF', bg='#0A2647').pack(pady=2)

    card = tk.Frame(root, bg='#E7F8F6', padx=24, pady=20)
    card.pack(fill='both', expand=True, padx=16, pady=(0, 16))

    tk.Label(card, text='Get your FREE key from:',
             font=('Arial', 10, 'bold'), fg='#0A2647', bg='#E7F8F6').pack(anchor='w')
    tk.Label(card,
             text='https://aistudio.google.com → Sign in → Get API Key\n(You will never be asked again after this)',
             font=('Arial', 9), fg='#1B2A3A', bg='#E7F8F6', justify='left').pack(anchor='w', pady=(4, 10))

    key_var = tk.StringVar()
    key_entry = tk.Entry(card, textvariable=key_var, font=('Arial', 10),
                         show='*', width=55, relief='solid', bd=1)
    key_entry.pack(fill='x', ipady=6)
    key_entry.focus()

    show_var = tk.BooleanVar()
    def _toggle():
        key_entry.config(show='' if show_var.get() else '*')
    tk.Checkbutton(card, text='Show key', variable=show_var, command=_toggle,
                   bg='#E7F8F6', fg='#1B2A3A', font=('Arial', 9)).pack(anchor='w', pady=4)

    err = tk.Label(card, text='', fg='#CC0000', bg='#E7F8F6', font=('Arial', 9))
    err.pack()

    def _save():
        key = key_var.get().strip()
        if not key or len(key) < 15:
            err.config(text='Please enter a valid API key.')
            return
        result['key'] = key
        _save_config({'api_key': key})
        root.destroy()

    root.bind('<Return>', lambda e: _save())
    tk.Button(card, text='Save & Start (runs silently from now on)',
              command=_save, bg='#0A2647', fg='#FFFFFF',
              font=('Arial', 10, 'bold'), relief='flat',
              padx=20, pady=8, cursor='hand2').pack(pady=(10, 0))

    root.mainloop()
    return result['key']


def _get_api_key() -> str:
    config = _load_config()
    if config.get('api_key'):
        return config['api_key']

    logger.info('No API key found — showing setup window.')
    key = _ask_api_key_with_window()
    if not key:
        logger.error('No API key provided. Exiting.')
        sys.exit(1)
    logger.info('API key saved.')
    return key


def _process_file(file_path: Path, api_key: str):
    logger.info(f'>> Processing: {file_path.name}')
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from reader import read_file
        from generator import generate_proposal
        from reviewer import review_proposal
        from builder import build_pptx

        text = read_file(file_path)
        if not text or len(text.strip()) < 30:
            logger.warning(f'  File too short: {file_path.name}')
            return

        logger.info(f'  Read {len(text):,} chars. Calling Gemini...')
        proposal_data = generate_proposal(text, api_key)
        logger.info(f'  Generated: {proposal_data.get("program_title", "?")}')

        stem = file_path.stem
        out_path = COMPLETED_DIR / f'Zawad_Proposal_{stem}.pptx'
        build_pptx(proposal_data, str(out_path))
        logger.info(f'  OK PPTX: {out_path.name}')

        try:
            report = review_proposal(proposal_data, text, api_key)
            review_path = REVIEW_DIR / f'Review_{stem}.txt'
            with open(review_path, 'w', encoding='utf-8') as f:
                f.write(report)
            logger.info(f'  OK Review: {review_path.name}')
        except Exception as rev_exc:
            logger.warning(f'  Reviewer skipped (Gemini busy): {rev_exc}')
        logger.info(f'DONE: {file_path.name}')

    except Exception as exc:
        logger.error(f'ERROR: {file_path.name} - {exc}', exc_info=True)


class _Handler:
    def __init__(self, api_key: str):
        self._api_key = api_key
        self._active = set()

    def dispatch(self, event):
        from watchdog.events import FileCreatedEvent, FileMovedEvent
        if isinstance(event, FileCreatedEvent):
            self._on_file(event.src_path)
        elif isinstance(event, FileMovedEvent):
            self._on_file(event.dest_path)

    def _on_file(self, path_str: str):
        fp = Path(path_str)
        try:
            rel = fp.relative_to(WATCH_FOLDER)
            if len(rel.parts) > 1:
                return
        except ValueError:
            return

        if fp.suffix.lower() not in SUPPORTED_EXT:
            return
        if str(fp) in self._active:
            return
        self._active.add(str(fp))

        def _run():
            time.sleep(3)
            if fp.exists():
                _process_file(fp, self._api_key)
            self._active.discard(str(fp))

        threading.Thread(target=_run, daemon=True).start()


def _scan_existing(api_key: str):
    """Process any files already sitting in the watch folder at startup."""
    for fp in WATCH_FOLDER.iterdir():
        if fp.is_file() and fp.suffix.lower() in SUPPORTED_EXT:
            logger.info(f'Startup scan — found existing file: {fp.name}')
            threading.Thread(target=_process_file, args=(fp, api_key), daemon=True).start()
            time.sleep(2)


def main():
    _ensure_folders()
    api_key = _get_api_key()
    logger.info('Zawad Agent started — watching folder.')

    _scan_existing(api_key)

    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    handler = _Handler(api_key)

    class _Bridge(FileSystemEventHandler):
        def on_created(self, event):
            handler.dispatch(event)
        def on_moved(self, event):
            handler.dispatch(event)

    observer = Observer()
    observer.schedule(_Bridge(), str(WATCH_FOLDER), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(5)
    except (KeyboardInterrupt, SystemExit):
        observer.stop()
    observer.join()
    logger.info('Agent stopped.')


if __name__ == '__main__':
    main()
