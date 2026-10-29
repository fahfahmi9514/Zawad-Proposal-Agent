"""
Zawad Proposal Agent — Main entry point.
Watches Downloads/Proposals_2026/ and auto-generates proposals when a draft is dropped.
"""

import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import logging

# Log to both file and console
log_path = Path(__file__).parent / 'agent.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s — %(levelname)s — %(message)s',
    handlers=[
        logging.FileHandler(str(log_path), encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
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


# ── Config ──────────────────────────────────────────────────────

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


# ── Setup window (first run) ────────────────────────────────────

def _show_setup_window() -> str | None:
    result = {'key': None}

    root = tk.Tk()
    root.title('Zawad Proposal Agent — First Time Setup')
    root.geometry('500x360')
    root.resizable(False, False)
    root.configure(bg='#0A2647')

    # Header
    hdr = tk.Frame(root, bg='#0A2647')
    hdr.pack(fill='x', pady=(20, 10))
    tk.Label(hdr, text='زواد للتطوير والاستشارات',
             font=('Arial', 15, 'bold'), fg='#2EC4B6', bg='#0A2647').pack()
    tk.Label(hdr, text='Zawad Proposal Agent — Setup',
             font=('Arial', 10), fg='#FFFFFF', bg='#0A2647').pack(pady=2)

    # Content card
    card = tk.Frame(root, bg='#E7F8F6', padx=24, pady=20)
    card.pack(fill='both', expand=True, padx=16, pady=(0, 16))

    tk.Label(card, text='Enter your Gemini API Key',
             font=('Arial', 12, 'bold'), fg='#0A2647', bg='#E7F8F6').pack(anchor='w')
    tk.Label(card,
             text='Get your FREE key from:\nhttps://aistudio.google.com → Sign in → Get API Key',
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
        if not key:
            err.config(text='Please enter your API key.')
            return
        if len(key) < 15:
            err.config(text="That doesn't look like a valid API key (too short).")
            return
        result['key'] = key
        _save_config({'api_key': key})
        root.destroy()

    root.bind('<Return>', lambda e: _save())

    btn = tk.Button(card, text='Save & Start Agent', command=_save,
                    bg='#0A2647', fg='#FFFFFF', font=('Arial', 10, 'bold'),
                    relief='flat', padx=20, pady=8, cursor='hand2',
                    activebackground='#0E7C86', activeforeground='#FFFFFF')
    btn.pack(pady=(10, 0))

    root.mainloop()
    return result['key']


def _get_api_key() -> str:
    config = _load_config()
    if config.get('api_key'):
        return config['api_key']
    key = _show_setup_window()
    if not key:
        logger.error('No API key provided. Exiting.')
        sys.exit(1)
    return key


# ── File processing ─────────────────────────────────────────────

def _process_file(file_path: Path, api_key: str):
    """Full pipeline: read → generate → build PPTX → review → save."""
    logger.info(f'▶ Processing: {file_path.name}')

    try:
        from reader import read_file
        from generator import generate_proposal
        from reviewer import review_proposal
        from builder import build_pptx

        # 1. Read draft
        text = read_file(file_path)
        if not text or len(text.strip()) < 30:
            logger.warning(f'  File is empty or too short: {file_path.name}')
            return
        logger.info(f'  Read {len(text):,} characters.')

        # 2. Generate proposal content via Gemini
        logger.info('  Calling Gemini Generator...')
        proposal_data = generate_proposal(text, api_key)
        logger.info(f'  Generated: {proposal_data.get("program_title", "(no title)")}')

        # 3. Build PPTX
        stem = file_path.stem
        out_name = f'Zawad_Proposal_{stem}.pptx'
        out_path = COMPLETED_DIR / out_name
        logger.info(f'  Building PPTX → {out_name}')
        build_pptx(proposal_data, str(out_path))
        logger.info(f'  ✓ PPTX saved: {out_path}')

        # 4. Review
        logger.info('  Calling Gemini Reviewer...')
        report = review_proposal(proposal_data, text, api_key)
        review_name = f'Review_{stem}.txt'
        review_path = REVIEW_DIR / review_name
        with open(review_path, 'w', encoding='utf-8') as f:
            f.write(report)
        logger.info(f'  ✓ Review saved: {review_path}')

        logger.info(f'✅ Done: {file_path.name}')

    except Exception as exc:
        logger.error(f'❌ Error processing {file_path.name}: {exc}', exc_info=True)


# ── Watchdog ─────────────────────────────────────────────────────

class _Handler:
    def __init__(self, api_key: str):
        self._api_key = api_key
        self._active = set()

    def dispatch(self, event):
        from watchdog.events import FileCreatedEvent, FileMovedEvent
        if isinstance(event, FileCreatedEvent):
            self._on_created(event.src_path)
        elif isinstance(event, FileMovedEvent):
            self._on_created(event.dest_path)

    def _on_created(self, path_str: str):
        fp = Path(path_str)
        # Ignore files inside subfolders
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
            time.sleep(3)  # wait for file to finish copying
            if fp.exists():
                _process_file(fp, self._api_key)
            self._active.discard(str(fp))

        threading.Thread(target=_run, daemon=True).start()


def _scan_existing(api_key: str):
    """Process files already in the watch folder at startup (retry after failed runs)."""
    for fp in WATCH_FOLDER.iterdir():
        if fp.is_file() and fp.suffix.lower() in SUPPORTED_EXT:
            logger.info(f'Startup scan — processing existing file: {fp.name}')
            threading.Thread(target=_process_file, args=(fp, api_key), daemon=True).start()
            time.sleep(2)


def _run_watcher(api_key: str):
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    _scan_existing(api_key)

    handler = _Handler(api_key)

    class _Bridge(FileSystemEventHandler):
        def on_created(self, event):
            handler.dispatch(event)
        def on_moved(self, event):
            handler.dispatch(event)

    observer = Observer()
    observer.schedule(_Bridge(), str(WATCH_FOLDER), recursive=False)
    observer.start()
    logger.info(f'Watching: {WATCH_FOLDER}')

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


# ── Status window ────────────────────────────────────────────────

def _show_status_window(api_key: str):
    import subprocess
    from tkinter import filedialog

    root = tk.Tk()
    root.title('Zawad Proposal Agent — Running')
    root.geometry('420x295')
    root.resizable(False, False)
    root.configure(bg='#0A2647')

    tk.Label(root, text='زواد للتطوير والاستشارات',
             font=('Arial', 13, 'bold'), fg='#2EC4B6', bg='#0A2647').pack(pady=(18, 2))
    tk.Label(root, text='Proposal Agent is Running',
             font=('Arial', 10), fg='#FFFFFF', bg='#0A2647').pack()

    card = tk.Frame(root, bg='#E7F8F6', padx=18, pady=14)
    card.pack(fill='both', expand=True, padx=14, pady=12)

    status_lbl = tk.Label(card, text='● Agent is active — watching your folder',
                          font=('Arial', 10), fg='#0E7C86', bg='#E7F8F6')
    status_lbl.pack(anchor='w')
    tk.Label(card, text='Drop any draft (PDF, Word, PPT) into:',
             font=('Arial', 9), fg='#1B2A3A', bg='#E7F8F6').pack(anchor='w', pady=(6, 0))
    tk.Label(card, text='  Downloads → Proposals_2026',
             font=('Arial', 9, 'bold'), fg='#0A2647', bg='#E7F8F6').pack(anchor='w')
    tk.Label(card, text='Output PPTX → Proposals_2026/completed/',
             font=('Arial', 9), fg='#1B2A3A', bg='#E7F8F6').pack(anchor='w', pady=(4, 0))
    tk.Label(card, text='Review report → Proposals_2026/review_needed/',
             font=('Arial', 9), fg='#1B2A3A', bg='#E7F8F6').pack(anchor='w')

    btn_frame = tk.Frame(card, bg='#E7F8F6')
    btn_frame.pack(fill='x', pady=(10, 0))

    def _open_folder():
        subprocess.Popen(f'explorer "{WATCH_FOLDER}"')

    def _manual_trigger():
        fp_str = filedialog.askopenfilename(
            title='Select a draft file to process',
            initialdir=str(WATCH_FOLDER),
            filetypes=[
                ('Supported files', '*.pdf *.docx *.doc *.pptx *.ppt *.txt'),
                ('All files', '*.*'),
            ]
        )
        if not fp_str:
            return
        fp = Path(fp_str)
        status_lbl.config(text=f'⏳ Processing: {fp.name}', fg='#B87A00')
        root.update_idletasks()

        def _run():
            _process_file(fp, api_key)
            root.after(0, lambda: status_lbl.config(
                text='✅ Done — agent watching for next file', fg='#0A7C2F'))

        threading.Thread(target=_run, daemon=True).start()

    tk.Button(btn_frame, text='Open Proposals_2026 folder',
              command=_open_folder, bg='#0A2647', fg='#FFFFFF',
              font=('Arial', 9), relief='flat', padx=10, pady=5,
              cursor='hand2').pack(side='left', padx=(0, 6))

    tk.Button(btn_frame, text='⚡ Process a File Now',
              command=_manual_trigger, bg='#0E7C86', fg='#FFFFFF',
              font=('Arial', 9, 'bold'), relief='flat', padx=10, pady=5,
              cursor='hand2').pack(side='left')

    # Start watcher thread
    threading.Thread(target=_run_watcher, args=(api_key,), daemon=True).start()

    root.mainloop()


# ── Entry point ───────────────────────────────────────────────

def main():
    _ensure_folders()
    api_key = _get_api_key()
    logger.info('Agent starting...')
    _show_status_window(api_key)


if __name__ == '__main__':
    main()
