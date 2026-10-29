====================================================================
  Zawad Proposal Agent
  زواد للتطوير والاستشارات — نظام إنشاء العروض الآلي
====================================================================

HOW IT WORKS
────────────
1. You drop a draft file (PDF, Word, or PPT) into:
   Downloads\Proposals_2026\

2. The agent auto-detects the file and starts processing immediately.
   No buttons. No windows to click.

3. Within a few minutes:
   ✓ Your branded proposal PPTX is saved in:
       Downloads\Proposals_2026\completed\

   ✓ A review report is saved in:
       Downloads\Proposals_2026\review_needed\


SETUP (First Time Only)
────────────────────────
1. Run install.bat (one time — installs Python packages)

2. Double-click run.bat to start the agent.
   A small window opens asking for your Gemini API key.

3. Get your FREE API key from:
   https://aistudio.google.com
   → Sign in with Google
   → Click "Get API Key"
   → Copy the key

4. Paste it in the setup window and press Save.
   The key is saved on your computer — you'll never be asked again.


STARTING THE AGENT
───────────────────
Every time you want to use the agent:
→ Double-click run.bat

The small status window shows the agent is running.
Minimize it — it can run in the background.
You do NOT need to keep the window open (closing it stops the agent).


FOR EACH PERSON
────────────────
Each person on the team uses their OWN:
- Gemini API key (from their own Google account — it's free)
- Downloads\Proposals_2026\ folder (on their own computer)

No shared accounts or shared servers needed.


SUPPORTED INPUT FORMATS
────────────────────────
- PDF (.pdf)
- Microsoft Word (.docx, .doc)
- PowerPoint (.pptx, .ppt)
- Plain text (.txt)

The draft can be rough notes, a bullet list, or a partial outline.
The agent fills in all 9 proposal sections automatically.


OUTPUT FILES
─────────────
Proposal PPTX : Downloads\Proposals_2026\completed\Zawad_Proposal_[name].pptx
Review report : Downloads\Proposals_2026\review_needed\Review_[name].txt

The review report checks:
- Arabic language quality
- Pricing math (rate card + 5% VAT)
- Section completeness
- Compliance items (VAT, Circular 13/2024, performance bond)


TROUBLESHOOTING
────────────────
• Nothing happening after dropping a file?
  → Check agent.log in this folder for error messages
  → Make sure the agent (run.bat) is still running

• "API key" error?
  → Delete config\config.json and restart run.bat
  → Enter your key again

• Python not found?
  → Download Python from https://python.org
  → During install: check "Add Python to PATH"
  → Then run install.bat again


CONTACT
────────
fahmi@zawad.om
ghasan@zawad.om | +968 9777 3444
www.zawad.om

====================================================================
