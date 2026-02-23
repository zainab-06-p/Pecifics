"""
Pecifics AI Desktop Assistant - LangChain Backend
=====================================================
No GPU required. Runs on free cloud APIs.

Supported LLM providers (set LLM_PROVIDER env var):
  groq         - Groq cloud FREE tier (llama-3.3-70b)         ← DEFAULT ★
  ollama       - Local Ollama (llama3, mistral, qwen, phi3)
  huggingface  - HuggingFace Inference API free tier
  transformers - Direct local HF model (GPU, works on Kaggle/Colab)
  openai       - OpenAI GPT-4o (paid, optional)

Vision provider (set VISION_PROVIDER env var):
  gemini       - Google Gemini 2.0 Flash (free tier)          ← DEFAULT ★
  openai       - GPT-4o vision (paid)
  ollama       - Local llava/moondream
  transformers - Qwen2-VL on GPU

Quick start (recommended — no GPU, no ngrok, fully free):
  # 1. Get free Groq key → https://console.groq.com
  # 2. Get free Gemini key → https://aistudio.google.com/app/apikey
  set GROQ_API_KEY=your_groq_key
  set GEMINI_API_KEY=your_gemini_key
  python langchain_backend.py --port 8000

  # Option B — Ollama (fully local)
  ollama pull llama3
  set LLM_PROVIDER=ollama
  set VISION_PROVIDER=ollama
  python langchain_backend.py --port 8000

  # Option C — Kaggle/Colab GPU (legacy)
  # Set LLM_PROVIDER=transformers in the notebook, model loads inline
"""

import os, re, json, base64, logging, traceback, gc
from io import BytesIO
from typing import Dict, List, Optional
from datetime import datetime

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import uvicorn

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Config (from .env or environment)
# ─────────────────────────────────────────────────────────────────────────────

try:
    from dotenv import load_dotenv; load_dotenv()
except ImportError:
    pass

LLM_PROVIDER    = os.getenv("LLM_PROVIDER",    "groq")            # ← default: Groq (free, fast)
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL    = os.getenv("OLLAMA_MODEL",    "llama3")
GROQ_API_KEY    = os.getenv("GROQ_API_KEY",    "")
GROQ_MODEL      = os.getenv("GROQ_MODEL",      "llama-3.3-70b-versatile")   # free tier
HF_API_KEY      = os.getenv("HF_API_KEY",      "")
HF_MODEL        = os.getenv("HF_MODEL",        "mistralai/Mistral-7B-Instruct-v0.3")
HF_LOCAL_MODEL  = os.getenv("HF_LOCAL_MODEL",  "Qwen/Qwen2.5-7B-Instruct")  # transformers provider
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY",  "")
OPENAI_MODEL    = os.getenv("OPENAI_MODEL",    "gpt-4o")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY",  "")               # ← default vision provider
GEMINI_MODEL    = os.getenv("GEMINI_MODEL",    "gemini-2.0-flash")  # free tier
VISION_PROVIDER = os.getenv("VISION_PROVIDER", "gemini")         # ← default: Gemini vision
MAX_RETRIES     = int(os.getenv("MAX_RETRIES", "3"))
MAX_ITERATIONS  = int(os.getenv("MAX_ITERATIONS", "12"))

# ─────────────────────────────────────────────────────────────────────────────
# Optional dependency flags
# ─────────────────────────────────────────────────────────────────────────────

HAS_OPENAI       = False
HAS_OLLAMA       = False
HAS_GROQ         = False
HAS_HF_HUB       = False
HAS_TRANSFORMERS = False
HAS_PIL          = False
HAS_GEMINI       = False

try:
    from langchain_openai import ChatOpenAI; HAS_OPENAI = True
except ImportError: pass

try:
    from langchain_community.chat_models import ChatOllama; HAS_OLLAMA = True
except ImportError:
    try:
        from langchain_ollama import ChatOllama; HAS_OLLAMA = True
    except ImportError: pass

try:
    from langchain_groq import ChatGroq; HAS_GROQ = True
except ImportError: pass

try:
    from langchain_community.llms import HuggingFaceEndpoint
    from langchain_community.chat_models.huggingface import ChatHuggingFace
    HAS_HF_HUB = True
except ImportError: pass

try:
    import torch
    from transformers import AutoModelForCausalLM, AutoTokenizer, BitsAndBytesConfig, pipeline
    from langchain_community.llms import HuggingFacePipeline
    HAS_TRANSFORMERS = True
except ImportError: pass

try:
    from PIL import Image; HAS_PIL = True
except ImportError: pass

try:
    import google.generativeai as genai; HAS_GEMINI = True
except ImportError: pass

from langchain_core.messages import HumanMessage

# ─────────────────────────────────────────────────────────────────────────────
# Session store (in-memory, per-request retry state)
# ─────────────────────────────────────────────────────────────────────────────

sessions: Dict[str, Dict] = {}

# ─────────────────────────────────────────────────────────────────────────────
# Request / Response models
# ─────────────────────────────────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message:              str
    screenshot:           Optional[str] = None
    conversation_history: Optional[List[Dict]] = []
    screen_width:         Optional[int] = 1920
    screen_height:        Optional[int] = 1080
    user_choice:          Optional[Dict] = None
    session_id:           Optional[str] = None

class ContinueRequest(BaseModel):
    session_id:     str
    action_results: List[Dict]
    screenshot:     Optional[str] = None

class AnalyzeScreenRequest(BaseModel):
    screenshot: Optional[str] = None
    question:   Optional[str] = "What do you see on screen?"

class VerifyRequest(BaseModel):
    screenshot:      str
    task:            str
    expected_result: Optional[str] = ""

# ─────────────────────────────────────────────────────────────────────────────
# Action catalog — everything the Electron side can execute
# ─────────────────────────────────────────────────────────────────────────────

ACTION_CATALOG = """
=== VISION / SCREEN ANALYSIS ===
get_screenshot(question)           — capture the current screen and use AI vision to answer the question
describe_screen()                  — describe everything visible on screen right now
analyze_screen(question)           — ask a specific question about screen content


create_file(file_path, content) | create_folder(folder_path) | delete_file(file_path)
read_file(file_path)            — read full content of any file (txt, md, csv, py, docx, etc.)
find_in_file(file_path, search_text)                          — find text in file, returns matching lines/paragraphs/slides
replace_in_file(file_path, search_text, replacement_text)     — replace text in any file (plain text, Word, or PPT)
append_to_file(file_path, content)                            — append content to end of file
list_directory(path) | copy_file(source, destination)
move_file(source, destination) | rename_file(old_path, new_name)
search_files(pattern, location) | search_file_content(search_text, location, file_pattern)
open_file(file_path) | show_in_explorer(file_path)

=== APPLICATION CONTROL ===
open_application(app_name) | close_application(app_name) | open_url(url)
search_web(query) | focus_app(app_name) | download_file(url, filename)

=== MOUSE / KEYBOARD ===
click_at(x, y, click_type) | move_mouse(x, y) | scroll(direction, amount)
drag(start_x, start_y, end_x, end_y) | type_text(text) | type_into_app(app_name, text)
press_key(key) | hotkey(key) | wait(duration)

=== SYSTEM CORE ===
set_volume(volume) | set_brightness(brightness) | toggle_wifi(enable)
toggle_bluetooth(enable) | toggle_night_light(enable) | set_wallpaper(image_path)
lock_computer() | sleep_computer() | get_battery_status() | get_system_info()
get_disk_space() | get_network_status() | empty_recycle_bin() | run_disk_cleanup()
set_resolution(width, height) | speak(message)

=== OS TASKS (EXTENDED) ===
clear_cache(cache_type)            - type: temp/dns/browser/arp/icon/chrome/edge/firefox/all
flush_dns() | reset_network() | ping_host(host, count) | check_port(host, port) | get_public_ip()
run_winr(command)                  - execute any Win+R dialog command
open_msconfig() | open_services() | open_device_manager() | open_regedit()
open_event_viewer() | open_task_scheduler() | open_firewall() | open_network_connections()
open_power_options() | open_windows_update() | open_task_manager() | open_control_panel()
open_cmd() | open_powershell() | open_cmd_admin() | open_powershell_admin()
manage_service(service_name, action)  - start/stop/restart/status/enable/disable
list_services(filter) | get_running_processes() | kill_process(name_or_pid)
list_startup_programs() | toggle_startup_program(name, enable)
check_windows_update() | get_update_history()
get_event_logs(log_type, count, level) | clear_event_log(log_type)
create_restore_point(description) | list_restore_points()
get_installed_apps() | uninstall_app(app_name)
get_env_variable(name) | set_env_variable(name, value, scope) | delete_env_variable(name)
read_registry(key_path, value_name) | write_registry(key_path, value_name, value)
get_disk_health() | analyze_storage(path) | run_disk_cleanup_silent()
get_power_plan() | set_power_plan(plan)   - balanced/high_performance/power_saver
hibernate() | restart_computer(delay) | shutdown_computer(delay) | cancel_shutdown()
windows_defender_scan() | get_defender_status() | check_firewall()
list_scheduled_tasks() | run_scheduled_task(name) | disable_scheduled_task(name)
toggle_dark_mode(enable) | enable_dark_mode() | light_mode() | refresh_desktop()
get_clipboard() | set_clipboard(text) | show_notification(title, message) | list_fonts()
get_system_health()

=== BROWSER AUTOMATION ===
browser_open(url, browser) | browser_navigate(url) | browser_click(selector)
browser_type(selector, text) | browser_get_text(selector) | browser_screenshot()
browser_scroll(direction, amount) | browser_wait_for(selector) | browser_close()
browser_go_back() | browser_new_tab(url) | browser_reload() | browser_fill_form(fields)
browser_smart_login(url, name, email, password, is_new_user)  — login OR signup, auto-detects form type
browser_signup(url, name, email, password)                    — explicitly create a new account
browser_search_in_page(text)                                  — type into search bar / ChatGPT prompt and submit
browser_chat(text)                                            — alias for browser_search_in_page (for AI chat sites)
browser_detect_page()                                         — returns: login | signup | chat | search | main
open_gmail(email)                                             — open Gmail for a SPECIFIC account email (auto-detects which slot u/0, u/1, etc.)
send_gmail(to, subject, body, account_email)                  — send email; account_email selects which Gmail account to send FROM
google_search(query) | youtube_search(query)
browser_execute_script(script) | install_playwright()

=== MS OFFICE — WORD ===
word_create_document(title, content) | word_open_document(filepath)
word_read_content()                                           — read all paragraphs from open Word doc
word_find_replace(search_text, replacement_text)              — find & replace text in open Word doc
word_add_paragraph(text, bold, italic, font_size, color, alignment)
word_add_heading(text, level) | word_insert_table(rows, cols)
word_apply_theme(theme_name) | word_save(filename)

=== MS OFFICE — EXCEL ===
excel_create_workbook() | excel_open_workbook(filepath)
excel_write_cell(row, col, value) | excel_write_data(data)
excel_add_worksheet(name) | excel_create_chart(chart_type) | excel_save(filename)

=== MS OFFICE — POWERPOINT ===
create_presentation(title, save_path, slides_content)         — CREATE a full PPT from scratch with all slides at once
powerpoint_new_slide() | powerpoint_add_title(title) | powerpoint_add_content(content)
ppt_apply_theme(theme_name) | ppt_add_animation(animation_type) | ppt_change_layout(layout)
ppt_find_slide(search_text)                                   — find slides containing text, returns slide numbers
ppt_get_slide_content(slide_number)                           — read all text shapes from a slide
ppt_update_slide_text(slide_number, old_text, new_text)       — replace text on a specific slide

TYPICAL WORKFLOW for modifying an existing file:
  1. open_file(file_path)            — open the file in its app
  2. read_file(file_path)            — read content so you know what's there
  3. find_in_file(file_path, text)   — locate the section to change
  4. replace_in_file(file_path, old, new) — make the change
     OR for Word: word_find_replace(old, new)
     OR for PPT:  ppt_find_slide(text) → ppt_update_slide_text(num, old, new)
"""

# ─────────────────────────────────────────────────────────────────────────────
# System prompt
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = f"""You are Pecifics, an intelligent AI desktop assistant for Windows 11.
You operate exactly like GitHub Copilot / Claude Sonnet — you think step-by-step, ask clarifying questions before acting on ambiguous tasks, execute actions autonomously, and provide complete human-readable summaries.

AVAILABLE ACTIONS:
{ACTION_CATALOG}

═══════════════════════════════════════════════════════
BEHAVIOR RULES (follow these EXACTLY like Claude would)
═══════════════════════════════════════════════════════

1. ALWAYS ASK BEFORE CREATING/SAVING FILES:
   - If the user says "save this", "create a doc", "make a file", "write a report" etc. and has NOT provided a filename or save location → set clarification_needed=true and ask for both.
   - Fields to ask: filename (with extension), save_path (default suggestion: Desktop or Documents).

2. ALWAYS ASK BEFORE CREATING PRESENTATIONS (PPT):
   - Ask: title, topic/content, number of slides, save location.
   - Use action: create_presentation(title, slides_content, save_path) where slides_content is a JSON array of {{title, content}} objects.
   - create_presentation automatically saves AND opens PowerPoint to show the result — do NOT add a separate open_file action.

3. TASKS THAT ALWAYS NEED CLARIFICATION:
   - Save/create any document → ask filename + location
   - Send email → ask recipient, subject, body if not provided; if user has multiple Gmail accounts ask WHICH one to send from (account_email)
   - Open Gmail → if the user has said they have multiple accounts OR the message is ambiguous, ask: "Which Gmail account? (e.g. work@gmail.com or personal@gmail.com)"
   - Login to a website → ask email + password if not provided
   - Sign up / create account → ask name, email, password (+ confirm what site)
   - Create PPT/presentation → ask title, topic, slides, save path
   - Delete files → ask for confirmation (use requires_choice=true with Yes/No)
   - If user says "search X on ChatGPT" but not logged in → ask: "Do you need me to log in first? If yes, provide email and password."

4. TASKS THAT DON'T NEED CLARIFICATION (just do them):
   - Open applications, websites, system tools (just open — no login needed)
   - Volume, brightness, bluetooth, wifi controls
   - Browser navigation: open site → search/type something on it (no account required)
   - System info, battery, disk space
   - Take screenshot, show notifications
   - "Open ChatGPT in browser" with no login request → just browser_open + browser_search_in_page

5. TASK SUMMARY:
   - ALWAYS include a task_summary field — a clear human-readable summary of what you are about to do (or did), written in first person past tense once actions are listed.
   - Example: "I'll create a 5-slide PowerPoint presentation titled 'Q1 Sales Report', save it to your Desktop, and open it in PowerPoint."

6. BROWSER RULES — GMAIL MULTI-ACCOUNT:
   - Gmail supports multiple signed-in accounts at /mail/u/0/, /mail/u/1/, etc.
   - When user says "open my Gmail" or "open Gmail" and a specific email is given or implied: use open_gmail(email=\"their@email.com\")
   - open_gmail scans all account slots, finds the matching one, and navigates directly to it.
   - If no email is specified AND context implies they may have multiple accounts: ask which one.
   - When sending email, always set account_email to the sender’s address: send_gmail(to, subject, body, account_email=\"sender@gmail.com\")
   - NEVER use browser_open(\"https://mail.google.com\") without account_email when a specific account was requested.

   EMAIL FLOW EXAMPLE — \"Open my work Gmail and send an email to X\":
     clarify: which work account address? (if not already known)
     1. open_gmail(email=\"work@gmail.com\")
     2. send_gmail(to=\"x@example.com\", subject=\"...\", body=\"...\", account_email=\"work@gmail.com\")

7. BROWSER RULES — GENERAL:
   - Use browser_open with DIRECT URL — never use search_web to navigate to a known site.
   - For LOGIN to existing account: browser_smart_login(url, email, password, is_new_user=false)
   - For SIGNUP / creating a new account: browser_smart_login(url, name, email, password, is_new_user=true)
     → It auto-clicks the “Sign up” link, fills name/email/password fields, handles confirm-password
   - For SEARCHING on any site or CHATTING with ChatGPT/AI: browser_search_in_page(text)
     → Works with ChatGPT prompt, YouTube search bar, Google search, Bing, site search boxes
   - For TYPING into a page’s chat/search without submitting: browser_type(selector, text)
   - Selectors: use input[type="email"], input[type="password"], button[type="submit"] — never generic class names.
   - Complex browser flows (open site → login/signup → search/use it) should be described as: clarify credentials if not provided, then use browser_smart_login + browser_search_in_page

   EXAMPLE FLOWS:
   “Open ChatGPT and ask it about climate change”:
     1. browser_open(url="https://chat.openai.com")
     2. browser_search_in_page(text="Tell me about climate change")

   “Open ChatGPT, sign me in, then search X” (ask for credentials first):
     1. browser_smart_login(url="https://chat.openai.com", email=EMAIL, password=PASS, is_new_user=false)
     2. browser_search_in_page(text="X")

   “Create a ChatGPT account” (ask name/email/password first):
     1. browser_smart_login(url="https://chat.openai.com", name=NAME, email=EMAIL, password=PASS, is_new_user=true)

   “Open Google and search for AI news”:
     1. browser_open(url="https://www.google.com")
     2. browser_search_in_page(text="AI news 2026")

8. OUTPUT FORMAT (ALWAYS valid JSON, no markdown code blocks):
{{
  "message": "Brief one-line description of what you're doing",
  "task_summary": "Complete human-readable explanation of what will happen, step by step",
  "clarification_needed": false,
  "clarification_question": "",
  "clarification_fields": [],
  "actions": [
    {{"name": "action_name", "parameters": {{"param": "value"}}}}
  ],
  "expected_result": "What should be true after all actions complete",
  "requires_choice": false,
  "choice_type": null,
  "choice_options": [],
  "pending_task": null,
  "session_continues": false
}}

CLARIFICATION FIELDS FORMAT:
{{
  "clarification_needed": true,
  "clarification_question": "To save your document, I need a couple of details:",
  "clarification_fields": [
    {{"key": "filename", "label": "File Name", "placeholder": "e.g. MyReport.docx", "default": "Document.docx"}},
    {{"key": "save_path", "label": "Save Location", "placeholder": "e.g. Desktop, C:\\\\Users\\\\...", "default": "Desktop"}}
  ],
  "actions": [],
  "task_summary": "Once you provide details, I'll create and save the document.",
  "message": "I need a few details before I can save the file."
}}

PPT CREATION EXAMPLE (single action — it saves and opens the file automatically):
  create_presentation(title="Q1 Report", save_path="C:\\\\Users\\\\USERNAME\\\\Desktop\\\\Q1Report.pptx",
    slides_content=[
      {{"slide_title": "Introduction", "slide_content": "Overview of Q1 performance"}},
      {{"slide_title": "Revenue", "slide_content": "Revenue grew 20% YoY. Key drivers: product A and B."}}
    ])
"""

# ─────────────────────────────────────────────────────────────────────────────
# LLM factory
# ─────────────────────────────────────────────────────────────────────────────

_cached_transformers_llm = None


def _load_transformers_llm():
    """Load a local HuggingFace checkpoint and wrap it as a LangChain LLM.
    Reuses the global vision/text model if already loaded by the Kaggle notebook."""
    global _cached_transformers_llm
    if _cached_transformers_llm is not None:
        return _cached_transformers_llm

    if not HAS_TRANSFORMERS:
        raise ImportError("pip install transformers accelerate bitsandbytes")

    # Check if the Kaggle notebook already loaded a text model
    import sys
    g = sys._getframe(0).f_globals
    text_model     = g.get("text_model")
    text_tokenizer = g.get("text_tokenizer")

    if text_model and text_tokenizer:
        logger.info("Reusing already-loaded text_model from notebook globals.")
        pipe = pipeline(
            "text-generation", model=text_model, tokenizer=text_tokenizer,
            max_new_tokens=1024, temperature=0.1, do_sample=True,
            pad_token_id=text_tokenizer.eos_token_id,
        )
        _cached_transformers_llm = HuggingFacePipeline(pipeline=pipe)
        return _cached_transformers_llm

    model_id = HF_LOCAL_MODEL
    logger.info(f"Loading {model_id} …")
    gc.collect()
    if torch.cuda.is_available():
        torch.cuda.empty_cache()

    bnb_cfg = None
    if torch.cuda.is_available():
        try:
            bnb_cfg = BitsAndBytesConfig(
                load_in_4bit=True, bnb_4bit_quant_type="nf4",
                bnb_4bit_compute_dtype=torch.float16, bnb_4bit_use_double_quant=True,
            )
        except Exception:
            pass

    tokenizer = AutoTokenizer.from_pretrained(model_id, trust_remote_code=True)
    model = AutoModelForCausalLM.from_pretrained(
        model_id,
        torch_dtype=torch.float16 if torch.cuda.is_available() else torch.float32,
        device_map="auto" if torch.cuda.is_available() else None,
        quantization_config=bnb_cfg,
        trust_remote_code=True,
        low_cpu_mem_usage=True,
    ).eval()

    pipe = pipeline(
        "text-generation", model=model, tokenizer=tokenizer,
        max_new_tokens=1024, temperature=0.1, do_sample=True,
        pad_token_id=tokenizer.eos_token_id,
    )
    _cached_transformers_llm = HuggingFacePipeline(pipeline=pipe)
    logger.info("Local model ready.")
    return _cached_transformers_llm


def create_llm():
    provider = LLM_PROVIDER.lower()

    if provider == "ollama":
        if not HAS_OLLAMA:
            raise ImportError("pip install langchain-community  (or langchain-ollama)")
        logger.info(f"Using Ollama — {OLLAMA_MODEL} @ {OLLAMA_BASE_URL}")
        return ChatOllama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL, temperature=0.1)

    if provider == "groq":
        if not HAS_GROQ:
            raise ImportError("pip install langchain-groq")
        if not GROQ_API_KEY:
            raise ValueError("GROQ_API_KEY not set. Free key → https://console.groq.com")
        logger.info(f"Using Groq — {GROQ_MODEL}")
        return ChatGroq(api_key=GROQ_API_KEY, model_name=GROQ_MODEL, temperature=0.1)

    if provider == "huggingface":
        if not HAS_HF_HUB:
            raise ImportError("pip install langchain-community")
        if not HF_API_KEY:
            raise ValueError("HF_API_KEY not set. Free token → https://huggingface.co/settings/tokens")
        logger.info(f"Using HuggingFace Inference API — {HF_MODEL}")
        llm = HuggingFaceEndpoint(
            repo_id=HF_MODEL, huggingfacehub_api_token=HF_API_KEY,
            temperature=0.1, max_new_tokens=1024,
        )
        return ChatHuggingFace(llm=llm)

    if provider == "transformers":
        logger.info(f"Using local transformers — {HF_LOCAL_MODEL}")
        return _load_transformers_llm()

    if provider == "openai":
        if not HAS_OPENAI:
            raise ImportError("pip install langchain-openai")
        if not OPENAI_API_KEY:
            raise ValueError("OPENAI_API_KEY not set.")
        logger.info(f"Using OpenAI — {OPENAI_MODEL}")
        return ChatOpenAI(model=OPENAI_MODEL, api_key=OPENAI_API_KEY, temperature=0.1, max_retries=3)

    raise ValueError(
        f"Unknown LLM_PROVIDER='{provider}'. "
        "Valid: ollama | groq | huggingface | transformers | openai"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Vision helpers
# ─────────────────────────────────────────────────────────────────────────────

def describe_screenshot(b64: str, llm) -> str:
    if not b64:
        return "No screenshot available."

    vision = VISION_PROVIDER.lower()

    # ── Gemini vision (default, free tier) ──────────────────────────────────
    if vision == "gemini" and HAS_GEMINI and GEMINI_API_KEY:
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            model = genai.GenerativeModel(GEMINI_MODEL)
            img_data = base64.b64decode(b64)
            img_part = {"mime_type": "image/jpeg", "data": img_data}
            resp = model.generate_content([
                "Describe this Windows desktop screenshot in ≤100 words: open windows, active app, visible text, current state.",
                img_part,
            ])
            return resp.text.strip()
        except Exception as e:
            logger.warning(f"Gemini vision failed: {e} — falling back.")

    # ── GPT-4o vision ────────────────────────────────────────────────────────
    if vision == "openai" and HAS_OPENAI and OPENAI_API_KEY:
        try:
            vis = ChatOpenAI(model="gpt-4o", api_key=OPENAI_API_KEY, temperature=0, max_tokens=300)
            resp = vis.invoke([HumanMessage(content=[
                {"type": "text", "text": "Describe this Windows desktop screenshot in ≤100 words: open windows, active app, visible text, current state."},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}", "detail": "low"}},
            ])])
            return resp.content
        except Exception as e:
            logger.warning(f"GPT-4V: {e}")

    # ── Ollama vision (llava / moondream / bakllava) ─────────────────────────
    if vision == "ollama" or LLM_PROVIDER == "ollama":
        for vm in ("llava", "moondream", "bakllava", "llava-llama3"):
            try:
                import requests as _r
                r = _r.post(f"{OLLAMA_BASE_URL}/api/generate",
                            json={"model": vm, "prompt": "Describe this Windows screen briefly.", "images": [b64], "stream": False},
                            timeout=15)
                if r.ok:
                    return r.json().get("response", "Screen captured.")
            except Exception:
                continue

    # ── Local transformers vision (Qwen2-VL etc.) ──────────────────────────
    if (vision == "transformers" or LLM_PROVIDER == "transformers") and HAS_TRANSFORMERS and HAS_PIL:
        try:
            import sys
            g = sys._getframe(0).f_globals
            vm = g.get("vision_model") or g.get("LOADED_VISION_MODEL")
            vp = g.get("vision_processor") or g.get("LOADED_VISION_PROCESSOR")
            if vm and vp:
                img = Image.open(BytesIO(base64.b64decode(b64))).convert("RGB")
                msgs = [{"role": "user", "content": [
                    {"type": "image", "image": img},
                    {"type": "text", "text": "Describe what's on this Windows screen in 80 words."},
                ]}]
                text = vp.apply_chat_template(msgs, tokenize=False, add_generation_prompt=True)
                inp = vp(text=[text], images=[img], return_tensors="pt").to(vm.device)
                with torch.no_grad():
                    out = vm.generate(**inp, max_new_tokens=200)
                return vp.batch_decode(
                    [o[inp.input_ids.shape[1]:] for o in out], skip_special_tokens=True
                )[0].strip()
        except Exception as e:
            logger.warning(f"Transformers vision: {e}")

    return "Screenshot captured. Vision analysis unavailable for this provider — proceeding with text-only planning."


# ─────────────────────────────────────────────────────────────────────────────
# LLM invocation wrapper (supports both chat models and text pipelines)
# ─────────────────────────────────────────────────────────────────────────────

def _call(llm, prompt: str) -> str:
    try:
        resp = llm.invoke([HumanMessage(content=prompt)])
        return resp.content if hasattr(resp, "content") else str(resp)
    except TypeError:
        return llm.invoke(prompt)


# ─────────────────────────────────────────────────────────────────────────────
# Core planning / verification
# ─────────────────────────────────────────────────────────────────────────────

def plan_actions(
    user_message: str,
    screenshot_b64: Optional[str],
    conversation_history: List[Dict],
    screen_width: int,
    screen_height: int,
    user_choice: Optional[Dict],
    llm,
    retry_count: int = 0,
    previous_error: Optional[str] = None,
) -> Dict:

    history_txt = "\n".join(
        f"{'User' if t.get('role')=='user' else 'Assistant'}: {t.get('content','')}"
        for t in conversation_history[-6:]
    )
    screen_desc = describe_screenshot(screenshot_b64, llm) if screenshot_b64 else "No screenshot."
    retry_ctx = (
        f"\n\nPREVIOUS ATTEMPT FAILED: {previous_error}\nUse a DIFFERENT approach this time."
        if retry_count > 0 and previous_error else ""
    )
    if user_choice:
        user_message += f"\n[User selected: {user_choice.get('type')} = {user_choice.get('value')}]"

    prompt = f"""{SYSTEM_PROMPT}

CURRENT SCREEN: {screen_desc}
SCREEN SIZE: {screen_width}x{screen_height}

CONVERSATION HISTORY:
{history_txt}
{retry_ctx}

USER REQUEST: {user_message}

Reply ONLY with valid JSON — no markdown, no extra text."""

    try:
        raw = _call(llm, prompt).strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        result = json.loads(raw)
        result.setdefault("message", "Processing…")
        result.setdefault("task_summary", "")
        result.setdefault("actions", [])
        result.setdefault("expected_result", "")
        result.setdefault("task_count", len(result["actions"]))
        result.setdefault("requires_choice", False)
        result.setdefault("choice_options", [])
        result.setdefault("clarification_needed", False)
        result.setdefault("clarification_question", "")
        result.setdefault("clarification_fields", [])
        result.setdefault("session_continues", False)
        return result

    except json.JSONDecodeError as e:
        logger.error(f"JSON parse error (attempt {retry_count+1}): {e}")
        if retry_count < MAX_RETRIES:
            return plan_actions(user_message, screenshot_b64, conversation_history,
                                screen_width, screen_height, user_choice, llm,
                                retry_count + 1, f"JSON parse error: {e}. Return ONLY raw JSON.")
        return {"message": "I had trouble planning that. Please rephrase.", "actions": [],
                "task_count": 0, "requires_choice": False, "error": str(e)}

    except Exception as e:
        logger.error(f"LLM failed: {e}")
        if retry_count < MAX_RETRIES:
            return plan_actions(user_message, screenshot_b64, conversation_history,
                                screen_width, screen_height, user_choice, llm,
                                retry_count + 1, str(e))
        return {"message": f"Error: {e}", "actions": [], "task_count": 0,
                "requires_choice": False, "error": str(e)}


def verify_completion(screenshot_b64: str, task: str, expected: str, llm) -> Dict:
    screen = describe_screenshot(screenshot_b64, llm) if screenshot_b64 else "No screenshot."
    prompt = f"""{SYSTEM_PROMPT}

Task just executed: "{task}"
Expected result: "{expected}"
Current screen: {screen}

Was the task completed? Reply ONLY with valid JSON:
{{
  "success": true,
  "observation": "What I see on screen",
  "should_retry": false,
  "retry_actions": []
}}"""
    try:
        raw = _call(llm, prompt).strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        return json.loads(raw)
    except Exception as e:
        return {"success": False, "observation": str(e), "should_retry": False, "retry_actions": []}


# ─────────────────────────────────────────────────────────────────────────────
# FastAPI app
# ─────────────────────────────────────────────────────────────────────────────

app = FastAPI(title="Pecifics AI Backend", version="3.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

_llm = None

def get_llm():
    global _llm
    if _llm is None:
        _llm = create_llm()
    return _llm


@app.get("/health")
async def health():
    status = "ok"
    try:
        get_llm()
    except Exception as e:
        status = f"error: {e}"
    model = {
        "ollama": OLLAMA_MODEL, "groq": GROQ_MODEL,
        "huggingface": HF_MODEL, "transformers": HF_LOCAL_MODEL, "openai": OPENAI_MODEL,
    }.get(LLM_PROVIDER, "?")
    return {"status": "ok", "version": "3.0.0",
            "llm_provider": LLM_PROVIDER, "llm_model": model,
            "llm_status": status, "timestamp": datetime.now().isoformat()}


@app.post("/chat")
async def chat(req: ChatRequest):
    try:
        llm = get_llm()
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"LLM unavailable: {e}")

    try:
        result = plan_actions(
            user_message=req.message,
            screenshot_b64=req.screenshot,
            conversation_history=req.conversation_history or [],
            screen_width=req.screen_width or 1920,
            screen_height=req.screen_height or 1080,
            user_choice=req.user_choice,
            llm=llm,
        )
        if req.session_id:
            sessions[req.session_id] = {
                "task": req.message,
                "expected": result.get("expected_result", ""),
                "history": req.conversation_history or [],
                "retry_count": 0,
            }
        return JSONResponse(content=result)
    except Exception as e:
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/continue")
async def continue_task(req: ContinueRequest):
    try: llm = get_llm()
    except Exception as e: raise HTTPException(status_code=503, detail=str(e))

    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")

    results_summary = "\n".join(
        f"- {r.get('action','?')}: {'✅' if r.get('result',{}).get('success') else '❌ '+str(r.get('result',{}).get('error',''))}"
        for r in req.action_results
    )
    failed = [r for r in req.action_results if not r.get("result", {}).get("success")]

    if not failed:
        vr = verify_completion(req.screenshot or "", session["task"], session["expected"], llm)
        if vr.get("should_retry") and session["retry_count"] < MAX_RETRIES:
            session["retry_count"] += 1
            return {"done": False, "message": f"🔄 Retrying… {vr.get('observation','')}",
                    "actions": vr.get("retry_actions", [])}
        sessions.pop(req.session_id, None)
        done_msg = "✅ Task completed!" if vr.get("success") else f"⚠️ {vr.get('observation','')}"
        return {"done": True, "message": done_msg, "observation": vr.get("observation", "")}

    if session["retry_count"] >= MAX_RETRIES:
        sessions.pop(req.session_id, None)
        return {"done": True, "message": f"❌ Failed after {MAX_RETRIES} retries.\n{results_summary}"}

    session["retry_count"] += 1
    failed_desc = "; ".join(f"{r.get('action')}: {r.get('result',{}).get('error','?')}" for r in failed)
    replan = plan_actions(
        user_message=f"ORIGINAL TASK: {session['task']}\nDONE:\n{results_summary}\nFAILURES: {failed_desc}\nPlan ONLY remaining/retry steps using a DIFFERENT approach for what failed.",
        screenshot_b64=req.screenshot,
        conversation_history=session.get("history", []),
        screen_width=1920, screen_height=1080, user_choice=None, llm=llm,
        retry_count=session["retry_count"], previous_error=failed_desc,
    )
    return {"done": False, "message": replan.get("message", "Retrying…"),
            "actions": replan.get("actions", [])}


@app.post("/verify")
async def verify(req: VerifyRequest):
    try:
        return JSONResponse(content=verify_completion(req.screenshot, req.task, req.expected_result, get_llm()))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/analyze_screen")
async def analyze_screen(req: AnalyzeScreenRequest):
    try:
        llm = get_llm()
        screenshot = req.screenshot or ""
        question   = req.question or "What do you see?"
        desc = describe_screenshot(screenshot, llm)
        # For OpenAI, use full vision Q&A
        if LLM_PROVIDER == "openai" and HAS_OPENAI and OPENAI_API_KEY and screenshot:
            try:
                vis = ChatOpenAI(model="gpt-4o", api_key=OPENAI_API_KEY, temperature=0, max_tokens=600)
                resp = vis.invoke([HumanMessage(content=[
                    {"type": "text", "text": question},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{screenshot}", "detail": "high"}},
                ])])
                return {"answer": resp.content, "description": desc}
            except Exception:
                pass
        return {"answer": f"Screen: {desc}\n\nQuestion: {question}", "description": desc}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class BrowserStateRequest(BaseModel):
    screenshot: str = ""


@app.post("/check_browser_state")
async def check_browser_state(req: BrowserStateRequest):
    """Analyze a browser page screenshot and detect any blocking states."""
    screenshot = req.screenshot or ""
    if not screenshot:
        return {"state": "clear", "description": "No screenshot", "can_auto_handle": False,
                "auto_selector": None, "needs_user": False, "user_message": None}
    try:
        prompt = """You are analyzing a browser screenshot to detect blocking states.

Identify the state and return ONLY valid JSON (no markdown, no extra text):

States:
- clear          : Page loaded fine, no blockers
- cookie_consent : Cookie / privacy banner with Accept/Allow/OK button visible
- confirm_button : A Continue / Next / Got it / Dismiss / OK popup blocking the page
- verify_email   : Page says \"check your email\", \"verify your email\", \"confirmation sent\"
- 2fa            : 2-factor auth / OTP / authenticator code entry required
- captcha        : reCAPTCHA or image captcha is shown
- error          : Login failed / wrong password / account issue / blocking error
- loading        : Page is still loading or mid-redirect

Return exactly:
{\n  \"state\": \"<one of the above>\",\n  \"description\": \"One sentence of what you see\",\n  \"can_auto_handle\": true or false,\n  \"auto_selector\": \"CSS selector to click or null\",\n  \"needs_user\": true or false,\n  \"user_message\": \"Message to show user or null\"\n}

Rules:
- cookie_consent / confirm_button: can_auto_handle=true, provide the most specific CSS selector for the button
- verify_email: needs_user=true, user_message=\"Check your email inbox and click the verification link, then come back.\"
- 2fa: needs_user=true, user_message=\"Please enter the 2FA/OTP code shown in your authenticator app or SMS.\"
- captcha: needs_user=true, user_message=\"A CAPTCHA is shown. Please solve it manually in the browser.\"
- error: needs_user=true, describe the error in user_message
- clear / loading: can_auto_handle=false, needs_user=false"""

        if HAS_GEMINI and GEMINI_API_KEY:
            import google.generativeai as genai
            genai.configure(api_key=GEMINI_API_KEY)
            model = genai.GenerativeModel(GEMINI_MODEL)
            img_data = base64.b64decode(screenshot)
            img_part = {"mime_type": "image/jpeg", "data": img_data}
            resp = model.generate_content([prompt, img_part])
            raw = resp.text.strip()
            # Strip markdown code fences if present
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"): raw = raw[4:]
            import json as _json
            return _json.loads(raw.strip())
        # Fallback: return clear if no vision
        return {"state": "clear", "description": "Vision not available",
                "can_auto_handle": False, "auto_selector": None,
                "needs_user": False, "user_message": None}
    except Exception as e:
        logger.warning(f"check_browser_state error: {e}")
        return {"state": "clear", "description": f"Check failed: {str(e)}",
                "can_auto_handle": False, "auto_selector": None,
                "needs_user": False, "user_message": None}


@app.get("/capabilities")
async def capabilities():
    providers = {
        "ollama":       {"cost": "FREE (local)", "needs_key": False, "model": OLLAMA_MODEL,
                         "setup": "ollama pull llama3"},
        "groq":         {"cost": "FREE tier",    "needs_key": True,  "model": GROQ_MODEL,
                         "key_url": "https://console.groq.com"},
        "huggingface":  {"cost": "FREE tier",    "needs_key": True,  "model": HF_MODEL,
                         "key_url": "https://huggingface.co/settings/tokens"},
        "transformers": {"cost": "FREE (GPU)",   "needs_key": False, "model": HF_LOCAL_MODEL,
                         "setup": "uses local GPU"},
        "openai":       {"cost": "PAID",         "needs_key": True,  "model": OPENAI_MODEL,
                         "key_url": "https://platform.openai.com/api-keys"},
    }
    return {
        "version": "3.0.0",
        "active_provider": LLM_PROVIDER,
        "active_model": providers.get(LLM_PROVIDER, {}).get("model", "?"),
        "all_providers": providers,
        "features": [
            "Natural language → action planning",
            "Multi-turn retry with automatic fallback approaches",
            "Vision screen analysis (Ollama llava / GPT-4o / Qwen2-VL)",
            "Task completion verification",
            "50+ Windows OS operations",
            "Playwright browser automation",
            "MS Office COM automation (Word, Excel, PowerPoint)",
            "Full file system control",
            "System management: volume, WiFi, Bluetooth, power, registry…",
        ],
    }


# ─────────────────────────────────────────────────────────────────────────────
# CLI entrypoint
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Pecifics AI Backend")
    parser.add_argument("--host",     default="0.0.0.0")
    parser.add_argument("--port",     type=int, default=8000)
    parser.add_argument("--ngrok",    action="store_true", help="Expose via ngrok tunnel")
    parser.add_argument("--provider", default=None,
                        help="Override LLM_PROVIDER (ollama|groq|huggingface|transformers|openai)")
    args = parser.parse_args()

    if args.provider:
        os.environ["LLM_PROVIDER"] = args.provider
        LLM_PROVIDER = args.provider  # update module-level var

    if args.ngrok:
        try:
            from pyngrok import ngrok
            token = os.getenv("NGROK_TOKEN", "")
            if token: ngrok.set_auth_token(token)
            tunnel = ngrok.connect(args.port)
            url = tunnel.public_url
            logger.info("┌───────────────────────────────────────────────────┐")
            logger.info(f"│  Pecifics Backend  ►  {url:<30}│")
            logger.info("│  Paste this URL into Pecifics Settings            │")
            logger.info("└───────────────────────────────────────────────────┘")
        except Exception as e:
            logger.warning(f"ngrok failed: {e}")

    model_label = {
        "ollama": OLLAMA_MODEL, "groq": GROQ_MODEL, "huggingface": HF_MODEL,
        "transformers": HF_LOCAL_MODEL, "openai": OPENAI_MODEL,
    }.get(LLM_PROVIDER, "?")
    logger.info(f"Provider: {LLM_PROVIDER}  |  Model: {model_label}")
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")
