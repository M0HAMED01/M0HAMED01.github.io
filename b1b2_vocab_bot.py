# b1b2_vocab_bot.py — local persistent state, fixed paths, 24h time
# Build with: pyinstaller --onefile --noconsole b1b2_vocab_bot.py

import json
import shutil
from pathlib import Path
from datetime import time
from typing import Optional

import pandas as pd

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

# ========= FIXED SETTINGS (as requested) =========
BOT_TOKEN     = "8057373800:AAGdyoVer8-cayMeFuxqevpnrXm-TzbdMCo"
TARGET_CHAT_ID = 1150807213
XLSX_PATH     = Path(r"E:\courses\German\bot\german_B1_to_B2_vocab.xlsx")
STATE_FILE    = Path(r"E:\courses\German\bot\progress_state.json")  # <- human-readable, durable
SLANG_SHEET   = "Slang_Colloquial"

DEFAULT_BATCH_SIZE = 10
DEFAULT_HOUR = 9   # 09:00 (24h)
DEFAULT_MIN  = 0
# ================================================

# ---------- state helpers (robust & atomic) ----------
def _safe_read_json(p: Path, default):
    try:
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        # Backup the bad/corrupt file once
        try:
            shutil.copyfile(p, p.with_suffix(".json.bak"))
        except Exception:
            pass
    return default


def _safe_write_json(p: Path, data):
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)


def read_state():
    default = {
        "index": 0,                    # next vocab row to send
        "batch_size": DEFAULT_BATCH_SIZE,
        "paused": False,
        "hour": DEFAULT_HOUR,
        "minute": DEFAULT_MIN
    }
    st = _safe_read_json(STATE_FILE, default)
    for k, v in default.items():
        st.setdefault(k, v)
    return st


def write_state(st):
    _safe_write_json(STATE_FILE, st)


# ---------- data loading ----------
def load_vocab(xlsx_path: Path):
    """Load all sheets except slang; return list of dict rows."""
    xl = pd.ExcelFile(xlsx_path)
    sheets = [s for s in xl.sheet_names if s != SLANG_SHEET]
    rows = []
    for sheet in sheets:
        df = pd.read_excel(xlsx_path, sheet_name=sheet)
        # Expected columns in your XLSX:
        # German | English | Example (DE) | Small-talk prompt (DE)
        for _, r in df.iterrows():
            rows.append({
                "theme": sheet,
                "german": str(r["German"]),
                "english": str(r["English"]),
                "example": str(r["Example (DE)"]),
                "prompt": str(r.get("Small-talk prompt (DE)", "")) if "Small-talk prompt (DE)" in df.columns else "",
            })
    return rows


def load_slang(xlsx_path: Path):
    if SLANG_SHEET not in pd.ExcelFile(xlsx_path).sheet_names:
        return []
    df = pd.read_excel(xlsx_path, sheet_name=SLANG_SHEET)
    out = []
    for _, r in df.iterrows():
        out.append({
            "expression": str(r["Expression (DE)"]),
            "english": str(r["English"]),
            "example": str(r["Example (DE)"]),
        })
    return out


# ---------- formatting ----------
def _format_row(r: dict) -> str:
    p = r.get("prompt", "")
    pmsg = f"\n   💬 {p}" if p and str(p).strip() and p != "nan" else ""
    return (
        f"• <b>{r['german']}</b> — <i>{r['english']}</i>\n"
        f"   📝 {r['example']}{pmsg}\n"
        f"   <code>{r['theme']}</code>"
    )


def _render_batch(batch) -> str:
    return "<b>Heutige Vokabeln</b>\n\n" + "\n\n".join(_format_row(r) for r in batch)


async def _send_batch(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    batch,
    *,
    prefix: Optional[str] = None,
):
    if not batch:
        await context.bot.send_message(chat_id=chat_id, text="🎉 Keine neuen Wörter verfügbar.")
        return
    text = _render_batch(batch)
    if prefix:
        text = f"{prefix}\n\n{text}"
    await context.bot.send_message(
        chat_id=chat_id,
        text=text,
        parse_mode=ParseMode.HTML,
        disable_web_page_preview=True
    )


# ---------- daily job ----------
async def daily_job(context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    if st.get("paused"):
        return
    vocab = load_vocab(XLSX_PATH)
    if not vocab:
        await context.bot.send_message(
            chat_id=TARGET_CHAT_ID,
            text="📭 In der Vokabelliste wurden keine Einträge gefunden."
        )
        return
    idx = st["index"]
    size = st.get("batch_size", DEFAULT_BATCH_SIZE)
    reset_message = None
    if idx >= len(vocab):
        idx = 0
        st["index"] = 0
        reset_message = "🔄 Alle Wörter wurden bereits versendet – beginne wieder von vorne."
    batch = vocab[idx: idx + size]
    st["index"] = idx + len(batch)
    write_state(st)  # persist immediately
    await _send_batch(context, TARGET_CHAT_ID, batch, prefix=reset_message)


def _reschedule_daily(app, hour: int, minute: int):
    # remove any existing schedule
    for job in app.job_queue.get_jobs_by_name("daily_vocab"):
        job.schedule_removal()
    # uses Windows system timezone; 24h time
    app.job_queue.run_daily(daily_job, time=time(hour=hour, minute=minute), name="daily_vocab")


# ---------- commands ----------
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    await update.message.reply_text(
        f"⏱️ Time (24h): {st['hour']:02d}:{st['minute']:02d}\n"
        f"📦 Batch: {st['batch_size']}\n"
        f"⏸️ Paused: {st['paused']}\n"
        f"💾 State: {STATE_FILE}"
    )


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    vocab_len = len(load_vocab(XLSX_PATH))
    await update.message.reply_text(
        f"📊 Index: {st['index']} / {vocab_len}\n"
        f"📦 Batch: {st['batch_size']}\n"
        f"⏱️ Time (24h): {st['hour']:02d}:{st['minute']:02d}\n"
        f"⏸️ Paused: {st['paused']}\n"
        f"💾 State file: {STATE_FILE}"
    )


async def cmd_next(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    vocab = load_vocab(XLSX_PATH)
    if not vocab:
        return await update.message.reply_text("📭 Keine Vokabeln zum Versenden gefunden.")
    idx = st["index"]
    size = st.get("batch_size", DEFAULT_BATCH_SIZE)
    reset_message = None
    if idx >= len(vocab):
        idx = 0
        st["index"] = 0
        reset_message = "🔄 Alle Wörter wurden bereits versendet – starte wieder am Anfang."
    batch = vocab[idx: idx + size]
    st["index"] = idx + len(batch)
    write_state(st)
    await _send_batch(context, update.effective_chat.id, batch, prefix=reset_message)


async def cmd_setbatch(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        return await update.message.reply_text("Usage: /setbatch 8")
    try:
        n = int(context.args[0])
        assert 1 <= n <= 100
    except Exception:
        return await update.message.reply_text("Enter a number 1–100.")
    st = read_state()
    st["batch_size"] = n
    write_state(st)
    await update.message.reply_text(f"✅ Batch set to {n}.")


async def cmd_settime(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # /settime HH:MM (24h)
    parts = update.message.text.strip().split()
    if len(parts) < 2:
        return await update.message.reply_text("Usage: /settime HH:MM (24h). Example: /settime 09:00")
    try:
        hh, mm = parts[1].split(":")
        hh, mm = int(hh), int(mm)
        assert 0 <= hh < 24 and 0 <= mm < 60
    except Exception:
        return await update.message.reply_text("Use HH:MM in 24h format. Example: /settime 15:30")

    st = read_state()
    st["hour"] = hh
    st["minute"] = mm
    write_state(st)                       # save first
    _reschedule_daily(context.application, hh, mm)  # then reschedule
    await update.message.reply_text(f"⏰ Time set to {hh:02d}:{mm:02d} (24h). Saved to {STATE_FILE.name}")


async def cmd_pause(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    st["paused"] = True
    write_state(st)
    await update.message.reply_text("⏸️ Paused (saved).")


async def cmd_resume(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = read_state()
    st["paused"] = False
    write_state(st)
    await update.message.reply_text("▶️ Resumed (saved).")


async def cmd_slang(update: Update, context: ContextTypes.DEFAULT_TYPE):
    slang = load_slang(XLSX_PATH)
    if not slang:
        return await update.message.reply_text("No slang sheet found.")
    from random import choice
    r = choice(slang)
    await update.message.reply_text(
        f"🗯️ <b>{r['expression']}</b> — <i>{r['english']}</i>\n   📝 {r['example']}",
        parse_mode=ParseMode.HTML
    )


# ---------- main ----------
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("status",   cmd_status))
    app.add_handler(CommandHandler("next",     cmd_next))
    app.add_handler(CommandHandler("setbatch", cmd_setbatch))
    app.add_handler(CommandHandler("settime",  cmd_settime))
    app.add_handler(CommandHandler("pause",    cmd_pause))
    app.add_handler(CommandHandler("resume",   cmd_resume))
    app.add_handler(CommandHandler("slang",    cmd_slang))

    st = read_state()
    _reschedule_daily(app, st["hour"], st["minute"])  # schedule from saved state
    app.run_polling()  # blocking; background if packaged with --noconsole


if __name__ == "__main__":
    main()
