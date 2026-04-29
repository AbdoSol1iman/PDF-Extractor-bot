import os
import logging
import base64
import time
import io
from collections import defaultdict
import requests
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation

# ===== التوكنات من Environment Variables =====
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")

# ===== حدود الاستخدام =====
MAX_PER_USER_PER_DAY = 15
MAX_PER_MINUTE_GLOBAL = 18  # الحد الفعلي 20، خلينا هامش أمان

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "meta-llama/llama-3.3-70b-instruct:free"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===== تتبع الاستخدام =====
user_usage = defaultdict(lambda: {"count": 0, "reset_time": time.time() + 86400})
global_minute_usage = {"count": 0, "reset_time": time.time() + 60}


def check_global_rate() -> bool:
    now = time.time()
    if now > global_minute_usage["reset_time"]:
        global_minute_usage["count"] = 0
        global_minute_usage["reset_time"] = now + 60
    if global_minute_usage["count"] >= MAX_PER_MINUTE_GLOBAL:
        return False
    global_minute_usage["count"] += 1
    return True


def check_user_rate(user_id: int):
    now = time.time()
    user = user_usage[user_id]
    if now > user["reset_time"]:
        user["count"] = 0
        user["reset_time"] = now + 86400
    remaining = MAX_PER_USER_PER_DAY - user["count"]
    if remaining <= 0:
        return False, 0
    user["count"] += 1
    return True, remaining - 1


def extract_text_from_pptx(pptx_bytes: bytes) -> str:
    prs = Presentation(io.BytesIO(pptx_bytes))
    text_parts = []
    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = "".join(run.text for run in para.runs).strip()
                    if line:
                        slide_texts.append(line)
        if slide_texts:
            text_parts.append(f"=== Slide {slide_num} ===\n" + "\n".join(slide_texts))
    return "\n\n".join(text_parts)


def ask_openrouter_with_pdf(pdf_base64: str) -> str:
    """بيبعت الـ PDF لـ OpenRouter باستخدام cloudflare-ai engine المجاني"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": MODEL,
        "plugins": [{"id": "file-parser", "pdf": {"engine": "cloudflare-ai"}}],
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "file",
                        "file": {
                            "filename": "document.pdf",
                            "content_type": "application/pdf",
                            "data": pdf_base64,
                        }
                    },
                    {
                        "type": "text",
                        "text": (
                            "Extract all programming code from this file.\n"
                            "For each code block:\n"
                            "1. Write the language name\n"
                            "2. Write the complete code inside a code block\n"
                            "3. Write a very short description\n\n"
                            "If there is no code, say: No programming code found in this file."
                        )
                    }
                ]
            }
        ]
    }
    response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]


def ask_openrouter_with_text(text: str) -> str:
    """بيبعت النص لـ OpenRouter (للـ PPTX)"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": MODEL,
        "messages": [
            {
                "role": "user",
                "content": (
                    f"The following content is extracted from a PowerPoint file:\n\n{text}\n\n"
                    "Extract all programming code found.\n"
                    "For each code block:\n"
                    "1. Write the language name\n"
                    "2. Write the complete code inside a code block\n"
                    "3. Write a very short description\n\n"
                    "If there is no code, say: No programming code found in this file."
                )
            }
        ]
    }
    response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً!\n\n"
        "📄 ابعتلي PDF أو PPTX وأنا هستخرجلك كل الكود اللي فيه.\n\n"
        f"⚠️ الحد المسموح: {MAX_PER_USER_PER_DAY} ملفات لكل شخص في اليوم.\n\n"
        "يدعم: Python, Java, C++, JavaScript, وأي لغة تانية."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    user_id = message.from_user.id
    doc = message.document
    file_name = doc.file_name.lower()

    is_pdf = file_name.endswith(".pdf") or doc.mime_type == "application/pdf"
    is_pptx = file_name.endswith(".pptx") or "presentationml" in (doc.mime_type or "")

    if not is_pdf and not is_pptx:
        await message.reply_text("⚠️ بس بدعم PDF و PPTX!")
        return

    allowed, remaining = check_user_rate(user_id)
    if not allowed:
        await message.reply_text("❌ وصلت للحد اليومي بتاعك!\n⏳ هيتجدد الحد بكره.")
        return

    if not check_global_rate():
        await message.reply_text("⏳ البوت مشغول دلوقتي، جرب تاني بعد دقيقة!")
        user_usage[user_id]["count"] -= 1
        return

    file_type = "PDF" if is_pdf else "PPTX"
    await message.reply_text(
        f"⏳ جاري استخراج الكود من الـ {file_type}... انتظر لحظة.\n"
        f"📊 باقيلك {remaining} ملفات النهارده."
    )

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())

        if is_pdf:
            pdf_base64 = base64.b64encode(file_bytes).decode("utf-8")
            result_text = ask_openrouter_with_pdf(pdf_base64)
        else:
            pptx_text = extract_text_from_pptx(file_bytes)
            if not pptx_text.strip():
                await message.reply_text("⚠️ الملف فاضي أو مش فيه نصوص!")
                user_usage[user_id]["count"] -= 1
                return
            result_text = ask_openrouter_with_text(pptx_text)

        if len(result_text) <= 4000:
            await message.reply_text(result_text, parse_mode="Markdown")
        else:
            chunks = split_text(result_text, 4000)
            for i, chunk in enumerate(chunks):
                header = f"📦 *الجزء {i+1}/{len(chunks)}*\n\n" if len(chunks) > 1 else ""
                await message.reply_text(header + chunk, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"Error: {e}")
        await message.reply_text(f"❌ حصل خطأ: {str(e)}")
        user_usage[user_id]["count"] -= 1


def split_text(text: str, max_length: int) -> list:
    chunks = []
    while len(text) > max_length:
        split_at = text.rfind("\n", 0, max_length)
        if split_at == -1:
            split_at = max_length
        chunks.append(text[:split_at])
        text = text[split_at:].lstrip("\n")
    if text:
        chunks.append(text)
    return chunks


async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📄 ابعتلي ملف PDF أو PPTX وأنا هستخرجلك الكود!")


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("✅ البوت شغال...")
    app.run_polling()


if __name__ == "__main__":
    main()
