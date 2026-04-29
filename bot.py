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
from pptx.util import Inches
from PIL import Image

# ===== التوكنات من Environment Variables =====
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")

# ===== حدود الاستخدام =====
MAX_PER_USER_PER_DAY = 15
MAX_PER_MINUTE_GLOBAL = 18

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "openrouter/free"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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


def extract_pptx_content(pptx_bytes: bytes):
    """
    بيستخرج النصوص والصور من الـ PPTX
    بيرجع: (نص كامل, قايمة صور base64)
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    text_parts = []
    images_b64 = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_texts = []
        for shape in slide.shapes:
            # استخراج النصوص
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = "".join(run.text for run in para.runs).strip()
                    if line:
                        slide_texts.append(line)

            # استخراج الصور
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                try:
                    img_bytes = shape.image.blob
                    # تحويل لـ PNG وتصغير لو كبيرة
                    img = Image.open(io.BytesIO(img_bytes))
                    img = img.convert("RGB")
                    # تصغير الصورة لو أكبر من 1000px
                    max_size = 1000
                    if max(img.size) > max_size:
                        ratio = max_size / max(img.size)
                        new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                        img = img.resize(new_size, Image.LANCZOS)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
                    images_b64.append((slide_num, img_b64))
                except Exception as e:
                    logger.warning(f"فشل استخراج صورة من slide {slide_num}: {e}")

        if slide_texts:
            text_parts.append(f"=== Slide {slide_num} ===\n" + "\n".join(slide_texts))

    return "\n\n".join(text_parts), images_b64


def call_openrouter(payload: dict, retries: int = 3) -> str:
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    for attempt in range(retries):
        response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
        if response.status_code == 429:
            wait = 15 * (attempt + 1)
            logger.warning(f"429 rate limit، بستنى {wait} ثانية...")
            time.sleep(wait)
            continue
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    raise Exception("الـ API مشغول دلوقتي، جرب تاني بعد شوية.")


def ask_openrouter_with_pdf(pdf_base64: str) -> str:
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
    return call_openrouter(payload)


def ask_openrouter_with_pptx(text: str, images_b64: list) -> str:
    """بيبعت النص + الصور لـ OpenRouter"""
    content = []

    # أضف النص الأول
    prompt = (
        "Extract all programming code from the following PowerPoint content.\n"
        "The content includes both extracted text and images from the slides.\n"
        "For each code block found (in text OR in images):\n"
        "1. Write the language name\n"
        "2. Write the complete code inside a code block\n"
        "3. Write a very short description\n\n"
        "If there is no code, say: No programming code found in this file."
    )

    if text.strip():
        prompt += f"\n\n=== Extracted Text ===\n{text}"

    content.append({"type": "text", "text": prompt})

    # أضف الصور
    for slide_num, img_b64 in images_b64:
        content.append({
            "type": "text",
            "text": f"=== Image from Slide {slide_num} ==="
        })
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{img_b64}"
            }
        })

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": content}]
    }
    return call_openrouter(payload)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً!\n\n"
        "📄 ابعتلي PDF أو PPTX وأنا هستخرجلك كل الكود اللي فيه.\n"
        "🖼️ بيدعم الكود اللي في الصور جوا الـ PPTX كمان!\n\n"
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
            text, images_b64 = extract_pptx_content(file_bytes)
            if not text.strip() and not images_b64:
                await message.reply_text("⚠️ الملف فاضي أو مش فيه محتوى!")
                user_usage[user_id]["count"] -= 1
                return
            img_count = len(images_b64)
            if img_count > 0:
                await message.reply_text(f"🖼️ لقيت {img_count} صورة في الـ PPTX، بحللهم...")
            result_text = ask_openrouter_with_pptx(text, images_b64)

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
