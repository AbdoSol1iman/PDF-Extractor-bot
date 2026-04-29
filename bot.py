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
from PIL import Image
from pypdf import PdfReader

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")

MAX_PER_USER_PER_DAY = 15
MAX_PER_MINUTE_GLOBAL = 18
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# قايمة موديلات نصية — لو الأول مشغول بنجرب التاني تلقائياً
TEXT_MODELS = [
    "meta-llama/llama-3.3-70b-instruct:free",
    "mistralai/mistral-small-3.1-24b-instruct:free",
    "google/gemma-4-31b-it:free",
    "deepseek/deepseek-r1:free",
]

# قايمة موديلات vision للصور
VISION_MODELS = [
    "google/gemma-4-31b-it:free",
    "google/gemma-4-26b-a4b-it:free",
]

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


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text_parts = []
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text()
        if text and text.strip():
            text_parts.append(f"=== Page {i} ===\n{text.strip()}")
    return "\n\n".join(text_parts)


def extract_pptx_content(pptx_bytes: bytes):
    prs = Presentation(io.BytesIO(pptx_bytes))
    text_parts = []
    images_b64 = []
    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = "".join(run.text for run in para.runs).strip()
                    if line:
                        slide_texts.append(line)
            if shape.shape_type == 13:
                try:
                    img = Image.open(io.BytesIO(shape.image.blob)).convert("RGB")
                    if max(img.size) > 800:
                        ratio = 800 / max(img.size)
                        img = img.resize((int(img.size[0]*ratio), int(img.size[1]*ratio)), Image.LANCZOS)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    images_b64.append((slide_num, base64.b64encode(buf.getvalue()).decode("utf-8")))
                except Exception as e:
                    logger.warning(f"فشل استخراج صورة slide {slide_num}: {e}")
        if slide_texts:
            text_parts.append(f"=== Slide {slide_num} ===\n" + "\n".join(slide_texts))
    return "\n\n".join(text_parts), images_b64


def call_openrouter_with_fallback(models: list, payload_fn) -> str:
    """
    بيجرب الموديلات واحد واحد — لو الأول بيجيب 429 أو error بيجرب التاني
    payload_fn: function بتاخد اسم الموديل وبترجع الـ payload
    """
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    last_error = None
    for model in models:
        payload = payload_fn(model)
        for attempt in range(2):  # محاولتين لكل موديل
            try:
                logger.info(f"بجرب موديل: {model}")
                response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=90)
                if response.status_code == 429:
                    wait = 10 * (attempt + 1)
                    logger.warning(f"429 على {model}، بستنى {wait}s...")
                    time.sleep(wait)
                    continue
                if response.status_code in (404, 501, 400):
                    logger.warning(f"موديل {model} مش متاح: {response.status_code}")
                    break  # جرب الموديل الجاي
                response.raise_for_status()
                return response.json()["choices"][0]["message"]["content"]
            except Exception as e:
                last_error = e
                logger.warning(f"خطأ مع {model}: {e}")
                break
    raise Exception(f"كل الموديلات المتاحة فشلت. آخر خطأ: {last_error}")


CODE_PROMPT = (
    "Extract all programming code from the following content.\n"
    "For each code block:\n"
    "1. Write the language name\n"
    "2. Write the complete code inside a code block\n"
    "3. Write a very short description\n\n"
    "If there is no code, say: No programming code found."
)


def ask_with_text(text: str) -> str:
    def make_payload(model):
        return {
            "model": model,
            "messages": [{"role": "user", "content": f"{CODE_PROMPT}\n\n{text}"}]
        }
    return call_openrouter_with_fallback(TEXT_MODELS, make_payload)


def ask_with_text_and_images(text: str, images_b64: list) -> str:
    def make_payload(model):
        content = []
        full_prompt = CODE_PROMPT
        if text.strip():
            full_prompt += f"\n\n=== Extracted Text ===\n{text}"
        content.append({"type": "text", "text": full_prompt})
        for slide_num, img_b64 in images_b64[:10]:
            content.append({"type": "text", "text": f"--- Image from Slide {slide_num} ---"})
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_b64}"}
            })
        return {"model": model, "messages": [{"role": "user", "content": content}]}
    return call_openrouter_with_fallback(VISION_MODELS, make_payload)


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
            text = extract_text_from_pdf(file_bytes)
            if not text.strip():
                await message.reply_text("⚠️ الـ PDF مش فيه نص قابل للاستخراج.")
                user_usage[user_id]["count"] -= 1
                return
            result_text = ask_with_text(text)
        else:
            text, images_b64 = extract_pptx_content(file_bytes)
            if not text.strip() and not images_b64:
                await message.reply_text("⚠️ الملف فاضي!")
                user_usage[user_id]["count"] -= 1
                return
            if images_b64:
                await message.reply_text(f"🖼️ لقيت {len(images_b64)} صورة، بحلل أول 10...")
            result_text = ask_with_text_and_images(text, images_b64)

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
