import os
import logging
import base64
import time
import io
from collections import defaultdict
from openai import OpenAI  # المكتبة الجديدة
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from PIL import Image
from pypdf import PdfReader

# الإعدادات من البيئة
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

# إعدادات الحدود
MAX_PER_USER_PER_DAY = 15
MAX_PER_MINUTE_GLOBAL = 18

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# إدارة الاستهلاك
user_usage = defaultdict(lambda: {"count": 0, "reset_time": time.time() + 86400})
global_minute_usage = {"count": 0, "reset_time": time.time() + 60}

# تهيئة عميل GitHub Models
client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=GITHUB_TOKEN,
)

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
                    if line: slide_texts.append(line)
            if shape.shape_type == 13: # Image
                try:
                    img = Image.open(io.BytesIO(shape.image.blob)).convert("RGB")
                    if max(img.size) > 800:
                        ratio = 800 / max(img.size)
                        img = img.resize((int(img.size[0]*ratio), int(img.size[1]*ratio)), Image.LANCZOS)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    images_b64.append((slide_num, base64.b64encode(buf.getvalue()).decode("utf-8")))
                except Exception as e:
                    logger.warning(f"فشل صورة شريحة {slide_num}: {e}")
        if slide_texts:
            text_parts.append(f"=== Slide {slide_num} ===\n" + "\n".join(slide_texts))
    return "\n\n".join(text_parts), images_b64

CODE_PROMPT = (
    "Extract all programming code from the following content.\n"
    "For each code block:\n"
    "1. Write the language name\n"
    "2. Write the complete code inside a code block\n"
    "3. Write a very short description\n\n"
    "If there is no code, say: No programming code found."
)

def call_github_models(prompt_content):
    """الاتصال بـ GitHub Copilot Models"""
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are a professional code extractor."},
                {"role": "user", "content": prompt_content}
            ],
            model="gpt-4o-mini",
            temperature=0.1
        )
        return response.choices[0].message.content
    except Exception as e:
        logger.error(f"GitHub API Error: {e}")
        raise Exception(f"فشل الاتصال بـ GitHub Models: {str(e)}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً بك في بوت مستخرج الكود (نسخة Copilot)!\n\n"
        "📄 ابعتلي PDF أو PPTX وهستخرجلك الكود بجودة عالية.\n"
        f"⚠️ الحد اليومي: {MAX_PER_USER_PER_DAY} ملفات."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    user_id = message.from_user.id
    doc = message.document
    file_name = doc.file_name.lower()

    is_pdf = file_name.endswith(".pdf") or doc.mime_type == "application/pdf"
    is_pptx = file_name.endswith(".pptx") or "presentationml" in (doc.mime_type or "")

    if not is_pdf and not is_pptx:
        await message.reply_text("⚠️ نعتذر، ارسل ملفات PDF أو PPTX فقط.")
        return

    allowed, remaining = check_user_rate(user_id)
    if not allowed:
        await message.reply_text("❌ وصلت للحد اليومي!")
        return

    if not check_global_rate():
        await message.reply_text("⏳ ضغط مؤقت، جرب كمان دقيقة.")
        return

    await message.reply_text(f"⏳ جاري معالجة {file_name} عبر GPT-4o-mini...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())

        if is_pdf:
            extracted_text = extract_text_from_pdf(file_bytes)
            if not extracted_text.strip():
                await message.reply_text("⚠️ لم يتم العثور على نص في الـ PDF.")
                return
            final_prompt = f"{CODE_PROMPT}\n\nContent:\n{extracted_text}"
        else:
            extracted_text, images = extract_pptx_content(file_bytes)
            # GitHub Models تدعم الصور في الـ vision ولكن للتبسيط سنرسل النص 
            # وإذا أردت دعم الصور (Vision) يتم تعديل الـ prompt_content ليكون List
            final_prompt = f"{CODE_PROMPT}\n\nContent:\n{extracted_text}"

        result_text = call_github_models(final_prompt)

        # تقسيم الرسالة إذا كانت طويلة (تليجرام حدوده 4096 حرف)
        if len(result_text) <= 4000:
            await message.reply_text(result_text, parse_mode="Markdown")
        else:
            for i in range(0, len(result_text), 4000):
                await message.reply_text(result_text[i:i+4000])

    except Exception as e:
        logger.error(f"Error: {e}")
        await message.reply_text(f"❌ خطأ فني: {str(e)}")

def main():
    if not TELEGRAM_TOKEN or not GITHUB_TOKEN:
        print("❌ تأكد من ضبط TELEGRAM_TOKEN و GITHUB_TOKEN في الإعدادات!")
        return

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("✅ البوت شغال الآن باستخدام GitHub Models...")
    app.run_polling()

if __name__ == "__main__":
    main()
