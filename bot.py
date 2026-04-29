import os
import logging
import time
import io
import asyncio
from collections import defaultdict
from openai import OpenAI
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from pypdf import PdfReader

# --- الإعدادات ---
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

MAX_PER_USER_PER_DAY = 30
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

user_usage = defaultdict(lambda: {"count": 0, "reset_time": time.time() + 86400})
processing_lock = asyncio.Lock()

client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=GITHUB_TOKEN,
)

# --- وظائف الاستخراج ---
def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text_parts = []
    for page in reader.pages:
        text = page.extract_text()
        if text: text_parts.append(text.strip())
    return "\n".join(text_parts)

def extract_pptx_content(pptx_bytes: bytes):
    prs = Presentation(io.BytesIO(pptx_bytes))
    text_parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text_parts.append(shape.text.strip())
    return "\n".join(text_parts)

# --- معالجة الرسائل ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 ابعت ملفات PDF أو PPTX وهستخرجلك الكود فوراً!")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    doc = message.document
    file_name = doc.file_name.lower()

    if doc.file_size > 20 * 1024 * 1024:
        await message.reply_text(f"❌ الملف '{file_name}' كبير جداً (أقصى حد 20MB).")
        return

    async with processing_lock:
        status_msg = await message.reply_text(f"⚙️ جاري معالجة '{file_name}'...")
        try:
            file = await context.bot.get_file(doc.file_id)
            file_bytes = bytes(await file.download_as_bytearray())

            if file_name.endswith(".pdf"):
                extracted_content = extract_text_from_pdf(file_bytes)
            else:
                extracted_content = extract_pptx_content(file_bytes)

            if not extracted_content.strip():
                await status_msg.edit_text("⚠️ لم أجد نصاً داخل الملف.")
                return

            # طلب الـ AI مع تنبيهه ينسق الكود بشكل بسيط
            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": "Extract programming code. Use simple plain text. No complex markdown symbols."},
                    {"role": "user", "content": extracted_content[:15000]}
                ],
                model="gpt-4o-mini",
            )
            result = response.choices[0].message.content

            # الحل النهائي: نبعت الرسالة بدون أي Parse Mode تماماً
            # كدا تليجرام هيتعامل مع الكلام كأنه نص عادي ومستحيل يديك خطأ Entities
            header = f"✅ كود ملف: {file_name}\n" + ("-"*20) + "\n\n"
            full_response = header + result

            if len(full_response) <= 4000:
                await message.reply_text(full_response) # مفيش parse_mode هنا
            else:
                for i in range(0, len(full_response), 4000):
                    await message.reply_text(full_response[i:i+4000])

        except Exception as e:
            logger.error(f"Error: {e}")
            # في حالة الفشل بنبعت الخطأ كرسالة عادية
            await message.reply_text(f"❌ خطأ في '{file_name}': {str(e)}")

def main():
    if not TELEGRAM_TOKEN or not GITHUB_TOKEN:
        print("Missing Tokens!")
        return
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()
