import os
import logging
import base64
import time
import io
import asyncio  # مكتبة التعامل مع العمليات المتزامنة
from collections import defaultdict
from openai import OpenAI
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from PIL import Image
from pypdf import PdfReader

# --- الإعدادات ---
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

MAX_PER_USER_PER_DAY = 15
MAX_PER_MINUTE_GLOBAL = 18

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# إدارة الاستهلاك والقفل العالمي
user_usage = defaultdict(lambda: {"count": 0, "reset_time": time.time() + 86400})
global_minute_usage = {"count": 0, "reset_time": time.time() + 60}
# هذا القفل يضمن معالجة ملف واحد فقط في كل مرة على مستوى البوت بالكامل
processing_lock = asyncio.Lock()

# تهيئة عميل GitHub Models
client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=GITHUB_TOKEN,
)

# --- وظائف المساعدة ---

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
    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = "".join(run.text for run in para.runs).strip()
                    if line: slide_texts.append(line)
        if slide_texts:
            text_parts.append(f"=== Slide {slide_num} ===\n" + "\n".join(slide_texts))
    return "\n\n".join(text_parts)

CODE_PROMPT = (
    "Extract all programming code from the following content.\n"
    "For each code block:\n"
    "1. Write the language name\n"
    "2. Write the complete code inside a code block\n"
    "3. Write a very short description\n\n"
    "If there is no code, say: No programming code found."
)

def call_github_models(prompt_content):
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
        raise Exception(f"GitHub Models Error: {str(e)}")

# --- معالجات التليجرام ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً بك! (نسخة الطابور المنظم 🚦)\n\n"
        "📄 ابعتلي ملفاتك وسأقوم بمعالجتها واحدًا تلو الآخر لضمان أفضل جودة.\n"
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
        await message.reply_text("⚠️ يرجى إرسال PDF أو PPTX فقط.")
        return

    allowed, remaining = check_user_rate(user_id)
    if not allowed:
        await message.reply_text("❌ عذراً، لقد استنفدت حدك اليومي.")
        return

    if not check_global_rate():
        await message.reply_text("⏳ البوت عليه ضغط حالياً، جرب كمان دقيقة.")
        return

    # إبلاغ المستخدم أن الملف دخل الطابور
    status_msg = await message.reply_text(f"📥 استلمت '{file_name}'.. تم وضعه في الطابور 🚦")

    # --- بداية الطابور ---
    async with processing_lock:
        try:
            await status_msg.edit_text(f"⚙️ جاري معالجة '{file_name}' الآن... يرجى الانتظار.")
            
            file = await context.bot.get_file(doc.file_id)
            file_bytes = bytes(await file.download_as_bytearray())

            if is_pdf:
                extracted_text = extract_text_from_pdf(file_bytes)
                if not extracted_text.strip():
                    await message.reply_text(f"⚠️ الملف '{file_name}' لا يحتوي على نص قابل للقراءة.")
                    return
                final_prompt = f"{CODE_PROMPT}\n\nContent:\n{extracted_text}"
            else:
                extracted_text = extract_pptx_content(file_bytes)
                final_prompt = f"{CODE_PROMPT}\n\nContent:\n{extracted_text}"

            # طلب الـ AI من GitHub Models
            result_text = call_github_models(final_prompt)

            # إرسال النتيجة مع مراعاة طول الرسالة
            if len(result_text) <= 4000:
                await message.reply_text(f"✅ كود ملف: {file_name}\n\n{result_text}", parse_mode="Markdown")
            else:
                chunks = [result_text[i:i+4000] for i in range(0, len(result_text), 4000)]
                for i, chunk in enumerate(chunks):
                    header = f"📦 كود ملف: {file_name} (جزء {i+1})\n\n"
                    await message.reply_text(header + chunk)
            
            # تأخير بسيط (ثانية) لراحة السيرفر قبل الملف القادم في الطابور
            await asyncio.sleep(1.5)

        except Exception as e:
            logger.error(f"Error processing {file_name}: {e}")
            await message.reply_text(f"❌ حصل خطأ أثناء معالجة '{file_name}':\n{str(e)}")
    # --- نهاية الطابور ---

def main():
    if not TELEGRAM_TOKEN or not GITHUB_TOKEN:
        print("❌ خطأ: لم يتم العثور على المتغيرات TELEGRAM_TOKEN أو GITHUB_TOKEN")
        return

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("✅ البوت شغال بنظام الطابور المطور...")
    app.run_polling()

if __name__ == "__main__":
    main()
