import os
import logging
import base64
import time
from collections import defaultdict
import google.generativeai as genai
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ===== التوكنات من Environment Variables =====
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# ===== حدود الاستخدام =====
MAX_PER_USER_PER_DAY = 10      # كل يوزر يقدر يبعت 10 PDFs في اليوم
MAX_PER_MINUTE_GLOBAL = 12     # البوت كله مش أكتر من 12 request/دقيقة (أمان من حد الـ 15)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

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


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً!\n\n"
        "📄 ابعتلي أي PDF وأنا هستخرجلك كل الكود اللي فيه.\n\n"
        f"⚠️ الحد المسموح: {MAX_PER_USER_PER_DAY} ملفات لكل شخص في اليوم.\n\n"
        "يدعم: Python, Java, C++, JavaScript, وأي لغة تانية."
    )


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    user_id = message.from_user.id
    doc = message.document

    if not doc.file_name.lower().endswith(".pdf") and doc.mime_type != "application/pdf":
        await message.reply_text("⚠️ ابعتلي ملف PDF بس!")
        return

    allowed, remaining = check_user_rate(user_id)
    if not allowed:
        await message.reply_text(
            "❌ وصلت للحد اليومي بتاعك!\n"
            "⏳ هيتجدد الحد بكره."
        )
        return

    if not check_global_rate():
        await message.reply_text("⏳ البوت مشغول دلوقتي، جرب تاني بعد دقيقة!")
        user_usage[user_id]["count"] -= 1
        return

    await message.reply_text(
        f"⏳ جاري استخراج الكود... انتظر لحظة.\n"
        f"📊 باقيلك {remaining} ملفات النهارده."
    )

    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        response = model.generate_content([
            {
                "inline_data": {
                    "mime_type": "application/pdf",
                    "data": pdf_base64
                }
            },
            (
                "استخرج كل الكود البرمجي الموجود في هذا الملف.\n"
                "لكل كود:\n"
                "1. اكتب اسم اللغة\n"
                "2. اكتب الكود كاملاً داخل code block\n"
                "3. اكتب وصف قصير جداً للكود\n\n"
                "إذا لم يكن هناك كود، قل: لا يوجد كود برمجي في هذا الملف."
            )
        ])

        result_text = response.text

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
    await update.message.reply_text("📄 ابعتلي ملف PDF وأنا هستخرجلك الكود!")


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("✅ البوت شغال...")
    app.run_polling()


if __name__ == "__main__":
    main()
