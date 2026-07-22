# -*- coding: utf-8 -*-
"""
موقع تحويل PDF (رياضيات/فيزياء) إلى Word بمعادلات قابلة للتعديل.

الفكرة:
1) نستخرج النص الخام من PDF (عبر PyMuPDF).
2) نرسل النص إلى Gemini API ليعيد تنظيمه كـ Markdown:
   عناوين بـ #  ، ومعادلات بصيغة LaTeX بين $...$ أو $$...$$.
3) نبني ملف Word "مرجعي" (reference.docx) نطبّق فيه اختيارات
   المستخدم (نوع/حجم/لون خط العنوان والنص).
4) نستعمل pandoc لدمج نص Markdown مع القالب المرجعي، فينتج
   ملف Word نهائي فيه معادلات حقيقية (OMML) قابلة للنقر والتعديل.
"""

import os
import re
import uuid
import shutil
import subprocess

from flask import Flask, request, render_template, send_file, jsonify
import fitz  # PyMuPDF
import requests
from docx import Document
from docx.shared import Pt, RGBColor

app = Flask(__name__)

# مجلد مؤقت لكل عملية معالجة
BASE_TMP = "/tmp/pdf2word_jobs"
os.makedirs(BASE_TMP, exist_ok=True)

# مفتاح Gemini API يُقرأ من متغير بيئة (لا تكتبه في الكود مباشرة)
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)

# حد أقصى لعدد الأحرف المرسلة لكل استدعاء (حماية من تجاوز الحصة المجانية)
MAX_CHARS_PER_CALL = 60000


def extract_text_from_pdf(pdf_path: str) -> str:
    """يستخرج النص الخام من كل صفحات PDF."""
    doc = fitz.open(pdf_path)
    pages_text = []
    for page in doc:
        pages_text.append(page.get_text("text"))
    doc.close()
    return "\n\n".join(pages_text)


def clean_with_gemini(raw_text: str) -> str:
    """
    يرسل النص الخام إلى Gemini ويطلب منه إعادة تنظيمه كـ Markdown
    مع تمييز العناوين وتحويل المعادلات إلى LaTeX صحيح.
    """
    if not GEMINI_API_KEY:
        raise RuntimeError("لم يتم ضبط GEMINI_API_KEY في متغيرات البيئة")

    prompt = (
        "أنت مساعد متخصص في تنسيق المحتوى العلمي (رياضيات/فيزياء).\n"
        "سأعطيك نصًا خامًا مستخرجًا من ملف PDF قد يحتوي أخطاء استخراج "
        "أو رموز رياضية مشوّهة.\n"
        "أعد كتابته بصيغة Markdown نظيفة وفق القواعد التالية بدقة:\n"
        "- العناوين الرئيسية تبدأ بـ '# '، والعناوين الفرعية بـ '## '.\n"
        "- كل معادلة رياضية (سواء كانت مستقلة بسطر أو داخل الجملة) "
        "يجب كتابتها بصيغة LaTeX صحيحة:\n"
        "  المعادلات المستقلة بين $$ $$ والمعادلات داخل السطر بين $ $.\n"
        "- صحّح أي رمز رياضي مشوّه أو مفقود قدر الإمكان بالاعتماد على السياق.\n"
        "- حافظ على النص العادي كما هو دون تلخيص أو حذف.\n"
        "- لا تضف أي شرح أو تعليق منك، أعد فقط نص Markdown النهائي.\n\n"
        "النص الخام:\n"
        "-----\n"
        f"{raw_text}\n"
        "-----"
    )

    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1},
    }
    headers = {"Content-Type": "application/json", "x-goog-api-key": GEMINI_API_KEY}

    resp = requests.post(GEMINI_URL, json=payload, headers=headers, timeout=180)
    resp.raise_for_status()
    data = resp.json()

    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        raise RuntimeError(f"استجابة غير متوقعة من Gemini: {data}")


def chunk_text(text: str, max_chars: int):
    """يقسّم نصًا طويلاً إلى أجزاء لا تتجاوز max_chars حرفًا (بحدود الفقرات)."""
    paragraphs = text.split("\n\n")
    chunks, current = [], ""
    for p in paragraphs:
        if len(current) + len(p) + 2 > max_chars and current:
            chunks.append(current)
            current = p
        else:
            current = current + "\n\n" + p if current else p
    if current:
        chunks.append(current)
    return chunks


def build_reference_docx(path: str, style: dict):
    """يبني ملف Word مرجعي بأنماط (Heading 1 / Normal) حسب اختيارات المستخدم."""
    doc = Document()

    h1 = doc.styles["Heading 1"]
    h1.font.name = style["title_font"]
    h1.font.size = Pt(style["title_size"])
    h1.font.color.rgb = RGBColor.from_string(style["title_color"])

    h2 = doc.styles["Heading 2"]
    h2.font.name = style["title_font"]
    h2.font.size = Pt(max(style["title_size"] - 4, 10))
    h2.font.color.rgb = RGBColor.from_string(style["title_color"])

    normal = doc.styles["Normal"]
    normal.font.name = style["text_font"]
    normal.font.size = Pt(style["text_size"])
    normal.font.color.rgb = RGBColor.from_string(style["text_color"])

    doc.save(path)


def hex_or_default(value: str, default: str = "000000") -> str:
    """ينظّف قيمة اللون القادمة من <input type=color> (تكون بصيغة #rrggbb)."""
    if not value:
        return default
    return value.lstrip("#").upper()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    job_id = str(uuid.uuid4())
    job_dir = os.path.join(BASE_TMP, job_id)
    os.makedirs(job_dir, exist_ok=True)

    try:
        uploaded = request.files.get("pdf_file")
        if not uploaded or uploaded.filename == "":
            return jsonify({"error": "الرجاء رفع ملف PDF"}), 400

        pdf_path = os.path.join(job_dir, "input.pdf")
        uploaded.save(pdf_path)

        style = {
            "title_font": request.form.get("title_font", "Arial"),
            "title_size": int(request.form.get("title_size", 20)),
            "title_color": hex_or_default(request.form.get("title_color"), "C00000"),
            "text_font": request.form.get("text_font", "Calibri"),
            "text_size": int(request.form.get("text_size", 13)),
            "text_color": hex_or_default(request.form.get("text_color"), "000000"),
        }

        # 1) استخراج النص الخام
        raw_text = extract_text_from_pdf(pdf_path)
        if not raw_text.strip():
            return jsonify({
                "error": "لم أتمكن من استخراج نص من هذا الملف. "
                         "قد يكون PDF عبارة عن صور ممسوحة ضوئيًا فقط."
            }), 400

        # 2) تنظيفه وتحويل المعادلات عبر Gemini (بأجزاء إذا كان طويلًا)
        chunks = chunk_text(raw_text, MAX_CHARS_PER_CALL)
        cleaned_parts = [clean_with_gemini(c) for c in chunks]
        markdown_text = "\n\n".join(cleaned_parts)

        md_path = os.path.join(job_dir, "content.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown_text)

        # 3) بناء القالب المرجعي بالخطوط/الألوان المختارة
        ref_path = os.path.join(job_dir, "reference.docx")
        build_reference_docx(ref_path, style)

        # 4) الدمج عبر pandoc لإنتاج Word نهائي بمعادلات حقيقية
        output_path = os.path.join(job_dir, "output.docx")
        subprocess.run(
            ["pandoc", md_path, f"--reference-doc={ref_path}", "-o", output_path],
            check=True,
            capture_output=True,
            text=True,
        )

        download_name = os.path.splitext(uploaded.filename)[0] + "_محرر.docx"
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"فشل توليد ملف Word: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # تنظيف الملفات المؤقتة (بعد الإرسال؛ Flask يرسل الملف قبل هذا السطر
        # في حالة النجاح لأن send_file يقرأه بالكامل قبل الرجوع هنا غالبًا،
        # لذا نترك حذفًا مؤجلًا بسيطًا بدل حذف فوري لتفادي كسر التحميل)
        pass


@app.route("/health")
def health():
    return jsonify({"status": "ok", "gemini_key_set": bool(GEMINI_API_KEY)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
