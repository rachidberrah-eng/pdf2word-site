# -*- coding: utf-8 -*-
"""
موقع تحويل PDF (رياضيات/فيزياء) إلى Word بمعادلات قابلة للتعديل.

الميزات:
1) استخراج النص من PDF (نص أو صور)
2) تحويل المعادلات إلى LaTeX
3) دعم RTL و LTR
4) تنسيق مخصص (خط، حجم، لون)
"""

import os
import re
import uuid
import shutil
import subprocess
import tempfile
import base64

from flask import Flask, request, render_template, send_file, jsonify
import fitz  # PyMuPDF
import requests
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

# مجلد مؤقت لكل عملية معالجة
BASE_TMP = os.path.join(tempfile.gettempdir(), "pdf2word_jobs")
os.makedirs(BASE_TMP, exist_ok=True)

# Gemini API
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash")
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)

MAX_CHARS_PER_CALL = 60000


def extract_text_from_pdf(pdf_path: str) -> str:
    """يستخرج النص من PDF (نص أو صور)."""
    doc = fitz.open(pdf_path)
    pages_text = []
    
    for page_num, page in enumerate(doc):
        # محاولة استخراج النص العادي أولاً
        text = page.get_text("text")
        
        if text.strip():
            pages_text.append(f"[صفحة {page_num + 1}]\n{text}")
        else:
            # إذا لم يوجد نص، نحاول استخراج الصور وقراءة النص منها
            images = page.get_images(full=True)
            if images:
                page_images_text = []
                for img_index, img in enumerate(images):
                    xref = img[0]
                    pix = fitz.Pixmap(doc, xref)
                    
                    # تحويل الصورة إلى base64
                    if pix.n - pix.alpha > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    
                    img_bytes = pix.tobytes("png")
                    img_base64 = base64.b64encode(img_bytes).decode()
                    
                    # قراءة النص من الصورة عبر Gemini
                    try:
                        img_text = read_text_from_image(img_base64)
                        if img_text:
                            page_images_text.append(img_text)
                    except Exception as e:
                        print(f"خطأ في قراءة صورة: {e}")
                
                if page_images_text:
                    pages_text.append(f"[صفحة {page_num + 1} - صور]\n" + "\n".join(page_images_text))
    
    doc.close()
    return "\n\n".join(pages_text)


def read_text_from_image(image_base64: str) -> str:
    """يقرأ النص من صورة عبر Gemini Vision."""
    if not GEMINI_API_KEY:
        raise RuntimeError("لم يتم ضبط GEMINI_API_KEY")
    
    prompt = """هذه صورة من ملف PDF. استخرج كل النص الموجود فيها بدقة.
اكتب النص بالعربية أو الفرنسية أو الإنجليزية حسب اللغة الموجودة.
إذا كانت معادلات رياضية، اكتبها بصيغة LaTeX بين $ أو $$."""
    
    payload = {
        "contents": [{
            "role": "user",
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/png", "data": image_base64}}
            ]
        }],
        "generationConfig": {"temperature": 0.1}
    }
    headers = {"Content-Type": "application/json", "x-goog-api-key": GEMINI_API_KEY}
    
    resp = requests.post(GEMINI_URL, json=payload, headers=headers, timeout=180)
    resp.raise_for_status()
    data = resp.json()
    
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        return ""


def clean_with_gemini(raw_text: str, text_direction: str = "rtl") -> str:
    """ينظف النص ويعيد تنسيقه كـ Markdown."""
    if not GEMINI_API_KEY:
        raise RuntimeError("لم يتم ضبط GEMINI_API_KEY")
    
    dir_instruction = "من اليمين إلى اليسار (عربي)" if text_direction == "rtl" else "من اليسار إلى اليمين (فرنسي/إنجليزي)"
    
    prompt = (
        f"أنت مساعد متخصص في تنسيق المحتوى العلمي.\n"
        f"المحتوى يجب أن يكون {dir_instruction}.\n\n"
        f"سأعطيك نصًا خامًا من PDF:\n"
        f"- العناوين بـ '# '\n"
        f"- المعادلات الرياضية بـ LaTeX (سطر كامل بـ $$ $$، داخل السطر بـ $ $)\n"
        f"- حافظ على النص كما هو دون حذف\n"
        f"- أبقِ النص {dir_instruction}\n\n"
        f"النص:\n-----\n"
        f"{raw_text}\n"
        f"-----"
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
        raise RuntimeError(f"استجابة غير متوقعة: {data}")


def chunk_text(text: str, max_chars: int):
    """يقسّم النص إلى أجزاء."""
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


def set_rtl_paragraph(paragraph):
    """يضبط اتجاه الفقرة إلى RTL."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    dir = OxmlElement('w:dir')
    dir.set(qn('w:val'), 'rtl')
    pPr.append(dir)


def build_reference_docx(path: str, style: dict, text_direction: str = "rtl"):
    """يبني ملف Word مرجعي."""
    doc = Document()
    
    # ضبط اتجاه المستند
    if text_direction == "rtl":
        section = doc.sections[0]
        section.page_direction = 2  # RTL
    
    # أنماط العناوين
    h1 = doc.styles["Heading 1"]
    h1.font.name = style["title_font"]
    h1.font.size = Pt(style["title_size"])
    h1.font.color.rgb = RGBColor.from_string(style["title_color"])
    
    h2 = doc.styles["Heading 2"]
    h2.font.name = style["title_font"]
    h2.font.size = Pt(max(style["title_size"] - 4, 10))
    h2.font.color.rgb = RGBColor.from_string(style["title_color"])
    
    # النص العادي
    normal = doc.styles["Normal"]
    normal.font.name = style["text_font"]
    normal.font.size = Pt(style["text_size"])
    normal.font.color.rgb = RGBColor.from_string(style["text_color"])
    
    doc.save(path)


def hex_or_default(value: str, default: str = "000000") -> str:
    """ينظف قيمة اللون."""
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
        
        # قراءة الخيارات
        text_direction = request.form.get("text_direction", "rtl")
        alignment = request.form.get("alignment", "right" if text_direction == "rtl" else "left")
        line_spacing = float(request.form.get("line_spacing", 1.5))
        
        style = {
            "title_font": request.form.get("title_font", "Arial"),
            "title_size": int(request.form.get("title_size", 20)),
            "title_color": hex_or_default(request.form.get("title_color"), "C00000"),
            "text_font": request.form.get("text_font", "Calibri"),
            "text_size": int(request.form.get("text_size", 13)),
            "text_color": hex_or_default(request.form.get("text_color"), "000000"),
        }
        
        # 1) استخراج النص
        raw_text = extract_text_from_pdf(pdf_path)
        if not raw_text.strip():
            return jsonify({
                "error": "لم أتمكن من قراءة هذا الملف. تأكد أن الملف ليس فارغًا."
            }), 400
        
        # 2) التنظيف عبر Gemini
        chunks = chunk_text(raw_text, MAX_CHARS_PER_CALL)
        cleaned_parts = [clean_with_gemini(c, text_direction) for c in chunks]
        markdown_text = "\n\n".join(cleaned_parts)
        
        # إضافة توجيه RTL/LTR في Markdown
        if text_direction == "rtl":
            markdown_text = "{.dir=rtl}\n\n" + markdown_text
        else:
            markdown_text = "{.dir=ltr}\n\n" + markdown_text
        
        md_path = os.path.join(job_dir, "content.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown_text)
        
        # 3) بناء القالب
        ref_path = os.path.join(job_dir, "reference.docx")
        build_reference_docx(ref_path, style, text_direction)
        
        # 4) الدمج عبر Pandoc
        output_path = os.path.join(job_dir, "output.docx")
        pandoc_cmd = "pandoc.exe" if os.name == "nt" else "pandoc"
        
        # إضافة خيارات المحاذاة والمسافة
        align_map = {"left": "left", "center": "center", "right": "right", "justify": "justify"}
        extra_args = [
            f"--reference-doc={ref_path}",
            "--extract-media", job_dir
        ]
        
        subprocess.run(
            [pandoc_cmd, md_path, "-o", output_path] + extra_args,
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
        return jsonify({"error": f"فشل التحويل: {e.stderr or str(e)}"}), 500
    except FileNotFoundError as e:
        if "pandoc" in str(e).lower():
            return jsonify({"error": "Pandoc غير موجود. يرجى تثبيته."}), 500
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health")
def health():
    pandoc_available = shutil.which("pandoc") or shutil.which("pandoc.exe")
    return jsonify({
        "status": "ok",
        "gemini_key_set": bool(GEMINI_API_KEY),
        "pandoc_available": bool(pandoc_available),
        "ocr_available": True,  # Gemini Vision للـ OCR
        "temp_dir": BASE_TMP,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
