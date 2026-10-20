from flask import Flask, request, render_template_string, send_file, jsonify
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pypdf import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import uuid
import datetime
import json
import time
import threading
import urllib.request

# محاولة استيراد المترجم مع بديل احتياطي
try:
    from deep_translator import GoogleTranslator
    translator = GoogleTranslator(source="auto", target="ar")
    TRANSLATOR_OK = True
except Exception as e:
    print(f"⚠️ googletrans غير متاح: {e}")
    translator = None
    TRANSLATOR_OK = False

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
app.config['FONT_FOLDER'] = os.path.join(BASE_DIR, 'fonts')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['FONT_FOLDER'], exist_ok=True)

# ===== تحميل خط Amiri =====
FONT_PATH = os.path.join(BASE_DIR, 'fonts', 'Amiri-Regular.ttf')
if not os.path.exists(FONT_PATH) or os.path.getsize(FONT_PATH) < 10000:
    try:
        urllib.request.urlretrieve(
            "https://raw.githubusercontent.com/aliftype/amiri/master/Amiri-Regular.ttf",
            FONT_PATH
        )
        print("✅ تم تحميل الخط في وقت التشغيل")
    except Exception as e:
        print(f"⚠️ فشل تحميل الخط: {e}")

FONT_REGISTERED = False
try:
    if os.path.exists(FONT_PATH) and os.path.getsize(FONT_PATH) > 10000:
        pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
        FONT_REGISTERED = True
        print("✅ الخط Amiri مسجل بنجاح")
except Exception as e:
    print(f"⚠️ فشل تسجيل الخط: {e}")

STATS_FILE = os.path.join(BASE_DIR, 'stats.json')
progress_store = {}
stats_lock = threading.Lock()

# ===== دوال الإحصائيات =====
def load_stats():
    try:
        if os.path.exists(STATS_FILE):
            with open(STATS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {"visitors": 0, "translated_files": 0, "downloads": 0, "history": []}

def save_stats(stats):
    try:
        with open(STATS_FILE, 'w', encoding='utf-8') as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ فشل حفظ الإحصائيات: {e}")

def log_action(action, filename=""):
    with stats_lock:
        stats = load_stats()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if action == "visitor":
            stats["visitors"] = stats.get("visitors", 0) + 1
        elif action == "translate":
            stats["translated_files"] = stats.get("translated_files", 0) + 1
        elif action == "download":
            stats["downloads"] = stats.get("downloads", 0) + 1
        stats.setdefault("history", []).append({
            "time": timestamp, "action": action, "file": filename
        })
        if len(stats["history"]) > 500:
            stats["history"] = stats["history"][-500:]
        save_stats(stats)

# ===== دالة الترجمة مع retry =====
def safe_translate(text, retries=2):
    if not TRANSLATOR_OK or not translator:
        return text
    for attempt in range(retries):
        try:
            result = translator.translate(text, dest='ar')
            if result and result.text:
                return result.text
        except Exception as e:
            if attempt == retries - 1:
                print(f"⚠️ فشل الترجمة بعد {retries} محاولات: {e}")
            time.sleep(1)
    return text

# ===== PPTX =====
def translate_pptx(input_path, output_path, task_id):
    try:
        prs = Presentation(input_path)
        slides_list = list(prs.slides)
        total_slides = len(slides_list)
        if total_slides == 0:
            progress_store[task_id] = 100
            return

        for i, slide in enumerate(slides_list):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        original_text = paragraph.text
                        if original_text.strip():
                            translated = safe_translate(original_text)
                            new_p = shape.text_frame.add_paragraph()
                            new_p.text = translated
                            new_p.font.size = Pt(24)
                            new_p.font.name = 'Arial'
                            new_p.alignment = PP_ALIGN.RIGHT
                            new_p.space_before = Pt(10)
                            new_p.space_after = Pt(5)
            progress_store[task_id] = int(((i + 1) / total_slides) * 100)

        prs.save(output_path)
        progress_store[task_id] = 100
    except Exception as e:
        print(f"❌ خطأ في PPTX: {e}")
        progress_store[task_id] = -1

# ===== PDF =====
def translate_pdf(input_path, output_path, task_id):
    try:
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
        if total_pages == 0:
            progress_store[task_id] = 100
            return

        doc = SimpleDocTemplate(
            output_path, pagesize=letter,
            rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50
        )
        styles = getSampleStyleSheet()
        arabic_font = 'Amiri' if FONT_REGISTERED else 'Helvetica'

        original_style = ParagraphStyle(
            'OriginalStyle', parent=styles['Normal'],
            fontName='Helvetica', fontSize=20, leading=26,
            alignment=0, textColor='#000000'
        )
        arabic_style = ParagraphStyle(
            'ArabicStyle', parent=styles['Normal'],
            fontName=arabic_font, fontSize=20, leading=26,
            alignment=2, textColor='#000000'
        )

        story = []
        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
            if text:
                for line in text.split('\n'):
                    if line.strip():
                        translated = safe_translate(line)
                        story.append(Paragraph(line, original_style))
                        story.append(Spacer(1, 4))
                        story.append(Paragraph(f"الترجمة: {translated}", arabic_style))
                        story.append(Spacer(1, 12))
            progress_store[task_id] = int(((i + 1) / total_pages) * 100)

        doc.build(story)
        progress_store[task_id] = 100
    except Exception as e:
        print(f"❌ خطأ في PDF: {e}")
        progress_store[task_id] = -1

# ===== القوالب =====
HTML_PAGE = """
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>الريان للترجمة</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Tajawal', sans-serif; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); min-height: 100vh; display: flex; justify-content: center; align-items: center; padding: 20px; }
    .container { background: rgba(255,255,255,0.95); border-radius: 24px; padding: 50px 40px; max-width: 480px; width: 100%; text-align: center; box-shadow: 0 20px 60px rgba(0,0,0,0.4); border: 1px solid rgba(255,215,0,0.3); }
    .logo { font-size: 42px; font-weight: 700; color: #1a1a2e; margin-bottom: 10px; }
    .logo span { color: #d4af37; }
    .subtitle { font-size: 16px; color: #666; margin-bottom: 40px; }
    .upload-area { border: 2px dashed #d4af37; border-radius: 16px; padding: 40px 20px; margin-bottom: 30px; cursor: pointer; background: #fafafa; transition: 0.3s; }
    .upload-area:hover { background: #fff8e7; }
    .icon { font-size: 48px; margin-bottom: 10px; }
    .upload-btn { background: linear-gradient(135deg, #d4af37 0%, #b8960c 100%); color: white; padding: 16px 40px; border: none; border-radius: 12px; font-size: 18px; font-weight: 700; cursor: pointer; width: 100%; }
    .upload-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    #loader { display: none; margin-top: 20px; color: #d4af37; font-size: 18px; font-weight: 700; }
</style>
</head>
<body>
    <div class="container">
        <div class="logo">الريان <span>للترجمة</span></div>
        <p class="subtitle">جسرك نحو ملفات دقيقة ومحتوى احترافي</p>
        <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" id="fileInput" accept=".pptx,.pdf" required style="display:none;">
            <div class="upload-area" onclick="document.getElementById('fileInput').click();">
                <div class="icon">📄</div>
                <p id="fileName">اسحب ملف PPTX أو PDF هنا أو انقر للتحميل</p>
            </div>
            <button type="submit" class="upload-btn" id="submitBtn" disabled>ابدأ الترجمة</button>
            <div id="loader">⏳ جاري الترجمة... يرجى الانتظار</div>
        </form>
    </div>
    <script>
        document.getElementById('fileInput').addEventListener('change', function() {
            if (this.files.length > 0) {
                document.getElementById('fileName').innerText = this.files[0].name;
                document.getElementById('submitBtn').disabled = false;
            }
        });
        document.getElementById('uploadForm').onsubmit = function() {
            document.getElementById('loader').style.display = 'block';
            document.getElementById('submitBtn').disabled = true;
        };
    </script>
</body>
</html>
"""

PROGRESS_PAGE = """
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>جاري الترجمة...</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<style>
    body { font-family: 'Tajawal', sans-serif; background: #1a1a2e; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; text-align: center; padding: 20px; }
    .container { background: rgba(255,255,255,0.1); padding: 50px 30px; border-radius: 24px; box-shadow: 0 20px 60px rgba(0,0,0,0.4); border: 1px solid #d4af37; max-width: 400px; width: 100%; }
    h1 { color: #d4af37; margin-bottom: 30px; font-size: 28px; }
    .progress-bar-container { width: 100%; background-color: #16213e; border-radius: 10px; overflow: hidden; margin-bottom: 20px; height: 30px; }
    .progress-bar { height: 100%; background: linear-gradient(90deg, #d4af37, #f1c40f); width: 0%; transition: width 0.5s ease; }
    .percentage { font-size: 36px; font-weight: bold; color: #d4af37; }
    .error { color: #e74c3c; font-size: 18px; margin-top: 20px; }
</style>
</head>
<body>
    <div class="container">
        <h1>⏳ جاري ترجمة الملف</h1>
        <div class="progress-bar-container">
            <div class="progress-bar" id="progressBar"></div>
        </div>
        <div class="percentage" id="percentage">0%</div>
        <div class="error" id="errorBox" style="display:none;">❌ حدث خطأ في الترجمة</div>
    </div>
    <script>
        function updateProgress() {
            fetch('/progress/{{ task_id }}')
                .then(r => r.json())
                .then(data => {
                    if (data.progress === -1) {
                        document.getElementById('errorBox').style.display = 'block';
                        return;
                    }
                    document.getElementById('progressBar').style.width = data.progress + '%';
                    document.getElementById('percentage').innerText = data.progress + '%';
                    if (data.progress >= 100) {
                        window.location.href = "/download/{{ filename }}";
                    } else {
                        setTimeout(updateProgress, 1000);
                    }
                })
                .catch(() => setTimeout(updateProgress, 2000));
        }
        updateProgress();
    </script>
</body>
</html>
"""

STATS_PAGE = """
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>الإحصائيات</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<style>
    body { font-family: 'Tajawal', sans-serif; background: #1a1a2e; color: #fff; padding: 30px; }
    .container { max-width: 800px; margin: 0 auto; }
    h1 { color: #d4af37; text-align: center; margin-bottom: 30px; }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .card { background: rgba(255,255,255,0.1); padding: 30px; border-radius: 16px; text-align: center; border: 1px solid #d4af37; }
    .card .num { font-size: 42px; font-weight: bold; color: #d4af37; }
    .card .label { font-size: 16px; color: #ccc; margin-top: 10px; }
    .history { background: rgba(255,255,255,0.05); border-radius: 16px; padding: 20px; }
    .history h2 { color: #d4af37; margin-bottom: 15px; }
    .history-item { padding: 10px; border-bottom: 1px solid rgba(255,255,255,0.1); font-size: 14px; }
    .history-item:last-child { border-bottom: none; }
</style>
</head>
<body>
    <div class="container">
        <h1>📊 إحصائيات التطبيق</h1>
        <div class="cards">
            <div class="card"><div class="num">{{ stats.visitors }}</div><div class="label">👁️ الزوار</div></div>
            <div class="card"><div class="num">{{ stats.translated_files }}</div><div class="label">📄 ملفات مترجمة</div></div>
            <div class="card"><div class="num">{{ stats.downloads }}</div><div class="label">⬇️ تحميلات</div></div>
        </div>
        <div class="history">
            <h2>📜 آخر الأنشطة</h2>
            {% for h in stats.history[-30:]|reverse %}
            <div class="history-item">{{ h.time }} — {{ h.action }} {{ h.file }}</div>
            {% endfor %}
        </div>
    </div>
</body>
</html>
"""

# ===== المسارات =====
@app.route('/')
def index():
    log_action("visitor")
    return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "لم يتم رفع ملف", 400
    file = request.files['file']
    if file.filename == '':
        return "اسم الملف فارغ", 400

    task_id = str(uuid.uuid4())
    filename = task_id

    if file.filename.lower().endswith('.pptx'):
        filename += ".pptx"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        file.save(input_path)
        log_action("translate", filename)
        progress_store[task_id] = 0
        threading.Thread(target=translate_pptx, args=(input_path, output_path, task_id), daemon=True).start()
        return render_template_string(PROGRESS_PAGE, task_id=task_id, filename=filename)

    elif file.filename.lower().endswith('.pdf'):
        filename += ".pdf"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        file.save(input_path)
        log_action("translate", filename)
        progress_store[task_id] = 0
        threading.Thread(target=translate_pdf, args=(input_path, output_path, task_id), daemon=True).start()
        return render_template_string(PROGRESS_PAGE, task_id=task_id, filename=filename)

    return "الملف يجب أن يكون بصيغة .pptx أو .pdf", 400

@app.route('/progress/<task_id>')
def get_progress(task_id):
    return jsonify({"progress": progress_store.get(task_id, 0)})

@app.route('/download/<filename>')
def download_file(filename):
    log_action("download", filename)
    path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if not os.path.exists(path):
        return "الملف غير موجود", 404
    return send_file(path, as_attachment=True, conditional=True)

@app.route('/stats-riyadh-2026')
def stats_page():
    stats = load_stats()
    return render_template_string(STATS_PAGE, stats=stats)

@app.route('/health')
def health():
    return jsonify({
        "status": "ok",
        "translator": TRANSLATOR_OK,
        "font": FONT_REGISTERED
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print("=" * 50)
    print("🚀 بدء تشغيل الريان للترجمة")
    print(f"🔤 المترجم: {'✅ جاهز' if TRANSLATOR_OK else '❌ غير متاح'}")
    print(f"🅰️ الخط العربي: {'✅ جاهز' if FONT_REGISTERED else '❌ غير متاح'}")
    print("=" * 50)
    app.run(host='0.0.0.0', port=port)
