import os
import shutil
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --------------------------------------------------
# CONFIG & FOLDER SETUP
# --------------------------------------------------
project_dir = r"d:\Document\study\Degree\sem 5\PDS\projects\DocumentSimilarityDetector"
assets_dir = os.path.join(project_dir, "assets")
os.makedirs(assets_dir, exist_ok=True)

# Copy actual screenshots from gemini conversation directory if present
brain_dir = r"C:\Users\DELL\.gemini\antigravity-ide\brain\1671c805-0e4b-4d42-bd94-7ac6082ead66"
screenshot_mapping = {
    "ui_empty.png": "latest_ui_empty_1787324231499.png",
    "ui_sandbox.png": "latest_ui_results_1787324446876.png",
    "latest_ui_scan.png": "latest_ui_results_1787324446876.png",
    "diff_highlights.png": "latest_ui_diff_1787324684761.png",
    "ocr_scan.png": "latest_ui_single_ocr_1787325431264.png",
    "scanned_notes.jpg": ".user_uploaded/media_1786947808389.jpg"
}

copied_screenshots = {}
for dest_name, src_name in screenshot_mapping.items():
    src_path = os.path.join(brain_dir, src_name)
    if os.path.exists(src_path):
        dest_path = os.path.join(assets_dir, dest_name)
        shutil.copy(src_path, dest_path)
        copied_screenshots[dest_name] = dest_path
        print(f"Copied screenshot: {dest_name}")
    else:
        # Fallback to local placeholders or look for similar files
        print(f"Screenshot {src_name} not found in brain directory.")

# --------------------------------------------------
# INITIALIZE PPTX PRESENTATION (16:9 Widescreen)
# --------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

# Colors
DARK_BG = RGBColor(15, 23, 42)      # Deep Navy (#0f172a)
CYAN = RGBColor(56, 189, 248)       # Neon Cyan (#38bdf8)
WHITE = RGBColor(255, 255, 255)
SILVER = RGBColor(148, 163, 184)    # Soft Gray (#94a3b8)
SLATE_CARD = RGBColor(30, 41, 59)   # Slate Card (#1e293b)
CODE_BG = RGBColor(10, 15, 29)      # Charcoal Code (#0a0f1d)

def set_slide_background(slide, color=DARK_BG):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_title(slide, text, color=CYAN, size=32):
    title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.5), Inches(11.83), Inches(0.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = 'Segoe UI'
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = color
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

def add_bullets(slide, left, top, width, height, bullets, size=18, color=SILVER):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    for idx, bullet in enumerate(bullets):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.name = 'Segoe UI'
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(12)
        # Check if sub-bullet
        if bullet.strip().startswith("-"):
            p.level = 1
            p.font.size = Pt(size - 2)

def colorize_code_line(p, line, font_size):
    import re
    p.space_after = Pt(2)
    trimmed = line.strip()
    # Style comments completely in gray
    if trimmed.startswith("//") or trimmed.startswith("#"):
        run = p.add_run()
        run.text = line
        run.font.name = 'Consolas'
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(100, 116, 139)  # Slate Gray
        return

    # Regex to split line into strings, comments, numbers, keywords, and syntax tokens
    token_pattern = re.compile(r'(\".*?\"|\'.*?\'|\/\/.*|[\w]+|[^\w\s\x00-\x1f]+|\s+)')
    tokens = token_pattern.findall(line)
    
    keywords = {'async', 'function', 'try', 'catch', 'const', 'let', 'if', 'else', 'return', 'new', 'await', 'resolve', 'reject', 'then'}
    
    for token in tokens:
        run = p.add_run()
        run.text = token
        run.font.name = 'Consolas'
        run.font.size = Pt(font_size)
        
        # Color coding logic
        if token.startswith('"') or token.startswith("'"):
            run.font.color.rgb = RGBColor(253, 186, 116)  # Warm Orange for Strings
        elif token in keywords:
            run.font.color.rgb = RGBColor(244, 114, 182)  # Pink for Keywords
            run.font.bold = True
        elif token.startswith("//") or token.startswith("#"):
            run.font.color.rgb = RGBColor(100, 116, 139)  # Slate Gray for Comments
        elif token.isdigit():
            run.font.color.rgb = RGBColor(129, 140, 248)  # Indigo for Numbers
        elif token in {'document', 'window', 'Tesseract', 'pdfjsLib', 'mammoth', 'circle', 'prs'}:
            run.font.color.rgb = RGBColor(56, 189, 248)   # Cyan for built-in APIs
        else:
            run.font.color.rgb = RGBColor(241, 245, 249)  # Off-White for identifiers and punctuation

def add_code_block(slide, left, top, width, height, code_lines, font_size=11):
    # Draw a card background
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = CODE_BG
    shape.line.color.rgb = RGBColor(71, 85, 105)
    shape.line.width = Pt(1)
    
    # Add text and apply syntax coloring
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = Inches(0.15)
    for idx, line in enumerate(code_lines):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        colorize_code_line(p, line, font_size)

def add_screenshot(slide, left, top, width, height, filename):
    img_path = os.path.join(assets_dir, filename)
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, left, top, width=width, height=height)
        print(f"Placed screenshot {filename} on slide.")
    else:
        # Draw placeholder shape
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = SLATE_CARD
        shape.line.color.rgb = RGBColor(71, 85, 105)
        shape.line.width = Pt(1.5)
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = f"📷 [Actual Screenshot: {filename}]\nNot available in the current environment"
        p.font.name = 'Segoe UI'
        p.font.size = Pt(14)
        p.font.color.rgb = SILVER

# --------------------------------------------------
# SLIDE 1: TITLE SLIDE
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)

# Main Title
title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.2), Inches(11.83), Inches(1.4))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
p.text = "AI-Based Document Similarity & Duplicate Detection System Using NLP"
p.font.name = 'Segoe UI'
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = CYAN

# Subtitle
p_sub = tf.add_paragraph()
p_sub.alignment = PP_ALIGN.CENTER
p_sub.text = "PBL Task 2"
p_sub.font.name = 'Segoe UI'
p_sub.font.size = Pt(16)
p_sub.font.color.rgb = SILVER
p_sub.space_before = Pt(4)

# Subject Text Block (Lime-Yellow and White styling matching the screenshot)
subj_box = slide.shapes.add_textbox(Inches(0.75), Inches(1.6), Inches(11.83), Inches(0.8))
tf_subj = subj_box.text_frame
tf_subj.word_wrap = True
p_subj = tf_subj.paragraphs[0]
p_subj.alignment = PP_ALIGN.CENTER

run1 = p_subj.add_run()
run1.text = "Subject: "
run1.font.name = 'Segoe UI'
run1.font.size = Pt(20)
run1.font.bold = True
run1.font.color.rgb = RGBColor(163, 230, 53)  # Lime Yellow

run2 = p_subj.add_run()
run2.text = "Python for Data Science (BE05000231)"
run2.font.name = 'Segoe UI'
run2.font.size = Pt(20)
run2.font.bold = True
run2.font.color.rgb = WHITE

# Card block for student details
card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(2.5), Inches(11.83), Inches(4.3))
card.fill.solid()
card.fill.fore_color.rgb = SLATE_CARD
card.line.color.rgb = RGBColor(71, 85, 105)
card.line.width = Pt(1.5)

# Student 1 (Left Column)
s1_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.8), Inches(5.0), Inches(1.1))
tf_s1 = s1_box.text_frame
tf_s1.word_wrap = True
p_s1_n = tf_s1.paragraphs[0]
r1_n_k = p_s1_n.add_run()
r1_n_k.text = "NAME 1: "
r1_n_k.font.bold = True
r1_n_k.font.color.rgb = CYAN
r1_n_v = p_s1_n.add_run()
r1_n_v.text = "Nimavat Jaydeep M."
r1_n_v.font.color.rgb = WHITE

p_s1_e = tf_s1.add_paragraph()
p_s1_e.space_before = Pt(4)
r1_e_k = p_s1_e.add_run()
r1_e_k.text = "ENROLLMENT 1: "
r1_e_k.font.bold = True
r1_e_k.font.color.rgb = CYAN
r1_e_v = p_s1_e.add_run()
r1_e_v.text = "250043107032"
r1_e_v.font.color.rgb = WHITE

# Student 2 (Right Column)
s2_box = slide.shapes.add_textbox(Inches(6.8), Inches(2.8), Inches(5.0), Inches(1.1))
tf_s2 = s2_box.text_frame
tf_s2.word_wrap = True
p_s2_n = tf_s2.paragraphs[0]
r2_n_k = p_s2_n.add_run()
r2_n_k.text = "NAME 2: "
r2_n_k.font.bold = True
r2_n_k.font.color.rgb = CYAN
r2_n_v = p_s2_n.add_run()
r2_n_v.text = "Upadhyay Hitarth S."
r2_n_v.font.color.rgb = WHITE

p_s2_e = tf_s2.add_paragraph()
p_s2_e.space_before = Pt(4)
r2_e_k = p_s2_e.add_run()
r2_e_k.text = "ENROLLMENT 2: "
r2_e_k.font.bold = True
r2_e_k.font.color.rgb = CYAN
r2_e_v = p_s2_e.add_run()
r2_e_v.text = "250043107050"
r2_e_v.font.color.rgb = WHITE

# Format fonts for student runs
for r in [r1_n_k, r1_n_v, r1_e_k, r1_e_v, r2_n_k, r2_n_v, r2_e_k, r2_e_v]:
    r.font.name = 'Segoe UI'
    r.font.size = Pt(17)

# Class Details (Branch / Semester) - Centered
class_box = slide.shapes.add_textbox(Inches(1.2), Inches(4.3), Inches(10.9), Inches(0.6))
tf_class = class_box.text_frame
p_class = tf_class.paragraphs[0]
p_class.alignment = PP_ALIGN.CENTER
r_br_k = p_class.add_run()
r_br_k.text = "BRANCH: "
r_br_k.font.bold = True
r_br_k.font.color.rgb = CYAN
r_br_v = p_class.add_run()
r_br_v.text = "Computer Engineering     |     "
r_br_v.font.color.rgb = WHITE

r_sem_k = p_class.add_run()
r_sem_k.text = "SEMESTER: "
r_sem_k.font.bold = True
r_sem_k.font.color.rgb = CYAN
r_sem_v = p_class.add_run()
r_sem_v.text = "5th"
r_sem_v.font.color.rgb = WHITE

for r in [r_br_k, r_br_v, r_sem_k, r_sem_v]:
    r.font.name = 'Segoe UI'
    r.font.size = Pt(17)

# College (Centered)
coll_box = slide.shapes.add_textbox(Inches(1.2), Inches(5.1), Inches(10.9), Inches(0.6))
tf_coll = coll_box.text_frame
p_coll = tf_coll.paragraphs[0]
p_coll.alignment = PP_ALIGN.CENTER
r_coll_k = p_coll.add_run()
r_coll_k.text = "COLLEGE: "
r_coll_k.font.bold = True
r_coll_k.font.color.rgb = CYAN
r_coll_v = p_coll.add_run()
r_coll_v.text = "B H Gardi College Of Engineering And Technology"
r_coll_v.font.color.rgb = WHITE
for r in [r_coll_k, r_coll_v]:
    r.font.name = 'Segoe UI'
    r.font.size = Pt(17)

# --------------------------------------------------
# SLIDE 2: INTRODUCTION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "1. Introduction")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Document Plagiarism is a critical issue in digital classrooms and professional domains.",
    "• Traditional checkers cannot inspect scanned books, images, or offline note photos.",
    "• Data Privacy is often violated by third-party checkers uploading files to external servers.",
    "• Our project builds a 100% serverless, private tool to scan documents and check similarity directly in the browser."
])

# --------------------------------------------------
# SLIDE 3: PROJECT OVERVIEW
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "2. Project Overview")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• 100% Browser-Based Web Application.",
    "• Integrates local OCR, PDF, and DOCX parsers.",
    "• Runs NLP Text Cleaning and Tokenization on-device.",
    "• Computes similarity match indexes dynamically.",
    "• Renders a clean glassmorphic presentation dashboard."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "ui_empty.png")

# --------------------------------------------------
# SLIDE 4: PROBLEM STATEMENT
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "3. Problem Statement")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Inability to Parse Scanned Books/Photos: Students often upload photos of pages which standard text matchers ignore.",
    "• Server Dependencies & Network Timeouts: Local Python servers run slowly and fail behind university proxies (WinError 10060).",
    "• Security & Document Exposure: Commercial checkers keep copies of submitted assignments, creating data leak risks.",
    "• Complex UI: Traditional dashboards are cluttered and slow to load on mobile devices."
])

# --------------------------------------------------
# SLIDE 5: OBJECTIVES
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "4. Objectives")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Serverless Architecture: Build a completely client-side tool requiring zero backend execution.",
    "• Universal Scanned Ingestion: Enable reading text from images (Tesseract.js) and PDFs (PDF.js).",
    "• Fast Plagiarism Checks: Run standard similarity calculations (Cosine & Jaccard index) in milliseconds.",
    "• Safe & Private Checking: Ensure that files are parsed in the browser and never uploaded to any remote host."
])

# --------------------------------------------------
# SLIDE 6: PROJECT MOTIVATION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "5. Project Motivation")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Eliminating Installation Hurdles: Moving away from heavy local Python setups that require installing PyTorch/OpenCV.",
    "• Internet-Independent Scans: Caching language models inside browser storage for offline execution.",
    "• Cost-Free Deployment: Building static files (HTML/CSS/JS) that host for free on Netlify.",
    "• Responsive Presentability: Creating a tool that can be used on any smartphone in a live presentation room."
])

# --------------------------------------------------
# SLIDE 7: EXISTING APPROACH
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "6. Existing Approach")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Heavy Server-Side Frameworks: Standard custom solutions require Python backend processes.",
    "• Network Blockage Constraints: Downloading neural net models (EasyOCR) causes `WinError 10060` connection timeouts in terminals.",
    "• Proprietary Checking Subscriptions: Teachers and students depend on paid accounts (Turnitin, Copyleaks).",
    "• Heavy Resource Footprint: Spawning sub-processes slows down browsers and eats RAM."
])

# --------------------------------------------------
# SLIDE 8: PROPOSED SYSTEM
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "7. Proposed System")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Client-Side JavaScript Application: Runs entirely inside the local browser window.",
    "• CDN Integration: Loads stable parsers (Tesseract.js, PDF.js, Mammoth.js) directly.",
    "• On-Device OCR Execution: Uses Web Workers to scan photos locally.",
    "• Interactive Dashboard: Instant comparisons and sandbox files."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "ui_sandbox.png")

# --------------------------------------------------
# SLIDE 9: KEY FEATURES
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "8. Key Features")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Browser OCR (Tesseract.js): Scans text from JPG, PNG, and JPEG notes directly.",
    "• Binary File Parsers: Reads PDF (PDF.js) and DOCX (Mammoth.js) without backend parsers.",
    "• SVG Circular progress gauge: Animates dynamically based on matching score.",
    "• Turnitin-Style Diff highlighting: Colors matching segments in orange.",
    "• Citation Filter: Toggles ignoring quotes ('...') from the comparison."
])

# --------------------------------------------------
# SLIDE 10: TECHNOLOGY STACK
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "9. Technology Stack")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Core Language: JavaScript (ES6+), HTML5, CSS3 Variables.",
    "• Optical Character Recognition (OCR): Tesseract.js (v5, compiled in WebAssembly).",
    "• PDF Parser: Mozilla's PDF.js (v3.11.174).",
    "• Word Parser: Mammoth.js (v1.6.0).",
    "• Hosting: Netlify Continuous Deployment (Git-Synced)."
])

# --------------------------------------------------
# SLIDE 11: SYSTEM ARCHITECTURE
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "10. System Architecture")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Input Layer: HTML drag-and-drop file inputs or raw copy-paste text areas.",
    "• Native Extractor: FileReader API converts file bytes into array buffers.",
    "• Parsing Engine: Tesseract, PDF.js, or Mammoth.js processes buffers into clean string arrays.",
    "• Math Processor: JavaScript calculates Cosine vector distances and Jaccard intersections.",
    "• Visual Output: Updates SVG dash-array and highlighted spans dynamically."
])

# --------------------------------------------------
# SLIDE 12: SYSTEM WORKFLOW
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "11. System Workflow")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Start: User opens index.html in browser.",
    "• Load: CDNs fetch parsers and initialize Tesseract workers.",
    "• Input: User uploads document photos or pastes notes.",
    "• Click: Form submission triggers the extraction loader overlay.",
    "• Process: Client scans images, extracts text, runs comparison math.",
    "• Output: Displays gauge percentage and side-by-side matches."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "ocr_scan.png")

# --------------------------------------------------
# SLIDE 13: INPUT PROCESSING
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "12. Input Processing")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• FileReader API: Asynchronously reads local file bytes in the browser.",
    "• Array Buffer Extraction: Reads DOCX and PDF as raw byte arrays for Mammoth and PDF.js.",
    "• Data URL Extraction: Reads JPG and PNG files as base64 URLs for image display previews.",
    "• State Sync: Automatically fills text areas with the extracted text for editing."
])

# --------------------------------------------------
# SLIDE 14: DOCUMENT PROCESSING
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "13. Document Processing")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• PDF.js Processing:",
    "  - `pdfjsLib.getDocument()` parses PDF arrayBuffer into document objects.",
    "  - Iterates page-by-page to fetch text elements via `page.getTextContent()`.",
    "• Mammoth.js Processing:",
    "  - `mammoth.extractRawText()` converts Word DOCX XML blocks into clean paragraphs.",
    "  - Strips layout formatting metadata to retrieve pure string arrays."
])

# --------------------------------------------------
# SLIDE 15: IMAGE PREPROCESSING
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "14. Image Preprocessing")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Status: Not available in the current implementation.",
    "• Rationale: In our browser-native client app, images are fed directly to Tesseract.js in their raw state.",
    "• Tesseract.js handles thresholding and binarization internally within its compiled WebAssembly core.",
    "• Future Scope: Can integrate canvas-based threshold filters to improve scanning accuracy on blurry pages."
])

# --------------------------------------------------
# SLIDE 16: OCR TECHNOLOGY
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "15. OCR Technology")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Tesseract.js WebAssembly Port: Compiles Google's neural network OCR to browser-executable binary.",
    "• Web Worker Sandbox: Spawns offline worker threads, leaving the main UI thread lag-free.",
    "• Local Language Cache: Downloads the English language training set once and caches it in IndexedDB.",
    "• Real-time progress updates: Logs execution percentages dynamically to show scanning status."
])

# --------------------------------------------------
# SLIDE 17: OCR IMPLEMENTATION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "16. OCR Implementation")

add_bullets(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(4.8), [
    "• The OCR extraction is called inside a Promise.",
    "• Tesseract.recognize accepts the image file.",
    "• A logger callback tracks progress.",
    "• Displays percentage live in the loader overlay.",
    "• Extracted text is returned to the parser route."
])

code_ocr = [
    "Tesseract.recognize(",
    "    file,",
    "    'eng',",
    "    {",
    "        logger: m => {",
    "            if (m.status === 'recognizing text') {",
    "                const pct = Math.round(m.progress * 100);",
    "                updateStatus(`📷 Scanning Image: ${pct}%`);",
    "            }",
    "        }",
    "    }",
    ").then(({ data: { text } }) => {",
    "    resolve(text);",
    "});"
]
add_code_block(slide, Inches(6.5), Inches(1.8), Inches(6.08), Inches(4.2), code_ocr, font_size=12)

# --------------------------------------------------
# SLIDE 18: OCR OUTPUT
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "17. OCR Output")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• String Generation: Outputs a clean text string containing all recognized characters and symbols.",
    "• Syncing fields: Extracted text is automatically entered into the text areas for verification.",
    "• Visual preview: Displays a base64 preview of the original photographed note right above the results.",
    "• Stats counter: Calculates total words and characters found in the scanned image."
])

# --------------------------------------------------
# SLIDE 19: TEXT PREPROCESSING
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "18. Text Preprocessing")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Lowercasing: `text.toLowerCase()` avoids matches failing due to capitalization differences.",
    "• Punctuation Stripping: `.replace(/[^\w\s]/g, '')` strips periods, commas, and question marks.",
    "• Spaces Trimming: `.trim()` removes trailing spaces.",
    "• RegEx word boundary splits: `.split(/\s+/)` cleans double-spaces to generate a clean array of words."
])

# --------------------------------------------------
# SLIDE 20: NLP PROCESSING
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "19. NLP Processing")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Tokenization: Splits text strings into individual word token arrays.",
    "• Vocabulary Sets: Javascript `Set` creates unique list vectors of words.",
    "• Array transformations: Converts vocab lists into structured frequency arrays for math models.",
    "• Clean Input: Normalization removes duplicate spaces, preparing the data for similarity models."
])

# --------------------------------------------------
# SLIDE 21: FEATURE EXTRACTION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "20. Feature Extraction")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Vocabulary Union Vector: Compares word sets from both documents to create a master list of all unique words.",
    "• Word Count Maps: Creates key-value frequency objects for both Document A and Document B.",
    "• Indexing coordinates: Maps the frequency values matching the master list order to represent both documents as vectors.",
    "• Ready for similarity: High-performance representation built entirely dynamically in JavaScript."
])

# --------------------------------------------------
# SLIDE 22: TF-IDF / ACTUAL FEATURE METHOD
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "21. Term Frequency (TF) Vector Method")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Actual Method in Code: Dynamic Word Frequency Vectors.",
    "• Step 1: Combine both documents' word arrays to create a master Vocabulary Union Set.",
    "• Step 2: Initialize frequency map objects (frequency arrays) mapping each vocab word to 0.",
    "• Step 3: Populate counts by iterating through both text arrays.",
    "• Step 4: Map values to numerical arrays representing vectors, ready for Cosine Distance calculation."
])

# --------------------------------------------------
# SLIDE 23: COSINE SIMILARITY / ACTUAL SIMILARITY METHOD
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "22. Cosine Similarity Vector Math")

add_bullets(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(4.8), [
    "• Represents both texts as vectors.",
    "• Calculates Dot Product by multiplying matching frequencies.",
    "• Divides by the product of their Vector Magnitudes.",
    "• Outputs similarity score on scale 0% to 100%.",
    "• Keeps matching accurate regardless of text length differences."
])

code_cosine = [
    "let dotProduct = 0;",
    "let mag1 = 0, mag2 = 0;",
    "vocab.forEach(word => {",
    "    dotProduct += freq1[word] * freq2[word];",
    "    mag1 += freq1[word] * freq1[word];",
    "    mag2 += freq2[word] * freq2[word];",
    "});",
    "if (mag1 === 0 || mag2 === 0) return 0;",
    "return (dotProduct / (Math.sqrt(mag1) * Math.sqrt(mag2))) * 100;"
]
add_code_block(slide, Inches(6.5), Inches(1.8), Inches(6.08), Inches(4.2), code_cosine, font_size=12)

# --------------------------------------------------
# SLIDE 24: SIMILARITY DETECTION (JACCARD)
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "23. Similarity Detection (Jaccard Index)")

add_bullets(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(4.8), [
    "• Checks overall Vocabulary Set overlap.",
    "• Converts cleaned word arrays into unique Sets.",
    "• Intersection: Shared unique words.",
    "• Union: Total unique words combined.",
    "• Dividing Intersection by Union gives Jaccard percentage.",
    "• Highly effective at catching copied vocabulary sets."
])

code_jaccard = [
    "const words1 = new Set(t1.split(/\\s+/));",
    "const words2 = new Set(t2.split(/\\s+/));",
    "const intersection = new Set(",
    "    [...words1].filter(x => words2.has(x))",
    ");",
    "const union = new Set([...words1, ...words2]);",
    "if (union.size === 0) return 0;",
    "return (intersection.size / union.size) * 100;"
]
add_code_block(slide, Inches(6.5), Inches(1.8), Inches(6.08), Inches(4.2), code_jaccard, font_size=11)

# --------------------------------------------------
# SLIDE 25: DOCUMENT COMPARISON
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "24. Document Comparison & Highlighting")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Combination Index: Computes the average of Cosine Similarity and Jaccard Index score for balanced output.",
    "• Word-Level Match Highlighter:",
    "  - Converts strings to word lists.",
    "  - Checks if a word in Document A exists in the vocabulary set of Document B.",
    "  - Wraps matches in orange highlight spans.",
    "• Citation Filter: Ignores quoted phrases dynamically."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "diff_highlights.png")

# --------------------------------------------------
# SLIDE 26: SIMILARITY RESULTS
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "25. Similarity Results & Status")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Classification ranges based on combined score:",
    "  - 90%+ : Direct Copy / Duplicate (Red)",
    "  - 70% - 89% : Highly Similar (Orange)",
    "  - 40% - 69% : Moderately Similar (Yellow)",
    "  - 15% - 39% : Slightly Similar (Blue)",
    "  - <15% : Completely Unique (Green)",
    "• Renders matching warning descriptions automatically."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "latest_ui_scan.png")

# --------------------------------------------------
# SLIDE 27: DATA VISUALIZATION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "26. Data Visualization (Circle Gauge)")

add_bullets(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(4.8), [
    "• Interactive SVG Circular progress ring.",
    "• Circumference = 326.7px (radius = 52px).",
    "• JavaScript calculates matching stroke-dashoffset.",
    "• Smooth CSS transition filling animation.",
    "• Accent color changes dynamically matching status levels."
])

code_gauge = [
    "// Radial Gauge update logic in JS",
    "const circle = document.getElementById('radial-bar');",
    "const circumference = 326.7;",
    "const offset = circumference - ",
    "      (circumference * scoreInt / 100);",
    "circle.style.strokeDashoffset = offset;",
    "circle.style.stroke = color;",
    "circle.style.filter = ",
    "      `drop-shadow(0 0 8px ${color}80)`;"
]
add_code_block(slide, Inches(6.5), Inches(1.8), Inches(6.08), Inches(4.2), code_gauge, font_size=12)

# --------------------------------------------------
# SLIDE 28: USER INTERFACE
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "27. User Interface Design")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Modern Glassmorphic style card layout.",
    "• Background: Deep midnight slate-blue gradients.",
    "• UI Accents: Neon glowing borders for result levels.",
    "• Google Fonts integration (Plus Jakarta Sans).",
    "• Copy-to-clipboard and reset controls.",
    "• Fully responsive layouts scaling on mobile devices."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "latest_ui_scan.png")

# --------------------------------------------------
# SLIDE 29: CODE IMPLEMENTATION HIGHLIGHTS
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "28. Main Web Controller Logic")

add_bullets(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(4.8), [
    "• Client-side controller routing in JS.",
    "• Handles extraction, OCR, comparisons.",
    "• Single scanner vs Dual match flow check.",
    "• Shows loader overlay during calculations.",
    "• Catches errors and updates error container."
])

code_main = [
    "async function startScanner() {",
    "  try {",
    "    const t1 = await extractText('file1', 'text1', 'Doc A');",
    "    const t2 = await extractText('file2', 'text2', 'Doc B');",
    "    if (t1 && !t2) {",
    "      renderSingleScannerResult('A', t1);",
    "    } else {",
    "      renderComparisonResult(t1, t2);",
    "    }",
    "  } catch (err) {",
    "    showError(err);",
    "  }",
    "}"
]
add_code_block(slide, Inches(6.5), Inches(1.8), Inches(6.08), Inches(4.2), code_main, font_size=11)

# --------------------------------------------------
# SLIDE 30: TESTING & VALIDATION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "29. Testing & Validation")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(6.0), Inches(4.8), [
    "• Tested offline OCR extraction with handwritten note scans (Math mappings and functions).",
    "• Tested digital PDF extraction page-by-page to check text stream reconstruction.",
    "• Verified similarity scores using identical and modified notes in the sandbox selector.",
    "• Confirmed that network firewalls have zero impact on scanning speeds or execution."
])
add_screenshot(slide, Inches(7.2), Inches(1.8), Inches(5.38), Inches(4.2), "scanned_notes.jpg")

# --------------------------------------------------
# SLIDE 31: ADVANTAGES, LIMITATIONS & FUTURE SCOPE
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)
add_title(slide, "30. Advantages, Limitations & Future Scope")
add_bullets(slide, Inches(0.75), Inches(1.8), Inches(11.83), Inches(4.8), [
    "• Advantages: Serverless, completely cost-free hosting, total document privacy (local runs).",
    "• Limitations: Large files take longer to scan via browser CPU. Lack of database to store history.",
    "• Future Scope:",
    "  - Web Link compare support: input URL to check text against online web pages.",
    "  - Multi-language OCR: dropdown to select Spanish, Hindi, or French scanning models.",
    "  - PDF Report Export: client-side PDF compilation to save results offline."
])

# --------------------------------------------------
# SLIDE 32: CONCLUSION
# --------------------------------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide)

# Conclusion Header
title_box = slide.shapes.add_textbox(Inches(0.75), Inches(1.5), Inches(11.83), Inches(0.8))
tf = title_box.text_frame
p = tf.paragraphs[0]
p.text = "31. Conclusion"
p.font.name = 'Segoe UI'
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = CYAN

# Body
body_box = slide.shapes.add_textbox(Inches(0.75), Inches(2.5), Inches(11.83), Inches(2.2))
tf_body = body_box.text_frame
tf_body.word_wrap = True
p_b = tf_body.paragraphs[0]
p_b.text = "We have successfully built a fully serverless, lightweight browser-based document scanner and similarity detector using Tesseract.js, PDF.js, and Mammoth.js.\n\nThe project demonstrates that advanced AI parsing and NLP comparison algorithms can run completely locally on consumer browser engines, eliminating the need for expensive GPU-hosting setups."
p_b.font.name = 'Segoe UI'
p_b.font.size = Pt(20)
p_b.font.color.rgb = SILVER
p_b.space_after = Pt(15)

# Thank You Centered
thanks_box = slide.shapes.add_textbox(Inches(0.75), Inches(4.8), Inches(11.83), Inches(1.8))
tf_t = thanks_box.text_frame
p_t = tf_t.paragraphs[0]
p_t.alignment = PP_ALIGN.CENTER
p_t.text = "THANK YOU\n\nQuestions?"
p_t.font.name = 'Segoe UI'
p_t.font.size = Pt(36)
p_t.font.bold = True
p_t.font.color.rgb = CYAN

# --------------------------------------------------
# SAVE PRESENTATION (With robust fallback loops for locked files)
# --------------------------------------------------
import time
base_name = "AI_Document_OCR_Scanner_Similarity_Detector_PBL"
out_path = os.path.join(project_dir, f"{base_name}.pptx")
saved = False

try:
    prs.save(out_path)
    print(f"Presentation saved successfully to {out_path}")
    saved = True
except PermissionError:
    pass

if not saved:
    alt_path = os.path.join(project_dir, f"{base_name}_v2.pptx")
    try:
        prs.save(alt_path)
        print(f"WARNING: original file locked. Saved alternative version successfully to: {alt_path}")
        saved = True
    except PermissionError:
        pass

if not saved:
    timestamp = int(time.time())
    timestamp_path = os.path.join(project_dir, f"{base_name}_{timestamp}.pptx")
    prs.save(timestamp_path)
    print(f"WARNING: Both original and v2 files are locked! Saved with timestamp to: {timestamp_path}")
