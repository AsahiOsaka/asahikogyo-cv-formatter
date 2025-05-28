import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import fitz  # PyMuPDF
import re
from PIL import Image
from collections import defaultdict

def apply_modern_css():
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(90deg, #0f2027 0%, #2c5364 100%);
        color: #fff;
        padding: 1.5rem 1rem 1rem 1rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        text-align: center;
        font-size: 2.1rem;
        font-weight: 700;
        letter-spacing: 1px;
        box-shadow: 0 2px 8px rgba(44,83,100,0.12);
    }
    .emoji {
        font-size: 2.2rem;
        vertical-align: middle;
        margin-right: 0.5rem;
    }
    .footer {
        margin-top: 2rem;
        color: #888;
        font-size: 1rem;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

class PIIDetector:
    def __init__(self):
        self.patterns = {
            'email': re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'),
            'phone': re.compile(r'(?:\+?1[-.\s]?)?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}|\+\d{1,3}[-.\s]?\d{1,4}[-.\s]?\d{1,4}[-.\s]?\d{1,9}'),
            'address': re.compile(r'\d+\s+[\w\s,.-]+(?:street|st|avenue|ave|road|rd|drive|dr|lane|ln|boulevard|blvd|court|ct|place|pl)(?:\s+(?:apt|apartment|unit|#)\s*\w+)?', re.IGNORECASE),
            'zip_code': re.compile(r'\b\d{5}(?:-\d{4})?\b'),
            'height': re.compile(r'\b(?:\d+\'\s*\d+\"|\d+\s*ft\s*\d+\s*in|\d+\.\d+\s*m|\d+\s*cm)\b', re.IGNORECASE),
            'weight': re.compile(r'\b\d+(?:\.\d+)?\s*(?:lbs?|pounds?|kg|kilograms?)\b', re.IGNORECASE),
            'date_of_birth': re.compile(r'\b(?:DOB|Date of Birth|Born):?\s*(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})', re.IGNORECASE),
            'ssn': re.compile(r'\b\d{3}-\d{2}-\d{4}\b'),
            'linkedin': re.compile(r'linkedin\.com/in/[\w-]+', re.IGNORECASE),
        }
        self.personal_keywords = [
            'home address', 'residential address', 'current address', 'permanent address',
            'contact number', 'mobile number', 'cell phone', 'telephone', 'date of birth', 'dob',
            'born on', 'age:', 'years old', 'marital status', 'married', 'single', 'divorced',
            'nationality', 'citizen', 'passport', 'visa status', 'height:', 'weight:', 'blood type',
            'emergency contact'
        ]

    def detect_names(self, text):
        detected_names = set()
        name_patterns = [
            re.compile(r'^([A-Z][a-z]+(?:\s+[A-Z][a-z]*)*)\s*$', re.MULTILINE),
            re.compile(r'(?:Name|Full Name|Candidate):?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]*)+)', re.IGNORECASE),
            re.compile(r'^([A-Z][a-z]+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)', re.MULTILINE),
        ]
        for pattern in name_patterns:
            matches = pattern.findall(text)
            for match in matches:
                if isinstance(match, tuple):
                    match = match[0] if match[0] else match[1]
                if not any(word.lower() in ['university', 'college', 'company', 'corporation', 'inc', 'ltd', 'experience', 'education', 'skills'] for word in match.split()):
                    if len(match.split()) >= 2:
                        detected_names.add(match.strip())
        return list(detected_names)

    def detect_all_pii(self, text):
        detected_pii = defaultdict(list)
        detected_pii['names'] = self.detect_names(text)
        for pii_type, pattern in self.patterns.items():
            matches = pattern.findall(text)
            if matches:
                detected_pii[pii_type] = list(set(matches))
        personal_info_lines = []
        lines = text.split('\n')
        for line in lines:
            if any(keyword in line.lower() for keyword in self.personal_keywords):
                personal_info_lines.append(line.strip())
        if personal_info_lines:
            detected_pii['personal_info_lines'] = personal_info_lines
        return dict(detected_pii)

    def remove_pii(self, text, detected_pii):
        cleaned_text = text
        for name in detected_pii.get('names', []):
            if name and len(name.strip()) > 2:
                pattern = re.compile(re.escape(name), re.IGNORECASE)
                cleaned_text = pattern.sub('', cleaned_text)
        for pii_type, items in detected_pii.items():
            if pii_type == 'names':
                continue
            for item in items:
                if item and len(str(item).strip()) > 1:
                    pattern = re.compile(re.escape(str(item)), re.IGNORECASE)
                    cleaned_text = pattern.sub('[REDACTED]', cleaned_text)
        lines = cleaned_text.split('\n')
        filtered_lines = []
        for line in lines:
            line_lower = line.lower()
            should_remove = False
            for keyword in self.personal_keywords:
                if keyword in line_lower and not any(work_keyword in line_lower for work_keyword in ['experience', 'work', 'employment', 'company', 'project', 'skill']):
                    should_remove = True
                    break
            if not should_remove and line.strip():
                filtered_lines.append(line)
        cleaned_text = '\n'.join(filtered_lines)
        cleaned_text = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_text)
        cleaned_text = cleaned_text.strip()
        return cleaned_text, 0

def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text()
    return text

def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def abbreviate_name(full_name):
    name_parts = [part for part in full_name.strip().split() if part]
    if len(name_parts) == 0:
        return "NA"
    return ''.join([part[0].upper() for part in name_parts])

def add_header_with_logo(doc, logo_img):
    section = doc.sections[0]
    header = section.header
    for paragraph in header.paragraphs:
        paragraph.clear()
    logo_para = header.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    logo_run = logo_para.add_run()
    image_stream = BytesIO()
    logo_img.save(image_stream, format='PNG')
    image_stream.seek(0)
    logo_run.add_picture(image_stream, width=Inches(2.634), height=Inches(0.508))
    section.header_distance = Inches(0.4)

def generate_asahi_cv(cleaned_text, logo_img, candidate_name):
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.2)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    add_header_with_logo(doc, logo_img)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    name_paragraph = doc.add_paragraph()
    name_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_paragraph.paragraph_format.space_after = Pt(24)
    name_run = name_paragraph.add_run(abbreviate_name(candidate_name))
    name_run.font.name = 'ＭＳ 明朝'
    name_run.font.size = Pt(16)
    name_run.font.bold = True
    doc.add_paragraph()
    content_lines = [line.strip() for line in cleaned_text.strip().split("\n") if line.strip()]
    for line in content_lines:
        if line.strip():
            doc.add_paragraph(line.strip())
    return doc

def main():
    st.set_page_config(
        page_title="Asahi CV Formatter",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    apply_modern_css()
    st.markdown(
        '<div class="main-header"><span class="emoji">📝</span>Asahi CV Formatter <span class="emoji">🔒</span></div>',
        unsafe_allow_html=True
    )
    st.write("Format and anonymize CVs with privacy protection.")

    uploaded_file = st.file_uploader("Upload a CV file (PDF or DOCX)", type=["pdf", "docx"])
    logo_img = Image.new("RGB", (400, 80), color=(44, 83, 100))

    if uploaded_file:
        if uploaded_file.type == "application/pdf":
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            text = extract_text_from_docx(uploaded_file)
        else:
            st.error("Unsupported file type.")
            return

        if not text.strip():
            st.error("Could not extract text from the file.")
            return

        detector = PIIDetector()
        pii = detector.detect_all_pii(text)
        cleaned_text, _ = detector.remove_pii(text, pii)
        candidate_name = pii.get('names', ['Candidate'])[0] if pii.get('names') else "Candidate"
        abbreviation = abbreviate_name(candidate_name)
        doc = generate_asahi_cv(cleaned_text, logo_img, candidate_name)
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        file_name = f"Asahi_CV_{abbreviation}.docx"
        st.download_button(
            label=f"⬇️ Download Abbreviated CV",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    st.markdown('<div class="footer">Made with ❤️</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
