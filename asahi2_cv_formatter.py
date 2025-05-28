# Professional Asahi CV Formatter - Clean & Simple Design
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from io import BytesIO
import fitz  # PyMuPDF
import re
from PIL import Image
from collections import defaultdict
import time

# --- Professional Clean CSS Design ---
def apply_professional_css():
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

# --- Advanced PII Detection Class ---
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
        removal_count = 0
        for name in detected_pii.get('names', []):
            if name and len(name.strip()) > 2:
                pattern = re.compile(re.escape(name), re.IGNORECASE)
                if pattern.search(cleaned_text):
                    cleaned_text = pattern.sub('', cleaned_text)
                    removal_count += 1
        for pii_type, items in detected_pii.items():
            if pii_type == 'names':
                continue
            for item in items:
                if item and len(str(item).strip()) > 1:
                    pattern = re.compile(re.escape(str(item)), re.IGNORECASE)
                    if pattern.search(cleaned_text):
                        cleaned_text = pattern.sub('[REDACTED]', cleaned_text)
                        removal_count += 1
        lines = cleaned_text.split('\n')
        filtered_lines = []
        for line in lines:
            line_lower = line.lower()
            should_remove = False
            for keyword in self.personal_keywords:
                if keyword in line_lower and not any(work_keyword in line_lower for work_keyword in ['experience', 'work', 'employment', 'company', 'project', 'skill']):
                    should_remove = True
                    removal_count += 1
                    break
            if not should_remove and line.strip():
                filtered_lines.append(line)
        cleaned_text = '\n'.join(filtered_lines)
        cleaned_text = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_text)
        cleaned_text = cleaned_text.strip()
        return cleaned_text, removal_count

# --- Helper Functions ---
def extract_text_from_pdf(file):
    text = ""
    try:
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            for page in doc:
                text += page.get_text()
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return ""
    return text

def extract_text_from_docx(file):
    try:
        doc = Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Error reading DOCX: {str(e)}")
        return ""

def abbreviate_name_age(full_name, age):
    try:
        name_parts = [part.strip() for part in full_name.strip().split() if part.strip()]
        if not name_parts:
            return f"N.A.{age}yrs"
        initials = ''.join([part[0].upper() + '.' for part in name_parts])
        return f"{initials} {age}yrs"
    except Exception:
        return f"N.A.{age}yrs"

def add_header_with_logo(doc, logo_img):
    section = doc.sections[0]
    header = section.header
    for paragraph in header.paragraphs:
        paragraph.clear()
    logo_para = header.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tab_stops = logo_para.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(6.5), WD_ALIGN_PARAGRAPH.RIGHT)
    logo_run = logo_para.add_run("\t")
    image_stream = BytesIO()
    logo_img.save(image_stream, format='PNG')
    image_stream.seek(0)
    logo_run.add_picture(image_stream, width=Inches(2.634), height=Inches(0.508))
    section.header_distance = Inches(0.4)

def generate_asahi_cv(cleaned_text, logo_img, candidate_name, age):
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
    name_run = name_paragraph.add_run(abbreviate_name_age(candidate_name, age))
    name_run.font.name = 'ＭＳ 明朝'
    name_run.font.size = Pt(16)
    name_run.font.bold = True
    doc.add_paragraph()
    content_lines = [line.strip() for line in cleaned_text.strip().split("\n") if line.strip()]
    for line in content_lines:
        if line.strip():
            doc.add_paragraph(line.strip())
    return doc

# --- Main Application ---
def main():
    st.set_page_config(
        page_title="Asahi CV Formatter",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    apply_professional_css()

    # Clean professional header
    st.markdown("""
    <div class="main-header">
        <span class="emoji">📝</span>Professional CV formatting with automatic privacy protection
    </div>
    """, unsafe_allow_html=True)
    st.write("Upload your CV for professional formatting and PII protection.")

    # Age input
    age = st.number_input("Enter Candidate Age", min_value=18, max_value=99, value=25)

    # File uploader
    uploaded_file = st.file_uploader("Upload a CV file (PDF or DOCX)", type=["pdf", "docx"])

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
        cleaned_text, removal_count = detector.remove_pii(text, pii)

        candidate_name = pii['names'][0] if pii.get('names') else "Candidate Name"

        # Load logo image
        try:
            logo_img = Image.open("asahi_logo.png")
        except FileNotFoundError:
            st.error("asahi_logo.png not found.  Place in the same directory.")
            return
        except Exception as e:
            st.error(f"Error loading logo: {e}")
            return

        doc = generate_asahi_cv(cleaned_text, logo_img, candidate_name, age)

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        file_name = f"Asahi_CV_{abbreviate_name_age(candidate_name,age)}.docx"

        st.download_button(
            label="Download Formatted CV",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    st.markdown('<div class="footer">©Asahi Kogyo Co., Ltd. Osaka Office</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
