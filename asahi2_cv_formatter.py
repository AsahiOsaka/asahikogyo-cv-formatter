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
        cursor: pointer;
    }
    .main-header:hover {
        background: linear-gradient(90deg, #1a2b32 0%, #3a6374 100%);
        transition: all 0.3s ease;
    }
    .emoji {
        font-size: 2.2rem;
        vertical-align: middle;
        margin-right: 0.5rem;
    }
    .supported-formats {
        background: linear-gradient(135deg, #f0f4f8 0%, #e8f2ff 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        border-left: 4px solid #2c5364;
        box-shadow: 0 2px 10px rgba(44,83,100,0.08);
    }
    .supported-formats h3 {
        color: #2c5364;
        margin-bottom: 1rem;
        font-weight: 600;
        font-size: 1.1rem;
    }
    .format-list {
        margin: 0;
        padding-left: 1.2rem;
        color: #4a5568;
    }
    .format-list li {
        margin-bottom: 0.5rem;
        font-weight: 500;
    }
    .pdf-icon {
        color: #e25555;
        font-size: 1.1rem;
        margin-left: 0.5rem;
    }
    .docx-icon {
        color: #4a90e2;
        font-size: 1.1rem;
        margin-left: 0.5rem;
    }
    .clean-card {
        background: #ffffff;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        border: 1px solid #e2e8f0;
        margin-bottom: 1.5rem;
    }
    .clean-card h3 {
        color: #374151;
        font-weight: 600;
        margin-bottom: 1.5rem;
        font-size: 1.2rem;
        border-bottom: 2px solid #f1f5f9;
        padding-bottom: 0.5rem;
    }
    .upload-area {
        border: 2px dashed #cbd5e0;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        background: #f7fafc;
        transition: all 0.3s ease;
    }
    .upload-area:hover {
        border-color: #2c5364;
        background: #edf2f7;
    }
    .status-success {
        color: #2d5016;
        padding: 0.5rem 0;
        margin: 1rem 0;
        font-weight: 500;
        font-size: 0.95rem;
    }
    .status-info {
        background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
        box-shadow: 0 2px 8px rgba(66,153,225,0.2);
    }
    .status-warning {
        background: linear-gradient(135deg, #ed8936 0%, #dd6b20 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
        box-shadow: 0 2px 8px rgba(237,137,54,0.2);
    }
    .progress-step {
        margin: 0.5rem 0;
        color: #4a5568;
        font-weight: 500;
    }
    .step-check {
        color: #48bb78;
        margin-right: 0.5rem;
        font-weight: bold;
    }
    .metric-item {
        text-align: center;
        background: #f7fafc;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    .metric-number {
        font-size: 1.8rem;
        font-weight: bold;
        color: #2c5364;
    }
    .metric-label {
        color: #718096;
        font-size: 0.9rem;
        margin-top: 0.2rem;
    }
    .detection-summary {
        background: #fffbeb;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #f59e0b;
    }
    .detection-item {
        margin: 0.5rem 0;
        color: #92400e;
        font-size: 0.9rem;
    }
    .footer {
        margin-top: 3rem;
        padding-top: 2rem;
        border-top: 1px solid #e2e8f0;
        color: #718096;
        font-size: 0.9rem;
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
            'contact number', 'mobile number', 'cell phone', 'telephone',
            'date of birth', 'dob', 'born on', 'age:', 'years old',
            'marital status', 'married', 'single', 'divorced',
            'nationality', 'citizen', 'passport', 'visa status',
            'height:', 'weight:', 'blood type', 'emergency contact'
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
        # Don't include names in the returned PII
        
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
        
        # Still detect and remove names internally, but don't show them in PII report
        detected_names = self.detect_names(text)
        for name in detected_names:
            if name and len(name.strip()) > 2:
                pattern = re.compile(re.escape(name), re.IGNORECASE)
                if pattern.search(cleaned_text):
                    cleaned_text = pattern.sub('', cleaned_text)
                    removal_count += 1
        
        for pii_type, items in detected_pii.items():
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
    
    # Clickable header that scrolls to upload area
    st.markdown("""
    <div class="main-header" onclick="document.querySelector('.upload-area').scrollIntoView({behavior: 'smooth'});">
        <span class="emoji">📝</span>Asahi CV Formatter
    </div>
    """, unsafe_allow_html=True)
    st.write("Professional CV formatting with automatic privacy protection")

    # Supported formats section with improved design
    st.markdown("""
    <div class="supported-formats">
        <h3>📄 Supported Formats</h3>
        <ul class="format-list">
            <li>PDF documents <span class="pdf-icon">📄</span></li>
            <li>DOCX documents <span class="docx-icon">📄</span></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    pii_detector = PIIDetector()
    
    # Upload section without separate card
    with st.container():
        st.markdown('<div class="upload-area">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Choose CV file (PDF or DOCX)", 
            type=["docx", "pdf"],
            help="Upload the candidate's resume in PDF or Word format"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Age input only, no separate card
    age = st.number_input("Enter age", min_value=18, max_value=99, step=1)
    
    # Processing section
    if uploaded_file and age:
        # Load logo
        try:
            logo_img = Image.open("asahi_logo-04.jpg")
        except FileNotFoundError:
            st.markdown("""
            <div class="status-warning">
                <strong>Warning:</strong> Logo file 'asahi_logo-04.jpg' not found. Please ensure it's in the same directory.
            </div>
            """, unsafe_allow_html=True)
            st.stop()
        
        # Extract text
        if uploaded_file.name.lower().endswith(".pdf"):
            raw_text = extract_text_from_pdf(uploaded_file)
        else:
            raw_text = extract_text_from_docx(uploaded_file)
        
        if not raw_text.strip():
            st.markdown("""
            <div class="status-warning">
                <strong>Error:</strong> No text could be extracted from the file. Please check the file format.
            </div>
            """, unsafe_allow_html=True)
            st.stop()
        
        # Show file loaded successfully in plain text
        st.markdown(f"""
        <div class="status-success">
            File loaded successfully: {uploaded_file.name} ({len(raw_text.split())} words)
        </div>
        """, unsafe_allow_html=True)
        
        # Single process button
        if st.button("Process CV", use_container_width=True):
            with st.spinner("Processing CV..."):
                # Detect PII
                detected_pii = pii_detector.detect_all_pii(raw_text)
                total_pii_items = sum(len(items) for items in detected_pii.values())
                
                # Show progress
                progress_html = """
                <div style="margin: 1.5rem 0;">
                    <div class="progress-step"><span class="step-check">✓</span> File uploaded and text extracted</div>
                    <div class="progress-step"><span class="step-check">✓</span> Personal information detected and removed</div>
                    <div class="progress-step"><span class="step-check">✓</span> Professional formatting applied</div>
                    <div class="progress-step"><span class="step-check">✓</span> Document ready for download</div>
                </div>
                """
                st.markdown(progress_html, unsafe_allow_html=True)
                
                # Clean the text
                cleaned_text, removal_count = pii_detector.remove_pii(raw_text, detected_pii)
                
                # Generate document using extracted name from CV
                detected_names = pii_detector.detect_names(raw_text)
                candidate_name = detected_names[0] if detected_names else "Candidate"
                
                final_doc = generate_asahi_cv(cleaned_text, logo_img, candidate_name, age)
                buffer = BytesIO()
                final_doc.save(buffer)
                buffer.seek(0)
                
                # Show completion message
                st.markdown(f"""
                <div class="status-success">
                    CV Processing Complete! Personal information removed: {removal_count} items. Final document ready with professional Asahi formatting.
                </div>
                """, unsafe_allow_html=True)
                
                # Simple metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"""
                    <div class="metric-item">
                        <div class="metric-number">{removal_count}</div>
                        <div class="metric-label">Items Removed</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.markdown(f"""
                    <div class="metric-item">
                        <div class="metric-number">{len(cleaned_text.split())}</div>
                        <div class="metric-label">Final Words</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col3:
                    st.markdown(f"""
                    <div class="metric-item">
                        <div class="metric-number">{len(detected_pii)}</div>
                        <div class="metric-label">PII Categories</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Download with filename format
                st.markdown("<br/>", unsafe_allow_html=True)
                file_name = f"Asahi_CV_{abbreviate_name_age(candidate_name, age)}.docx"
                st.download_button(
                    label="Download Formatted CV",
                    data=buffer,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Optional: Show what was removed (collapsible) - without showing names
                if total_pii_items > 0:
                    with st.expander("View what was removed (optional)"):
                        st.markdown("""
                        <div class="detection-summary">
                            <h4 style="margin-bottom: 1rem; color: #92400e;">Detected Personal Information:</h4>
                        """, unsafe_allow_html=True)
                        
                        for pii_type, items in detected_pii.items():
                            if items:
                                pii_type_display = pii_type.replace('_', ' ').title()
                                st.markdown(f"""
                                <div class="detection-item">
                                    <strong>{pii_type_display}:</strong> {len(items)} item(s) found
                                </div>
                                """, unsafe_allow_html=True)
                        
                        st.markdown("</div>", unsafe_allow_html=True)
    
    elif uploaded_file or age:
        st.markdown("""
        <div class="status-info">
            <strong>Ready to process:</strong> Please provide all required information above to continue.
        </div>
        """, unsafe_allow_html=True)

    # Footer
    st.markdown('<div class="footer">Made with ❤️ for professional CV formatting</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
