#AI提示词：请以HTML格式输出，保留纯文本内容
#环境设置：


# Python Script: Convert AI-Generated HTML (from TXT) to PDF using ReportLab and HTMLParser

# Goal: Read a TXT file (first line "【Title】", rest is HTML body).
# Parse the HTML content using Python's html.parser.
# Convert HTML elements to ReportLab Flowables with defined styles.
# Generate a PDF visually approximating the HTML, avoiding WeasyPrint/GTK+.
# Auto-open the PDF.

# ----- Configuration -----
import os
import re
import platform
import subprocess
from pathlib import Path
from html.parser import HTMLParser
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, ListFlowable, ListItem
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors

# Input TXT file path
txt_file_path = r'C:\Users\fucha\PycharmProjects\python\公司\cursor项目\4.文件夹处理\AI格式转化\text.txt'
# Script directory
script_dir = Path(__file__).parent

# --- Font Configuration (Requires Noto Sans SC in ./fonts folder) ---
# [Keep the font registration block as it was in the previous ReportLab versions,
#  ensuring NotoSansSC-Regular and NotoSansSC-Bold are registered.]
print("Registering Noto Sans SC font...")
try:
    font_dir = script_dir / "fonts"
    font_path_regular = font_dir / "NotoSansSC-Regular.ttf"
    font_path_bold = font_dir / "NotoSansSC-Bold.ttf"

    if not font_path_regular.exists(): raise FileNotFoundError(f"Font not found: {font_path_regular}")
    if not font_path_bold.exists(): raise FileNotFoundError(f"Font not found: {font_path_bold}")

    registered_font_name_regular = 'NotoSansSC-Regular'
    registered_font_name_bold = 'NotoSansSC-Bold'

    pdfmetrics.registerFont(TTFont(registered_font_name_regular, font_path_regular))
    pdfmetrics.registerFont(TTFont(registered_font_name_bold, font_path_bold))
    pdfmetrics.registerFontFamily('NotoSansSC', normal=registered_font_name_regular, bold=registered_font_name_bold)
    print(f"Successfully registered fonts and family 'NotoSansSC'")

except FileNotFoundError as e: print(f"Error: {e}\nPlease ensure Noto fonts are in './fonts'."); exit(1)
except Exception as e: print(f"Error registering font: {e}"); exit(1)


# --- Define ReportLab Paragraph Styles (Adjust as needed) ---
styles = getSampleStyleSheet()
# Base font for styles will be NotoSansSC if correctly registered and used
default_font = 'NotoSansSC' # This should match the family name registered

# H1 Style (for major task titles if AI uses <h1> or similar via <p><strong>)
styles.add(ParagraphStyle(
    name='PDF_H1', parent=styles['Heading1'], fontName=registered_font_name_bold,
    fontSize=16, leading=16*1.4, spaceBefore=10*mm, spaceAfter=5*mm,
    textColor=colors.HexColor('#111111'), alignment=TA_LEFT
))
# H2 Style (for sub-titles like "干什么:")
styles.add(ParagraphStyle(
    name='PDF_H2', parent=styles['Heading2'], fontName=registered_font_name_bold,
    fontSize=13, leading=13*1.4, spaceBefore=6*mm, spaceAfter=3*mm,
    textColor=colors.HexColor('#222222'), alignment=TA_LEFT, leftIndent=5*mm
))
# Paragraph Style
styles.add(ParagraphStyle(
    name='PDF_P', parent=styles['Normal'], fontName=registered_font_name_regular,
    fontSize=11, leading=11*1.5, spaceAfter=3*mm,
    textColor=colors.HexColor('#333333'), alignment=TA_LEFT, leftIndent=0*mm # Base paragraphs no extra indent beyond H2
))
# Paragraph Style for content under H2 that is not a list
styles.add(ParagraphStyle(
    name='PDF_P_Indent1', parent=styles['PDF_P'], leftIndent=10*mm # Further indent for these
))
# List Item Style
styles.add(ParagraphStyle(
    name='PDF_LI', parent=styles['Normal'], fontName=registered_font_name_regular,
    fontSize=11, leading=11*1.5, spaceAfter=2*mm, leftIndent=15*mm, bulletIndent=10*mm
))

# --- Custom HTML Parser to build ReportLab Story ---
class ReportLabHTMLParser(HTMLParser):
    def __init__(self, styles):
        super().__init__()
        self.styles = styles
        self.story = []
        self.current_data = [] # Accumulate text data for current element
        self.tag_stack = []    # Keep track of nested tags for context

    def handle_starttag(self, tag, attrs):
        self.tag_stack.append(tag)
        # For ReportLab Paragraph, <b> and <i> are usually handled directly within the text.
        # More complex tags might need specific handling here.
        if tag in ('p', 'h1', 'h2', 'h3', 'li'): # Flush previous data before starting new block
            self._flush_data_as_paragraph()

    def handle_endtag(self, tag):
        if self.tag_stack and self.tag_stack[-1] == tag:
            self.tag_stack.pop()
        else:
            # This can happen with malformed HTML, ignore for now or log
            pass

        if tag in ('p', 'h1', 'h2', 'h3', 'li', 'body'): # Block level tags that trigger paragraph creation
            self._flush_data_as_paragraph()
        elif tag == 'br':
            self.current_data.append("<br/>") # Pass <br/> to Paragraph

    def handle_data(self, data):
        # Prepend tags like <b>, <i> to the data if they are active
        prefix = ""
        if 'strong' in self.tag_stack or 'b' in self.tag_stack:
            # Check if data already starts with <b>, avoid double tagging
            if not (data.strip().lower().startswith("<b>") and self.current_data and self.current_data[-1].strip().lower().endswith("<b>")):
                 prefix = "<b>"
        # Add more for <i>, <u> etc. if needed

        self.current_data.append(prefix + data.replace('\n', ' ').strip()) # Normalize whitespace

        if 'strong' in self.tag_stack or 'b' in self.tag_stack:
            if not (data.strip().lower().endswith("</b>") and self.current_data and self.current_data[-1].strip().lower().startswith("</b>")):
                self.current_data.append("</b>")


    def _flush_data_as_paragraph(self):
        if not self.current_data:
            return

        text_content = "".join(self.current_data).strip()
        self.current_data = []

        if not text_content:
            return

        # Determine style based on the most recent block tag in stack or context
        # This part is a simplification and might need much more sophisticated logic
        # for a generic HTML parser. For the given AI HTML, it might work.
        applied_style = self.styles['PDF_P'] # Default to paragraph
        last_block_tag = None
        for t in reversed(self.tag_stack): # Check parent tags
            if t in ('p', 'h1', 'h2', 'h3', 'li'):
                last_block_tag = t
                break
        
        # Rudimentary style mapping based on HTML structure in the example
        # The AI HTML example mostly uses <p> and <strong>.
        # We need to infer structure (e.g. a <p> with only <strong> inside is a heading)
        if "<strong>任务" in text_content: # Heuristic for Task titles
            applied_style = self.styles['PDF_H1']
            text_content = text_content.replace("<strong>", "").replace("</strong>", "") # Remove strong if style is already bold
        elif re.match(r"^\s*<strong>(干什么|怎么干|怎么算完成|需要记录什么/怎么给你反馈):", text_content, re.IGNORECASE):
            applied_style = self.styles['PDF_H2']
            # text_content = text_content # Keep <strong> for partial bolding, H2 style is bold
        elif re.match(r"^\s*\d+\.", text_content): # Heuristic for list items
            applied_style = self.styles['PDF_LI']
            # For numbered lists, ReportLab's Paragraph doesn't auto-number.
            # We'd need ListFlowable for proper lists, or keep numbers in text.
            # For simplicity, keeping numbers in text.
        elif last_block_tag == 'p' and self.story and isinstance(self.story[-1], Paragraph) and self.story[-1].style.name == 'PDF_H2':
            # If previous was H2, this paragraph under it should be indented
            applied_style = self.styles['PDF_P_Indent1']


        # Handle <p>---</p> as a spacer or horizontal rule
        if text_content.strip() == "---":
            self.story.append(Spacer(1, 8*mm)) # Add a visual separator
            return

        # Remove leftover/empty <b></b> pairs if any
        text_content = re.sub(r"<b>\s*</b>", "", text_content).strip()
        if text_content:
            self.story.append(Paragraph(text_content, applied_style))

    def get_story(self):
        self._flush_data_as_paragraph() # Ensure any remaining data is flushed
        return self.story


# --- Function to Sanitize Filename ---
# [Keep sanitize_filename function as before]
def sanitize_filename(name):
    name = name.strip('【】')
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    return name

# --- Function to automatically open the PDF ---
# [Keep open_file function as before]
def open_file(filepath):
    try:
        filepath_str = str(filepath)
        if platform.system() == 'Darwin': subprocess.call(('open', filepath_str))
        elif platform.system() == 'Windows': os.startfile(os.path.abspath(filepath_str))
        else: subprocess.call(('xdg-open', filepath_str))
        print(f"Attempting to open {filepath_str}...")
    except FileNotFoundError: print(f"Error: Could not find {filepath_str} to open.")
    except Exception as e: print(f"Error opening file {filepath_str}: {e}")

# --- Main Execution Logic ---
if __name__ == "__main__":
    print(f"Reading text file: {txt_file_path}")
    output_filename = "ai_html_parsed_output.pdf" # New default
    html_body_content_from_txt = ""

    try:
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            first_line = f.readline().strip()
            if first_line.startswith("【") and first_line.endswith("】"):
                 pdf_title_text = sanitize_filename(first_line)
                 output_filename = f"{pdf_title_text}.pdf"
                 html_body_content_from_txt = f.read()
                 print(f"Using PDF title and filename: {output_filename}")
            else:
                 print("First line not in 【Title】 format. Using default filename.")
                 f.seek(0)
                 html_body_content_from_txt = f.read()
            output_pdf_path = script_dir / output_filename
    except FileNotFoundError: print(f"Error: Input file not found at {txt_file_path}"); exit(1)
    except Exception as e: print(f"Error reading file: {e}"); exit(1)

    print("Parsing HTML content and preparing ReportLab story...")
    parser = ReportLabHTMLParser(styles)
    parser.feed(html_body_content_from_txt) # Feed the HTML body
    story = parser.get_story()

    # --- Sanity check: Print the generated story structure (for debugging) ---
    # print("\n--- Generated Story ---")
    # for item in story:
    #     if isinstance(item, Paragraph):
    #         print(f"Style: {item.style.name}, Text: '{item.text[:50]}...'")
    #     else:
    #         print(type(item))
    # print("--- End Story ---\n")


    print(f"Generating PDF: {output_pdf_path}")
    try:
        doc = SimpleDocTemplate(str(output_pdf_path),
                                leftMargin=15*mm, rightMargin=15*mm,
                                topMargin=20*mm, bottomMargin=20*mm)
        doc.build(story)
        print("PDF generated successfully!")
        open_file(output_pdf_path)
    except Exception as e:
        print(f"\nError generating PDF with ReportLab: {e}")
        # print("\n--- Story Content at Error ---") # For debugging if build fails
        # for i, item in enumerate(story):
        #     if isinstance(item, Paragraph):
        #         print(f"Item {i}: Style={item.style.name}, Font={item.style.fontName}, Text='{item.text[:100]}...'")
        #     else:
        #         print(f"Item {i}: {type(item)}")
        exit(1)