
from docx import Document
from docx.shared import Pt

# Paragraph-Level Formatting: 

# Applies to the entire paragraph and affects aspects like alignment, line spacing, and indentation. 
# It does not directly change the formatting of text within the paragraph's 
# runs, such as font size, bold, italic, etc., unless those runs are explicitly formatted to "inherit" from the paragraph.


# Run-Level Formatting: 

# Applies to specific segments of text within a paragraph, controlling font type, size, bold, italic, underline, and color. 
# Run-level formatting takes precedence over paragraph-level formatting for these attributes.
    
def paragraph_and_runs_formatting(file_path):
    doc = Document(file_path)
    text_and_formatting = {}

    for para in doc.paragraphs:
        para_text = para.text

        format = para.paragraph_format
        para_format = {
            "alignment": para.alignment,
            "style": para.style,
            "left_indent": format.left_indent,
            "right_indent":format.right_indent,
            "first_line_indent":format.first_line_indent
        }
        runs = []
        for run in para.runs:
            run_details = {
                "text": run.text,
                "bold": run.bold,
                "italic": run.italic,
                "underline": run.underline,
                "font": run.font.name,
                "size": run.font.size,
            }
            runs.append(run_details)
        text_and_formatting[para_text] = [para_format, runs]
    return text_and_formatting

def create_docx(paragraph_runs):
    doc = Document()
    for text, run_styles in paragraph_runs.items():
        paragraph = doc.add_paragraph()
        for run_sequence in run_styles:

            run = paragraph.add_run(run_sequence['text'])
            if run_sequence['bold'] is not None:
                run.bold = run_sequence['bold']
            if run_sequence['italic'] is not None:
                run.italic = run_sequence['italic']
            if run_sequence['underline'] is not None:
                run.underline = run_sequence['underline']
            if run_sequence['font'] is not None:
                run.font.name = run_sequence['font']
            if run_sequence['size'] is not None:
                run.font.size = Pt(run_sequence['size'] / 12700)

        if paragraph['style'] in doc.styles:
            paragraph.style = doc.styles[paragraph['style']]
        if paragraph['alignment'] is not None:
            paragraph.alignment = paragraph['alignment']
        if paragraph['left_indent'] is not None:
            paragraph.paragraph_format.left_indent = paragraph['left_indent']
        if paragraph['right_indent'] is not None:
            paragraph.paragraph_format.right_indent = paragraph['right_indent']
        if paragraph['first_line_indent'] is not None:
            paragraph.paragraph_format.first_line_indent = paragraph['first_line_indent']

        doc.save()
    doc.save()




if (__name__ == "__main__"):
    input_file ="input_file/reformatting env(3).docx"
    text_formatting = paragraph_and_runs_formatting(input_file)
    for i, k in text_formatting.items():
        print(f"{i}: {k}\n")