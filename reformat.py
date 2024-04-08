from docx import Document

doc = Document('input_file/reformatting env(1).docx')
run_style = []

def check_footnote(paragraph):
    # Possibilities
    # No footnotes, one footnotes, multiple footnotes
    paragraph_xml = paragraph._element.xml
    footnote = False
    if '<w:footnoteReference' in paragraph_xml: footnote = True
    return footnote

def check_indentation(paragraph):
    pass

for p in doc.paragraphs:

    para_text = p.text
    para_format = {
        "alignment":p.alignment,
        "style":p.style
    }
    for runs in p.runs:
        run_style.append([p.text, p.style, {"size":runs.font.size, "bold":runs.bold, "italic":runs.italic, "underline":runs.underline, "footnote":check_footnote(p)}])

for i in run_style:print(f"{i}\n")



