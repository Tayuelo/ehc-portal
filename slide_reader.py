from utils import has_hyperlink

def get_slide_type(slide):
    paragraphs = get_slide_paragraphs(slide.shapes)
    runs = get_slide_runs(paragraphs)

    for run in runs:
        if run.font.bold:
            return "FIRST"
        if has_hyperlink(run):
            return "LAST"
    return "REGULAR"

def get_slide_text(slide):
    paragraphs = get_slide_paragraphs(slide.shapes)
    runs = get_slide_runs(paragraphs)

    return " ".join(run.text.strip() for run in runs).replace("  ", " ")

def get_slide_paragraphs(shapes):
    return [
        paragraph
        for shape in shapes if shape.has_text_frame
        for paragraph in shape.text_frame.paragraphs
    ]

def get_slide_runs(paragraphs):
    return [run for p in paragraphs for run in p.runs]
