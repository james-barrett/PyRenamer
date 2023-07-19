import fitz


def pdf_text_finder(file):
    doc = fitz.open(file)
    for page in doc:
        wlist = page.get_text_words()
        return wlist

