import fitz


def pdf_text_finder(file):
    doc = fitz.open(file)
    for page in doc:
        wlist = page.get_text_words()
        return wlist


# print(dev_helper.pdf_text_finder(os.path.join(folder_path, file)))
