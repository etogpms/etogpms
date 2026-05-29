import PyPDF2
import sys

def extract_text(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text() + "\n"
        with open('output.txt', 'w', encoding='utf-8') as out_file:
            out_file.write(text)
        print("Done")

if __name__ == '__main__':
    extract_text('../assets/20260310 MWC Daily Service Update Report (1).pdf')
