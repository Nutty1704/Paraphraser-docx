import requests
from config import headers
import os
from docx.api import Document as DocumentRead
from docx import Document as DocumentWrite


def papaphrase(text: str) -> str:
    url = "https://rewriter-paraphraser-text-changer-multi-language.p.rapidapi.com/rewrite"

    payload = {
        "language": "en",
        "strength": 3,
        "text": text
    }

    response = requests.post(url, json=payload, headers=headers)
    try:
        ret = response.json()['rewrite']
    except:
        ret = text
    return ret
    

def get_word_file() -> str:
    while True:
        file_path = input('Enter the file path: ')
        if os.path.exists(file_path):
            if file_path.endswith('.docx'):
                return file_path
            else:
                print('Invalid file extension. Try again.')
        else:
            print('Invalid file path. Try again.')


def main():
    file_path = get_word_file()
    # read the file
    doc = DocumentRead(file_path)
        
    new_path = file_path.rstrip('.docx') + f'-paraphrase.docx'
    new_doc = DocumentWrite()

    # paraphrase the file
    for p in doc.paragraphs:
        paraphrased = papaphrase(p.text)

        new_p = new_doc.add_paragraph(paraphrased)
        new_p.style = p.style

    new_doc.save(new_path)
    print(f'Paraphrased file saved to {new_path}')


if __name__ == "__main__":
    main()

