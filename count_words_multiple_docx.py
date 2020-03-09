import sys
import re
import docx
from collections import Counter
from pathlib import Path
from tqdm import tqdm


def count_docx_file(path):
    '''Returns a tuple consisting of the document length and a frequency list.'''
    document = docx.Document(path)
    text = []

    for p in document.paragraphs:
        text.append(p.text)

    text = '\n'.join(text)

    # Very crude tokenization
    words = re.split('\s+', text)
    frequency_list = Counter(words)

    return (len(words), frequency_list)

if __name__ == "__main__":
    try:
        folder = Path(sys.argv[1])
    except IndexError:
        print('Please provide a path. python count_words_multiple_docx.py PATH')
        sys.exit()

    docx_files = list(folder.glob('*.docx'))
    if len(docx_files) == 0:
        print('There appear to be no .docx files in this folder.')
        sys.exit()


    frequency_list = Counter()
    overall_count = 0

    for docx_path in tqdm(docx_files):
        docx_count, docx_frequency_list = count_docx_file(docx_path)

        overall_count += docx_count
        frequency_list += docx_frequency_list


    with open('frequency_list.csv', 'w') as frequency_list_file:
        for k,v in frequency_list.most_common():
            frequency_list_file.write(f'{k.strip()};{v}\n')

    print(f'There are approximately {overall_count} words in {len(docx_files)} files.')