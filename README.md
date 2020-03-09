# count_words_multiple_docx
This is a very simple Python script which counts the overall amount of words in a set of .docx files.
It will also generate a CSV containing a frequency list over all files.

I tried to keep the dependencies to a minimum. If you need a better tokenizer, I I would recommend replacing the regular expression with [spaCy](https://spacy.io/).

```
pip install -r requirements.txt
python count_words_multiple_docx.py FOLDER_CONTAINING_DOCX_FILES
```
