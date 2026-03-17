# Sample Commands

```bash
docfreq myfile.docx
docfreq chapter1.docx chapter2.docx notes.txt
docfreq myfile.docx --top 20 --plot
docfreq myfile.docx --ngrams 1,2,3 --min-count 3
docfreq chapter1.docx chapter2.docx --ngrams 1,2 --plot
docfreq myfile.docx --dump-text extracted.txt
docfreq myfile.docx --dump-text
docfreq myfile.docx --dump-text -
docfreq --print-completion zsh
docfreq myfile.docx --csv out.csv
docfreq myfile.docx --keep-stopwords --no-lemma
DOCFREQ_PASSWORD='secret' docfreq encrypted.docx --plot
docfreq notes.txt --plot
```
