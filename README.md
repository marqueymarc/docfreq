# docfreq

`docfreq` is a macOS-first command-line tool for extracting text from Word documents and reporting smart word frequencies and n-gram counts. It uses Docling for `.docx` extraction, spaCy for tokenization and lemmatization, optional CSV export for Excel, and `termgraph` for terminal plots.

It also supports plain `.txt` and `.md` inputs, encrypted `.docx` files when you provide a password, and concatenating multiple inputs into one combined analysis.

## What "smart" means here

- lemmatization by default
- stopword filtering by default
- alpha-only token filtering
- tokens shorter than 3 characters removed
- optional unigram, bigram, and trigram counting

## Requirements

- macOS
- Homebrew
- Python 3.12 recommended

System packages used by the workflow:

```bash
brew update
brew install python pipx gnuplot youplot
pipx ensurepath
```

## Setup

Create the virtual environment and install dependencies:

```bash
/opt/homebrew/bin/python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
pip install https://github.com/explosion/spacy-models/releases/download/en_core_web_sm-3.8.0/en_core_web_sm-3.8.0-py3-none-any.whl
chmod +x docfreq docfreq.py
```

## Usage

Run the wrapper command from the project directory:

```bash
./docfreq myfile.docx
./docfreq chapter1.docx chapter2.docx notes.txt
```

Command options:

- positional input paths: one or more `.docx`, `.txt`, or `.md` files
- `--top N`: number of rows to print, default `30`
- `--ngrams 1,2,3`: choose unigram, bigram, and trigram counting, default `1`
- `--min-count N`: minimum count threshold, default `2`
- `--keep-stopwords`: keep stopwords
- `--no-lemma`: disable lemmatization
- `--password`: password for encrypted Office files
- `--csv PATH`: write results to CSV
- `--plot`: render terminal charts for the requested n-gram results

Example:

```bash
./docfreq myfile.docx mynotes.txt --top 25 --ngrams 1,2 --csv freq.csv --plot
```

For encrypted `.docx` files, either pass the password directly:

```bash
./docfreq secret.docx --password 'your-password' --plot
```

or use an environment variable so the password does not appear in shell history:

```bash
DOCFREQ_PASSWORD='your-password' ./docfreq secret.docx --plot
```

## Output

- ranked frequencies are printed to stdout from the concatenated input text
- CSV output is written with columns `kind`, `term`, and `count`
- terminal plotting uses `termgraph` for each requested n-gram kind that has results

If you want charts in Excel, open the generated CSV there and create a bar chart from the exported counts.

## Notes

- `.docx` extraction is handled through `docling.document_converter.DocumentConverter`
- when multiple files are passed, `docfreq` concatenates their extracted text before counting
- if `termgraph` is unavailable, run without `--plot`
- Python 3.13+ may currently resolve older Docling builds; the setup above uses Python 3.12 because that was the stable path during verification
