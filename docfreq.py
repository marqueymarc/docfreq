#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import getpass
import io
import os
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path
from typing import Iterable

import spacy


APP_NAME = "docfreq"
APP_VERSION = "1.1.9"
APP_BUILD_DATE = "2026.03.17.gmt00"
APP_VERSION_STRING = f"{APP_NAME} {APP_VERSION} ({APP_BUILD_DATE})"

SUPPORTED_EXTENSIONS = {".docx", ".txt", ".md"}
SUPPORTED_SHELLS = ("bash", "zsh")


class DocfreqError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract text from a document and report smart term frequencies."
    )
    parser.add_argument(
        "input_paths",
        nargs="*",
        help="One or more paths to .docx, .txt, or .md files",
    )
    parser.add_argument("--top", type=int, default=30, help="Number of top results to show")
    parser.add_argument(
        "--ngrams",
        default="1",
        help="Comma-separated n-gram sizes to count, for example 1,2,3",
    )
    parser.add_argument(
        "--min-count",
        type=int,
        default=2,
        help="Minimum count required for a term to be included",
    )
    parser.add_argument(
        "--keep-stopwords",
        action="store_true",
        help="Keep stopwords instead of filtering them out",
    )
    parser.add_argument(
        "--no-lemma",
        action="store_true",
        help="Use the original token text instead of spaCy lemmas",
    )
    parser.add_argument(
        "--password",
        help="Password for encrypted Office files. You can also set DOCFREQ_PASSWORD.",
    )
    parser.add_argument(
        "--keep-words-file",
        help="Path to a file containing words that should not be treated as stopwords. You can also set DOCFREQ_KEEP_WORDS_FILE.",
    )
    parser.add_argument(
        "--dump-text",
        dest="dump_text_path",
        nargs="?",
        const="-",
        help="Write the combined extracted text before counting. Use no value, '-' or /dev/stdout for stdout.",
    )
    parser.add_argument(
        "--print-completion",
        choices=SUPPORTED_SHELLS,
        help="Print a shell completion script for bash or zsh and exit.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=APP_VERSION_STRING,
    )
    parser.add_argument("--csv", dest="csv_path", help="Optional CSV output path")
    parser.add_argument(
        "--plot",
        action="store_true",
        help="Render terminal bar charts for the requested n-gram results",
    )
    if hasattr(parser, "parse_intermixed_args"):
        return parser.parse_intermixed_args()
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> list[int]:
    if args.top < 1:
        raise DocfreqError("--top must be at least 1")
    if args.min_count < 1:
        raise DocfreqError("--min-count must be at least 1")

    try:
        n_values = sorted({int(value.strip()) for value in args.ngrams.split(",") if value.strip()})
    except ValueError as exc:
        raise DocfreqError("--ngrams must be a comma-separated list of integers") from exc

    if not n_values:
        raise DocfreqError("--ngrams must include at least one value")
    invalid = [value for value in n_values if value < 1 or value > 3]
    if invalid:
        raise DocfreqError("Only unigram, bigram, and trigram counts are supported")

    return n_values


def resolve_password(password: str | None) -> str | None:
    if password:
        return password
    env_password = os.getenv("DOCFREQ_PASSWORD")
    if env_password:
        return env_password
    return None


def resolve_keep_words_file(path: str | None) -> str | None:
    if path:
        return path
    env_path = os.getenv("DOCFREQ_KEEP_WORDS_FILE")
    if env_path:
        return env_path
    return None


def print_completion(shell: str) -> None:
    scripts = {
        "zsh": """#compdef docfreq
_docfreq() {
  _arguments -s \\
    '--top[Number of top results to show]:count:' \\
    '--ngrams[Comma-separated n-gram sizes to count]:n-grams:' \\
    '--min-count[Minimum count required for a term]:count:' \\
    '--keep-stopwords[Keep stopwords instead of filtering them out]' \\
    '--no-lemma[Use the original token text instead of spaCy lemmas]' \\
    '--password[Password for encrypted Office files]:password:' \\
    '--keep-words-file[File containing words to keep out of stopword filtering]:word file:_files' \\
    '--dump-text=-[Write extracted text before counting]:output path:_files' \\
    '--csv[Write results to CSV]:csv file:_files' \\
    '--plot[Render terminal bar charts for requested n-grams]' \\
    '--print-completion[Print a shell completion script and exit]:shell:(bash zsh)' \\
    '--version[Print version information and exit]' \\
    '*:input file:_files'
}
compdef _docfreq docfreq
""",
        "bash": """_docfreq_completion() {
  local cur prev
  COMPREPLY=()
  cur="${COMP_WORDS[COMP_CWORD]}"
  prev="${COMP_WORDS[COMP_CWORD-1]}"

  case "$prev" in
    --top|--min-count)
      return 0
      ;;
    --ngrams)
      COMPREPLY=( $(compgen -W "1 1,2 1,2,3 2 3" -- "$cur") )
      return 0
      ;;
    --password|--keep-words-file|--dump-text|--csv)
      COMPREPLY=( $(compgen -f -- "$cur") )
      return 0
      ;;
    --print-completion)
      COMPREPLY=( $(compgen -W "bash zsh" -- "$cur") )
      return 0
      ;;
  esac

  if [[ "$cur" == -* ]]; then
    COMPREPLY=( $(compgen -W "--top --ngrams --min-count --keep-stopwords --no-lemma --password --keep-words-file --dump-text --csv --plot --print-completion --version" -- "$cur") )
    return 0
  fi

  COMPREPLY=( $(compgen -f -- "$cur") )
}
complete -F _docfreq_completion docfreq
""",
    }
    sys.stdout.write(scripts[shell])
    if not scripts[shell].endswith("\n"):
        sys.stdout.write("\n")
    sys.stdout.flush()


def decrypt_docx_if_needed(path: Path, password: str | None) -> tuple[Path, Path | None]:
    with path.open("rb") as handle:
        header = handle.read(8)
        if not header.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
            return path, None

    try:
        import msoffcrypto
    except ImportError as exc:
        raise DocfreqError(
            "This .docx appears to be encrypted. Install msoffcrypto-tool support first."
        ) from exc

    password = resolve_password(password)
    if password is None and sys.stdin.isatty():
        password = getpass.getpass("Password for encrypted DOCX: ")
    if not password:
        raise DocfreqError(
            "This .docx is encrypted. Re-run with --password or set DOCFREQ_PASSWORD."
        )

    try:
        with path.open("rb") as handle:
            office_file = msoffcrypto.OfficeFile(handle)
            office_file.load_key(password=password)
            decrypted_bytes = io.BytesIO()
            office_file.decrypt(decrypted_bytes)
    except Exception as exc:
        raise DocfreqError(f"Unable to decrypt encrypted DOCX: {exc}") from exc

    with tempfile.NamedTemporaryFile("wb", suffix=".docx", delete=False) as handle:
        handle.write(decrypted_bytes.getvalue())
        temp_path = Path(handle.name)

    return temp_path, temp_path


def extract_text(path: Path, password: str | None = None) -> str:
    if not path.exists():
        raise DocfreqError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise DocfreqError(f"Unsupported file type '{suffix}'. Supported types: {supported}")

    if suffix in {".txt", ".md"}:
        return path.read_text(encoding="utf-8")

    docling_path, temp_path = decrypt_docx_if_needed(path, password)
    try:
        from docling.document_converter import DocumentConverter

        converter = DocumentConverter()
        result = converter.convert(docling_path)
        if hasattr(result, "document") and hasattr(result.document, "export_to_text"):
            return result.document.export_to_text()
        if hasattr(result, "render_as_text"):
            return result.render_as_text()
        raise DocfreqError("Docling returned a result object without a text export method")
    except Exception as exc:  # pragma: no cover - depends on third-party conversion errors
        raise DocfreqError(f"Docling failed to convert {path}: {exc}") from exc
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def extract_combined_text(paths: list[Path], password: str | None = None) -> str:
    return "\n\n".join(extract_text(path, password=password) for path in paths)


def load_keep_words(path: str | None) -> set[str]:
    resolved = resolve_keep_words_file(path)
    if not resolved:
        return set()

    keep_words_path = Path(resolved).expanduser()
    if not keep_words_path.exists():
        raise DocfreqError(f"Keep-words file not found: {keep_words_path}")

    words = {
        word.lower()
        for word in keep_words_path.read_text(encoding="utf-8").split()
        if word.strip()
    }
    return words


def load_nlp(keep_words_file: str | None = None):
    try:
        nlp = spacy.load("en_core_web_sm", disable=["ner", "parser"])
        for word in load_keep_words(keep_words_file):
            nlp.vocab[word].is_stop = False
        return nlp
    except OSError as exc:
        raise DocfreqError(
            "spaCy model 'en_core_web_sm' is not installed. "
            "Run: python -m spacy download en_core_web_sm"
        ) from exc


def normalize_tokens(doc, keep_stopwords: bool, no_lemma: bool) -> list[str]:
    tokens: list[str] = []
    for token in doc:
        if token.is_space or token.is_punct or not token.is_alpha:
            continue
        if not keep_stopwords and token.is_stop:
            continue

        value = token.text if no_lemma else token.lemma_
        value = value.lower().strip()
        if not value.isalpha() or len(value) < 3:
            continue
        tokens.append(value)

    return tokens


def make_ngrams(tokens: list[str], n: int) -> list[str]:
    if n == 1:
        return tokens
    return [" ".join(tokens[index : index + n]) for index in range(len(tokens) - n + 1)]


def count_terms(tokens: list[str], n_values: Iterable[int], min_count: int, top: int) -> dict[str, list[tuple[str, int]]]:
    results: dict[str, list[tuple[str, int]]] = {}
    for n in n_values:
        label = {1: "unigram", 2: "bigram", 3: "trigram"}[n]
        counter = Counter(make_ngrams(tokens, n))
        pairs = [
            (term, count)
            for term, count in sorted(counter.items(), key=lambda item: (-item[1], item[0]))
            if count >= min_count
        ]
        results[label] = pairs[:top]
    return results


def print_results(results: dict[str, list[tuple[str, int]]]) -> None:
    for kind, pairs in results.items():
        print(f"\n{kind.title()}s")
        print("-" * len(f"{kind.title()}s"))
        if not pairs:
            print("No results")
            continue
        for term, count in pairs:
            print(f"{count:>5}  {term}")


def write_csv(results: dict[str, list[tuple[str, int]]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["kind", "term", "count"])
        for kind, pairs in results.items():
            for term, count in pairs:
                writer.writerow([kind, term, count])


def write_text_dump(text: str, destination: str) -> None:
    if destination in {"-", "/dev/stdout"}:
        sys.stdout.write(text)
        if not text.endswith("\n"):
            sys.stdout.write("\n")
        sys.stdout.flush()
        return

    output_path = Path(destination).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")


def plot_counts(pairs: list[tuple[str, int]]) -> None:
    if not pairs:
        print("No results available to plot.", file=sys.stderr)
        return

    termgraph = shutil.which("termgraph")
    if not termgraph:
        raise DocfreqError("The 'termgraph' command was not found. Install it or omit --plot.")

    with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, encoding="utf-8") as handle:
        temp_path = Path(handle.name)
        for term, count in pairs:
            handle.write(f"{term},{count}\n")

    try:
        sys.stdout.flush()
        subprocess.run(
            [termgraph, str(temp_path), "--delim", ",", "--width", "50"],
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        raise DocfreqError(f"termgraph failed: {exc}") from exc
    finally:
        temp_path.unlink(missing_ok=True)


def main() -> int:
    args = parse_args()

    try:
        if args.print_completion:
            print_completion(args.print_completion)
            return 0

        if not args.input_paths:
            raise DocfreqError("Provide at least one input path, or use --print-completion.")

        n_values = validate_args(args)
        input_paths = [Path(value).expanduser().resolve() for value in args.input_paths]
        text = extract_combined_text(input_paths, password=args.password)

        if args.dump_text_path is not None:
            write_text_dump(text, args.dump_text_path)
            if args.dump_text_path not in {"-", "/dev/stdout"}:
                print(f"Wrote extracted text to {args.dump_text_path}")

        nlp = load_nlp(args.keep_words_file)
        doc = nlp(text)
        tokens = normalize_tokens(doc, args.keep_stopwords, args.no_lemma)
        results = count_terms(tokens, n_values, args.min_count, args.top)

        print_results(results)

        if args.csv_path:
            write_csv(results, Path(args.csv_path).expanduser())
            print(f"\nWrote CSV to {args.csv_path}")

        if args.plot:
            plotted = False
            for kind in ("unigram", "bigram", "trigram"):
                pairs = results.get(kind, [])
                if not pairs:
                    continue
                plotted = True
                print(f"\nTop {kind} chart")
                print("-" * len(f"Top {kind} chart"))
                plot_counts(pairs)
            if not plotted:
                print("\nSkipping plot because no requested n-gram results were found.")
    except DocfreqError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
