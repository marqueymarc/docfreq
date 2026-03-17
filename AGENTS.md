# Repo Instructions

## Versioning Policy

- Use semantic versions in the form `major.minor.rev`.
- The current major is `1`.
- The current minor is `1`.
- Start the revision track at `9` and increment `rev` for small changes.
- Stamp each shipped change with a UTC/GMT build string in the form `YYYY.MM.DD.HHMM`.
- Keep the package version in `pyproject.toml` and the CLI version output in `docfreq.py` aligned.
- `docfreq --version` must print both the semantic version and the UTC build string.
