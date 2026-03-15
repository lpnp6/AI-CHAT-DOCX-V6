# AI-CHAT-DOCX-V6

Minimal DOCX editing MVP that asks an LLM for XPath-based text replacement instructions and applies them to `word/document.xml`.

## Requirements

- Python 3.11+
- `OPENAI_API_KEY`
- Optional: `OPENAI_BASE_URL`, `OPENAI_MODEL`

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Usage

Put the source `.docx` file under `input/`, then run:

```bash
docx-edit-mvp demo.docx "replace the company name"
```

Output files are written under `output/<name>/`.

## Tests

```bash
python -m unittest
```
