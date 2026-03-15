from __future__ import annotations

import argparse
from pathlib import Path

from .workflow import edit_docx


ROOT = Path.cwd()
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"


def ask(value: str | None, label: str) -> str:
    return value if value is not None else input(f"{label}: ").strip()


def build_paths(filename: str, output_name: str | None) -> tuple[Path, Path]:
    input_path = INPUT_DIR / filename
    name = output_name or filename
    output_path = OUTPUT_DIR / name / name
    return input_path, output_path


def build_log_path(output_path: Path) -> Path:
    return output_path.parent / f"{output_path.name}.log"


def make_logger(log_path: Path):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text("", encoding="utf-8")

    def log(message: str) -> None:
        print(message)
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"{message}\n")

    return log


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("filename", nargs="?")
    parser.add_argument("prompt", nargs="?")
    parser.add_argument("-o", "--output-name")
    parser.add_argument("--max-rounds", type=int, default=3)
    args = parser.parse_args()
    filename = ask(args.filename, "input filename")
    prompt = ask(args.prompt, "prompt")
    input_path, output_path = build_paths(filename, args.output_name)
    if not input_path.exists():
        raise FileNotFoundError(f"input file not found: {input_path}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log = make_logger(build_log_path(output_path))
    log(f"input: {input_path}")
    log(f"output: {output_path}")
    log(f"prompt: {prompt}")
    failures = edit_docx(str(input_path), str(output_path), prompt, args.max_rounds, log=log)
    log(f"written: {output_path}")
    if failures:
        log(f"remaining failures: {failures}")


if __name__ == "__main__":
    main()
