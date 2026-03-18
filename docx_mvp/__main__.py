from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from .workflow import edit_docx


ROOT = Path.cwd()
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
CONSOLE_LOG_MAX_CHARS = 240


def build_paths(filename: str, output_name: str | None) -> tuple[Path, Path]:
    raw_input_path = Path(filename).expanduser()
    input_path = raw_input_path if raw_input_path.is_absolute() else INPUT_DIR / filename
    name = output_name or input_path.name
    output_path = OUTPUT_DIR / name / name
    return input_path, output_path


def build_log_path(output_path: Path) -> Path:
    return output_path.parent / f"{output_path.name}.log"


def format_console_message(message: str, max_chars: int = CONSOLE_LOG_MAX_CHARS) -> str:
    single_line = " ".join(message.splitlines())
    if len(single_line) <= max_chars:
        return single_line
    return f"{single_line[: max_chars - 3]}..."


def make_logger(log_path: Path):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text("", encoding="utf-8")

    def log(message: str) -> None:
        timestamped = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
        print(format_console_message(timestamped))
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"{timestamped}\n")

    return log


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("filename")
    parser.add_argument("prompt")
    parser.add_argument("-o", "--output-name")
    parser.add_argument("--max-rounds", type=int, default=20)
    args = parser.parse_args()
    input_path, output_path = build_paths(args.filename, args.output_name)
    if not input_path.exists():
        raise FileNotFoundError(f"input file not found: {input_path}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log = make_logger(build_log_path(output_path))
    log(f"input: {input_path.resolve()}")
    log(f"prompt: {args.prompt}")
    failures = edit_docx(str(input_path), str(output_path), args.prompt, args.max_rounds, log=log)
    log(f"written: {output_path.resolve()}")
    if failures:
        log(f"remaining failures: {failures}")
    log(f"output: {output_path.resolve()}")


if __name__ == "__main__":
    main()
