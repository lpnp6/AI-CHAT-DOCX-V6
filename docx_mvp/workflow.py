from __future__ import annotations

from collections.abc import Callable

from .llm import InstructionFailure, LLM
from .executor import execute
from .package import DocxPackage


def instruction_key(instruction: object) -> tuple[object, object, object] | None:
    if not hasattr(instruction, "type") or not hasattr(instruction, "xpath") or not hasattr(instruction, "text"):
        return None
    return (instruction.type, instruction.xpath, instruction.text)


def failure_key(failure: InstructionFailure) -> tuple[object, object, object]:
    instruction = failure.instruction
    return (instruction.get("type"), instruction.get("xpath"), instruction.get("text"))


def edit_docx(
    input_path: str,
    output_path: str,
    prompt: str,
    max_rounds: int = 3,
    llm: LLM | None = None,
    log: Callable[[str], None] | None = None,
) -> list[InstructionFailure]:
    package = DocxPackage.load(input_path)
    failures: list[InstructionFailure] = []
    locked_xpaths: set[str] = set()
    llm = llm or LLM()
    for round_no in range(1, max_rounds + 1):
        if log:
            log(f"round {round_no}: generating instructions")
        instructions = llm.generate(
            package.document_xml,
            prompt,
            failures,
            locked_xpaths=sorted(locked_xpaths),
        )
        if log:
            if getattr(llm, "last_raw_output", ""):
                log(f"round {round_no}: model_output={llm.last_raw_output}")
            log(f"round {round_no}: instructions={instructions}")
        updated_xml, failures = execute(package.document_xml, instructions, locked_xpaths=locked_xpaths)
        failed_keys = {failure_key(item) for item in failures}
        for instruction in instructions:
            key = instruction_key(instruction)
            if key is None or key in failed_keys:
                continue
            locked_xpaths.add(instruction.xpath)
        package.document_xml = updated_xml
        if log:
            log(f"round {round_no}: failures={failures}")
        if not failures:
            break
    package.dump(output_path)
    return failures
