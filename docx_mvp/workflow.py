from __future__ import annotations

from collections.abc import Callable

from .agent import Agent, InstructionFailure
from .executor import execute
from .package import DocxPackage


def edit_docx(
    input_path: str,
    output_path: str,
    prompt: str,
    max_rounds: int = 3,
    agent: Agent | None = None,
    log: Callable[[str], None] | None = None,
) -> list[InstructionFailure]:
    package = DocxPackage.load(input_path)
    failures: list[InstructionFailure] = []
    agent = agent or Agent()
    for round_no in range(1, max_rounds + 1):
        if log:
            log(f"round {round_no}: generating instructions")
        instructions = agent.generate(package.document_xml, prompt, failures)
        if log:
            if getattr(agent, "last_raw_output", ""):
                log(f"round {round_no}: model_output={agent.last_raw_output}")
            log(f"round {round_no}: instructions={instructions}")
        updated_xml, failures = execute(package.document_xml, instructions)
        package.document_xml = updated_xml
        if log:
            log(f"round {round_no}: failures={failures}")
        if not failures:
            break
    package.dump(output_path)
    return failures
