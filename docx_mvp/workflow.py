from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, asdict

from lxml import etree

from .llm import InstructionFailure, LLM
from .executor import execute
from .package import DocxPackage

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass
class FieldCandidate:
    label: str
    xpath: str
    context: str
    current_text: str
    confidence: float


def normalize_text(value: str) -> str:
    return " ".join(value.split())


def element_text(element: etree._Element) -> str:
    return normalize_text("".join(element.xpath(".//w:t/text()", namespaces=NS)))


def element_xpath(root: etree._Element, element: etree._Element) -> str:
    path = root.getroottree().getpath(element)
    return "." + path.removeprefix("/w:document")


def label_score(label: str) -> float:
    if not label or len(label) > 20:
        return 0.0
    if any(char in label for char in "。；，,.!?！？()（）[]【】"):
        return 0.0
    return 0.95 if any(char in label for char in ":：") else 0.8


def extract_fields(document_xml: str) -> list[FieldCandidate]:
    root = etree.fromstring(document_xml.encode("utf-8"))
    candidates: list[FieldCandidate] = []
    seen: set[tuple[str, str]] = set()

    for row in root.xpath(".//w:tr", namespaces=NS):
        cells = row.xpath("./w:tc", namespaces=NS)
        if len(cells) < 2:
            continue
        label = element_text(cells[0]).strip().strip(":：")
        confidence = label_score(label)
        if confidence == 0.0:
            continue
        target_cell = cells[1]
        text_nodes = target_cell.xpath(".//w:t", namespaces=NS)
        target_node = text_nodes[0] if text_nodes else None
        if target_node is None:
            paragraphs = target_cell.xpath("./w:p[1]", namespaces=NS)
            if not paragraphs:
                continue
            target_node = paragraphs[0]
        if target_node is None:
            continue
        xpath = element_xpath(root, target_node)
        key = (label, xpath)
        if key in seen:
            continue
        seen.add(key)
        candidates.append(
            FieldCandidate(
                label=label,
                xpath=xpath,
                context=normalize_text(f"{element_text(cells[0])} | {element_text(target_cell)}"),
                current_text=element_text(target_cell),
                confidence=confidence,
            )
        )

    for paragraph in root.xpath(".//w:p", namespaces=NS):
        text = element_text(paragraph)
        if not text.endswith((":", "：")):
            continue
        label = text.rstrip(":：").strip()
        confidence = label_score(text)
        if confidence == 0.0:
            continue
        next_paragraph = paragraph.getnext()
        while next_paragraph is not None and next_paragraph.tag != f"{{{NS['w']}}}p":
            next_paragraph = next_paragraph.getnext()
        if next_paragraph is None:
            continue
        text_nodes = next_paragraph.xpath(".//w:t", namespaces=NS)
        target = text_nodes[0] if text_nodes else next_paragraph
        xpath = element_xpath(root, target)
        key = (label, xpath)
        if key in seen:
            continue
        seen.add(key)
        candidates.append(
            FieldCandidate(
                label=label,
                xpath=xpath,
                context=normalize_text(f"{text} | {element_text(next_paragraph)}"),
                current_text=element_text(next_paragraph),
                confidence=confidence,
            )
        )

    return candidates


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
    field_candidates = extract_fields(package.document_xml)
    serialized_fields = [asdict(item) for item in field_candidates]
    failures: list[InstructionFailure] = []
    locked_xpaths: set[str] = set()
    llm = llm or LLM()
    if log:
        log(f"field_candidates={serialized_fields}")
    for round_no in range(1, max_rounds + 1):
        if log:
            log(f"round {round_no}: generating instructions")
            log(f"round {round_no}: requesting model")
        instructions = llm.generate(
            package.document_xml,
            prompt,
            field_candidates=serialized_fields,
            failures=failures,
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
