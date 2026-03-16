from __future__ import annotations

from dataclasses import asdict
from lxml import etree

from .llm import InstructionFailure, SetText


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def w_tag(name: str) -> str:
    return f"{W}{name}"


def first_child(element: etree._Element, name: str) -> etree._Element | None:
    for child in element:
        if child.tag == w_tag(name):
            return child
    return None


def append_text_to_run(run: etree._Element, text: str) -> None:
    text_node = first_child(run, "t")
    if text_node is None:
        text_node = etree.Element(w_tag("t"))
        run.append(text_node)
    text_node.text = text
    if text.strip() != text:
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif "{http://www.w3.org/XML/1998/namespace}space" in text_node.attrib:
        del text_node.attrib["{http://www.w3.org/XML/1998/namespace}space"]


def set_paragraph_text(paragraph: etree._Element, text: str) -> None:
    text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
    if text_nodes:
        text_nodes[0].text = text
        if text.strip() != text:
            text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        elif "{http://www.w3.org/XML/1998/namespace}space" in text_nodes[0].attrib:
            del text_nodes[0].attrib["{http://www.w3.org/XML/1998/namespace}space"]
        for node in text_nodes[1:]:
            node.text = ""
        return
    run = etree.Element(w_tag("r"))
    append_text_to_run(run, text)
    p_pr = first_child(paragraph, "pPr")
    if p_pr is None:
        paragraph.insert(0, run)
    else:
        p_pr.addnext(run)


def set_cell_text(cell: etree._Element, text: str) -> None:
    paragraph = first_child(cell, "p")
    if paragraph is None:
        paragraph = etree.Element(w_tag("p"))
        cell.append(paragraph)
    set_paragraph_text(paragraph, text)


def set_node_text(node: etree._Element, text: str) -> None:
    if node.tag == w_tag("t"):
        node.text = text
        if text.strip() != text:
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        elif "{http://www.w3.org/XML/1998/namespace}space" in node.attrib:
            del node.attrib["{http://www.w3.org/XML/1998/namespace}space"]
        return
    if node.tag == w_tag("r"):
        append_text_to_run(node, text)
        return
    if node.tag == w_tag("p"):
        set_paragraph_text(node, text)
        return
    if node.tag == w_tag("tc"):
        set_cell_text(node, text)
        return
    raise ValueError(f"unsupported target node: {node.tag}")


def try_repair_xpath(root: etree._Element, xpath: str, text: str) -> bool:
    if xpath.endswith("/w:r[1]/w:t[1]"):
        paragraphs = root.xpath(xpath[: -len("/w:r[1]/w:t[1]")], namespaces=NS)
        if len(paragraphs) == 1 and paragraphs[0].tag == w_tag("p"):
            set_paragraph_text(paragraphs[0], text)
            return True
    if xpath.endswith("/w:t[1]"):
        runs = root.xpath(xpath[: -len("/w:t[1]")], namespaces=NS)
        if len(runs) == 1 and runs[0].tag == w_tag("r"):
            append_text_to_run(runs[0], text)
            return True
    if xpath.endswith("/w:p[1]"):
        cells = root.xpath(xpath[: -len("/w:p[1]")], namespaces=NS)
        if len(cells) == 1 and cells[0].tag == w_tag("tc"):
            set_cell_text(cells[0], text)
            return True
    return False


def execute(
    document_xml: str,
    instructions: list[SetText],
    locked_xpaths: set[str] | None = None,
) -> tuple[str, list[InstructionFailure]]:
    root = etree.fromstring(document_xml.encode("utf-8"))
    failures: list[InstructionFailure] = []
    locked_xpaths = locked_xpaths or set()
    for instruction in instructions:
        if instruction.type != "set_text":
            failures.append(InstructionFailure(asdict(instruction), f"unsupported type: {instruction.type}"))
            continue
        if instruction.xpath in locked_xpaths:
            failures.append(
                InstructionFailure(
                    asdict(instruction),
                    f'xpath is locked and cannot be overwritten: "{instruction.xpath}"',
                )
            )
            continue
        try:
            matches = root.xpath(instruction.xpath, namespaces=NS)
        except etree.XPathError as exc:
            failures.append(InstructionFailure(asdict(instruction), f"invalid xpath: {exc}"))
            continue
        if not matches:
            if try_repair_xpath(root, instruction.xpath, instruction.text):
                continue
            failures.append(
                InstructionFailure(
                    asdict(instruction),
                    f'xpath matched 0 nodes: "{instruction.xpath}"',
                )
            )
            continue
        if len(matches) != 1:
            failures.append(
                InstructionFailure(
                    asdict(instruction),
                    f'xpath matched {len(matches)} nodes, expected 1: "{instruction.xpath}"',
                )
            )
            continue
        try:
            set_node_text(matches[0], instruction.text)
        except ValueError as exc:
            failures.append(InstructionFailure(asdict(instruction), str(exc)))
    return etree.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8"), failures
