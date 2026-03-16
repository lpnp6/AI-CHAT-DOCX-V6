from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from urllib.error import HTTPError
from urllib.request import Request, urlopen


SYSTEM_PROMPT = """You edit Microsoft Word document.xml content.
Return only JSON in this shape:
{"instructions":[{"type":"set_text","xpath":"./w:body/w:p[1]/w:r[1]/w:t[1]","text":"replacement text"}]}
Rules:
- Only use type "set_text".
- xpath must be an ElementTree-compatible XPath relative to the <w:document> root.
- Use the `w:` namespace prefix exactly as it appears in document.xml.
- Prefer pointing xpath to an existing <w:t> node.
- If a target cell or paragraph is empty and has no <w:t>, you may point to its existing <w:p> or <w:tc> node instead.
- If `locked_xpaths` is provided, never reuse or overwrite those XPath targets.
- When `failed_instructions` is provided, only repair those failed targets instead of moving content onto unrelated fields.
- Never wrap JSON in markdown.
"""


@dataclass
class SetText:
    type: str
    xpath: str
    text: str


@dataclass
class InstructionFailure:
    instruction: dict
    error: str


class LLM:
    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
    ) -> None:
        self.api_key = api_key or os.environ["OPENAI_API_KEY"]
        self.model = model or os.environ.get("OPENAI_MODEL", "gpt-5.2")
        self.base_url = (base_url or os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")).rstrip("/")
        self.last_raw_output = ""

    def generate(
        self,
        document_xml: str,
        prompt: str,
        failures: list[InstructionFailure] | None = None,
        locked_xpaths: list[str] | None = None,
    ) -> list[SetText]:
        payload = {
            "model": self.model,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "prompt": prompt,
                            "document_xml": document_xml,
                            "failed_instructions": [asdict(item) for item in failures or []],
                            "locked_xpaths": locked_xpaths or [],
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        }
        request = Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urlopen(request) as response:
                body = json.loads(response.read())
        except HTTPError as exc:
            raise RuntimeError(exc.read().decode("utf-8")) from exc
        content = body["choices"][0]["message"]["content"]
        self.last_raw_output = content
        data = json.loads(content)
        return [SetText(**item) for item in data["instructions"]]
