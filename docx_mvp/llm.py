from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from urllib.error import HTTPError
from urllib.error import URLError
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
- When `field_candidates` are provided, prefer selecting targets from them first.
- If a required target is not in `field_candidates` but is clearly supported by `document_xml`, you may still return that XPath.
- Do not extract unrelated information just because it appears in `document_xml`.
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


def load_dotenv() -> Path | None:
    project_root = Path(__file__).resolve().parent.parent
    env_path = project_root / ".env"
    if not env_path.exists():
        return None
    for line in env_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        if stripped.startswith("export "):
            stripped = stripped[len("export ") :].strip()
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        os.environ.setdefault(key, value)
    return env_path


class LLM:
    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
        timeout: float | None = None,
    ) -> None:
        load_dotenv()
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        if not self.api_key:
            raise RuntimeError("OPENAI_API_KEY is not set. Add it to .env or export it before running docx-edit-mvp.")
        self.model = model or os.environ.get("OPENAI_MODEL", "gpt-5.2")
        self.base_url = (base_url or os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")).rstrip("/")
        self.timeout = timeout or float(os.environ.get("OPENAI_TIMEOUT_SECONDS", "300"))
        self.last_raw_output = ""

    def generate(
        self,
        document_xml: str,
        prompt: str,
        failures: list[InstructionFailure] | None = None,
        field_candidates: list[dict] | None = None,
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
                            "field_candidates": field_candidates or [],
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
            with urlopen(request, timeout=self.timeout) as response:
                body = json.loads(response.read())
        except HTTPError as exc:
            raise RuntimeError(exc.read().decode("utf-8")) from exc
        except TimeoutError as exc:
            raise RuntimeError(
                f"request to model timed out after {self.timeout:g}s: {self.base_url}/chat/completions"
            ) from exc
        except URLError as exc:
            raise RuntimeError(f"request to model failed: {exc.reason}") from exc
        content = body["choices"][0]["message"]["content"]
        self.last_raw_output = content
        data = json.loads(content)
        return [SetText(**item) for item in data["instructions"]]
