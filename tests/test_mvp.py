from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile
from unittest.mock import patch

from docx_mvp.llm import InstructionFailure, LLM, SetText
from docx_mvp.__main__ import build_log_path, build_paths, main
from docx_mvp.executor import execute
from docx_mvp.package import DocxPackage
from docx_mvp.workflow import edit_docx


XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>"""


class MvpTest(unittest.TestCase):
    def test_cli_paths_use_input_output_dirs(self) -> None:
        input_path, output_path = build_paths("demo.docx", None)
        self.assertTrue(str(input_path).endswith("/input/demo.docx"))
        self.assertTrue(str(output_path).endswith("/output/demo.docx/demo.docx"))
        self.assertTrue(str(build_log_path(output_path)).endswith("/output/demo.docx/demo.docx.log"))

    def test_package_and_execute(self) -> None:
        with TemporaryDirectory() as tmp:
            source = f"{tmp}/in.docx"
            target = f"{tmp}/out.docx"
            with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")
                zf.writestr("word/document.xml", XML)
            package = DocxPackage.load(source)
            updated_xml, failures = execute(
                package.document_xml,
                [SetText(type="set_text", xpath="./w:body/w:p[2]/w:r[1]/w:t[1]", text="World")],
            )
            self.assertEqual([], failures)
            package.document_xml = updated_xml
            package.dump(target)
            self.assertIn("World", DocxPackage.load(target).document_xml)

    def test_execute_creates_text_for_empty_paragraph_targeting_text_xpath(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p/></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""
        updated_xml, failures = execute(
            xml,
            [SetText(type="set_text", xpath="./w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]/w:r[1]/w:t[1]", text="World")],
        )
        self.assertEqual([], failures)
        self.assertIn("<w:t>World</w:t>", updated_xml)

    def test_failure_is_returned(self) -> None:
        _, failures = execute(XML, [SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="World")])
        self.assertEqual(
            [InstructionFailure({"type": "set_text", "xpath": "./w:body/w:p[9]/w:r[1]/w:t[1]", "text": "World"}, 'xpath matched 0 nodes: "./w:body/w:p[9]/w:r[1]/w:t[1]"')],
            failures,
        )

    def test_failure_is_returned_for_non_unique_xpath(self) -> None:
        _, failures = execute(XML, [SetText(type="set_text", xpath=".//w:t", text="World")])
        self.assertEqual(
            [InstructionFailure({"type": "set_text", "xpath": ".//w:t", "text": "World"}, 'xpath matched 2 nodes, expected 1: ".//w:t"')],
            failures,
        )

    def test_failure_is_returned_for_locked_xpath(self) -> None:
        _, failures = execute(
            XML,
            [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="World")],
            locked_xpaths={"./w:body/w:p[1]/w:r[1]/w:t[1]"},
        )
        self.assertEqual(
            [
                InstructionFailure(
                    {"type": "set_text", "xpath": "./w:body/w:p[1]/w:r[1]/w:t[1]", "text": "World"},
                    'xpath is locked and cannot be overwritten: "./w:body/w:p[1]/w:r[1]/w:t[1]"',
                )
            ],
            failures,
        )

    def test_llm_loads_api_key_from_dotenv(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {}, clear=True):
            Path(tmp, ".env").write_text(
                "OPENAI_API_KEY=test-key\nOPENAI_MODEL=test-model\nOPENAI_BASE_URL=https://example.com/v1\n",
                encoding="utf-8",
            )
            with patch("docx_mvp.llm.Path.cwd", return_value=Path(tmp)):
                llm = LLM()
            self.assertEqual("test-key", llm.api_key)
            self.assertEqual("test-model", llm.model)
            self.assertEqual("https://example.com/v1", llm.base_url)

    def test_llm_raises_clear_error_when_api_key_missing(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {}, clear=True):
            with patch("docx_mvp.llm.Path.cwd", return_value=Path(tmp)):
                with self.assertRaisesRegex(RuntimeError, "OPENAI_API_KEY is not set"):
                    LLM()

    def test_retry_loop_stops_after_clean_round(self) -> None:
        class FakeLLM:
            def __init__(self) -> None:
                self.calls = 0

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                locked_xpaths: list[str] | None = None,
            ) -> list[SetText]:
                self.calls += 1
                if self.calls == 1:
                    return [SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="World")]
                return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="World")]

        with TemporaryDirectory() as tmp:
            source = f"{tmp}/in.docx"
            target = f"{tmp}/out.docx"
            with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")
                zf.writestr("word/document.xml", XML)
            llm = FakeLLM()
            failures = edit_docx(source, target, "replace hello", llm=llm)
            self.assertEqual(2, llm.calls)
            self.assertEqual([], failures)
            self.assertIn("World", DocxPackage.load(target).document_xml)

    def test_retry_loop_still_writes_output_when_failures_remain(self) -> None:
        class FakeLLM:
            def __init__(self) -> None:
                self.done = False

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                locked_xpaths: list[str] | None = None,
            ) -> list[SetText]:
                if not self.done:
                    self.done = True
                    return [
                        SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="World"),
                        SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="Nope"),
                    ]
                return [SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="Nope")]

        with TemporaryDirectory() as tmp:
            source = f"{tmp}/in.docx"
            target = f"{tmp}/out.docx"
            with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")
                zf.writestr("word/document.xml", XML)
            failures = edit_docx(source, target, "replace hello", max_rounds=2, llm=FakeLLM())
            self.assertEqual(1, len(failures))
            self.assertIn("World", DocxPackage.load(target).document_xml)

    def test_retry_loop_does_not_overwrite_locked_xpath(self) -> None:
        case = self

        class FakeLLM:
            def __init__(self) -> None:
                self.calls = 0

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                locked_xpaths: list[str] | None = None,
            ) -> list[SetText]:
                self.calls += 1
                if self.calls == 1:
                    return [
                        SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="Gold"),
                        SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="Missing"),
                    ]
                case.assertEqual(["./w:body/w:p[1]/w:r[1]/w:t[1]"], locked_xpaths)
                return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="Overwrite")]

        with TemporaryDirectory() as tmp:
            source = f"{tmp}/in.docx"
            target = f"{tmp}/out.docx"
            with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")
                zf.writestr("word/document.xml", XML)
            failures = edit_docx(source, target, "replace hello", max_rounds=2, llm=FakeLLM())
            self.assertEqual(1, len(failures))
            self.assertEqual(
                'xpath is locked and cannot be overwritten: "./w:body/w:p[1]/w:r[1]/w:t[1]"',
                failures[0].error,
            )
            self.assertIn("Gold", DocxPackage.load(target).document_xml)
            self.assertNotIn("Overwrite", DocxPackage.load(target).document_xml)

    def test_cli_prompts_and_writes_default_output_name(self) -> None:
        with TemporaryDirectory() as tmp:
            input_dir = Path(tmp) / "input"
            output_dir = Path(tmp) / "output"
            with patch("docx_mvp.__main__.INPUT_DIR", input_dir), patch("docx_mvp.__main__.OUTPUT_DIR", output_dir):
                input_dir.mkdir()
                source = input_dir / "demo.docx"
                with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                    zf.writestr("[Content_Types].xml", "<Types/>")
                    zf.writestr("word/document.xml", XML)

                class FakeLLM:
                    def __init__(self) -> None:
                        self.last_raw_output = '{"instructions":[{"type":"set_text","xpath":"./w:body/w:p[1]/w:r[1]/w:t[1]","text":"CLI"}]}'

                    def generate(
                        self,
                        document_xml: str,
                        prompt: str,
                        failures: list[InstructionFailure],
                        locked_xpaths: list[str] | None = None,
                    ) -> list[SetText]:
                        return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="CLI")]

                with patch("docx_mvp.workflow.LLM", return_value=FakeLLM()), patch(
                    "builtins.input", side_effect=["demo.docx", "replace hello"]
                ), patch("sys.argv", ["docx-edit-mvp"]):
                    main()

                self.assertIn("CLI", DocxPackage.load(str(output_dir / "demo.docx" / "demo.docx")).document_xml)
                self.assertIn("written:", (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8"))
                self.assertIn("model_output=", (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8"))

    def test_cli_does_not_raise_when_failures_remain(self) -> None:
        with TemporaryDirectory() as tmp:
            input_dir = Path(tmp) / "input"
            output_dir = Path(tmp) / "output"
            with patch("docx_mvp.__main__.INPUT_DIR", input_dir), patch("docx_mvp.__main__.OUTPUT_DIR", output_dir):
                input_dir.mkdir()
                source = input_dir / "demo.docx"
                with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                    zf.writestr("[Content_Types].xml", "<Types/>")
                    zf.writestr("word/document.xml", XML)

                class FakeLLM:
                    def __init__(self) -> None:
                        self.last_raw_output = '{"instructions":[{"type":"set_text","xpath":"./w:body/w:p[9]/w:r[1]/w:t[1]","text":"CLI"}]}'

                    def generate(
                        self,
                        document_xml: str,
                        prompt: str,
                        failures: list[InstructionFailure],
                        locked_xpaths: list[str] | None = None,
                    ) -> list[SetText]:
                        return [SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="CLI")]

                with patch("docx_mvp.workflow.LLM", return_value=FakeLLM()), patch(
                    "builtins.input", side_effect=["demo.docx", "replace hello"]
                ), patch("sys.argv", ["docx-edit-mvp"]):
                    main()

                self.assertTrue((output_dir / "demo.docx" / "demo.docx").exists())
                self.assertIn("remaining failures:", (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8"))

    def test_workflow_logs_to_callback(self) -> None:
        class FakeLLM:
            def __init__(self) -> None:
                self.last_raw_output = '{"instructions":[{"type":"set_text","xpath":"./w:body/w:p[1]/w:r[1]/w:t[1]","text":"World"}]}'

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                locked_xpaths: list[str] | None = None,
            ) -> list[SetText]:
                return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="World")]

        with TemporaryDirectory() as tmp:
            source = f"{tmp}/in.docx"
            target = f"{tmp}/out.docx"
            with ZipFile(source, "w", compression=ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")
                zf.writestr("word/document.xml", XML)
            messages: list[str] = []
            edit_docx(source, target, "replace hello", llm=FakeLLM(), log=messages.append)
            self.assertTrue(any(message.startswith("round 1: generating instructions") for message in messages))


if __name__ == "__main__":
    unittest.main()
