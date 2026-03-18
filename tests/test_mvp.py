from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile
from unittest.mock import patch

from docx_mvp.llm import InstructionFailure, LLM, SetText
from docx_mvp.__main__ import build_log_path, build_paths, format_console_message, main, make_logger
from docx_mvp.executor import execute
from docx_mvp.package import DocxPackage
from docx_mvp.workflow import edit_docx, extract_fields


XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>"""

FIELD_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>张三</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>电话：</w:t></w:r></w:p>
    <w:p><w:r><w:t>123456</w:t></w:r></w:p>
  </w:body>
</w:document>"""


class MvpTest(unittest.TestCase):
    def assert_timestamped_lines(self, text: str) -> None:
        for line in text.strip().splitlines():
            self.assertRegex(line, r"^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] ")

    def test_format_console_message_flattens_and_truncates(self) -> None:
        message = "line 1\n" + ("x" * 300)
        formatted = format_console_message(message, max_chars=40)
        self.assertNotIn("\n", formatted)
        self.assertTrue(formatted.endswith("..."))
        self.assertLessEqual(len(formatted), 40)

    def test_make_logger_keeps_full_message_in_log_file(self) -> None:
        with TemporaryDirectory() as tmp:
            log_path = Path(tmp) / "demo.log"
            logger = make_logger(log_path)
            message = "line 1\n" + ("x" * 300)
            with patch("builtins.print") as mock_print:
                logger(message)
            mock_print.assert_called_once()
            self.assertNotIn("\n", mock_print.call_args.args[0])
            log_text = log_path.read_text(encoding="utf-8")
            self.assertRegex(log_text.splitlines()[0], r"^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] ")
            self.assertIn(message, log_text)

    def test_cli_paths_use_input_output_dirs(self) -> None:
        input_path, output_path = build_paths("demo.docx", None)
        self.assertTrue(str(input_path).endswith("/input/demo.docx"))
        self.assertTrue(str(output_path).endswith("/output/demo.docx/demo.docx"))
        self.assertTrue(str(build_log_path(output_path)).endswith("/output/demo.docx/demo.docx.log"))

    def test_cli_paths_accept_absolute_input_path(self) -> None:
        input_path, output_path = build_paths("/tmp/demo.docx", None)
        self.assertEqual(Path("/tmp/demo.docx"), input_path)
        self.assertTrue(str(output_path).endswith("/output/demo.docx/demo.docx"))

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
            project_dir = Path(tmp) / "project"
            module_dir = project_dir / "docx_mvp"
            module_dir.mkdir(parents=True)
            (project_dir / ".env").write_text(
                "OPENAI_API_KEY=test-key\nOPENAI_MODEL=test-model\nOPENAI_BASE_URL=https://example.com/v1\n",
                encoding="utf-8",
            )
            fake_module_file = module_dir / "llm.py"
            fake_module_file.write_text("", encoding="utf-8")
            with patch("docx_mvp.llm.Path.cwd", return_value=Path("/tmp")), patch("docx_mvp.llm.__file__", str(fake_module_file)):
                llm = LLM()
            self.assertEqual("test-key", llm.api_key)
            self.assertEqual("test-model", llm.model)
            self.assertEqual("https://example.com/v1", llm.base_url)

    def test_llm_loads_dotenv_with_export_prefix(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {}, clear=True):
            project_dir = Path(tmp) / "project"
            module_dir = project_dir / "docx_mvp"
            module_dir.mkdir(parents=True)
            (project_dir / ".env").write_text(
                'export OPENAI_API_KEY="test-key"\nexport OPENAI_MODEL=test-model\n',
                encoding="utf-8",
            )
            fake_module_file = module_dir / "llm.py"
            fake_module_file.write_text("", encoding="utf-8")
            with patch("docx_mvp.llm.Path.cwd", return_value=Path("/tmp")), patch("docx_mvp.llm.__file__", str(fake_module_file)):
                llm = LLM()
            self.assertEqual("test-key", llm.api_key)
            self.assertEqual("test-model", llm.model)

    def test_llm_loads_dotenv_from_module_parent_when_cwd_does_not_contain_project_env(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {}, clear=True):
            project_dir = Path(tmp) / "project"
            module_dir = project_dir / "docx_mvp"
            module_dir.mkdir(parents=True)
            (project_dir / ".env").write_text("OPENAI_API_KEY=test-key\n", encoding="utf-8")
            fake_module_file = module_dir / "llm.py"
            fake_module_file.write_text("", encoding="utf-8")
            with patch("docx_mvp.llm.Path.cwd", return_value=Path("/tmp")), patch("docx_mvp.llm.__file__", str(fake_module_file)):
                llm = LLM()
            self.assertEqual("test-key", llm.api_key)

    def test_llm_prefers_shell_environment_over_project_dotenv(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {"OPENAI_API_KEY": "shell-key"}, clear=True):
            project_dir = Path(tmp) / "project"
            module_dir = project_dir / "docx_mvp"
            module_dir.mkdir(parents=True)
            (project_dir / ".env").write_text("OPENAI_API_KEY=dotenv-key\nOPENAI_MODEL=dotenv-model\n", encoding="utf-8")
            fake_module_file = module_dir / "llm.py"
            fake_module_file.write_text("", encoding="utf-8")
            with patch("docx_mvp.llm.Path.cwd", return_value=Path("/tmp")), patch("docx_mvp.llm.__file__", str(fake_module_file)):
                llm = LLM()
            self.assertEqual("shell-key", llm.api_key)
            self.assertEqual("dotenv-model", llm.model)

    def test_llm_raises_clear_error_when_api_key_missing(self) -> None:
        with TemporaryDirectory() as tmp, patch.dict("os.environ", {}, clear=True):
            project_dir = Path(tmp) / "project"
            module_dir = project_dir / "docx_mvp"
            module_dir.mkdir(parents=True)
            fake_module_file = module_dir / "llm.py"
            fake_module_file.write_text("", encoding="utf-8")
            with patch("docx_mvp.llm.Path.cwd", return_value=Path(tmp)), patch("docx_mvp.llm.__file__", str(fake_module_file)):
                with self.assertRaisesRegex(RuntimeError, "OPENAI_API_KEY is not set"):
                    LLM()

    def test_extract_fields_returns_table_and_following_paragraph_candidates(self) -> None:
        fields = extract_fields(FIELD_XML)
        self.assertEqual(
            [
                {
                    "label": "姓名",
                    "xpath": "./w:body/w:tbl/w:tr/w:tc[2]/w:p/w:r/w:t",
                    "context": "姓名 | 张三",
                    "current_text": "张三",
                    "confidence": 0.8,
                },
                {
                    "label": "电话",
                    "xpath": "./w:body/w:p[2]/w:r/w:t",
                    "context": "电话： | 123456",
                    "current_text": "123456",
                    "confidence": 0.95,
                },
            ],
            [field.__dict__ for field in fields],
        )

    def test_retry_loop_stops_after_clean_round(self) -> None:
        class FakeLLM:
            def __init__(self) -> None:
                self.calls = 0

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                field_candidates: list[dict] | None = None,
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
                field_candidates: list[dict] | None = None,
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
                field_candidates: list[dict] | None = None,
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

    def test_cli_writes_default_output_name(self) -> None:
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
                        field_candidates: list[dict] | None = None,
                        locked_xpaths: list[str] | None = None,
                    ) -> list[SetText]:
                        return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="CLI")]

                with patch("docx_mvp.workflow.LLM", return_value=FakeLLM()), patch(
                    "sys.argv", ["docx-edit-mvp", "demo.docx", "replace hello"]
                ):
                    main()

                log_text = (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8")
                self.assertIn("CLI", DocxPackage.load(str(output_dir / "demo.docx" / "demo.docx")).document_xml)
                self.assert_timestamped_lines(log_text)
                self.assertIn("written:", log_text)
                self.assertIn("model_output=", log_text)
                self.assertIn(str((output_dir / "demo.docx" / "demo.docx").resolve()), log_text)
                self.assertTrue(
                    log_text.strip().splitlines()[-1].endswith(str((output_dir / "demo.docx" / "demo.docx").resolve()))
                )
                self.assertRegex(log_text.strip().splitlines()[-1], r"\] output: ")

    def test_cli_accepts_absolute_input_path_and_logs_absolute_paths(self) -> None:
        with TemporaryDirectory() as tmp:
            output_dir = Path(tmp) / "output"
            source = Path(tmp) / "demo.docx"
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
                    field_candidates: list[dict] | None = None,
                    locked_xpaths: list[str] | None = None,
                ) -> list[SetText]:
                    return [SetText(type="set_text", xpath="./w:body/w:p[1]/w:r[1]/w:t[1]", text="CLI")]

            with patch("docx_mvp.__main__.OUTPUT_DIR", output_dir), patch("docx_mvp.workflow.LLM", return_value=FakeLLM()), patch(
                "sys.argv", ["docx-edit-mvp", str(source), "replace hello"]
            ):
                main()

            log_text = (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8")
            self.assert_timestamped_lines(log_text)
            self.assertIn(str(source.resolve()), log_text)
            self.assertIn(str((output_dir / "demo.docx" / "demo.docx").resolve()), log_text)

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
                        field_candidates: list[dict] | None = None,
                        locked_xpaths: list[str] | None = None,
                    ) -> list[SetText]:
                        return [SetText(type="set_text", xpath="./w:body/w:p[9]/w:r[1]/w:t[1]", text="CLI")]

                with patch("docx_mvp.workflow.LLM", return_value=FakeLLM()), patch(
                    "sys.argv", ["docx-edit-mvp", "demo.docx", "replace hello"]
                ):
                    main()

                self.assertTrue((output_dir / "demo.docx" / "demo.docx").exists())
                log_text = (output_dir / "demo.docx" / "demo.docx.log").read_text(encoding="utf-8")
                self.assert_timestamped_lines(log_text)
                self.assertIn("remaining failures:", log_text)

    def test_workflow_logs_to_callback(self) -> None:
        class FakeLLM:
            def __init__(self) -> None:
                self.last_raw_output = '{"instructions":[{"type":"set_text","xpath":"./w:body/w:p[1]/w:r[1]/w:t[1]","text":"World"}]}'

            def generate(
                self,
                document_xml: str,
                prompt: str,
                failures: list[InstructionFailure],
                field_candidates: list[dict] | None = None,
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
