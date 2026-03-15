from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile


DOCUMENT_XML = "word/document.xml"


@dataclass
class DocxPackage:
    files: dict[str, bytes]

    @classmethod
    def load(cls, path: str) -> "DocxPackage":
        with ZipFile(path) as zf:
            return cls({name: zf.read(name) for name in zf.namelist()})

    @property
    def document_xml(self) -> str:
        return self.files[DOCUMENT_XML].decode("utf-8")

    @document_xml.setter
    def document_xml(self, value: str) -> None:
        self.files[DOCUMENT_XML] = value.encode("utf-8")

    def dump(self, path: str) -> None:
        buffer = BytesIO()
        with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as zf:
            for name, data in self.files.items():
                zf.writestr(name, data)
        with open(path, "wb") as fh:
            fh.write(buffer.getvalue())
