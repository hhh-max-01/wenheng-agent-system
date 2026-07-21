from __future__ import annotations

import base64
import io
import sys
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from docx import Document


PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import server


def make_docx(text: str) -> bytes:
    output = io.BytesIO()
    document = Document()
    document.add_paragraph(text)
    document.save(output)
    return output.getvalue()


def damage_document_xml(data: bytes) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data), "r") as source, zipfile.ZipFile(output, "w") as target:
        for info in source.infolist():
            payload = source.read(info)
            if info.filename == "word/document.xml":
                payload = payload.replace(b"</w:document>", b"<broken>")
            target.writestr(info, payload)
    return output.getvalue()


class DocumentCompatibilityTests(unittest.TestCase):
    def test_corrupt_docx_recovers_text_and_marks_incomplete(self) -> None:
        damaged = damage_document_xml(make_docx("可恢复的项目正文" * 30))
        text, warnings = server.extract_docx_with_status(damaged)
        self.assertIn("可恢复的项目正文", text)
        self.assertTrue(any("仅恢复到部分" in warning for warning in warnings))
        checks = server.deterministic_checks(text, "立项申请书", warnings)
        self.assertTrue(any(check["rule_id"] == "R-C02" for check in checks))

    def test_old_doc_is_converted_without_changing_uploaded_bytes(self) -> None:
        original = b"legacy-doc-content"
        converted = make_docx("转换后可审核正文")
        item = {"name": "sample.doc", "base64": base64.b64encode(original).decode("ascii")}
        with patch.object(server, "find_office_converter", return_value="fake-office"), patch.object(
            server, "_run_office_conversion", return_value=converted
        ) as convert:
            record = server.records_from_file(item)[0]
        self.assertIn("转换后可审核正文", record.text)
        self.assertEqual(original, b"legacy-doc-content")
        convert.assert_called_once_with(original, ".doc")

    def test_blank_approval_field_alone_is_not_a_deterministic_rejection(self) -> None:
        text = "研究内容完整\n技术关键点及创新点完整\n应用前景明确\n申请部门/单位意见：（公章） 年 月 日"
        checks = server.deterministic_checks(text, "立项申请书")
        self.assertFalse(any(check["rule_id"] == "R-C09" for check in checks))


if __name__ == "__main__":
    unittest.main()
