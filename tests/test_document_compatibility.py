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
        with patch.object(server, "oxide_extract_text", None), patch.object(
            server, "oxide_to_markdown", None
        ), patch.object(server, "find_office_converter", return_value="fake-office"), patch.object(
            server, "_run_office_conversion", return_value=converted
        ) as convert:
            record = server.records_from_file(item)[0]
        self.assertIn("转换后可审核正文", record.text)
        self.assertTrue(any("LibreOffice" in warning for warning in record.extraction_warnings))
        self.assertEqual(original, b"legacy-doc-content")
        convert.assert_called_once_with(original, ".doc")

    def test_office_oxide_is_preferred_for_old_doc(self) -> None:
        original = b"legacy-doc-content"
        oxide_text = "研究内容\n" + "Office Oxide直接读取的完整正文" * 20
        with patch.object(server, "oxide_to_markdown", return_value=oxide_text), patch.object(
            server, "oxide_extract_text", return_value=""
        ), patch.object(server, "_run_office_conversion") as convert:
            text, warnings = server.extract_doc_with_status(original)
        self.assertIn("Office Oxide直接读取", text)
        self.assertTrue(any("Office Oxide直接读取" in warning for warning in warnings))
        convert.assert_not_called()

    def test_xml_fallback_reads_docx_without_table_grid(self) -> None:
        text = server._extract_docx_xml(make_docx("底层XML可以读取的正文"))
        self.assertIn("底层XML可以读取的正文", text)

    def test_compact_document_keeps_key_sections_and_both_ends(self) -> None:
        text = "文档开头\n" + ("普通说明内容\n" * 5000) + "预算总额8万元，材料费重复列支\n文档结尾"
        compacted = server.compact_document(text, "计划任务书", max_chars=4000)
        self.assertLessEqual(len(compacted), 4000)
        self.assertIn("文档开头", compacted)
        self.assertIn("预算总额8万元，材料费重复列支", compacted)
        self.assertIn("文档结尾", compacted)

    def test_blank_approval_field_alone_is_not_a_deterministic_rejection(self) -> None:
        text = "研究内容完整\n技术关键点及创新点完整\n应用前景明确\n申请部门/单位意见：（公章） 年 月 日"
        checks = server.deterministic_checks(text, "立项申请书")
        self.assertFalse(any(check["rule_id"] == "R-C09" for check in checks))

    def test_both_blank_approval_opinions_are_rejected(self) -> None:
        text = (
            "研究内容完整\n技术关键点及创新点完整\n应用前景明确\n"
            "申请部门/单位意见：（公章） 年 月 日\n"
            "直属单位科技管理部门意见：（公章） 年 月 日\n"
            "注：审批通过后打印。"
        )
        checks = server.deterministic_checks(text, "立项申请书")
        self.assertTrue(any(check["rule_id"] == "R-C09" for check in checks))

    def test_explicit_approval_opinions_are_not_rejected(self) -> None:
        text = (
            "研究内容完整\n技术关键点及创新点完整\n应用前景明确\n"
            "申请部门/单位意见：经审核，同意立项申报。2026年3月20日\n"
            "直属单位科技管理部门意见：经审核，同意该项目立项申报。2026年3月25日\n"
            "注：审批通过后打印。"
        )
        checks = server.deterministic_checks(text, "立项申请书")
        self.assertFalse(any(check["rule_id"] == "R-C09" for check in checks))

    def test_application_unit_approval_heading_is_supported(self) -> None:
        text = (
            "研究内容完整\n技术关键点及创新点完整\n应用前景明确\n"
            "申请部门/单位意见：（公章） 年 月 日\n"
            "申请单位科技管理部门意见：（公章） 年 月 日\n"
            "注：审批通过后打印。"
        )
        checks = server.deterministic_checks(text, "立项申请书")
        self.assertTrue(any(check["rule_id"] == "R-C09" for check in checks))

    def test_plan_pass_receives_independent_consistency_review(self) -> None:
        record = server.Record("sample", "sample.doc", "项目材料正文完整")
        first_pass = {
            "label": "通过",
            "matched_rules": [],
            "reason": "字段齐全",
            "confidence": 0.95,
        }
        conflict = {
            "label": "不通过",
            "matched_rules": [{
                "rule_id": "R-C03",
                "rule_name": "跨章节信息一致性",
                "evidence": "负责人信息为工程师，项目组表为高级技师",
            }],
            "reason": "同一人员职称前后不一致",
            "confidence": 0.96,
        }
        with patch.dict(server.os.environ, {"LLM_API_KEY": "test-key"}), patch.object(
            server, "deterministic_checks", return_value=[]
        ), patch.object(server, "extract_review_spec", return_value={"conflicts": []}), patch.object(
            server, "call_llm", side_effect=[first_pass, conflict, conflict]
        ) as mocked:
            result = server.judge_one(record, "计划任务书", "判断项目是否通过")

        self.assertEqual(result["label"], "不通过")
        self.assertEqual(result["matched_rules"][0]["rule_id"], "R-C03")
        self.assertIn("年龄、职称、单位", mocked.call_args_list[1].kwargs["review_note"])


if __name__ == "__main__":
    unittest.main()
