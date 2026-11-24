"""Offline checks: python verify.py. No API key or paid requests."""
import copy
import io
import json
import sys
import unittest
import urllib.error
from pathlib import Path
from unittest.mock import patch

from docx import Document
from PIL import Image
from pptx import Presentation
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

from core import (DEFAULT_STYLE, SourceDoc, UserError, ai_outline, extract_style,
                  make_outline, normalized_image, read_document, validate_outline)
from renderer import build_scenes, export_pptx, scene_svg

ROOT = Path(__file__).parent


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.doc = read_document("sample.md", (
            "# 传感器汇报\n\n## 背景\n腕关节传感器用于记录与运动相关的信号。"
            "研究中应保留单位与原始数据。\n\n## 方法\n先固定佩戴位置，再记录参考动作。"
            "校准后需要独立验证，不能把预期结果写成实测结论。").encode())

    def test_document_title_uses_first_heading(self):
        self.assertEqual(self.doc.title, "传感器汇报")

    def test_rule_source_coverage_and_notes(self):
        out = make_outline(self.doc, 8)
        validate_outline(out, {b["id"] for b in self.doc.blocks})
        used = {v for s in out["slides"] for v in s["source_ids"]}
        self.assertEqual(used, {b["id"] for b in self.doc.blocks})
        self.assertIn("独立验证", "\n".join(s["notes"] for s in out["slides"]))

    def test_docx_table_and_image(self):
        d = Document()
        d.add_heading("验证方案", level=1)
        d.add_paragraph("这是测试文档正文，包含用于来源核对的中文内容。")
        t = d.add_table(rows=1, cols=2)
        t.cell(0, 0).text, t.cell(0, 1).text = "单位", "毫米"
        pic = io.BytesIO()
        Image.new("RGB", (120, 60), "blue").save(pic, "PNG")
        pic.seek(0)
        d.add_picture(pic)
        data = io.BytesIO()
        d.save(data)
        doc = read_document("test.docx", data.getvalue())
        self.assertTrue(any("毫米" in b["text"] for b in doc.blocks))
        self.assertEqual(len(doc.images), 1)

    def test_pdf_text_and_scanned_rejection(self):
        writer = PdfWriter()
        p = writer.add_blank_page(612, 792)
        font = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"), NameObject("/BaseFont"): NameObject("/Helvetica")})
        p[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)})})
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 14 Tf 30 700 Td (Sensor validation uses independent reference measurements.) Tj ET")
        p[NameObject("/Contents")] = writer._add_object(stream)
        buf = io.BytesIO()
        writer.write(buf)
        result = read_document("test.pdf", buf.getvalue())
        self.assertIn("independent", result.text)
        blank = PdfWriter()
        blank.add_blank_page(612, 792)
        b = io.BytesIO()
        blank.write(b)
        with self.assertRaisesRegex(UserError, "OCR"):
            read_document("scan.pdf", b.getvalue())

    def test_encoding_and_empty_input(self):
        doc = read_document("test.txt", "腕关节运动记录。测量方法需要校准，结果需要独立验证。".encode("gb18030"))
        self.assertIn("独立验证", doc.text)
        with self.assertRaises(UserError):
            read_document("x.txt", b"")
        with self.assertRaises(UserError):
            read_document("x.docx", b"broken zip")

    def test_editable_pptx_and_sources(self):
        out = make_outline(self.doc, 6)
        scenes = build_scenes(out, DEFAULT_STYLE, sources={b["id"]: b["source"] for b in self.doc.blocks})
        data = export_pptx(scenes, out["title"])
        ppt = Presentation(io.BytesIO(data))
        self.assertEqual(len(ppt.slides), len(scenes))
        texts = [s.text for slide in ppt.slides for s in slide.shapes if s.has_text_frame]
        self.assertTrue(any("独立验证" in t for t in texts))
        self.assertIn("原文定位", ppt.slides[1].notes_slide.notes_text_frame.text)
        profile = extract_style(data)
        self.assertEqual(profile["primary"].upper(), DEFAULT_STYLE["primary"])
        self.assertEqual(profile["ratio"], "16:9")

    def test_long_bullets_paginate_without_loss(self):
        out = make_outline(self.doc, 6)
        original = "验证" * 240
        out["slides"] = [{"title":"较长的验证内容", "bullets":[original]*6, "source_ids":["S1"], "notes":"全量", "layout":"columns", "image":""}]
        scenes = build_scenes(out, DEFAULT_STYLE)
        self.assertGreater(len(scenes), 2)
        reconstructed = "".join(e["text"].replace("\n", "") for scene in scenes[1:]
            for e in scene["elements"] if e["kind"] == "text" and e["size"] == 26)
        self.assertEqual(reconstructed, original * 6)
        export_pptx(scenes)

    def test_all_layouts_ratios_images_and_bounds(self):
        pic = io.BytesIO()
        Image.new("RGB", (800, 1600), "green").save(pic, "PNG")
        assets = {"image": normalized_image(pic.getvalue(), "chart.png")}
        for ratio in ("16:9", "4:3"):
            for layout in ("bullets", "columns", "image"):
                out = make_outline(self.doc, 6)
                for s in out["slides"]:
                    s.update(layout=layout, image="image")
                scenes = build_scenes(out, DEFAULT_STYLE | {"ratio":ratio}, assets)
                for scene in scenes:
                    for e in scene["elements"]:
                        self.assertGreaterEqual(e["x"], 0)
                        self.assertGreaterEqual(e["y"], 0)
                        self.assertLessEqual(e["x"] + e["w"], scene["width"]+2)
                        self.assertLessEqual(e["y"] + e["h"], scene["height"]+2)
                ppt = Presentation(io.BytesIO(export_pptx(scenes)))
                self.assertEqual(len(ppt.slides), len(scenes))

    def test_extreme_cover_and_titles(self):
        out = make_outline(self.doc, 6, title="中"*120, subtitle="文"*150)
        out["slides"][0]["title"] = "文"*100
        scenes = build_scenes(out, DEFAULT_STYLE)
        self.assertGreater(len(scenes), 1)

    def test_svg_escapes_user_markup(self):
        out = make_outline(self.doc, 6, title='<script>alert("x")</script>')
        svg = scene_svg(build_scenes(out, DEFAULT_STYLE)[0]).decode()
        self.assertNotIn("<script>", svg)
        self.assertIn("&lt;script&gt;", svg)

    def test_validation_rejects_bad_sources_and_long_bullets(self):
        out = make_outline(self.doc, 6)
        out["slides"][0]["source_ids"] = ["S9999"]
        with self.assertRaises(UserError):
            validate_outline(out, {"S1", "S2"})
        out["slides"][0]["source_ids"] = ["S1"]
        out["slides"][0]["bullets"] = ["x"*501]
        with self.assertRaises(UserError):
            validate_outline(out, {"S1", "S2"})

    def test_ai_mock_only_no_network(self):
        out = make_outline(self.doc, 6)
        payload = json.dumps({"choices":[{"message":{"content":json.dumps(out)}}]}).encode()
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, *args): return payload
        with patch("urllib.request.urlopen", return_value=Response()) as request:
            result = ai_outline(self.doc, 6, "测试", "", "fake-key", "configured-model")
            self.assertEqual(result["mode"], "AI 整理")
            self.assertEqual(request.call_count, 1)
        with patch("urllib.request.urlopen", side_effect=urllib.error.HTTPError("https://example.com", 401, "bad", {}, None)):
            with self.assertRaisesRegex(UserError, "API Key"):
                ai_outline(self.doc, 6, "", "", "fake-key", "model")

    def test_streamlit_demo_edit_export_invalidation(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=40).run()
        self.assertFalse(at.exception, str(at.exception))
        next(b for b in at.button if b.label == "载入腕关节传感器示例").click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertFalse(at.error, [e.value for e in at.error])
        next(b for b in at.button if b.label == "生成 PPTX 下载文件").click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertTrue(at.session_state["export"].startswith(b"PK"))
        old_sig = at.session_state["export_sig"]
        next(t for t in at.text_input if t.label == "封面标题").set_value("修改后的课堂汇报").run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertFalse(at.error, [e.value for e in at.error])
        next(b for b in at.button if b.label == "生成 PPTX 下载文件").click().run()
        self.assertNotEqual(at.session_state["export_sig"], old_sig)
        ppt = Presentation(io.BytesIO(at.session_state["export"]))
        self.assertTrue(any("修改后的课堂汇报" in s.text for s in ppt.slides[0].shapes if s.has_text_frame))


if __name__ == "__main__":
    unittest.main(verbosity=2)
