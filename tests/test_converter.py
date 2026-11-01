from __future__ import annotations

import io
import json
import zipfile
from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt
from PIL import Image

from app import app
from converter import (
    Asset,
    ConversionError,
    SelectionRequired,
    _validate_public_host,
    build_docx,
    collect_project,
    prepare_images,
)


def png_bytes(color=(20, 50, 90)):
    buffer = io.BytesIO()
    Image.new("RGB", (40, 30), color).save(buffer, format="PNG")
    return buffer.getvalue()


def test_page_profile_headings_and_page_break():
    markdown = "# First\n\nText.\n\n###### Last\n\n<!-- pagebreak -->\n\nNext"
    content = build_docx(markdown, {}, header_mode="none", header_text="", footer_mode="none", footer_text="")
    document = Document(io.BytesIO(content))

    section = document.sections[0]
    assert abs(section.page_width - Mm(210)) < 500
    assert abs(section.left_margin - Mm(30)) < 500
    assert abs(section.right_margin - Mm(15)) < 500
    assert abs(section.top_margin - Mm(20)) < 500
    assert document.styles["Normal"].font.name == "Times New Roman"
    assert document.styles["Normal"].font.size == Pt(12)
    assert document.styles["Heading 1"].font.size == Pt(20)
    assert document.styles["Heading 6"].font.size == Pt(12)
    assert abs(document.styles["Heading 6"].paragraph_format.left_indent - Mm(25)) < 500
    assert any("w:type=\"page\"" in p._p.xml for p in document.paragraphs)


def test_long_table_is_split_with_repeated_headers_and_continuation_caption():
    data_rows = "\n".join(f"| Элемент {index} | Значение {index} |" for index in range(1, 45))
    content = build_docx(
        "Таблица 7 — Результаты\n\n| Имя | Значение |\n|---|---|\n" + data_rows,
        {},
        header_mode="none",
        header_text="",
        footer_mode="none",
        footer_text="",
    )
    document = Document(io.BytesIO(content))
    assert len(document.tables) > 1
    for table in document.tables:
        assert "w:tblHeader" in table.rows[0]._tr.xml
    assert "Таблица 7 — Результаты" in "\n".join(p.text for p in document.paragraphs)
    assert "Продолжение таблицы 7" in "\n".join(p.text for p in document.paragraphs)
    assert not document.paragraphs[0].paragraph_format.page_break_before


def test_large_table_keeps_its_section_heading_with_the_table():
    rows = "\n".join(f"| {index} | {index} |" for index in range(1, 45))
    content = build_docx(
        "Перед таблицей.\n\n# Результаты\n\nТаблица 8 — Данные\n\n| A | B |\n|---|---|\n" + rows,
        {},
        header_mode="none",
        header_text="",
        footer_mode="none",
        footer_text="",
    )
    document = Document(io.BytesIO(content))
    heading = next(p for p in document.paragraphs if p.text == "Результаты")
    assert heading.paragraph_format.page_break_before


def test_table_cells_follow_header_and_body_alignment_rules():
    content = build_docx(
        "Таблица 2 — Короткая таблица\n\n| Название | Значение |\n|---|---|\n| строка | 42 |",
        {},
        header_mode="none",
        header_text="",
        footer_mode="none",
        footer_text="",
    )
    document = Document(io.BytesIO(content))
    table = document.tables[0]
    header = table.rows[0].cells[0].paragraphs[0]
    body = table.rows[1].cells[0].paragraphs[0]
    assert header.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert header.runs[0].bold
    assert header.runs[0].font.name == "Times New Roman"
    assert body.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert sum(p.text == "Таблица 2 — Короткая таблица" for p in document.paragraphs) == 1


def test_header_footer_modes_insert_dynamic_page_field():
    content = build_docx(
        "# Report",
        {},
        header_mode="text",
        header_text="Моя работа",
        footer_mode="page",
        footer_text="",
    )
    document = Document(io.BytesIO(content))
    assert "Моя работа" in document.sections[0].header._element.xml
    assert "PAGE" in document.sections[0].footer._element.xml


def test_relative_and_absolute_image_paths_resolve_from_uploaded_assets():
    image_data = png_bytes()
    assets = [
        Asset("project:docs/images/pic.png", "docs/images/pic.png", image_data),
        Asset("resources:pic.png", "pic.png", image_data),
    ]
    relative = prepare_images("![Alt](images/pic.png)", "docs/report.md", assets, {})
    absolute_src = r"C:%5CUsers%5Cstudent%5CPictures%5Cpic.png"
    absolute = prepare_images(
        r"![Alt](C:\Users\student\Pictures\pic.png)",
        "report.md",
        assets,
        {absolute_src: "resources:pic.png"},
    )
    assert relative["images/pic.png"].startswith(b"\x89PNG")
    assert absolute[absolute_src].startswith(b"\x89PNG")


def test_relative_asset_paths_can_move_to_a_sibling_folder_without_leaving_archive():
    image = png_bytes()
    result = prepare_images(
        "![Plot](../images/plot.png)",
        "docs/report.md",
        [Asset("project:images/plot.png", "images/plot.png", image)],
        {},
    )
    assert result
    with pytest.raises(ConversionError, match="Не найдены изображения"):
        prepare_images("![Plot](../../outside.png)", "docs/report.md", [], {})


def test_https_image_flow_sanitizes_the_downloaded_file(monkeypatch):
    import converter

    image = png_bytes()
    monkeypatch.setattr(converter, "_fetch_https_image", lambda _url: image)
    result = prepare_images("![Remote](https://images.example.test/plot.png)", "report.md", [], {})
    assert result["https://images.example.test/plot.png"].startswith(b"\x89PNG")


def test_ambiguous_absolute_image_requires_user_selection():
    image_data = png_bytes()
    assets = [
        Asset("project:a/plot.png", "a/plot.png", image_data),
        Asset("project:b/plot.png", "b/plot.png", image_data),
    ]
    markdown = r"![Plot](C:\old\path\plot.png)"
    normalized_src = r"C:%5Cold%5Cpath%5Cplot.png"
    with pytest.raises(SelectionRequired) as error:
        prepare_images(markdown, "report.md", assets, {})
    assert error.value.payload["ambiguous_images"][0]["src"] == normalized_src

    selected = prepare_images(
        markdown,
        "report.md",
        assets,
        {normalized_src: "project:b/plot.png"},
    )
    assert selected


def test_missing_image_is_an_actionable_error():
    with pytest.raises(ConversionError, match="Не найдены изображения"):
        prepare_images("![Missing](images/missing.png)", "report.md", [], {})


def test_figure_caption_is_centered_and_unmatched_text_is_preserved():
    image = png_bytes()
    resolved = prepare_images(
        "![Diagram](diagram.png)\n\nРисунок 2 — Схема сети\n\nОписание рисунка 3 — это обычный текст.",
        "report.md",
        [Asset("resources:diagram.png", "diagram.png", image)],
        {},
    )
    document = Document(io.BytesIO(build_docx(
        "![Diagram](diagram.png)\n\nРисунок 2 — Схема сети\n\nОписание рисунка 3 — это обычный текст.",
        resolved,
        header_mode="none",
        header_text="",
        footer_mode="none",
        footer_text="",
    )))
    assert len(document.inline_shapes) == 1
    assert document.inline_shapes[0]._inline.docPr.get("descr") == "Diagram"
    caption = next(p for p in document.paragraphs if p.text == "Рисунок 2 — Схема сети")
    assert caption.alignment == 1
    assert "Описание рисунка 3 — это обычный текст." in [p.text for p in document.paragraphs]


def test_project_zip_requires_main_markdown_selection_if_multiple():
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("a.md", "# A")
        zf.writestr("b.md", "# B")
    archive.seek(0)

    with pytest.raises(SelectionRequired) as error:
        collect_project(
            SimpleNamespace(filename="project.zip", read=archive.read),
            None,
            "",
        )
    assert error.value.payload["markdown_files"] == ["a.md", "b.md"]


def test_private_and_local_hosts_are_rejected_for_remote_images(monkeypatch):
    monkeypatch.setattr(
        "converter.socket.getaddrinfo",
        lambda *_args, **_kwargs: [(2, 1, 6, "", ("127.0.0.1", 443))],
    )
    with pytest.raises(ConversionError, match="закрытую или локальную сеть"):
        _validate_public_host("localhost")


def test_http_endpoint_returns_editable_docx_with_headers_and_footer():
    client = app.test_client()
    response = client.post(
        "/api/convert",
        data={
            "source": (io.BytesIO("# Пример\n\nТекст".encode()), "report.md"),
            "header_mode": "text",
            "header_text": "Кафедра",
            "footer_mode": "page",
            "footer_text": "",
        },
        content_type="multipart/form-data",
    )
    assert response.status_code == 200
    assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    document = Document(io.BytesIO(response.data))
    assert "Кафедра" in document.sections[0].header._element.xml
    assert "PAGE" in document.sections[0].footer._element.xml


def test_api_converts_zip_project_with_relative_image():
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("docs/report.md", "# Проект\n\n![Схема](../images/diagram.png)")
        zf.writestr("images/diagram.png", png_bytes())
    archive.seek(0)

    response = app.test_client().post(
        "/api/convert",
        data={
            "source": (archive, "project.zip"),
            "header_mode": "none",
            "footer_mode": "none",
        },
        content_type="multipart/form-data",
    )
    assert response.status_code == 200
    document = Document(io.BytesIO(response.data))
    assert len(document.inline_shapes) == 1
    assert document.inline_shapes[0]._inline.docPr.get("descr") == "Схема"


def test_api_returns_json_error_for_missing_assets():
    client = app.test_client()
    response = client.post(
        "/api/convert",
        data={
            "source": (io.BytesIO(b"![Missing](missing.png)"), "report.md"),
            "header_mode": "none",
            "footer_mode": "none",
            "asset_choices": json.dumps({}),
        },
        content_type="multipart/form-data",
    )
    assert response.status_code == 400
    assert "Не найдены изображения" in response.json["error"]
