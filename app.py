from __future__ import annotations

import io
import json
from pathlib import PurePath
from zipfile import BadZipFile

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from converter import (
    ConversionError,
    SelectionRequired,
    build_docx,
    collect_project,
    prepare_images,
)

app = Flask(__name__)
app.config.update(
    MAX_CONTENT_LENGTH=50 * 1024 * 1024,
    MAX_FORM_MEMORY_SIZE=256 * 1024,
    MAX_FORM_PARTS=32,
)


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/favicon.ico")
def favicon():
    return "", 204


@app.post("/api/convert")
def convert():
    source = request.files.get("source")
    if source is None or not source.filename:
        return jsonify(error="Выберите Markdown-файл или ZIP-архив проекта."), 400

    try:
        asset_choices = json.loads(request.form.get("asset_choices", "{}"))
        if not isinstance(asset_choices, dict):
            raise ConversionError("Некорректный выбор файлов изображений.")

        markdown_text, markdown_name, markdown_path, assets = collect_project(
            source,
            request.files.get("assets"),
            request.form.get("main_path", ""),
        )
        images = prepare_images(markdown_text, markdown_path, assets, asset_choices)
        document = build_docx(
            markdown_text,
            images,
            header_mode=request.form.get("header_mode", "none"),
            header_text=request.form.get("header_text", ""),
            footer_mode=request.form.get("footer_mode", "none"),
            footer_text=request.form.get("footer_text", ""),
        )
    except SelectionRequired as exc:
        return jsonify(error=exc.message, **exc.payload), 422
    except (ConversionError, BadZipFile, UnicodeError, ValueError) as exc:
        return jsonify(error=str(exc)), 400
    except RequestEntityTooLarge:
        return jsonify(error="Файлы слишком большие. Максимальный общий размер — 50 МБ."), 413

    stem = PurePath(markdown_name).stem or "document"
    filename = f"{stem}-formatted.docx".replace("\r", "").replace("\n", "") or "document-formatted.docx"
    response = send_file(
        io.BytesIO(document),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename,
        max_age=0,
    )
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


@app.errorhandler(RequestEntityTooLarge)
def too_large(_error):
    return jsonify(error="Файлы слишком большие. Максимальный общий размер — 50 МБ."), 413


if __name__ == "__main__":
    app.run(debug=False)
