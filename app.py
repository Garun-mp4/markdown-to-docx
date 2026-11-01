from __future__ import annotations

import io
import json
from pathlib import PurePath
from zipfile import BadZipFile

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from converter import (
    MAX_REQUEST_BYTES,
    ConversionError,
    SelectionRequired,
    build_docx,
    collect_project,
    prepare_images,
)
from docx_profile import ProfileValidationError, default_profile, validate_profile

MAX_MULTIPART_OVERHEAD_BYTES = 1024 * 1024

app = Flask(__name__)
app.config.update(
    MAX_CONTENT_LENGTH=MAX_REQUEST_BYTES + MAX_MULTIPART_OVERHEAD_BYTES,
    MAX_FORM_MEMORY_SIZE=1024 * 1024,
    MAX_FORM_PARTS=1520,
)


def _total_uploaded_file_bytes(uploads) -> int:
    total = 0
    for upload in uploads:
        stream = upload.stream
        position = stream.tell()
        stream.seek(0, io.SEEK_END)
        total += stream.tell()
        stream.seek(position)
    return total


@app.get("/")
def index():
    return render_template("index.html", default_profile=default_profile())


@app.get("/favicon.ico")
def favicon():
    return "", 204


@app.post("/api/convert")
def convert():
    source = request.files.get("source")
    if source is None or not source.filename:
        return jsonify(error="Выберите Markdown-файл или ZIP-архив проекта.", field="source"), 400

    try:
        resource_uploads = request.files.getlist("assets")
        all_uploads = [upload for _, uploads in request.files.lists() for upload in uploads]
        total_upload_bytes = _total_uploaded_file_bytes(all_uploads)
        if total_upload_bytes > MAX_REQUEST_BYTES:
            field = "assets" if resource_uploads else "source"
            return jsonify(error="Общий размер загружаемых файлов не должен превышать 200 МБ.", field=field), 413
        try:
            asset_choices = json.loads(request.form.get("asset_choices", "{}"))
        except json.JSONDecodeError as exc:
            raise ConversionError("Не удалось прочитать выбор изображений. Повторите выбор ресурсов.", "assets") from exc
        try:
            profile = json.loads(request.form["profile"]) if request.form.get("profile") else None
        except json.JSONDecodeError as exc:
            raise ConversionError("Не удалось прочитать настройки оформления. Верните их к значениям по умолчанию.", "profile") from exc
        if not isinstance(asset_choices, dict):
            raise ConversionError("Некорректный выбор файлов изображений.", "assets")
        try:
            profile = validate_profile(profile)
        except ProfileValidationError as exc:
            raise ConversionError(str(exc), "profile") from exc

        try:
            numbering_start_page = int(request.form.get("numbering_start_page", "1"))
        except ValueError as exc:
            raise ConversionError("Укажите целую страницу, с которой показывать нумерацию.", "headers") from exc
        if not 1 <= numbering_start_page <= 1000:
            raise ConversionError("Страница начала нумерации должна быть от 1 до 1000.", "headers")
        numbering_mode = request.form.get("numbering_mode", "continue")
        if numbering_mode not in {"continue", "restart"}:
            raise ConversionError("Выберите корректный режим нумерации страниц.", "headers")

        markdown_text, markdown_name, markdown_path, assets = collect_project(
            source,
            resource_uploads,
            request.form.get("main_path", ""),
        )
        try:
            images = prepare_images(markdown_text, markdown_path, assets, asset_choices)
        except ConversionError as exc:
            raise ConversionError(str(exc), "assets") from exc
        document = build_docx(
            markdown_text,
            images,
            header_mode=request.form.get("header_mode", "none"),
            header_text=request.form.get("header_text", ""),
            footer_mode=request.form.get("footer_mode", "none"),
            footer_text=request.form.get("footer_text", ""),
            profile=profile,
            numbering_start_page=numbering_start_page,
            numbering_mode=numbering_mode,
        )
    except SelectionRequired as exc:
        return jsonify(error=exc.message, field=exc.field, **exc.payload), 422
    except ConversionError as exc:
        return jsonify(error=str(exc), field=exc.field), 400
    except (BadZipFile, UnicodeError, ValueError) as exc:
        return jsonify(error=str(exc), field="source"), 400
    except RequestEntityTooLarge:
        return jsonify(error="Файлы слишком большие. Максимальный общий размер — 200 МБ."), 413

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
    return jsonify(error="Файлы слишком большие. Максимальный общий размер — 200 МБ.", field="source"), 413


if __name__ == "__main__":
    app.run(debug=False)
