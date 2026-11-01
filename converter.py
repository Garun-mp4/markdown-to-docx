from __future__ import annotations

import http.client
import io
import ipaddress
import re
import socket
import ssl
import urllib.parse
import warnings
import zipfile
from dataclasses import dataclass
from pathlib import PurePosixPath, PureWindowsPath
from typing import Any

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Mm, Pt, RGBColor
from markdown_it import MarkdownIt
from PIL import Image, UnidentifiedImageError

MAX_ARCHIVE_ENTRIES = 1_500
MAX_ARCHIVE_UNPACKED = 100 * 1024 * 1024
MAX_IMAGE_BYTES = 10 * 1024 * 1024
MAX_TOTAL_IMAGE_BYTES = 30 * 1024 * 1024
MAX_REMOTE_REDIRECTS = 3
IMAGE_FORMATS = {"PNG", "JPEG", "GIF", "BMP", "TIFF", "WEBP"}
Image.MAX_IMAGE_PIXELS = 20_000_000
PAGEBREAK_MARKER = "DOCX_PAGEBREAK_4cf49205"
TABLE_CAPTION = re.compile(r"^\s*Таблица\s+([\w.-]+)\s*[—\-:]\s*(\S.*)\s*$", re.IGNORECASE)
FIGURE_CAPTION = re.compile(r"^\s*Рисунок\s+([\w.-]+)\s*[—\-:]\s*(\S.*)\s*$", re.IGNORECASE)
PAGEBREAK_DIRECTIVE = re.compile(r"(?m)^[ \t]*<!--[ \t]*pagebreak[ \t]*-->[ \t]*$", re.IGNORECASE)
SIMPLE_TABLE_BORDER = re.compile(r"^\s*-{12,}\s*$")
SIMPLE_TABLE_SEPARATOR = re.compile(r"^\s*-{3,}(?:\s+-{3,})+\s*$")
WINDOWS_ABSOLUTE = re.compile(r"^[A-Za-z]:[\\/]")


class ConversionError(Exception):
    pass


class SelectionRequired(Exception):
    def __init__(self, message: str, payload: dict[str, Any]):
        super().__init__(message)
        self.message = message
        self.payload = payload


@dataclass(frozen=True)
class Asset:
    asset_id: str
    archive_path: str
    data: bytes


@dataclass
class Block:
    kind: str
    value: Any = None
    level: int = 0
    children: list[Block] | None = None
    ordered: bool = False
    start_number: int = 1
    rows: list[list[Any]] | None = None
    caption: tuple[str, str] | None = None


def _normal_archive_path(path: str) -> str | None:
    path = path.replace("\\", "/")
    if path.startswith("/") or WINDOWS_ABSOLUTE.match(path):
        return None
    parts: list[str] = []
    for part in path.split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if not parts:
                return None
            parts.pop()
        else:
            parts.append(part)
    return "/".join(parts) or None


def _zip_assets(upload, label: str) -> tuple[list[Asset], list[tuple[str, bytes]]]:
    if upload is None or not upload.filename:
        return [], []
    if not upload.filename.lower().endswith(".zip"):
        raise ConversionError("Архив ресурсов должен иметь формат .zip.")

    raw = upload.read()
    if len(raw) > 50 * 1024 * 1024:
        raise ConversionError("ZIP-архив превышает лимит 50 МБ.")
    try:
        archive = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile as exc:
        raise ConversionError("Не удалось открыть ZIP-архив. Проверьте файл.") from exc

    infos = archive.infolist()
    if len(infos) > MAX_ARCHIVE_ENTRIES:
        raise ConversionError("В архиве слишком много файлов (максимум 1500).")

    assets: list[Asset] = []
    markdown_files: list[tuple[str, bytes]] = []
    total = 0
    with archive:
        for info in infos:
            if info.is_dir():
                continue
            safe_path = _normal_archive_path(info.filename)
            if safe_path is None:
                raise ConversionError("В архиве найден небезопасный путь к файлу.")
            mode = (info.external_attr >> 16) & 0o170000
            if mode == 0o120000:
                raise ConversionError("Символические ссылки в ZIP-архивах не поддерживаются.")
            total += info.file_size
            if info.file_size < 0 or total > MAX_ARCHIVE_UNPACKED:
                raise ConversionError("Распакованный архив превышает лимит 100 МБ.")
            if info.file_size > MAX_IMAGE_BYTES and not safe_path.lower().endswith(".md"):
                raise ConversionError(f"Файл {PurePosixPath(safe_path).name} превышает лимит 10 МБ.")
            try:
                data = archive.read(info)
            except (RuntimeError, zipfile.BadZipFile) as exc:
                raise ConversionError("Не удалось прочитать ZIP-архив.") from exc
            if safe_path.lower().endswith(".md"):
                markdown_files.append((safe_path, data))
            else:
                assets.append(Asset(f"{label}:{safe_path}", safe_path, data))
    return assets, markdown_files


def _decode_markdown(data: bytes) -> str:
    if len(data) > 5 * 1024 * 1024:
        raise ConversionError("Markdown-файл превышает лимит 5 МБ.")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ConversionError("Markdown-файл должен быть сохранён в кодировке UTF-8.") from exc


def collect_project(source, resources_upload, main_path: str) -> tuple[str, str, str, list[Asset]]:
    name = source.filename or ""
    if name.lower().endswith(".md"):
        markdown_text = _decode_markdown(source.read())
        markdown_name = PurePosixPath(name.replace("\\", "/")).name
        markdown_path = markdown_name
        project_assets: list[Asset] = []
    elif name.lower().endswith(".zip"):
        project_assets, markdown_files = _zip_assets(source, "project")
        if not markdown_files:
            raise ConversionError("В ZIP-архиве не найден Markdown-файл.")
        available = [path for path, _ in markdown_files]
        if main_path:
            selected = next(((path, data) for path, data in markdown_files if path == main_path), None)
            if selected is None:
                raise ConversionError("Выбранный Markdown-файл отсутствует в архиве.")
        elif len(markdown_files) == 1:
            selected = markdown_files[0]
        else:
            raise SelectionRequired(
                "В архиве несколько Markdown-файлов. Выберите основной файл.",
                {"markdown_files": available},
            )
        markdown_path, markdown_data = selected
        markdown_text = _decode_markdown(markdown_data)
        markdown_name = PurePosixPath(markdown_path).name
    else:
        raise ConversionError("Загрузите файл .md или ZIP-архив проекта.")

    resource_assets, resource_markdown = _zip_assets(resources_upload, "resources")
    if resource_markdown:
        raise ConversionError("В ZIP ресурсов не должно быть Markdown-файлов; загрузите их как основной источник.")
    return markdown_text, markdown_name, markdown_path, project_assets + resource_assets


def _markdown_parser():
    return MarkdownIt("commonmark", {"html": False}).enable(["table", "strikethrough"])


def _image_tokens(markdown: str):
    parser = _markdown_parser()
    prepared = PAGEBREAK_DIRECTIVE.sub(PAGEBREAK_MARKER, markdown)
    tokens = parser.parse(prepared)
    refs: list[str] = []
    for token in tokens:
        if token.type == "inline" and token.children:
            refs.extend(child.attrGet("src") for child in token.children if child.type == "image")
    return parser, tokens, [ref for ref in refs if ref]


def _simple_table_column_starts(header: str) -> list[int]:
    starts: list[int] = []
    previous_end = 0
    for match in re.finditer(r"\S+", header):
        if not starts or len(header[previous_end:match.start()]) >= 2:
            starts.append(match.start())
        previous_end = match.end()
    return starts


def _split_simple_table_row(line: str, starts: list[int]) -> list[str]:
    boundaries = [*starts, len(line)]
    return [line[boundaries[index]:boundaries[index + 1]].strip() for index in range(len(starts))]


def _normalize_pandoc_simple_tables(markdown: str) -> str:
    """Convert Pandoc simple tables into pipe tables before CommonMark parsing."""
    lines = markdown.splitlines()
    normalized: list[str] = []
    index = 0
    in_fence = False

    while index < len(lines):
        line = lines[index]
        if re.match(r"^\s{0,3}(```+|~~~+)", line):
            in_fence = not in_fence
            normalized.append(line)
            index += 1
            continue
        if in_fence or not SIMPLE_TABLE_BORDER.fullmatch(line) or index + 2 >= len(lines):
            normalized.append(line)
            index += 1
            continue

        header = lines[index + 1]
        separator = lines[index + 2]
        starts = _simple_table_column_starts(header)
        dash_groups = list(re.finditer(r"-{3,}", separator))
        if (
            not SIMPLE_TABLE_SEPARATOR.fullmatch(separator)
            or len(starts) < 2
            or len(dash_groups) != len(starts)
            or any(abs(group.start() - start) > 2 for group, start in zip(dash_groups, starts))
        ):
            normalized.append(line)
            index += 1
            continue

        bottom = next(
            (
                candidate
                for candidate in range(index + 3, len(lines))
                if SIMPLE_TABLE_BORDER.fullmatch(lines[candidate])
            ),
            None,
        )
        if bottom is None:
            normalized.append(line)
            index += 1
            continue

        body_lines = lines[index + 3:bottom]
        data_rows: list[list[str]] = []
        current_row = [""] * len(starts)
        physical_rows = 0

        for body_line in body_lines:
            if not body_line.strip():
                if any(current_row):
                    data_rows.append(current_row)
                    current_row = [""] * len(starts)
                continue
            physical_rows += 1
            for cell_index, value in enumerate(_split_simple_table_row(body_line, starts)):
                if value:
                    current_row[cell_index] = f"{current_row[cell_index]} {value}".strip()
        if any(current_row):
            data_rows.append(current_row)

        if not data_rows or (physical_rows > 1 and len(data_rows) == 1):
            normalized.append(line)
            index += 1
            continue

        header_cells = _split_simple_table_row(header, starts)
        table_rows = [header_cells, *data_rows]
        for row in table_rows:
            escaped = [cell.replace("|", r"\|") for cell in row]
            normalized.append("| " + " | ".join(escaped) + " |")
            if row is header_cells:
                normalized.append("| " + " | ".join("---" for _ in starts) + " |")
        index = bottom + 1

    return "\n".join(normalized)


def _is_absolute_local_path(src: str) -> bool:
    return WINDOWS_ABSOLUTE.match(src) is not None or src.startswith(("/", "\\\\"))


def _archive_index(assets: list[Asset]) -> dict[str, list[Asset]]:
    index: dict[str, list[Asset]] = {}
    for asset in assets:
        normalized = _normal_archive_path(asset.archive_path)
        if normalized:
            index.setdefault(normalized.casefold(), []).append(asset)
    return index


def _resolve_choice(src: str, candidates: list[Asset], choices: dict[str, Any]) -> Asset:
    choice = choices.get(src)
    if isinstance(choice, str):
        selected = next((candidate for candidate in candidates if candidate.asset_id == choice), None)
        if selected:
            return selected
    if len(candidates) == 1:
        return candidates[0]
    raise SelectionRequired(
        "Несколько загруженных файлов подходят для одной ссылки. Выберите нужный файл.",
        {
            "ambiguous_images": [
                {
                    "src": src,
                    "matches": [
                        {"id": item.asset_id, "path": item.archive_path}
                        for item in candidates
                    ],
                }
            ]
        },
    )


def _validate_public_host(host: str, port: int = 443) -> tuple[str, int]:
    if not host or port != 443:
        raise ConversionError("Для удалённых изображений поддерживаются только HTTPS-ссылки на стандартном порту.")
    try:
        infos = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
    except socket.gaierror as exc:
        raise ConversionError("Не удалось найти сервер удалённого изображения.") from exc
    addresses: list[tuple[str, int]] = []
    for family, _socktype, _proto, _canonname, sockaddr in infos:
        ip_text = sockaddr[0]
        try:
            ip = ipaddress.ip_address(ip_text)
        except ValueError:
            continue
        if not ip.is_global:
            raise ConversionError("Ссылка на изображение ведёт в закрытую или локальную сеть.")
        addresses.append((ip_text, family))
    if not addresses:
        raise ConversionError("У сервера изображения нет общедоступного IP-адреса.")
    return addresses[0]


def _fetch_https_image(url: str) -> bytes:
    current = url
    for redirect_count in range(MAX_REMOTE_REDIRECTS + 1):
        parsed = urllib.parse.urlsplit(current)
        if parsed.scheme.lower() != "https" or parsed.username or parsed.password:
            raise ConversionError("Удалённые изображения должны использовать безопасную HTTPS-ссылку.")
        host = parsed.hostname or ""
        host = host.encode("idna").decode("ascii")
        port = parsed.port or 443
        address, family = _validate_public_host(host, port)
        path = urllib.parse.quote(parsed.path or "/", safe="/%:@!$&'()*+,;=-._~")
        query = urllib.parse.quote(parsed.query, safe="/%?:@!$&'()*+,;=-._~")
        target = f"{path}?{query}" if query else path

        raw = socket.socket(family, socket.SOCK_STREAM)
        raw.settimeout(6)
        try:
            sockaddr = (address, port, 0, 0) if family == socket.AF_INET6 else (address, port)
            raw.connect(sockaddr)
            tls = ssl.create_default_context().wrap_socket(raw, server_hostname=host)
        except (OSError, ssl.SSLError) as exc:
            raw.close()
            raise ConversionError("Не удалось загрузить удалённое изображение.") from exc

        response = None
        try:
            request = (
                f"GET {target} HTTP/1.1\r\n"
                f"Host: {host}\r\n"
                "Accept: image/*\r\n"
                "Accept-Encoding: identity\r\n"
                "Connection: close\r\n\r\n"
            ).encode("ascii", errors="strict")
            tls.sendall(request)
            response = http.client.HTTPResponse(tls, method="GET")
            response.begin()
            if response.status in {301, 302, 303, 307, 308}:
                location = response.getheader("Location")
                if not location or redirect_count == MAX_REMOTE_REDIRECTS:
                    raise ConversionError("Удалённое изображение перенаправляется слишком много раз.")
                current = urllib.parse.urljoin(current, location)
                continue
            if response.status < 200 or response.status >= 300:
                raise ConversionError(f"Сервер изображения вернул код {response.status}.")
            content_length = response.getheader("Content-Length")
            if content_length and int(content_length) > MAX_IMAGE_BYTES:
                raise ConversionError("Удалённое изображение превышает лимит 10 МБ.")
            data = response.read(MAX_IMAGE_BYTES + 1)
            if len(data) > MAX_IMAGE_BYTES:
                raise ConversionError("Удалённое изображение превышает лимит 10 МБ.")
            return data
        except (OSError, http.client.HTTPException, UnicodeEncodeError, ValueError) as exc:
            if isinstance(exc, ConversionError):
                raise
            raise ConversionError("Не удалось прочитать удалённое изображение.") from exc
        finally:
            if response is not None:
                response.close()
            else:
                tls.close()
    raise ConversionError("Не удалось загрузить удалённое изображение.")


def _sanitize_image(data: bytes, src: str) -> bytes:
    if len(data) > MAX_IMAGE_BYTES:
        raise ConversionError(f"Изображение {src} превышает лимит 10 МБ.")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(data)) as image:
                if image.format not in IMAGE_FORMATS:
                    raise ConversionError(f"Формат изображения {src} не поддерживается.")
                dpi = image.info.get("dpi", (96, 96))
                image.seek(0)
                image.load()
                image = image.convert("RGBA") if "A" in image.getbands() else image.convert("RGB")
                output = io.BytesIO()
                image.save(output, format="PNG", optimize=True, dpi=dpi)
                sanitized = output.getvalue()
                if len(sanitized) > MAX_IMAGE_BYTES:
                    raise ConversionError(f"Обработанное изображение {src} превышает лимит 10 МБ.")
                return sanitized
    except UnidentifiedImageError as exc:
        raise ConversionError(f"Файл по ссылке {src} не является поддерживаемым изображением.") from exc
    except (Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise ConversionError(f"Размер изображения {src} слишком велик для безопасной обработки.") from exc
    except OSError as exc:
        raise ConversionError(f"Не удалось обработать изображение {src}.") from exc


def prepare_images(markdown: str, markdown_path: str, assets: list[Asset], choices: dict[str, Any]) -> dict[str, bytes]:
    _parser, _tokens, sources = _image_tokens(markdown)
    index = _archive_index(assets)
    base_dir = str(PurePosixPath(markdown_path).parent)
    if base_dir == ".":
        base_dir = ""

    resolved: dict[str, bytes] = {}
    unresolved: list[str] = []
    total_bytes = 0
    for src in dict.fromkeys(sources):
        decoded_src = urllib.parse.unquote(src)
        parsed = urllib.parse.urlsplit(src)
        if _is_absolute_local_path(decoded_src):
            basename = PureWindowsPath(decoded_src).name if "\\" in decoded_src else PurePosixPath(decoded_src).name
            matches = [asset for asset in assets if PurePosixPath(asset.archive_path).name.casefold() == basename.casefold()]
            if not matches:
                unresolved.append(src)
                continue
            selected = _resolve_choice(src, matches, choices)
            raw = selected.data
        elif parsed.scheme.lower() in {"http", "https"}:
            if parsed.scheme.lower() != "https":
                raise ConversionError("Удалённые изображения должны использовать HTTPS.")
            raw = _fetch_https_image(src)
        elif parsed.scheme:
            raise ConversionError(f"Тип ссылки на изображение {src} не поддерживается.")
        else:
            relative = urllib.parse.unquote(parsed.path).replace("\\", "/")
            joined = f"{base_dir}/{relative}" if base_dir else relative
            normalized = _normal_archive_path(joined)
            matches = index.get(normalized.casefold(), []) if normalized else []
            if not matches:
                unresolved.append(src)
                continue
            selected = _resolve_choice(src, matches, choices)
            raw = selected.data

        sanitized = _sanitize_image(raw, src)
        total_bytes += len(sanitized)
        if total_bytes > MAX_TOTAL_IMAGE_BYTES:
            raise ConversionError("Суммарный размер изображений превышает лимит 30 МБ.")
        resolved[src] = sanitized

    if unresolved:
        raise ConversionError(
            "Не найдены изображения: " + ", ".join(unresolved) + ". Добавьте их в ZIP ресурсов и повторите конвертацию."
        )
    return resolved


def _matching_close(tokens, start: int, end: int) -> int:
    depth = 0
    for index in range(start, end):
        depth += tokens[index].nesting
        if depth == 0:
            return index
    return end - 1


def _plain_inline(token) -> str:
    if token.children is None:
        return token.content
    parts = []
    for child in token.children:
        if child.type in {"softbreak", "hardbreak"}:
            parts.append(" ")
        elif child.type in {"text", "code_inline", "html_inline", "image"}:
            parts.append(child.content)
    return "".join(parts)


def _parse_table(tokens, start: int, end: int) -> list[list[Any]]:
    rows: list[list[Any]] = []
    index = start + 1
    while index < end:
        if tokens[index].type != "tr_open":
            index += 1
            continue
        index += 1
        row: list[Any] = []
        while index < end and tokens[index].type != "tr_close":
            if tokens[index].type in {"th_open", "td_open"}:
                cell_token = tokens[index + 1] if index + 1 < end and tokens[index + 1].type == "inline" else None
                row.append(cell_token)
                index += 1
            index += 1
        rows.append(row)
        index += 1
    return rows


def _parse_list(tokens, start: int, end: int, list_type: str) -> Block:
    items: list[Block] = []
    index = start + 1
    while index < end:
        if tokens[index].type != "list_item_open":
            index += 1
            continue
        item_end = _matching_close(tokens, index, end + 1)
        children = _parse_blocks(tokens, index + 1, item_end)
        items.append(Block("list_item", children=children))
        index = item_end + 1
    opening = tokens[start]
    return Block(
        "list",
        children=items,
        ordered=list_type == "ordered_list_open",
        start_number=int(opening.attrGet("start") or "1") if list_type == "ordered_list_open" else 1,
    )


def _parse_blocks(tokens, start: int = 0, end: int | None = None) -> list[Block]:
    if end is None:
        end = len(tokens)
    blocks: list[Block] = []
    index = start
    while index < end:
        token = tokens[index]
        if token.type in {"paragraph_open", "heading_open"}:
            inline = tokens[index + 1] if index + 1 < end and tokens[index + 1].type == "inline" else None
            if token.type == "paragraph_open":
                blocks.append(Block("paragraph", inline))
            else:
                level = int(token.tag[1:]) if token.tag.startswith("h") else 1
                blocks.append(Block("heading", inline, level=level))
            index += 3
        elif token.type in {"bullet_list_open", "ordered_list_open"}:
            close = _matching_close(tokens, index, end)
            blocks.append(_parse_list(tokens, index, close - 1, token.type))
            index = close + 1
        elif token.type == "blockquote_open":
            close = _matching_close(tokens, index, end)
            blocks.append(Block("quote", children=_parse_blocks(tokens, index + 1, close)))
            index = close + 1
        elif token.type == "table_open":
            close = _matching_close(tokens, index, end)
            blocks.append(Block("table", rows=_parse_table(tokens, index, close)))
            index = close + 1
        elif token.type in {"fence", "code_block"}:
            blocks.append(Block("code", token))
            index += 1
        elif token.type == "hr":
            blocks.append(Block("hr"))
            index += 1
        else:
            index += 1
    return blocks


def _tag_block_captions(blocks: list[Block]) -> None:
    for index, block in enumerate(blocks):
        if block.kind == "table" and index > 0 and blocks[index - 1].kind == "paragraph":
            caption_text = _plain_inline(blocks[index - 1].value) if blocks[index - 1].value else ""
            matched = TABLE_CAPTION.match(caption_text)
            if matched:
                block.caption = (matched.group(1), matched.group(2))
                blocks[index - 1].kind = "caption_consumed"
                blocks[index - 1].value = matched
        if block.kind == "list" or block.kind == "quote":
            for item in block.children or []:
                if item.kind == "list_item":
                    _tag_block_captions(item.children or [])
                else:
                    _tag_block_captions([item])


def _style_run(run, *, size=12, bold=False, italic=False, strike=False, code=False, underline=False):
    run.bold = bold
    run.italic = italic
    run.font.strike = strike
    run.underline = underline
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = "Consolas" if code else "Times New Roman"
    run.font.size = Pt(10 if code else size)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas" if code else "Times New Roman")
    run._element.rPr.rFonts.set(qn("w:cs"), "Consolas" if code else "Times New Roman")
    if code:
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "EEEEEE")
        run._element.get_or_add_rPr().append(shading)


def _add_hyperlink(paragraph, text: str, url: str, bold: bool, italic: bool, strike: bool):
    relationship_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "Times New Roman")
    properties.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "24")
    properties.append(size)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.append(underline)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    properties.append(color)
    if bold:
        properties.append(OxmlElement("w:b"))
    if italic:
        properties.append(OxmlElement("w:i"))
    if strike:
        properties.append(OxmlElement("w:strike"))
    run.append(properties)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


class DocxRenderer:
    def __init__(self, image_data: dict[str, bytes], header_mode: str, header_text: str, footer_mode: str, footer_text: str):
        self.document = Document()
        self.image_data = image_data
        self._configure_page()
        self._configure_styles()
        self._configure_field_update()
        self._configure_header_footer(header_mode, header_text, footer_mode, footer_text)
        self.content_width = self.document.sections[0].page_width - self.document.sections[0].left_margin - self.document.sections[0].right_margin

    def _configure_page(self):
        for section in self.document.sections:
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.left_margin = Mm(30)
            section.right_margin = Mm(15)
            section.top_margin = Mm(20)
            section.bottom_margin = Mm(20)
            section.header_distance = Mm(10)
            section.footer_distance = Mm(10)

    def _configure_styles(self):
        styles = self.document.styles
        normal = styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(12)
        normal.font.color.rgb = RGBColor(0, 0, 0)
        normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        normal._element.rPr.rFonts.set(qn("w:cs"), "Times New Roman")
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        normal.paragraph_format.first_line_indent = Mm(12.5)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.widow_control = True

        sizes = [20, 18, 16, 14, 13, 12]
        indents = [0, 5, 10, 15, 20, 25]
        for level, (size, indent) in enumerate(zip(sizes, indents), start=1):
            style = styles[f"Heading {level}"]
            style.font.name = "Times New Roman"
            style.font.size = Pt(size)
            style.font.bold = True
            style.font.color.rgb = RGBColor(0, 0, 0)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            style._element.rPr.rFonts.set(qn("w:cs"), "Times New Roman")
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            style.paragraph_format.left_indent = Mm(indent)
            style.paragraph_format.first_line_indent = Mm(0)
            style.paragraph_format.line_spacing = 1.5
            style.paragraph_format.space_before = Pt(12)
            style.paragraph_format.space_after = Pt(6)
            style.paragraph_format.keep_with_next = True
            style.paragraph_format.keep_together = True

    def _configure_field_update(self):
        settings = self.document.settings._element
        update = settings.find(qn("w:updateFields"))
        if update is None:
            update = OxmlElement("w:updateFields")
            settings.append(update)
        update.set(qn("w:val"), "true")

    def _configure_header_footer(self, header_mode, header_text, footer_mode, footer_text):
        for mode, text, paragraph in (
            (header_mode, header_text, self.document.sections[0].header.paragraphs[0]),
            (footer_mode, footer_text, self.document.sections[0].footer.paragraphs[0]),
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if mode == "page" else WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Mm(0)
            paragraph.paragraph_format.line_spacing = 1
            paragraph.paragraph_format.space_after = Pt(0)
            if mode == "text":
                _style_run(paragraph.add_run(text), size=10)
            elif mode == "page":
                _append_field(paragraph, "PAGE", "1")

    @staticmethod
    def _apply_quote_border(paragraph):
        p_pr = paragraph._p.get_or_add_pPr()
        borders = p_pr.find(qn("w:pBdr"))
        if borders is None:
            borders = OxmlElement("w:pBdr")
            p_pr.append(borders)
        left = borders.find(qn("w:left"))
        if left is None:
            left = OxmlElement("w:left")
            borders.append(left)
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "8")
        left.set(qn("w:space"), "6")
        left.set(qn("w:color"), "666666")

    def _new_ordered_list_id(self, start_number: int, depth: int, quote_depth: int) -> int:
        numbering = self.document.part.numbering_part.element
        abstract_ids = [int(node.get(qn("w:abstractNumId"))) for node in numbering.findall(qn("w:abstractNum"))]
        num_ids = [int(node.get(qn("w:numId"))) for node in numbering.findall(qn("w:num"))]
        abstract_id = max(abstract_ids, default=-1) + 1
        num_id = max(num_ids, default=0) + 1

        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        multi_level = OxmlElement("w:multiLevelType")
        multi_level.set(qn("w:val"), "singleLevel")
        abstract.append(multi_level)
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), "0")
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        level.append(start)
        number_format = OxmlElement("w:numFmt")
        number_format.set(qn("w:val"), "decimal")
        level.append(number_format)
        level_text = OxmlElement("w:lvlText")
        level_text.set(qn("w:val"), "%1.")
        level.append(level_text)
        justification = OxmlElement("w:lvlJc")
        justification.set(qn("w:val"), "left")
        level.append(justification)
        paragraph_properties = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "num")
        tab_position = 397 + depth * 397 + quote_depth * 567
        tab.set(qn("w:pos"), str(tab_position))
        tabs.append(tab)
        paragraph_properties.append(tabs)
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), str(tab_position))
        indent.set(qn("w:hanging"), "227")
        paragraph_properties.append(indent)
        level.append(paragraph_properties)
        abstract.append(level)

        first_num = numbering.find(qn("w:num"))
        insertion_index = numbering.index(first_num) if first_num is not None else len(numbering)
        numbering.insert(insertion_index, abstract)

        number = OxmlElement("w:num")
        number.set(qn("w:numId"), str(num_id))
        abstract_reference = OxmlElement("w:abstractNumId")
        abstract_reference.set(qn("w:val"), str(abstract_id))
        number.append(abstract_reference)
        if start_number != 1:
            override = OxmlElement("w:lvlOverride")
            override.set(qn("w:ilvl"), "0")
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), str(start_number))
            override.append(start_override)
            number.append(override)
        numbering.append(number)
        return num_id

    @staticmethod
    def _set_paragraph_numbering(paragraph, num_id: int):
        p_pr = paragraph._p.get_or_add_pPr()
        num_pr = OxmlElement("w:numPr")
        level = OxmlElement("w:ilvl")
        level.set(qn("w:val"), "0")
        number = OxmlElement("w:numId")
        number.set(qn("w:val"), str(num_id))
        num_pr.append(level)
        num_pr.append(number)
        p_pr.insert_element_before(
            num_pr,
            "w:suppressLineNumbers",
            "w:pBdr",
            "w:shd",
            "w:tabs",
            "w:spacing",
            "w:ind",
            "w:contextualSpacing",
            "w:jc",
            "w:outlineLvl",
            "w:rPr",
            "w:sectPr",
            "w:pPrChange",
        )

    def add_paragraph(self, inline, *, alignment=None, indent=0, first_line=12.5, keep=False, size=12, bold=False, italic=False, quote_depth=0, code_block=False, cell=False):
        paragraph = self.document.add_paragraph()
        self._format_paragraph(paragraph, indent=indent, first_line=first_line, keep=keep, quote_depth=quote_depth, cell=cell)
        if alignment is not None:
            paragraph.alignment = alignment
        if code_block:
            paragraph.paragraph_format.left_indent = Mm(10 + quote_depth * 10)
            paragraph.paragraph_format.first_line_indent = Mm(0)
            paragraph.paragraph_format.line_spacing = 1
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(4)
            p_pr = paragraph._p.get_or_add_pPr()
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "F2F2F2")
            p_pr.append(shading)
        self._append_inline(paragraph, inline, size=size, bold=bold, italic=italic, code_block=code_block)
        return paragraph

    def _format_paragraph(self, paragraph, *, indent=0, first_line=12.5, keep=False, quote_depth=0, cell=False):
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.left_indent = Mm(indent + quote_depth * 10)
        paragraph.paragraph_format.first_line_indent = Mm(0 if cell or quote_depth else first_line)
        paragraph.paragraph_format.right_indent = Mm(8 if quote_depth else 0)
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_before = Pt(4 if quote_depth else 0)
        paragraph.paragraph_format.space_after = Pt(4 if quote_depth else 0)
        paragraph.paragraph_format.keep_together = keep
        paragraph.paragraph_format.widow_control = True
        if quote_depth:
            self._apply_quote_border(paragraph)

    def _append_inline(
        self,
        paragraph,
        inline,
        *,
        size=12,
        bold=False,
        italic=False,
        code_block=False,
        omit_text_prefix: str | None = None,
    ):
        if inline is None or inline.children is None:
            if inline and inline.content:
                _style_run(paragraph.add_run(inline.content), size=size, bold=bold, italic=italic, code=code_block)
            return
        bold_stack = [bold]
        italic_stack = [italic]
        strike_stack = [False]
        link_stack: list[str | None] = []
        for child in inline.children:
            content = child.content
            if child.type == "text" and omit_text_prefix and content.startswith(omit_text_prefix):
                content = content[len(omit_text_prefix):]
                omit_text_prefix = None
            if child.type in {"strong_open"}:
                bold_stack.append(True)
            elif child.type in {"strong_close"}:
                if len(bold_stack) > 1:
                    bold_stack.pop()
            elif child.type in {"em_open"}:
                italic_stack.append(True)
            elif child.type in {"em_close"}:
                if len(italic_stack) > 1:
                    italic_stack.pop()
            elif child.type in {"s_open"}:
                strike_stack.append(True)
            elif child.type in {"s_close"}:
                if len(strike_stack) > 1:
                    strike_stack.pop()
            elif child.type == "link_open":
                link_stack.append(child.attrGet("href"))
            elif child.type == "link_close":
                if link_stack:
                    link_stack.pop()
            elif child.type in {"text", "code_inline"}:
                is_code = code_block or child.type == "code_inline"
                if link_stack and link_stack[-1]:
                    scheme = urllib.parse.urlsplit(link_stack[-1]).scheme.lower()
                    if scheme in {"http", "https", "mailto"}:
                        _add_hyperlink(
                            paragraph,
                            content,
                            link_stack[-1],
                            bold_stack[-1],
                            italic_stack[-1],
                            strike_stack[-1],
                        )
                    else:
                        _style_run(
                            paragraph.add_run(content),
                            size=size,
                            bold=bold_stack[-1],
                            italic=italic_stack[-1],
                            strike=strike_stack[-1],
                        )
                else:
                    _style_run(
                        paragraph.add_run(content),
                        size=size,
                        bold=bold_stack[-1],
                        italic=italic_stack[-1],
                        strike=strike_stack[-1],
                        code=is_code,
                    )
            elif child.type == "image":
                src = child.attrGet("src") or ""
                data = self.image_data[src]
                with Image.open(io.BytesIO(data)) as image:
                    width_px = image.width
                    dpi = image.info.get("dpi", (96, 96))[0]
                dpi_width = max(1, width_px) / (dpi if dpi and dpi > 0 else 96)
                width = min(Inches(dpi_width), self.content_width)
                run = paragraph.add_run()
                picture = run.add_picture(io.BytesIO(data), width=width)
                alt = child.attrGet("alt") or child.content or ""
                picture._inline.docPr.set("descr", alt)
            elif child.type == "softbreak":
                paragraph.add_run(" ")
            elif child.type == "hardbreak":
                paragraph.add_run().add_break()
            elif child.type == "html_inline":
                _style_run(
                    paragraph.add_run(child.content),
                    size=size,
                    bold=bold_stack[-1],
                    italic=italic_stack[-1],
                    strike=strike_stack[-1],
                )

    def _add_caption(self, caption: tuple[str, str], *, table: bool, page_break_before: bool = False):
        number, title = caption
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if table else WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Mm(0)
        paragraph.paragraph_format.left_indent = Mm(0)
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(3 if table else 6)
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.page_break_before = page_break_before
        _style_run(paragraph.add_run(f"{'Таблица' if table else 'Рисунок'} {number} — {title}"), bold=table)

    @staticmethod
    def _estimated_cell_lines(inline, characters_per_line: int) -> int:
        text = " ".join((_plain_inline(inline) if inline else "").split())
        words = text.split()
        if not words:
            return 1
        lines = 1
        occupied = 0
        for word in words:
            remaining = word
            while remaining:
                available = characters_per_line - occupied - (1 if occupied else 0)
                if available <= 0:
                    lines += 1
                    occupied = 0
                    available = characters_per_line
                take = min(len(remaining), available)
                occupied += take + (1 if occupied else 0)
                remaining = remaining[take:]
                if remaining:
                    lines += 1
                    occupied = 0
            if occupied >= characters_per_line:
                lines += 1
                occupied = 0
        return max(1, lines)

    def _table_column_widths(self, rows: list[list[Any]], column_count: int) -> list[float]:
        total_width = self.content_width / 36_000
        minimum = min(18.0, total_width / column_count)
        maximum = total_width if column_count <= 2 else total_width * 0.48
        weights: list[float] = []
        for column_index in range(column_count):
            cell_texts = [
                " ".join(_plain_inline(row[column_index]).split())
                for row in rows
                if column_index < len(row) and row[column_index]
                if " ".join(_plain_inline(row[column_index]).split())
            ]
            if not cell_texts:
                weights.append(12.0)
                continue
            lengths = sorted(len(text) for text in cell_texts)
            typical = lengths[int((len(lengths) - 1) * 0.75)]
            longest_word = max(len(word) for text in cell_texts for word in text.split())
            typical_weight = min(48, max(8, typical))
            word_weight = min(72, longest_word * 2)
            weights.append(float(max(typical_weight, word_weight)))

        widths = [0.0] * column_count
        active = set(range(column_count))
        remaining = total_width
        while active:
            total_weight = sum(weights[index] for index in active)
            constrained = False
            for index in list(active):
                proposed = remaining * weights[index] / total_weight
                if proposed < minimum:
                    widths[index] = minimum
                    remaining -= minimum
                    active.remove(index)
                    constrained = True
                elif proposed > maximum:
                    widths[index] = maximum
                    remaining -= maximum
                    active.remove(index)
                    constrained = True
            if not constrained:
                for index in active:
                    widths[index] = remaining * weights[index] / total_weight
                break
        return widths

    def _estimated_row_height(self, row_tokens: list[Any], column_widths: list[float]) -> float:
        estimated_lines = []
        for column_index, cell in enumerate(row_tokens):
            width_mm = column_widths[column_index] if column_index < len(column_widths) else min(column_widths)
            characters_per_line = max(8, int((width_mm - 3.2) / 1.9))
            estimated_lines.append(self._estimated_cell_lines(cell, characters_per_line))
        text_lines = max(estimated_lines, default=1)
        image_height = 35 if any(
            child.type == "image"
            for cell in row_tokens
            if cell and cell.children
            for child in cell.children
        ) else 0
        return max(text_lines * 6.35 + 2.5, image_height + 2.5)

    def _estimated_paragraph_height_mm(self, paragraph) -> float:
        font_size = paragraph.style.font.size or Pt(12)
        width_mm = self.content_width / 36_000
        characters_per_line = max(12, int(width_mm / (font_size.pt * 0.17)))
        text_length = len(" ".join(paragraph.text.split()))
        line_count = max(1, (text_length + characters_per_line - 1) // characters_per_line)
        spacing = paragraph.paragraph_format.line_spacing
        line_multiplier = spacing if isinstance(spacing, (float, int)) else 1.5
        line_height_mm = font_size.pt * line_multiplier / 2.83465
        before = paragraph.paragraph_format.space_before or paragraph.style.paragraph_format.space_before or Pt(0)
        after = paragraph.paragraph_format.space_after or paragraph.style.paragraph_format.space_after or Pt(0)
        return line_count * line_height_mm + before / 36_000 + after / 36_000

    def _table_page_capacities(self, previous, has_caption: bool) -> tuple[float, float]:
        section = self.document.sections[0]
        page_height = (section.page_height - section.top_margin - section.bottom_margin) / 36_000
        heading_height = (
            self._estimated_paragraph_height_mm(previous)
            if has_caption and previous and previous.style.name.startswith("Heading")
            else 0.0
        )
        first_capacity = page_height - 5.0 - (10.0 if has_caption else 0.0) - heading_height
        later_capacity = page_height - 5.0 - 8.0
        return max(30.0, first_capacity), max(30.0, later_capacity)

    def _paginate_table_rows(
        self,
        rows: list[list[Any]],
        column_widths: list[float],
        first_capacity: float,
        later_capacity: float,
    ) -> list[list[list[Any]]]:
        header_height = self._estimated_row_height(rows[0], column_widths) if rows else 0
        pages: list[list[list[Any]]] = []
        current: list[list[Any]] = []
        capacity = max(30, first_capacity - header_height)
        used = 0.0
        for row in rows[1:]:
            height = self._estimated_row_height(row, column_widths)
            if height + header_height > later_capacity:
                raise ConversionError("Одна строка таблицы слишком высокая для страницы. Разбейте её содержимое на несколько строк.")
            if current and used + height > capacity:
                pages.append(current)
                current = []
                capacity = max(30, later_capacity - header_height)
                used = 0.0
            if not current and height > capacity:
                raise ConversionError("Одна строка таблицы слишком высокая для страницы. Разбейте её содержимое на несколько строк.")
            current.append(row)
            used += height
        if current or not pages:
            pages.append(current)
        return pages

    def _add_continuation_label(self, number: str):
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.first_line_indent = Mm(0)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.keep_with_next = True
        _style_run(paragraph.add_run(f"Продолжение таблицы {number}"))

    def _add_table_rows(self, rows: list[list[Any]], column_widths: list[float]):
        if not rows:
            return
        column_count = len(column_widths)
        table = self.document.add_table(rows=0, cols=column_count)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        widths = [Emu(int(width_mm * 36_000)) for width_mm in column_widths]
        for column_index, column in enumerate(table.columns):
            column.width = widths[column_index]
        _set_table_borders(table)

        for row_index, row_tokens in enumerate(rows):
            row = table.add_row()
            _set_row_cant_split(row)
            if row_index == 0:
                _mark_repeat_header(row)
            for cell_index, inline in enumerate(row_tokens):
                cell = row.cells[cell_index]
                cell.width = widths[cell_index]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                paragraph = cell.paragraphs[0]
                self._format_paragraph(paragraph, first_line=0, cell=True)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 else WD_ALIGN_PARAGRAPH.LEFT
                self._append_inline(paragraph, inline, bold=row_index == 0)
                _set_cell_margins(cell)

    def add_table(self, block: Block):
        caption = block.caption
        rows = block.rows or []
        cols = max((len(row) for row in rows), default=0)
        if not cols:
            return
        column_widths = self._table_column_widths(rows, cols)
        previous_paragraphs = self.document.paragraphs
        previous = previous_paragraphs[-1] if previous_paragraphs else None
        first_capacity, later_capacity = self._table_page_capacities(previous, bool(caption))
        estimated_total = sum(self._estimated_row_height(row, column_widths) for row in rows)
        if estimated_total <= first_capacity:
            if caption:
                self._add_caption(caption, table=True)
            self._add_table_rows(rows, column_widths)
            return

        if caption:
            page_break_before = False
            if previous and previous.style.name.startswith("Heading"):
                if len(previous_paragraphs) > 1:
                    previous.paragraph_format.page_break_before = True
            elif previous:
                page_break_before = True
            self._add_caption(caption, table=True, page_break_before=page_break_before)
        else:
            self.document.add_page_break()
        pages = self._paginate_table_rows(rows, column_widths, first_capacity, later_capacity)
        for page_index, data_rows in enumerate(pages):
            if page_index:
                self.document.add_page_break()
                if caption:
                    self._add_continuation_label(caption[0])
            self._add_table_rows([rows[0], *data_rows], column_widths)

    def add_image_paragraph(self, inline):
        return self.add_paragraph(inline, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line=0, keep=True)

    def add_list(self, block: Block, depth: int = 0, quote_depth: int = 0):
        numbering_id = self._new_ordered_list_id(block.start_number, depth, quote_depth) if block.ordered else None
        for item_index, item in enumerate(block.children or []):
            item_blocks = item.children or []
            for child_index, child in enumerate(item_blocks):
                if child.kind == "paragraph":
                    style_name = "List Number" if block.ordered else "List Bullet"
                    paragraph = self.document.add_paragraph(style=style_name)
                    paragraph.paragraph_format.left_indent = Mm(7 + depth * 7 + quote_depth * 10)
                    paragraph.paragraph_format.first_line_indent = Mm(-4)
                    paragraph.paragraph_format.line_spacing = 1.5
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.right_indent = Mm(8 if quote_depth else 0)
                    if numbering_id is not None:
                        self._set_paragraph_numbering(paragraph, numbering_id)
                    if quote_depth:
                        self._apply_quote_border(paragraph)
                    task_prefix = None
                    task_match = None
                    if child.value and child.value.children:
                        first_text = next((part for part in child.value.children if part.type == "text"), None)
                        task_match = re.match(r"^(\s*\[([ xX])\]\s+)", first_text.content) if first_text else None
                    if task_match:
                        task_prefix = task_match.group(1)
                        checkbox = paragraph.add_run("☑ " if task_match.group(2).strip() else "☐ ")
                        _style_run(checkbox)
                        checkbox.font.name = "Segoe UI Symbol"
                        checkbox._element.rPr.rFonts.set(qn("w:eastAsia"), "Segoe UI Symbol")
                        checkbox._element.rPr.rFonts.set(qn("w:cs"), "Segoe UI Symbol")
                    self._append_inline(paragraph, child.value, omit_text_prefix=task_prefix)
                else:
                    self.add_block(child, quote_depth=quote_depth, list_depth=depth + 1)

    def add_block(self, block: Block, *, quote_depth: int = 0, list_depth: int = 0):
        if block.kind == "heading":
            paragraph = self.document.add_paragraph(style=f"Heading {max(1, min(6, block.level))}")
            if quote_depth:
                paragraph.paragraph_format.left_indent = (paragraph.paragraph_format.left_indent or Mm(0)) + Mm(quote_depth * 10)
                paragraph.paragraph_format.first_line_indent = Mm(0)
                paragraph.paragraph_format.right_indent = Mm(8)
                paragraph.paragraph_format.space_before = Pt(4)
                paragraph.paragraph_format.space_after = Pt(4)
                self._apply_quote_border(paragraph)
            self._append_inline(paragraph, block.value, size=[20, 18, 16, 14, 13, 12][max(0, min(5, block.level - 1))], bold=True)
        elif block.kind == "paragraph":
            if block.value and _plain_inline(block.value).strip() == PAGEBREAK_MARKER:
                self.document.add_page_break()
                return
            children = block.value.children if block.value else []
            if children and len(children) == 1 and children[0].type == "image":
                self.add_image_paragraph(block.value)
            else:
                alignment = WD_ALIGN_PARAGRAPH.CENTER if children and any(child.type == "image" for child in children) else None
                self.add_paragraph(block.value, alignment=alignment, quote_depth=quote_depth)
        elif block.kind == "table":
            self.add_table(block)
        elif block.kind == "list":
            self.add_list(block, list_depth, quote_depth)
        elif block.kind == "quote":
            for child in block.children or []:
                self.add_block(child, quote_depth=quote_depth + 1)
        elif block.kind == "code":
            text = block.value.content.rstrip("\n")
            for line in text.split("\n") or [""]:
                p = self.add_paragraph(None, code_block=True, quote_depth=quote_depth, first_line=0)
                _style_run(p.add_run(line), size=10, code=True)
        elif block.kind == "hr":
            paragraph = self.document.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Mm(0)
            paragraph.paragraph_format.space_before = Pt(5)
            paragraph.paragraph_format.space_after = Pt(5)
            p_bdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "4")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "000000")
            p_bdr.append(bottom)
            paragraph._p.get_or_add_pPr().append(p_bdr)

    def render(self, markdown: str) -> bytes:
        prepared = PAGEBREAK_DIRECTIVE.sub(PAGEBREAK_MARKER, markdown)
        # Reparse the prepared source so the explicit page-break directive remains a recognizable block.
        prepared = _normalize_pandoc_simple_tables(prepared)
        parser = _markdown_parser()
        blocks = _parse_blocks(parser.parse(prepared))
        _tag_block_captions(blocks)
        i = 0
        while i < len(blocks):
            block = blocks[i]
            if block.kind == "paragraph" and block.value and len(block.value.children or []) == 1 and block.value.children[0].type == "image":
                caption = FIGURE_CAPTION.match(_plain_inline(blocks[i + 1].value)) if i + 1 < len(blocks) and blocks[i + 1].kind == "paragraph" and blocks[i + 1].value else None
                self.add_image_paragraph(block.value)
                if caption:
                    self._add_caption((caption.group(1), caption.group(2)), table=False)
                    i += 2
                    continue
                i += 1
                continue
            self.add_block(block)
            i += 1
        output = io.BytesIO()
        self.document.save(output)
        return output.getvalue()


def build_docx(markdown: str, images: dict[str, bytes], *, header_mode: str, header_text: str, footer_mode: str, footer_text: str) -> bytes:
    if len(markdown.encode("utf-8")) > 5 * 1024 * 1024:
        raise ConversionError("Markdown-файл превышает лимит 5 МБ.")
    if not markdown.strip():
        raise ConversionError("Markdown-файл пустой.")
    header_mode, header_text = _validate_footer(header_mode, header_text, "верхнего")
    footer_mode, footer_text = _validate_footer(footer_mode, footer_text, "нижнего")
    renderer = DocxRenderer(images, header_mode, header_text, footer_mode, footer_text)
    return renderer.render(markdown)


def _validate_footer(mode: str, text: str, label: str) -> tuple[str, str]:
    if mode not in {"none", "text", "page"}:
        raise ConversionError(f"Выберите корректный режим {label} колонтитула.")
    text = text.strip()
    if len(text) > 240:
        raise ConversionError(f"Текст {label} колонтитула не должен превышать 240 символов.")
    if mode == "text" and not text:
        raise ConversionError(f"Введите текст {label} колонтитула или отключите его.")
    return mode, text


def _append_field(paragraph, instruction: str, result: str, hidden: bool = False):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    paragraph.add_run()._r.append(begin)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    paragraph.add_run()._r.append(instr)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    paragraph.add_run()._r.append(separate)
    text_run = paragraph.add_run(result)
    if hidden:
        text_run.font.hidden = True
        text_run.font.size = Pt(1)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)


def _mark_repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    table_header = OxmlElement("w:tblHeader")
    table_header.set(qn("w:val"), "true")
    tr_pr.append(table_header)


def _set_row_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def _set_cell_margins(cell, top=70, start=90, bottom=70, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for margin, value in (("top", top), ("left", start), ("bottom", bottom), ("right", end)):
        node = margins.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
