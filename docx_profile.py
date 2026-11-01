from __future__ import annotations

import math
import re
from copy import deepcopy
from typing import Any


def _heading(size: int, indent: int) -> dict[str, Any]:
    return {
        "font_family": "Times New Roman",
        "font_size_pt": size,
        "color": "#000000",
        "bold": True,
        "italic": False,
        "alignment": "left",
        "left_indent_mm": indent,
        "first_line_indent_mm": 0,
        "line_spacing": 1.5,
        "space_before_pt": 12,
        "space_after_pt": 6,
        "keep_with_next": True,
        "keep_together": True,
    }


DEFAULT_PROFILE: dict[str, Any] = {
    "page": {
        "width_mm": 210,
        "height_mm": 297,
        "left_margin_mm": 30,
        "right_margin_mm": 15,
        "top_margin_mm": 20,
        "bottom_margin_mm": 20,
        "header_distance_mm": 10,
        "footer_distance_mm": 10,
    },
    "body": {
        "font_family": "Times New Roman",
        "font_size_pt": 12,
        "color": "#000000",
        "alignment": "left",
        "line_spacing": 1.5,
        "first_line_indent_mm": 12.5,
        "space_before_pt": 0,
        "space_after_pt": 0,
    },
    "headings": [_heading(size, indent) for size, indent in zip((20, 18, 16, 14, 13, 12), (0, 5, 10, 15, 20, 25))],
    "lists": {
        "left_indent_mm": 7,
        "depth_indent_mm": 7,
        "hanging_indent_mm": 4,
        "line_spacing": 1.5,
        "space_after_pt": 0,
    },
    "quotes": {
        "left_indent_mm": 10,
        "depth_indent_mm": 10,
        "right_indent_mm": 8,
        "line_spacing": 1.5,
        "space_before_pt": 4,
        "space_after_pt": 4,
        "border_color": "#666666",
        "border_width_pt": 1,
    },
    "code": {
        "font_family": "Consolas",
        "font_size_pt": 10,
        "color": "#000000",
        "background_color": "#F2F2F2",
        "left_indent_mm": 10,
        "line_spacing": 1,
        "space_before_pt": 4,
        "space_after_pt": 4,
    },
    "tables": {
        "alignment": "left",
        "header_font_family": "Times New Roman",
        "header_font_size_pt": 12,
        "header_color": "#000000",
        "header_bold": True,
        "header_alignment": "center",
        "body_font_family": "Times New Roman",
        "body_font_size_pt": 12,
        "body_color": "#000000",
        "body_alignment": "left",
        "border_color": "#000000",
        "border_width_pt": 0.5,
        "vertical_alignment": "center",
        "padding_top_mm": 1.23,
        "padding_horizontal_mm": 1.59,
        "padding_bottom_mm": 1.23,
    },
    "images": {
        "alignment": "center",
        "max_width_percent": 100,
    },
    "captions": {
        "table": {
            "font_family": "Times New Roman",
            "font_size_pt": 12,
            "color": "#000000",
            "bold": True,
            "alignment": "left",
            "space_before_pt": 6,
            "space_after_pt": 3,
        },
        "figure": {
            "font_family": "Times New Roman",
            "font_size_pt": 12,
            "color": "#000000",
            "bold": False,
            "alignment": "center",
            "space_before_pt": 6,
            "space_after_pt": 6,
        },
        "continuation": {
            "font_family": "Times New Roman",
            "font_size_pt": 12,
            "color": "#000000",
            "bold": False,
            "alignment": "left",
            "space_before_pt": 0,
            "space_after_pt": 4,
        },
    },
    "links": {
        "font_family": "Times New Roman",
        "font_size_pt": 12,
        "color": "#000000",
        "underline": True,
    },
    "horizontal_rule": {
        "color": "#000000",
        "width_pt": 0.5,
        "space_before_pt": 5,
        "space_after_pt": 5,
    },
    "headers_footers": {
        "font_family": "Times New Roman",
        "font_size_pt": 10,
        "color": "#000000",
        "text_alignment": "left",
        "page_alignment": "center",
    },
}


class ProfileValidationError(ValueError):
    pass


def default_profile() -> dict[str, Any]:
    return deepcopy(DEFAULT_PROFILE)


def _bounds(path: str) -> tuple[float, float]:
    key = path.rsplit(".", 1)[-1]
    if key in {"width_mm", "height_mm"}:
        return 100, 420
    if key.endswith("_mm"):
        return 0, 10 if "padding" in key else 100
    if key == "line_spacing":
        return 1, 3
    if key == "max_width_percent":
        return 10, 100
    if key.endswith("_pt"):
        if "font_size" in key:
            return 6, 72
        if "width" in key:
            return 0.1, 3
        return 0, 72
    return 0, 100


def _validate(value: Any, default: Any, path: str) -> Any:
    if isinstance(default, dict):
        if not isinstance(value, dict):
            raise ProfileValidationError(f"Параметр «{path}» должен быть объектом настроек.")
        missing = default.keys() - value.keys()
        unknown = value.keys() - default.keys()
        if missing:
            name = f"{path}.{min(missing)}"
            raise ProfileValidationError(f"Не задан параметр оформления «{name}».")
        if unknown:
            name = f"{path}.{min(unknown)}"
            raise ProfileValidationError(f"Неизвестный параметр оформления «{name}».")
        return {key: _validate(value[key], default[key], f"{path}.{key}") for key in default}

    if isinstance(default, list):
        if not isinstance(value, list) or len(value) != len(default):
            raise ProfileValidationError(f"Параметр «{path}» должен содержать ровно {len(default)} уровней.")
        return [_validate(item, default[index], f"{path}.{index + 1}") for index, item in enumerate(value)]

    if isinstance(default, bool):
        if not isinstance(value, bool):
            raise ProfileValidationError(f"Параметр «{path}» должен быть логическим значением.")
        return value

    if isinstance(default, (int, float)):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise ProfileValidationError(f"Параметр «{path}» должен быть числом.")
        minimum, maximum = _bounds(path)
        if not minimum <= value <= maximum:
            raise ProfileValidationError(f"Параметр «{path}» должен быть от {minimum:g} до {maximum:g}.")
        return value

    if not isinstance(value, str):
        raise ProfileValidationError(f"Параметр «{path}» должен быть текстом.")
    value = value.strip()
    if path.endswith(("color", "background_color")):
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
            raise ProfileValidationError(f"Укажите цвет параметра «{path}» в формате #RRGGBB.")
        return value.upper()
    if path.endswith("vertical_alignment") and value not in {"top", "center", "bottom"}:
        raise ProfileValidationError(f"Выберите допустимое вертикальное выравнивание для «{path}».")
    if (
        path.endswith("alignment")
        and not path.endswith("vertical_alignment")
        and value not in ({"left", "center", "right"} if path.endswith("tables.alignment") else {"left", "center", "right", "justify"})
    ):
        raise ProfileValidationError(f"Выберите допустимое выравнивание для «{path}».")
    if path.endswith("font_family") and (not value or len(value) > 80 or any(ord(character) < 32 for character in value)):
        raise ProfileValidationError(f"Название шрифта в «{path}» должно содержать от 1 до 80 символов.")
    return value


def validate_profile(value: Any | None) -> dict[str, Any]:
    if value is None:
        return default_profile()
    profile = _validate(value, DEFAULT_PROFILE, "profile")
    page = profile["page"]
    if page["left_margin_mm"] + page["right_margin_mm"] >= page["width_mm"] - 30:
        raise ProfileValidationError(
            "Параметры «profile.page.left_margin_mm» и «profile.page.right_margin_mm» должны оставлять не менее 30 мм для текста."
        )
    if page["top_margin_mm"] + page["bottom_margin_mm"] >= page["height_mm"] - 30:
        raise ProfileValidationError(
            "Параметры «profile.page.top_margin_mm» и «profile.page.bottom_margin_mm» должны оставлять не менее 30 мм для текста."
        )
    return profile
