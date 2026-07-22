from __future__ import annotations

import base64
from pathlib import Path
from typing import Any, Dict


ASSET_DIR = Path(__file__).resolve().parent.parent / "assets"
FONT_DIR = ASSET_DIR / "fonts"
DEFAULT_FONT_NAME = "맑은 고딕"
FONT_EXTENSIONS = {".ttf", ".otf", ".woff", ".woff2"}


def _available_font_files() -> list[Path]:
    """Return font files actually present in assets/fonts.

    The dropdown is intentionally driven by the directory rather than a
    hard-coded JSON preset. Adding or removing a font file therefore changes
    the web options after the Streamlit app is restarted.
    """
    if not FONT_DIR.exists():
        return []
    return sorted(
        [
            path
            for path in FONT_DIR.iterdir()
            if path.is_file()
            and path.suffix.lower() in FONT_EXTENSIONS
            and not path.name.startswith("~$")
            and not path.name.startswith(".")
        ],
        key=lambda path: path.name.casefold(),
    )


def get_font_presets() -> Dict[str, Dict[str, str]]:
    """Build one selectable preset per physical font file.

    The dictionary key and display name are the exact filename requested by
    the user. The filename stem is used as the PowerPoint font name. This works
    best when the same font is installed in Windows under that name.
    """
    files = _available_font_files()
    if not files:
        return {
            "malgun_gothic": {
                "display_name": DEFAULT_FONT_NAME,
                "ppt_font_name": DEFAULT_FONT_NAME,
                "regular_file": "",
                "bold_file": "",
            }
        }

    presets: Dict[str, Dict[str, str]] = {}
    for path in files:
        presets[path.name] = {
            "display_name": path.name,
            "ppt_font_name": path.stem,
            # Selecting one file means that exact file is used throughout the
            # HTML preview. Register it for both normal and bold CSS weights.
            "regular_file": path.name,
            "bold_file": path.name,
        }
    return presets


def get_default_font_key() -> str:
    presets = get_font_presets()
    preferred = (
        "NanumSquareR.ttf",
        "NanumGothic.ttf",
    )
    for filename in preferred:
        if filename in presets:
            return filename
    return next(iter(presets), "malgun_gothic")


def get_font_preset(font_key: str | None = None) -> Dict[str, str]:
    presets = get_font_presets()
    key = str(font_key or get_default_font_key()).strip()
    if key not in presets:
        key = get_default_font_key()
    return dict(
        presets.get(key)
        or {
            "display_name": DEFAULT_FONT_NAME,
            "ppt_font_name": DEFAULT_FONT_NAME,
            "regular_file": "",
            "bold_file": "",
        }
    )


def get_ppt_font_name(font_key: str | None = None) -> str:
    return get_font_preset(font_key).get("ppt_font_name") or DEFAULT_FONT_NAME


def _font_face(file_name: str, weight: int, family_name: str) -> str:
    if not file_name:
        return ""
    path = FONT_DIR / file_name
    if not path.exists() or path.suffix.lower() not in FONT_EXTENSIONS:
        return ""
    mime = {
        ".ttf": "font/ttf",
        ".otf": "font/otf",
        ".woff": "font/woff",
        ".woff2": "font/woff2",
    }[path.suffix.lower()]
    payload = base64.b64encode(path.read_bytes()).decode("ascii")
    safe_family = family_name.replace("'", "")
    return (
        f"@font-face{{font-family:'{safe_family}';"
        f"src:url(data:{mime};base64,{payload});font-style:normal;"
        f"font-weight:{weight};font-display:swap;}}"
    )


def get_preview_font_css(font_key: str | None = None, family_name: str = "ReportAssetFont") -> str:
    preset = get_font_preset(font_key)
    regular = _font_face(preset.get("regular_file", ""), 400, family_name)
    bold = _font_face(preset.get("bold_file", ""), 700, family_name)
    return regular + bold


def get_preview_font_family(font_key: str | None = None, family_name: str = "ReportAssetFont") -> str:
    preset = get_font_preset(font_key)
    names = [family_name, preset.get("ppt_font_name", ""), DEFAULT_FONT_NAME, "Noto Sans KR"]
    unique = []
    for name in names:
        clean = str(name or "").strip().replace("'", "")
        if clean and clean not in unique:
            unique.append(clean)
    quoted = ",".join(f"'{name}'" for name in unique)
    return f"{quoted},sans-serif"


def get_font_status(font_key: str | None = None) -> Dict[str, Any]:
    preset = get_font_preset(font_key)
    regular_file = preset.get("regular_file", "")
    bold_file = preset.get("bold_file", "")
    regular_exists = not regular_file or (FONT_DIR / regular_file).exists()
    bold_exists = not bold_file or (FONT_DIR / bold_file).exists()
    return {
        "display_name": preset.get("display_name") or get_ppt_font_name(font_key),
        "ppt_font_name": get_ppt_font_name(font_key),
        "regular_file": regular_file,
        "bold_file": bold_file,
        "regular_exists": regular_exists,
        "bold_exists": bold_exists,
        "web_files_ready": regular_exists and bold_exists,
    }
