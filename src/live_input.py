from __future__ import annotations

from pathlib import Path
from typing import Optional

import streamlit.components.v1 as components


_FRONTEND_DIR = Path(__file__).resolve().parent / "live_input_frontend"
_component = components.declare_component("koos_live_text", path=str(_FRONTEND_DIR))


def live_text(
    label: str,
    value: str = "",
    *,
    key: str,
    multiline: bool = False,
    height: int = 110,
    debounce: int = 220,
    placeholder: str = "",
    disabled: bool = False,
    label_visibility: str = "visible",
    max_chars: Optional[int] = None,
) -> str:
    """Text input/textarea that sends its value to Streamlit while typing.

    This local component is bundled with the project, so it does not use an
    external CDN or require a separate login. `debounce` prevents a full app
    rerun for every individual keystroke while still updating the preview
    automatically after a short pause.
    """
    frame_height = max(45, int(height) + (30 if label_visibility == "visible" else 4)) if multiline else (
        73 if label_visibility == "visible" else 45
    )
    result = _component(
        label=label,
        value=str(value or ""),
        multiline=bool(multiline),
        input_height=max(48, int(height)),
        debounce=max(0, int(debounce)),
        placeholder=placeholder,
        disabled=bool(disabled),
        label_visibility=label_visibility,
        max_chars=max_chars,
        frame_height=frame_height,
        key=f"live_{key}",
        default=str(value or ""),
    )
    return str(result if result is not None else value or "")
