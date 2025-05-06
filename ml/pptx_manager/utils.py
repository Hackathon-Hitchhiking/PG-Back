from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.shapes.base import BaseShape
from pptx.slide import Slide
from pptx.util import Emu, Inches, Pt


def get_all_methods(cls):
    return [method for method in dir(cls) if not method.startswith('__')]


def px_to_emu(px: int) -> int:
    dpi = 96
    inches = px / dpi
    return Inches(inches).emu


def emu_to_px(emu: int, dpi: int = 96) -> int:
    inches = Emu(emu).inches
    return round(inches * dpi)


def hex_to_rgb(hex: str) -> tuple[int, ...]:
    hex_code = hex.lstrip('#')

    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))


def pt_to_px(pt) -> float:
    return Pt(pt) * (4 / 3)


def px_to_pt(px) -> Pt:
    return Pt(px * 0.75)


def get_slide_from_shape(shape: BaseShape) -> Slide:
    for i in range(100):
        shape = shape._parent
        if isinstance(shape, Slide):
            return shape
    raise ValueError('shape is not in a slide')


class CustomList:
    def __init__(self):
        self._data = []

    def insert(self, index, value):
        # If index is larger than the current length, fill the "gap" with None
        if index > len(self._data):
            self._data.extend([None] * (index - len(self._data)))

        # Perform normal list insert (shift everything at/after 'index' to the right)
        self._data.insert(index, value)

    def __getitem__(self, index):
        return self._data[index]

    def __setitem__(self, index, value):
        # If index is beyond length, fill up to that index with None
        if index >= len(self._data):
            self._data.extend([None] * (index - len(self._data) + 1))
        self._data[index] = value

    def __len__(self):
        return len(self._data)

    def __repr__(self):
        return repr(self._data)


_TEXT_ALIGN_MAP = {
    'left': PP_ALIGN.LEFT,
    'center': PP_ALIGN.CENTER,
    'right': PP_ALIGN.RIGHT,
    'justify': PP_ALIGN.JUSTIFY,
    'distribute': PP_ALIGN.DISTRIBUTE,
}
_TEXT_ALIGN_MAP_REV = {v: k for k, v in _TEXT_ALIGN_MAP.items()}

_TEXT_VERTICAL_ALIGN_MAP = {
    'top': MSO_VERTICAL_ANCHOR.TOP,
    'middle': MSO_VERTICAL_ANCHOR.MIDDLE,
    'bottom': MSO_VERTICAL_ANCHOR.BOTTOM,
}
_TEXT_VERTICAL_ALIGN_MAP_REV = {v: k for k, v in _TEXT_VERTICAL_ALIGN_MAP.items()}


def parse_text_align(value: str | None) -> PP_ALIGN | None:
    if value is None:
        return None
    return _TEXT_ALIGN_MAP.get(value.lower())


def parse_text_vertical_align(value: str | None) -> MSO_VERTICAL_ANCHOR | None:
    if value is None:
        return None
    return _TEXT_VERTICAL_ALIGN_MAP.get(value.lower())


def text_align_to_str(align: PP_ALIGN | None) -> str | None:
    if align is None:
        return None
    return _TEXT_ALIGN_MAP_REV.get(align)


def text_vertical_align_to_str(valign: MSO_VERTICAL_ANCHOR | None) -> str | None:
    if valign is None:
        return None
    return _TEXT_VERTICAL_ALIGN_MAP_REV.get(valign)
