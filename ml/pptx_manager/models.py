from enum import Enum
from typing import Any

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape
from pptx.text.text import Font, TextFrame
from pydantic import BaseModel


class TextAlign(str, Enum):
    LEFT = 'left'
    CENTER = 'center'
    RIGHT = 'right'
    JUSTIFY = 'justify'
    DISTRIBUTE = 'distribute'


class TextVerticalAlign(str, Enum):
    TOP = 'top'
    MIDDLE = 'middle'
    BOTTOM = 'bottom'


class TextFrameShape(BaseModel):
    shape_id: int
    text: str
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: tuple | None = None
    left: float | None = None
    top: float | None = None
    width: int | None = None
    height: int | None = None
    align: TextAlign | None = None
    vertical_align: TextVerticalAlign | None = None
    text_manager: TextFrame
    font_manager: Font
    shape_manager: Shape

    class Config:
        arbitrary_types_allowed = True


class TextFrameOpts(BaseModel):
    text: str | None = None
    font_name: str | None = None
    color: list[int] | None = None
    size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    align: TextAlign | None = None
    vertical_align: TextVerticalAlign | None = None


class UpdateTextFrameOpts(TextFrameOpts):
    left: int | None = None
    top: int | None = None
    width: int | None = None
    height: int | None = None


class CreateTextFrameOpts(TextFrameOpts):
    left: int
    top: int
    width: int
    height: int


class ShapeType(Enum):
    TEXT = MSO_SHAPE_TYPE.TEXT_BOX

    IMAGE = MSO_SHAPE_TYPE.PICTURE


class CreateShapeOpts(BaseModel):
    left: int | None = None
    top: int | None = None

    width: int | None = None
    height: int | None = None

    type: ShapeType


class ImageFrameOpts(BaseModel):
    width: float | None = None
    height: float | None = None

    left: float | None = None
    top: float | None = None


class CreateImageFrameOpts(ImageFrameOpts):
    image: bytes


class ImageFrameShape(BaseModel):
    shape_id: int

    width: float
    height: float

    left: float
    top: float

    blob: bytes | None = None

    shape_manager: Any


class SlideFrame(BaseModel):
    shape_count: int

    slide_manager: Any  # from pptx.slide import Slide

    background_color: list[int]


class AddShapeOnSlide(BaseModel):
    left: int
    top: int
    width: int
    height: int

    shape: Shape

    class Config:
        arbitrary_types_allowed = True
