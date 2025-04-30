import json
import os
import subprocess

from collections import defaultdict
from io import BytesIO
from tempfile import TemporaryDirectory
from textwrap import dedent

from dotenv import load_dotenv
from loguru import logger
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape

from ml.pptx_manager.figure import FigureManager
from ml.pptx_manager.image import ImageManager
from ml.pptx_manager.models import (
    CreateImageFrameOpts,
    CreateShapeOpts,
    CreateTextFrameOpts,
    ShapeType,
    UpdateTextFrameOpts,
)
from ml.pptx_manager.slide import SlideManager
from ml.pptx_manager.table import TableManager
from ml.pptx_manager.text import TextFrameManager

from ml.pptx_manager.shape import ShapeManager


class PPTXManager(
    TextFrameManager,
    ImageManager,
    ShapeManager,
    SlideManager,
    TableManager,
    FigureManager,
):
    def __init__(
        self, source: str | bytes | None, parse_slide_image: bool = False
    ) -> None:
        super().__init__()
        TextFrameManager.__init__(self)
        ImageManager.__init__(self)
        ShapeManager.__init__(self)
        SlideManager.__init__(self)
        TableManager.__init__(self)
        FigureManager.__init__(self)

        if isinstance(source, bytes):
            self._source_bytes = source
            self.pres = Presentation(BytesIO(source))
            self.source = None
        elif isinstance(source, str):
            self.pres = Presentation(source)
            self.source = source
        elif source is None:
            self.pres = Presentation()
            self.source = None
        else:
            raise TypeError("source must be str (filename), bytes (raw pptx) or None")

        self.parse_choice = {
            MSO_SHAPE_TYPE.PICTURE: self._parse_image_shape,
            MSO_SHAPE_TYPE.AUTO_SHAPE: self._parse_auto_shape,
            MSO_SHAPE_TYPE.TEXT_BOX: self._parse_text_shape,
            MSO_SHAPE_TYPE.GROUP: self.parse_group_shape,
        }

        self.slide_count = len(self.pres.slides)

        self.slide_image = {}

        self.parse_presentation()

        if parse_slide_image:
            self.parse_slide_as_images()

    def get_slide_image(self, slide_id: int) -> bytes:
        return self.slide_image[slide_id]

    def parse_slide_as_images(self):
        with TemporaryDirectory() as temp_dir:
            subprocess.run(
                [
                    os.environ.get("LIBREOFFICE_PATH", "libreoffice"),
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temp_dir,
                    self.source,
                ],
                check=True,
            )

            pptx_filename = os.path.basename(self.source)
            base_name = os.path.splitext(pptx_filename)[0]
            pdf_path = os.path.join(temp_dir, base_name + ".pdf")

            if not os.path.exists(pdf_path):
                raise FileNotFoundError(
                    f"PDF conversion failed; file not found at {pdf_path}"
                )

            pages = convert_from_path(pdf_path, dpi=200)

            for idx, page in enumerate(pages):
                image_buffer = BytesIO()
                page.save(image_buffer, format="PNG")

                self.slide_image[idx + 1] = image_buffer.getvalue()

    def parse_presentation(self) -> None:
        logger.debug(
            f"Анализ презентации вызван с количеством слайдов = {len(self.pres.slides)}"
        )
        slide_id = 1

        for slide in self.pres.slides:
            self.parse_slide(slide_id, slide)
            for shape in slide.shapes:
                parse_fn = self.parse_choice.get(shape.shape_type)
                if parse_fn is not None:
                    result = parse_fn(slide_id, self.get_shape_count(slide_id), shape)
                    if result is not None:
                        self.increase_shape_count(slide_id, 1)
            slide_id += 1

    def parse_group_shape(self, slide_id: int, shape_id: int, shape: Shape) -> None:
        for group_shape in shape.shapes:
            parse_fn = self.parse_choice.get(group_shape.shape_type)
            if parse_fn is not None:
                result = parse_fn(slide_id, self.get_shape_count(slide_id), group_shape)
                if result is not None:
                    self.increase_shape_count(slide_id, 1)

    def _parse_auto_shape(self, slide_id: int, shape_id: int, shape: Shape) -> None:
        result = self._parse_text_shape(slide_id, shape_id, shape)
        if result is not None:
            self.increase_shape_count(slide_id, 1)
            return

        match shape.auto_shape_type:
            case MSO_AUTO_SHAPE_TYPE.RECTANGLE:
                result = self._parse_figure_shape(slide_id, shape_id, shape)
                if result is not None:
                    self.increase_shape_count(slide_id, 1)
                    return
            case MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE:
                result = self._parse_figure_shape(slide_id, shape_id, shape)
                if result is not None:
                    self.increase_shape_count(slide_id, 1)
                    return

    def get_json_schema(self) -> dict:
        pres_json = {}
        for slide_id in range(1, self.slide_count + 1):
            text_json = self.get_text_frame_json(slide_id)
            image_json = self.get_image_json(slide_id)
            slide_json = self.get_slide_json(slide_id)
            figure_json = self.get_figure_frame_json(slide_id)

            pres_json[slide_id] = {
                "text": text_json,
                "image": image_json,
                "slide": slide_json,
                "figure": figure_json,
            }

        return pres_json

    def create_text_shape(self, slide_id: int, opts: CreateTextFrameOpts) -> str:
        """
        Создает новую текстовую фигуру на указанном слайде с заданными параметрами.

        Эта функция добавляет текстовую фигуру в презентацию и применяет указанные параметры форматирования, включая позицию, размер, текстовое содержание, стиль текста и выравнивание.

        Args:
            slide_id (int): ID слайда, на котором будет создана текстовая фигура.
            opts (CreateTextFrameOpts): Объект, содержащий параметры для создания текстовой фигуры.
                -   width (int): Ширина рамки текста.
                -   height (int): Высота рамки текста.
                -   left (int): Позиция рамки текста по оси X.
                -   top (int): Позиция рамки текста по оси Y.
                -   text (str | None, optional): Текстовое содержимое для фигуры.
                -   color (list[int] | None, optional): Цвет текста в формате RGB кортежа.
                -   size (float): Размер шрифта для текста.
                -   bold (bool | None, optional): Установить текст жирным или нет.
                -   italic (bool | None, optional): Установить текст курсивом или нет.
                -   underline (bool | None, optional): Установить подчеркивание текста или нет.
                -   align (str | None, optional): Горизонтальное выравнивание текста ("left", "center", "right", "justify", "distribute").
                -   vertical_align (str | None, optional): Вертикальное выравнивание текста ("top", "middle", "bottom").

        Returns:
            str: Сообщение о результате операции с подробным описанием созданной фигуры.
        """
        logger.debug(
            f"Вызов create_text_shape с параметрами slide_id={slide_id}, opts={opts}"
        )
        shape, shape_id = self._add_shape_on_slide(
            slide_id,
            CreateShapeOpts(
                left=opts.left,
                top=opts.top,
                height=opts.height,
                width=opts.width,
                type=ShapeType.TEXT,
            ),
        )

        self._parse_text_shape(slide_id, shape_id, shape)

        self.update_text_frame_shape(
            slide_id,
            shape_id,
            UpdateTextFrameOpts(
                text=dedent(opts.text),
                color=opts.color,
                size=opts.size,
                bold=opts.bold,
                italic=opts.italic,
                underline=opts.underline,
                font_name=opts.font_name,
                align=opts.align,
                vertical_align=opts.vertical_align,
            ),
        )

        return f'Создана текстовая фигура на слайде {slide_id} (shape_id: {shape_id}) с параметрами: позиция ({opts.left}, {opts.top}), размер ({opts.width}x{opts.height}), текст: "{opts.text}"'

    def delete_text_shape(self, slide_id: int, shape_id: int) -> str:
        """
        Удаляет текстовую фигуру с указанного слайда.

        Эта функция удаляет текстовую фигуру с заданным shape_id со слайда с заданным slide_id. Она пытается удалить фигуру как из внутренней системы отслеживания, так и из самого объекта презентации.

        Args:
            slide_id (int): ID слайда, содержащего текстовую область для удаления.
            shape_id (int): ID фигуры для удаления.

        Returns:
            str: Сообщение о результате операции с информацией об удаленной фигуре.
        """
        logger.debug(
            f"Вызов delete_text_shape с параметрами slide_id={slide_id}, shape_id={shape_id}"
        )
        shape = self._get_text_frame_shape(slide_id, shape_id)
        self._delete_text_frame_shape(slide_id, shape_id)

        try:
            el = shape.shape_manager.element
            parent = el.getparent()
            parent.remove(el)
        except AttributeError:
            logger.warning(f"Фигура с shape_id={shape_id} не найдена в shape_manager")
            return f"Неизвестный идентификатор фигуры: {shape_id} на слайде {slide_id}"

        return f"Текстовая фигура с ID {shape_id} успешно удалена со слайда {slide_id}"

    def delete_image_shape(self, slide_id: int, shape_id: int) -> str:
        """
        Удаляет фигуру с изображением с указанного слайда.

        Эта функция удаляет фигуру с изображением с заданным shape_id со слайда с заданным slide_id. Она пытается удалить фигуру как из внутренней системы отслеживания, так и из самого объекта презентации.

        Args:
            slide_id (int): ID слайда, содержащего изображение для удаления.
            shape_id (int): ID фигуры с изображением для удаления.

        Returns:
            str: Сообщение о результате операции с информацией об удаленном изображении.
        """
        logger.debug(
            f"Вызов delete_image_shape с параметрами slide_id={slide_id}, shape_id={shape_id}"
        )
        shape = self._get_image_frame_shape(slide_id, shape_id)
        self._delete_image_frame_shape(slide_id, shape_id)

        try:
            el = shape.shape_manager.element
            parent = el.getparent()
            parent.remove(el)
        except AttributeError:
            logger.warning(f"Фигура с shape_id={shape_id} не найдена в shape_manager")
            return f"Неизвестный идентификатор фигуры с изображением: {shape_id} на слайде {slide_id}"

        return f"Фигура с изображением (ID {shape_id}) успешно удалена со слайда {slide_id}"

    def delete_figure_shape(self, slide_id: int, shape_id: int) -> str:
        """
        Удаляет геометрическую фигуру (figure) с указанного слайда.

        Args:
            slide_id (int): ID слайда, содержащего фигуру для удаления.
            shape_id (int): ID фигуры для удаления.

        Returns:
            str: Сообщение о результате операции.
        """
        logger.debug(
            f"Вызов delete_figure_shape с параметрами slide_id={slide_id}, shape_id={shape_id}"
        )
        shape = self._get_shape(slide_id, shape_id)
        self._delete_shape(slide_id, shape_id)

        try:
            el = shape.shape_manager.element
            parent = el.getparent()
            parent.remove(el)
        except AttributeError:
            logger.warning(f"Фигура с shape_id={shape_id} не найдена в shape_manager")
            return f"Неизвестный идентификатор фигуры: {shape_id} на слайде {slide_id}"

        return f"Геометрическая фигура с ID {shape_id} успешно удалена со слайда {slide_id}"

    def delete_slide(self, slide_id: int) -> str:
        """
        Удаляет слайд из презентации.

        Args:
            slide_id (int): ID слайда для удаления.

        Returns:
            str: Сообщение о результате операции.
        """
        logger.debug(f"Вызов delete_slide с параметром slide_id={slide_id}")
        try:
            self._delete_slide(slide_id)
        except Exception as e:
            logger.warning(f"Ошибка при удалении слайда {slide_id}: {e}")
            return f"Ошибка при удалении слайда {slide_id}: {e}"
        return f"Слайд с ID {slide_id} успешно удалён"

    def create_image_shape(self, slide_id: int, opts: CreateImageFrameOpts) -> str:
        """
        Создает новую фигуру с изображением на указанном слайде с заданными параметрами.

        Эта функция добавляет изображение в презентацию и применяет указанные параметры форматирования, включая позицию и размер.

        Args:
            slide_id (int): ID слайда, на котором будет создана фигура с изображением.
            opts (CreateImageFrameOpts): Объект, содержащий параметры для создания фигуры с изображением.
                -   width (int): Ширина рамки изображения.
                -   height (int): Высота рамки изображения.
                -   left (int): Позиция рамки изображения по оси X.
                -   top (int): Позиция рамки изображения по оси Y.
                -   image (bytes): Байты изображения для размещения в фигуре.

        Returns:
            str: Сообщение о результате операции с подробным описанием созданной фигуры с изображением.
        """
        logger.debug(
            f"Вызов create_image_shape с параметрами slide_id={slide_id}, opts={opts.model_dump(exclude={'image'})}"
        )

        shape, shape_id = self._add_image_on_slide(
            slide_id,
            opts.image,
            CreateShapeOpts(
                left=opts.left,
                top=opts.top,
                height=opts.height,
                width=opts.width,
                type=ShapeType.IMAGE,
            ),
        )

        self._parse_image_shape(slide_id, shape_id, shape)

        return f"Создана фигура с изображением на слайде {slide_id} (shape_id: {shape_id}) с параметрами: позиция ({opts.left}, {opts.top}), размер ({opts.width}x{opts.height})"

    def copy_image_shape(
        self, slide_id_from: int, shape_id: int, slide_id_to: int
    ) -> str:
        """
        Копирует изображение с одного слайда на другой.

        Args:
            slide_id_from (int): ID исходного слайда.
            shape_id (int): ID фигуры с изображением на исходном слайде.
            slide_id_to (int): ID целевого слайда.

        Returns:
            str: Сообщение о результате операции.
        """
        logger.debug(
            f"Вызов copy_image_shape с параметрами slide_id_from={slide_id_from}, shape_id={shape_id}, slide_id_to={slide_id_to}"
        )

        shape = self._get_image_frame_shape(slide_id_from, shape_id)
        if shape is None:
            return f"Фигура с изображением (ID {shape_id}) не найдена на слайде {slide_id_from}."

        if shape.blob is None:
            return (
                f"Фигура с изображением (ID {shape_id}) не содержит данных изображения."
            )

        opts = CreateImageFrameOpts(
            left=shape.left,
            top=shape.top,
            width=shape.width,
            height=shape.height,
            image=shape.blob,
        )

        result = self.create_image_shape(slide_id_to, opts)
        return f"Изображение (shape_id={shape_id}) успешно скопировано со слайда {slide_id_from} на слайд {slide_id_to}. {result}"

    # TODO: работает хуево, нужно исправить.
    def copy_template_slide_to_presentation(
        self, slide_id: int, template_id: int, output_path: str
    ) -> str:
        """
        Копирует единственный слайд из шаблона templates/template_{template_id}.pptx
        и вставляет его в основную презентацию под номером slide_id (1-based).
        Сохраняет результат в output_path.

        Args:
            slide_id (int): Позиция для вставки нового слайда (1-based).
            template_id (int): ID шаблона (например, 1 или 2).
            output_path (str): Путь для сохранения итоговой презентации.

        Returns:
            str: Сообщение о результате операции.
        """
        template_path = f"templates/template_{template_id}.pptx"
        if not os.path.exists(template_path):
            return f"Файл шаблона не найден: {template_path}"

        # Используем PPTXManager для шаблона, чтобы получить корректную схему
        template_pr = PPTXManager(template_path, True)
        if template_pr.slide_count != 1:
            return f"В шаблоне {template_path} должен быть ровно один слайд"

        schema = template_pr.get_json_schema()
        slide_schema = schema[1]  # шаблон всегда один слайд, id=1

        # Вставляем новый слайд в нужную позицию
        self.add_slide_at_position(slide_id, layout_index=0)

        # Копируем фигуры (figure)
        for fig in slide_schema.get("figure", []):
            try:
                self.add_figure_shape(
                    slide_id,
                    fig["shape_type"],
                    fig["left"],
                    fig["top"],
                    fig["width"],
                    fig["height"],
                    color=fig.get("color"),
                    line_color=fig.get("line_color"),
                    line_width=fig.get("line_width"),
                    rounding=fig.get("rounding"),
                    transparency=fig.get("transparency"),
                    rotation=fig.get("rotation"),
                )
            except Exception as e:
                logger.warning(f"Ошибка при копировании фигуры: {e}")

        # Копируем текстовые блоки
        for text in slide_schema.get("text", []):
            try:
                self.create_text_shape(
                    slide_id,
                    CreateTextFrameOpts(
                        left=text["left"],
                        top=text["top"],
                        width=text["width"],
                        height=text["height"],
                        text=text.get("text", ""),
                        color=text.get("color"),
                        size=text.get("size"),
                        bold=text.get("bold"),
                        italic=text.get("italic"),
                        underline=text.get("underline"),
                        font_name=text.get("font_name"),
                        align=text.get("align"),
                        vertical_align=text.get("vertical_align"),
                    ),
                )
            except Exception as e:
                logger.warning(f"Ошибка при копировании текстового блока: {e}")

        # Копируем изображения
        for img in slide_schema.get("image", []):
            try:
                image_bytes = template_pr.get_image_frame_shape_blob(1, img["shape_id"])
                self.create_image_shape(
                    slide_id,
                    CreateImageFrameOpts(
                        left=img["left"],
                        top=img["top"],
                        width=img["width"],
                        height=img["height"],
                        image=image_bytes,
                    ),
                )
            except Exception as e:
                logger.warning(f"Ошибка при копировании изображения: {e}")

        # Можно добавить копирование таблиц и других элементов по аналогии

        # Сохраняем результат
        try:
            self.save(output_path)
        except Exception as e:
            logger.error(f"Ошибка при сохранении презентации: {e}")
            return f"Ошибка при сохранении презентации: {e}"

        return f"Слайд из шаблона template_{template_id}.pptx успешно вставлен как слайд #{slide_id} и сохранён в {output_path}"

    def get_tasks_from_slide(self):
        text_frames = self.get_all_text_frame_json()

        tasks: defaultdict[int, str] = defaultdict(str)

        for slide_id, texts in text_frames.items():
            for text in texts:
                if text["color"] == (255, 64, 0):
                    tasks[slide_id] += text["text"] + "\n"
                    self.delete_text_shape(slide_id, text["shape_id"])

        return tasks

    def save(self, path: str) -> None:
        self.pres.save(path)

    def to_bytes(self) -> bytes:
        buffer = BytesIO()
        self.pres.save(buffer)
        return buffer.getvalue()


if __name__ == "__main__":
    load_dotenv()
    pr = PPTXManager(os.environ.get("TEST_PRES_PATH"))

    # rect1 = pr.add_figure_shape(
    #     2,
    #     MSO_SHAPE.ROUNDED_RxwzECTANGLE,
    #     300,
    #     500,
    #     100,
    #     500,
    #     color=(255, 128, 255),
    #     line_color=(255, 0, 255),
    #     line_width=3.0,
    #     rounding=0.1,
    # )
    ## ДЛЯ КИРИЛЛА [update_shape_color, update_shape_position, update_shape_transparency, set_shape_rounding]

    # pr.update_shape_color(3, 1, color=[255, 255, 0])
    # pr.update_shape_line(3, rect1.shape_id, color=(255, 0, 0), width=3.0)
    # pr.update_shape_transparency(3, 1, transparency=1)
    # pr.update_shape_transparency(3, rect1.shape_id, transparency=0.2)
    # pr.update_shape_rotation(2, 1, rotation=45.0)
    # pr.update_shape_rotation(2, 2, rotation=45.0)
    # pr.update_shape_rotation(2, 3, rotation=45.0)
    # pr.update_shape_rotation(2, 4, rotation=45.0)
    #
    # pr.set_shape_rounding(3, 2, 0.5)
    # pr.set_shape_rounding(3, 1, 1)
    # pr.set_shape_rounding(3, 3, 0.1)

    # logger.debug(f'figure = {json.dumps(pr.get_figure_frame_json(2), indent=4)}')

    # pr.add_slide_at_position(9, 0)
    # result = pr.copy_figure_shape(2, 1, 8)
    # logger.debug(f'rsult = {result}')

    logger.debug(pr.get_tasks_from_slide())

    logger.debug(f"figure = {json.dumps(pr.get_text_frame_json(8), indent=4)}")

    pr.save("test_create.pptx")
