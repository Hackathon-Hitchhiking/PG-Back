from collections import defaultdict
from collections.abc import Iterator

from loguru import logger
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide
from pptx.util import Emu
from pydantic import BaseModel


class ShapeOpts(BaseModel):
    left: float | None = None
    top: float | None = None
    width: float | None = None
    height: float | None = None
    shape_type: int | None = None
    color: list[int] | None = None
    line_color: list[int] | None = None
    line_width: float | None = None
    transparency: float | None = None
    rotation: float | None = None
    adjustments: list[int] | None = None
    rounding: float | None = None


class GeometricShape(BaseModel):
    shape_id: int
    slide_id: int
    shape_type: int
    left: float
    top: float
    width: float
    height: float
    rounding: float
    color: list[int] = (255, 255, 255)
    line_color: list[int] = (0, 0, 0)
    line_width: float = 1.0
    transparency: float = 0.0
    rotation: float = 0.0
    adjustments: list[int] = []
    shape_manager: Shape | None = None

    class Config:
        arbitrary_types_allowed = True


EMUS_PER_PIXEL = 9525


def px_to_emu(px: float) -> int:
    return int(px * EMUS_PER_PIXEL)


def emu_to_px(emu: int) -> float:
    return emu / EMUS_PER_PIXEL


class FigureManager:
    def __init__(self):
        self.pres = None
        self.figure_shapes: dict[int, list[GeometricShape]] = defaultdict(list)
        self.shape_id_counter = defaultdict(int)

    def load_presentation(self, source_path: str) -> None:
        self.pres = Presentation(source_path)
        self._parse_existing_shapes()

    def create_presentation(self) -> None:
        self.pres = Presentation()

    def save_presentation(self, target_path: str) -> None:
        if self.pres:
            self.pres.save(target_path)
        else:
            raise ValueError('Презентация не загружена или не создана')

    def get_figure_frame_json(self, slide_id: int) -> list[dict]:
        return [shape.model_dump(exclude={'shape_manager'}) for shape in self.figure_shapes[slide_id]]

    def get_all_figure_frame_json(self) -> dict:
        figure_shape_json = {}
        for slide_id, shapes in self.figure_shapes.items():
            shapes_json = []
            for shape in shapes:
                shapes_json.append(shape.model_dump(exclude={'shape_manager'}))
            figure_shape_json[slide_id] = shapes_json

        return figure_shape_json

    def copy_figure_shape(self, slide_id_from: int, shape_id: int, slide_id_to: int) -> str:
        """
        Копирует фигуру с одного слайда на другой.

        Args:
            slide_id_from (int): ID исходного слайда.
            shape_id (int): ID фигуры на исходном слайде.
            slide_id_to (int): ID целевого слайда.

        Returns:
            str: Сообщение о результате операции.
        """
        logger.debug(
            f'Вызов copy_figure_shape с параметрами slide_id_from={slide_id_from}, shape_id={shape_id}, slide_id_to={slide_id_to}'
        )

        shape = self._get_shape(slide_id_from, shape_id)
        if shape is None:
            return f'Фигура (ID {shape_id}) не найдена на слайде {slide_id_from}.'

        result = self.add_figure_shape(
            slide_id=slide_id_to,
            shape_type=shape.shape_type,
            left=shape.left,
            top=shape.top,
            width=shape.width,
            height=shape.height,
            color=shape.color,
            line_color=shape.line_color,
            line_width=shape.line_width,
            rounding=shape.rounding,
            transparency=shape.transparency,
            rotation=shape.rotation,
            adjustments=shape.adjustments,
        )

        return f'Фигура (shape_id={shape_id}) успешно скопирована со слайда {slide_id_from} на слайд {slide_id_to}. {result}'

    def add_figure_shape(
        self,
        slide_id: int,
        shape_type: int,
        left: float,
        top: float,
        width: float,
        height: float,
        color: list[int] = None,
        line_color: list[int] = (0, 0, 0),
        line_width: float = 1.0,
        rounding: float = 0.0,
        transparency: float = 0.0,
        rotation: float = 0.0,
        adjustments: list[int] | None = None,
    ) -> str:
        """
        Добавляет новую фигуру на слайд и возвращает объект фигуры.

        Создает новую геометрическую фигуру указанного типа на заданном слайде
        с указанными параметрами положения и размера.

        Аргументы:
            slide_id (int): ID слайда, на который добавляется фигура.
            shape_type (int): Тип фигуры (константа из MSO_SHAPE, если необходимо закругление, то всегда использовать MSO_SHAPE.ROUNDED_RECTANGLE).
            left (float): Позиция левого края фигуры в пикселях.
            top (float): Позиция верхнего края фигуры в пикселях.
            width (float): Ширина фигуры в пикселях.
            height (float): Высота фигуры в пикселях.
            color (list[int]): Цвет заливки фигуры в формате RGB (по умолчанию белый: [255, 255, 255]).
            line_color (list[int]): Цвет контура фигуры в формате RGB (по умолчанию черный: (0, 0, 0)).
            line_width (float): Толщина контура фигуры в пикселях (по умолчанию 1.0).
            rounding (float): Значение закругления углов прямоугольника от 0.0 до 1.0 (по умолчанию 0.0).
            transparency (float): Прозрачность фигуры от 0.0 до 1.0 (по умолчанию 0.0).
            rotation (float): Угол поворота фигуры в градусах (по умолчанию 0.0).
            adjustments (Optional[list[int]]): Список параметров настройки для определенных типов фигур (по умолчанию None).

        Возвращает:
            Status: строку успеха или не успеха

        Вызывает:
            ValueError: Если презентация не загружена или не создана, или если указан неверный ID слайда.
        """

        if color is None:
            color = [255, 255, 255]

        logger.debug(
            f'add_figure_shape calls with parameters: {slide_id, shape_type, left, top, width, height, color, line_color, line_width, rounding, transparency, rotation, adjustments}'
        )

        if not self.pres:
            raise ValueError('Презентация не загружена или не создана')

        if slide_id <= 0 or slide_id > len(self.pres.slides):
            raise ValueError(f'Неверный ID слайда: {slide_id}')

        slide = self.pres.slides[slide_id - 1]

        left_emu = Emu(px_to_emu(left))
        top_emu = Emu(px_to_emu(top))
        width_emu = Emu(px_to_emu(width))
        height_emu = Emu(px_to_emu(height))

        # TODO delete this auto review
        # sending the figure to the background
        shape = slide.shapes.add_shape(shape_type, left_emu, top_emu, width_emu, height_emu)

        shape_tree = slide.shapes._spTree
        shape_element = shape_tree[slide.shapes.index(shape)]
        shape_tree.remove(shape_element)
        shape_tree.insert(0, shape_element)

        self.shape_id_counter[slide_id] += 1
        shape_id = self.shape_id_counter[slide_id]

        geometric_shape = GeometricShape(
            shape_id=shape_id,
            slide_id=slide_id,
            shape_type=shape_type,
            left=left,
            top=top,
            width=width,
            height=height,
            color=color,
            line_color=line_color,
            line_width=line_width,
            transparency=transparency,
            rotation=rotation,
            adjustments=adjustments or [],
            rounding=rounding,
            shape_manager=shape,
        )

        # Передаем все параметры явно!
        self._apply_shape_properties(
            geometric_shape,
            color=color,
            line_color=line_color,
            line_width=line_width,
            transparency=transparency,
            rotation=rotation,
            adjustments=adjustments or [],
            rounding=rounding,
        )

        self.figure_shapes[slide_id].append(geometric_shape)
        return 'Success'

    def update_shape(self, slide_id: int, shape_id: int, opts: ShapeOpts | dict) -> None:
        """
        Обновляет свойства существующей фигуры.

        Изменяет указанные свойства фигуры на заданном слайде. Свойства,
        которые не указаны в opts, остаются без изменений.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            opts (Union[ShapeOpts, Dict]): Объект или словарь с новыми свойствами фигуры.
                Может содержать следующие поля:
                - left (float): Новая позиция левого края в пикселях.
                - top (float): Новая позиция верхнего края в пикселях.
                - width (float): Новая ширина в пикселях.
                - height (float): Новая высота в пикселях.
                - color (List[int]): Новый цвет заливки (RGB).
                - line_color (List[int]): Новый цвет контура (RGB).
                - line_width (float): Новая ширина линии контура.
                - transparency (float): Новое значение прозрачности (0.0-1.0).
                - rotation (float): Новый угол поворота в градусах.
                - adjustments (List[int]): Новые значения настроек формы фигуры.
                - rounding (float): Новое значение скругления углов (0.0-1.0).

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        if isinstance(opts, dict):
            opts = ShapeOpts(**opts)

        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        if opts.left is not None or opts.top is not None or opts.width is not None or opts.height is not None:
            self._update_shape_position(shape, opts)

        if opts.color is not None:
            self._update_shape_fill_color(shape, opts.color)

        if opts.line_color is not None:
            self._update_shape_line_color(shape, opts.line_color)

        if opts.line_width is not None:
            self._update_shape_line_width(shape, opts.line_width)

        if opts.transparency is not None:
            self._update_shape_transparency(shape, opts.transparency)

        if opts.rotation is not None:
            self._update_shape_rotation(shape, opts.rotation)

        if opts.adjustments is not None:
            self._update_shape_adjustments(shape, opts.adjustments)

        if opts.rounding is not None:
            self.set_shape_rounding(slide_id, shape_id, opts.rounding)

    def update_shape_position(
        self,
        slide_id: int,
        shape_id: int,
        left: float = None,
        top: float = None,
        width: float = None,
        height: float = None,
    ) -> None:
        """
        Обновляет положение и размер фигуры.

        Изменяет положение и/или размер фигуры на указанные значения.
        Параметры, которые не указаны (None), остаются без изменений.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            left (float, optional): Новая позиция левого края в пикселях.
            top (float, optional): Новая позиция верхнего края в пикселях.
            width (float, optional): Новая ширина в пикселях.
            height (float, optional): Новая высота в пикселях.

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        logger.debug(f'update_shape_position calls with parameters: {slide_id, shape_id, left, top, width, height}')

        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        ppt_shape = shape.shape_manager

        if left is not None:
            ppt_shape.left = Emu(px_to_emu(left))
            shape.left = left

        if top is not None:
            ppt_shape.top = Emu(px_to_emu(top))
            shape.top = top

        if width is not None:
            ppt_shape.width = Emu(px_to_emu(width))
            shape.width = width

        if height is not None:
            ppt_shape.height = Emu(px_to_emu(height))
            shape.height = height

    def update_shape_color(self, slide_id: int, shape_id: int, color: list[int]) -> None:
        """
        Обновляет цвет заливки фигуры.

        Изменяет цвет заливки фигуры на указанный цвет RGB.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            color (list[int]): Новый цвет заливки (RGB).

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        logger.debug(f'update_shape_color calls with parameters: {slide_id, shape_id, color}')

        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        self._update_shape_fill_color(shape, color)

    def update_shape_line(self, slide_id: int, shape_id: int, color: list[int] = None, width: float = None) -> None:
        """
        Обновляет свойства линии контура фигуры.

        Изменяет цвет и/или ширину линии контура фигуры.
        Параметры, которые не указаны (None), остаются без изменений.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            color (list[int], optional): Новый цвет линии (RGB).
            width (float, optional): Новая ширина линии.

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        if color is not None:
            self._update_shape_line_color(shape, color)

        if width is not None:
            self._update_shape_line_width(shape, width)

    def update_shape_transparency(self, slide_id: int, shape_id: int, transparency: float) -> None:
        """
        Обновляет прозрачность фигуры.

        Изменяет прозрачность заливки фигуры на указанное значение.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            transparency (float): Новое значение прозрачности (0.0-1.0).

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде или значение прозрачности вне диапазона.
        """
        logger.debug(f'update_shape_transparency calls with parameters: {slide_id, shape_id, transparency}')

        if not 0.0 <= transparency <= 1.0:
            raise ValueError('Значение прозрачности должно быть в диапазоне от 0.0 до 1.0')

        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        self._update_shape_transparency(shape, transparency)

    def update_shape_rotation(self, slide_id: int, shape_id: int, rotation: float) -> None:
        """
        Обновляет поворот фигуры.

        Изменяет угол поворота фигуры на указанное значение.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            rotation (float): Новый угол поворота в градусах.

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        self._update_shape_rotation(shape, rotation)

    def set_shape_rounding(self, slide_id: int, shape_id: int, rounding_value: float) -> None:
        """
        Sets the rounding value for a shape.

        Sets the corner rounding radius for shapes that support this feature.
        Works with ROUNDED_RECTANGLE shapes and other shapes that have
        adjustment values. The rounding value is typically in the range
        from 0 to 1, where 0 means no rounding and 1 means maximum rounding.

        If the shape doesn't support rounding, the operation will be skipped.

        Args:
            slide_id (int): ID of the slide containing the shape.
            shape_id (int): ID of the shape to modify.
            rounding_value (float): Rounding value (from 0.0 to 1.0).

        Raises:
            ValueError: If the shape with the specified ID is not found on the slide or
                    the rounding value is outside the valid range.
        """
        logger.debug(f'set_shape_rounding calls with parameters: {slide_id, shape_id, rounding_value}')

        if not 0.0 <= rounding_value <= 1.0:
            raise ValueError('Rounding value must be between 0.0 and 1.0')

        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Shape with ID {shape_id} not found on slide {slide_id}')

        ppt_shape = shape.shape_manager

        # Calculate the adjustment value for PowerPoint (scale from 0-1 to 0-100000)
        adj_value = int(rounding_value * 100000)

        try:
            # Check if the shape has adjustment values and can be rounded
            if hasattr(ppt_shape, 'adjustment_values') and len(ppt_shape.adjustment_values) > 0:
                ppt_shape.adjustment_values[0] = adj_value

                # Update our internal representation
                if not shape.adjustments:
                    shape.adjustments = [adj_value]
                else:
                    shape.adjustments[0] = adj_value

                logger.info(f'Successfully set rounding {rounding_value} for shape with ID {shape_id}')
                return

            # Alternative way to access adjustments
            if hasattr(ppt_shape, 'adjustments') and ppt_shape.adjustments:
                ppt_shape.adjustments[0] = adj_value

                # Update our internal representation
                if not shape.adjustments:
                    shape.adjustments = [adj_value]
                else:
                    shape.adjustments[0] = adj_value

                logger.info(f'Successfully set rounding {rounding_value} for shape with ID {shape_id}')
                return

            # If we get here, the shape doesn't support rounding
            shape_type_name = 'unknown'
            if hasattr(ppt_shape, 'auto_shape_type'):
                shape_type_name = f'{ppt_shape.auto_shape_type} ({ppt_shape.auto_shape_type.value})'

            logger.info(f'Shape type {shape_type_name} does not support rounding. Operation skipped.')

        except IndexError:
            logger.info(
                'Shape does not have adjustment settings for rounding (no element with index 0). Operation skipped.'
            )
        except Exception as e:
            logger.warning(f'Failed to set rounding: {str(e)}')

    def delete_shape(self, slide_id: int, shape_id: int) -> None:
        """
        Удаляет фигуру со слайда.

        Удаляет указанную фигуру с заданного слайда как из презентации PowerPoint,
        так и из внутренней коллекции фигур менеджера.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для удаления.

        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f'Фигура с ID {shape_id} не найдена на слайде {slide_id}')

        sp = shape.shape_manager._element
        sp.getparent().remove(sp)

        shapes_list = self.figure_shapes[slide_id]
        for i, s in enumerate(shapes_list):
            if s.shape_id == shape_id:
                del shapes_list[i]
                break

    def get_shape_json(self, slide_id: int, shape_id: int = None) -> dict | list[dict]:
        """
        Получает JSON-представление фигуры или всех фигур на слайде.

        Возвращает словарь с данными указанной фигуры или список словарей
        со всеми фигурами на слайде, если shape_id не указан.

        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int, optional): ID фигуры. По умолчанию None.

        Возвращает:
            Union[Dict, List[Dict]]: Словарь с данными фигуры или список словарей с данными всех фигур.
        """
        if shape_id is not None:
            shape = self._get_shape_by_id(slide_id, shape_id)
            if not shape:
                return {}
            return shape.dict(exclude={'shape_manager'})
        return [s.dict(exclude={'shape_manager'}) for s in self.figure_shapes[slide_id]]

    def get_all_shapes_json(self) -> dict[int, list[dict]]:
        """
        Получает JSON-представление всех фигур в презентации.

        Возвращает словарь, где ключи - ID слайдов, а значения - списки словарей
        с данными всех фигур на соответствующих слайдах.

        Возвращает:
            Dict[int, List[Dict]]: Словарь с данными всех фигур в презентации.
        """
        result = {}
        for slide_id, shapes_list in self.figure_shapes.items():
            result[slide_id] = [s.dict(exclude={'shape_manager'}) for s in shapes_list]
        return result

    def get_shape_ids(self, slide_id: int) -> list[int]:
        """
        Получает список всех ID фигур на указанном слайде.

        Аргументы:
            slide_id (int): ID слайда.

        Возвращает:
            List[int]: Список ID фигур на слайде.
        """
        return [shape.shape_id for shape in self.figure_shapes[slide_id]]

    def find_shapes(self, slide_id: int, **criteria) -> list[GeometricShape]:
        """
        Находит фигуры на слайде, соответствующие указанным критериям.

        Аргументы:
            slide_id (int): ID слайда.
            **criteria: Критерии поиска в виде пар ключ-значение.
                Например: text="Заголовок", color=(255, 0, 0)

        Возвращает:
            List[GeometricShape]: Список найденных фигур.
        """
        result = []

        for shape in self.figure_shapes[slide_id]:
            match = True
            for key, value in criteria.items():
                if not hasattr(shape, key) or getattr(shape, key) != value:
                    match = False
                    break

            if match:
                result.append(shape)

        return result

    def _capture_shape_properties(self, ppt_shape) -> dict:
        """
        Сохраняет все свойства фигуры в словарь.

        Аргументы:
            ppt_shape: Объект фигуры python-pptx.

        Возвращает:
            dict: Словарь со всеми свойствами фигуры.
        """
        properties = {
            'left': ppt_shape.left,
            'top': ppt_shape.top,
            'width': ppt_shape.width,
            'height': ppt_shape.height,
            'rotation': ppt_shape.rotation if hasattr(ppt_shape, 'rotation') else 0,
        }

        # Сохраняем свойства заливки
        if hasattr(ppt_shape, 'fill'):
            fill_props = {}
            if hasattr(ppt_shape.fill, 'fore_color') and ppt_shape.fill.fore_color:
                fill_props['fore_color'] = ppt_shape.fill.fore_color.rgb
            if hasattr(ppt_shape.fill, 'back_color') and ppt_shape.fill.back_color:
                fill_props['back_color'] = ppt_shape.fill.back_color.rgb
            if hasattr(ppt_shape.fill, 'transparency'):
                fill_props['transparency'] = ppt_shape.fill.transparency
            if hasattr(ppt_shape.fill, 'type'):
                fill_props['type'] = ppt_shape.fill.type
            properties['fill'] = fill_props

        if hasattr(ppt_shape, 'line'):
            line_props = {}
            if hasattr(ppt_shape.line, 'color') and ppt_shape.line.color:
                line_props['color'] = ppt_shape.line.color.rgb
            if hasattr(ppt_shape.line, 'width'):
                line_props['width'] = ppt_shape.line.width
            if hasattr(ppt_shape.line, 'dash_style'):
                line_props['dash_style'] = ppt_shape.line.dash_style
            properties['line'] = line_props

        if hasattr(ppt_shape, 'text') and ppt_shape.text:
            properties['text'] = ppt_shape.text
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                text_props = []
                for p in ppt_shape.text_frame.paragraphs:
                    p_props = {
                        'text': p.text,
                        'alignment': p.alignment if hasattr(p, 'alignment') else None,
                        'runs': [],
                    }
                    for run in p.runs:
                        run_props = {'text': run.text}
                        if hasattr(run, 'font'):
                            font_props = {}
                            if hasattr(run.font, 'color') and run.font.color:
                                font_props['color'] = run.font.color.rgb
                            if hasattr(run.font, 'size'):
                                font_props['size'] = run.font.size
                            if hasattr(run.font, 'bold'):
                                font_props['bold'] = run.font.bold
                            if hasattr(run.font, 'italic'):
                                font_props['italic'] = run.font.italic
                            if hasattr(run.font, 'underline'):
                                font_props['underline'] = run.font.underline
                            run_props['font'] = font_props
                        p_props['runs'].append(run_props)
                    text_props.append(p_props)
                properties['text_frame'] = text_props

        return properties

    def _apply_shape_properties_from_dict(self, ppt_shape, properties: dict) -> None:
        if 'rotation' in properties and hasattr(ppt_shape, 'rotation'):
            ppt_shape.rotation = properties['rotation']

        if 'fill' in properties and hasattr(ppt_shape, 'fill'):
            fill_props = properties['fill']
            if 'type' in fill_props and fill_props['type'] != 0:
                ppt_shape.fill.solid()
                if 'fore_color' in fill_props and hasattr(ppt_shape.fill, 'fore_color'):
                    ppt_shape.fill.fore_color.rgb = fill_props['fore_color']
                if 'transparency' in fill_props and hasattr(ppt_shape.fill, 'transparency'):
                    ppt_shape.fill.transparency = fill_props['transparency']

        if 'line' in properties and hasattr(ppt_shape, 'line'):
            line_props = properties['line']
            if 'color' in line_props and hasattr(ppt_shape.line, 'color'):
                ppt_shape.line.color.rgb = line_props['color']
            if 'width' in line_props and hasattr(ppt_shape.line, 'width'):
                ppt_shape.line.width = line_props['width']
            if 'dash_style' in line_props and hasattr(ppt_shape.line, 'dash_style'):
                ppt_shape.line.dash_style = line_props['dash_style']

    def _get_slide_by_id(self, slide_id: int) -> Slide | None:
        """
        Получает объект слайда по его ID.

        Аргументы:
            slide_id (int): ID слайда.

        Возвращает:
            Optional[Slide]: Объект слайда или None, если слайд не найден.
        """
        if not self.pres:
            return None

        if slide_id <= 0 or slide_id > len(self.pres.slides):
            return None

        return self.pres.slides[slide_id - 1]

    def _parse_existing_shapes(self) -> None:
        """
        Анализирует существующие фигуры из загруженной презентации.

        Проходит по всем слайдам и фигурам в загруженной презентации,
        создает для них объекты GeometricShape и добавляет их в коллекцию фигур менеджера.
        """
        if not self.pres:
            return

        for slide_idx, slide in enumerate(self.pres.slides, 1):
            shape_idx = 1
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    self._parse_figure_shape(slide_idx, shape_idx, shape)
                    shape_idx += 1
                    self.shape_id_counter[slide_idx] = max(self.shape_id_counter[slide_idx], shape_idx)

    def _parse_figure_shape(self, slide_id: int, shape_id: int, shape: Shape) -> GeometricShape | None:
        """
        Анализирует фигуру из презентации и добавляет ее в коллекцию.

        Извлекает все свойства фигуры, создает объект GeometricShape
        и добавляет его в коллекцию фигур менеджера.

        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры.
            shape (Shape): Объект фигуры из python-pptx.

        Возвращает:
            Optional[GeometricShape]: Созданный объект фигуры или None в случае ошибки.
        """
        try:
            left = emu_to_px(shape.left)
            top = emu_to_px(shape.top)
            width = emu_to_px(shape.width)
            height = emu_to_px(shape.height)

            fill_color = [0, 0, 0]
            try:
                if hasattr(shape.fill, 'fore_color') and shape.fill.type != 0:
                    fill_color = tuple(shape.fill.fore_color.rgb)
            except Exception:
                pass

            line_color = (0, 0, 0)
            try:
                if hasattr(shape.line, 'color') and shape.line.color is not None:
                    line_color = tuple(shape.line.color.rgb)  ## CHANGE
            except Exception:
                pass

            line_width = 1.0
            try:
                if hasattr(shape.line, 'width') and shape.line.width is not None:
                    line_width = emu_to_px(shape.line.width)
            except Exception as e:
                logger.debug(f'Не удалось получить ширину линии: {e}')

            adjustments = []
            try:
                if hasattr(shape, 'adjustment_values') and shape.adjustment_values:
                    adjustments = list(shape.adjustment_values)
            except Exception as e:
                logger.debug(f'Не удалось получить настройки формы: {e}')

            # Extract rounding value from adjustments if available
            rounding = 0.0
            try:
                if adjustments and len(adjustments) > 0:
                    # Convert from PowerPoint's scale (0-100000) to normalized scale (0.0-1.0)
                    rounding = adjustments[0] / 100000.0
            except Exception as e:
                logger.debug(f'Не удалось получить значение скругления: {e}')

            geometric_shape = GeometricShape(
                shape_id=shape_id,
                slide_id=slide_id,
                shape_type=shape.auto_shape_type,
                left=left,
                top=top,
                width=width,
                height=height,
                color=fill_color,
                line_color=line_color,
                line_width=line_width,
                adjustments=adjustments,
                rounding=rounding,
                shape_manager=shape,
            )

            self.figure_shapes[slide_id].append(geometric_shape)
            return geometric_shape

        except Exception as e:
            logger.warning(f'Ошибка при анализе фигуры: {e}')
            return None

    def _get_shape_by_id(self, slide_id: int, shape_id: int) -> GeometricShape | None:
        """
        Получает фигуру по её ID с определенного слайда.

        Ищет фигуру с указанным ID на заданном слайде в коллекции фигур.
        В отличие от метода _get_shape, этот метод использует правильный
        подход к поиску фигуры по её shape_id, а не по индексу в коллекции.

        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int): ID фигуры.

        Возвращает:
            Optional[GeometricShape]: Найденная фигура или None, если фигура не найдена.
        """
        for shape in self.figure_shapes[slide_id]:
            if shape.shape_id == shape_id:
                return shape
        return None

    def _delete_shape(self, slide_id: int, shape_id: int) -> None:
        """
        Удаляет фигуру из внутренней коллекции по slide_id и shape_id.
        """
        frames = self.figure_shapes[slide_id]
        for index, frame in enumerate(frames):
            if frame.shape_id == shape_id:
                del frames[index]
                break

    def _get_shape(self, slide_id: int, shape_id: int) -> GeometricShape | None:
        """
        Получает фигуру по ее ID с определенного слайда.

        Ищет фигуру с указанным ID на заданном слайде в коллекции фигур менеджера.

        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int): ID фигуры.

        Возвращает:
            Optional[GeometricShape]: Найденная фигура или None, если фигура не найдена.
        """
        return self._get_shape_by_id(slide_id, shape_id)

    def _get_shapes(self, slide_id: int, shape_id: int | None = None) -> Iterator[GeometricShape]:
        """
        Получает фигуры с определенного слайда, опционально фильтруя по ID фигуры.

        Возвращает итератор по фигурам на заданном слайде, опционально фильтруя
        только фигуру с указанным ID.

        Аргументы:
            slide_id (int): ID слайда.
            shape_id (Optional[int], optional): ID фигуры для фильтрации. По умолчанию None.

        Возвращает:
            Iterator[GeometricShape]: Итератор по фигурам.
        """
        if slide_id < 0:
            slide_id = sorted(list(self.figure_shapes.keys()))[slide_id]

        for shape in self.figure_shapes[slide_id]:
            if shape_id is not None and shape.shape_id != shape_id:
                continue
            yield shape

    def _apply_shape_properties(self, shape: GeometricShape, **kwargs) -> None:
        """
        Применяет свойства к новой созданной фигуре.
        """
        ppt_shape = shape.shape_manager

        if 'color' in kwargs and kwargs['color'] is not None:
            color = kwargs['color']
            try:
                ppt_shape.fill.solid()
                ppt_shape.fill.fore_color.rgb = RGBColor(*color)
                shape.color = color
            except Exception as e:
                logger.warning(f'Не удалось установить цвет заливки: {e}')

        if 'line_color' in kwargs and kwargs['line_color'] is not None:
            line_color = kwargs['line_color']
            try:
                ppt_shape.line.color.rgb = RGBColor(*line_color)
                shape.line_color = line_color
            except Exception as e:
                logger.warning(f'Не удалось установить цвет линии: {e}')

        if 'line_width' in kwargs and kwargs['line_width'] is not None:
            line_width = kwargs['line_width']
            try:
                ppt_shape.line.width = Emu(px_to_emu(line_width))
                shape.line_width = line_width
            except Exception as e:
                logger.warning(f'Не удалось установить ширину линии: {e}')

        if 'transparency' in kwargs and kwargs['transparency'] is not None:
            transparency = kwargs['transparency']
            try:
                if hasattr(ppt_shape.fill, 'transparency'):
                    ppt_shape.fill.transparency = transparency
                shape.transparency = transparency
            except Exception as e:
                logger.warning(f'Не удалось установить прозрачность: {e}')

        if 'rotation' in kwargs and kwargs['rotation'] is not None:
            rotation = kwargs['rotation']
            try:
                if hasattr(ppt_shape, 'rotation'):
                    ppt_shape.rotation = rotation
                shape.rotation = rotation
            except Exception as e:
                logger.warning(f'Не удалось установить поворот: {e}')

        if 'adjustments' in kwargs and kwargs['adjustments']:
            adjustments = kwargs['adjustments']
            try:
                if hasattr(ppt_shape, 'adjustment_values'):
                    for i, value in enumerate(adjustments):
                        if i < len(ppt_shape.adjustment_values):
                            ppt_shape.adjustment_values[i] = value
                shape.adjustments = adjustments
            except Exception as e:
                logger.warning(f'Не удалось установить настройки формы: {e}')

        if 'rounding' in kwargs and kwargs['rounding'] is not None:
            rounding = kwargs['rounding']
            try:
                if hasattr(ppt_shape, 'adjustment_values') and len(ppt_shape.adjustment_values) > 0:
                    adj_value = int(rounding * 100000)
                    ppt_shape.adjustment_values[0] = adj_value
                    if not shape.adjustments:
                        shape.adjustments = [adj_value]
                    else:
                        shape.adjustments[0] = adj_value
                    shape.rounding = rounding
            except Exception as e:
                logger.warning(f'Не удалось установить скругление: {e}')

    def _update_shape_position(self, shape: GeometricShape, opts: ShapeOpts) -> None:
        ppt_shape = shape.shape_manager

        if opts.left is not None:
            ppt_shape.left = Emu(px_to_emu(opts.left))
            shape.left = opts.left

        if opts.top is not None:
            ppt_shape.top = Emu(px_to_emu(opts.top))
            shape.top = opts.top

        if opts.width is not None:
            ppt_shape.width = Emu(px_to_emu(opts.width))
            shape.width = opts.width

        if opts.height is not None:
            ppt_shape.height = Emu(px_to_emu(opts.height))
            shape.height = opts.height

    def _update_shape_fill_color(self, shape: GeometricShape, color: list[int]) -> None:
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.fill.solid()
            ppt_shape.fill.fore_color.rgb = RGBColor(color[0], color[1], color[2])
            shape.color = color
        except Exception as e:
            logger.warning(f'Не удалось обновить цвет заливки: {e}')

    def _update_shape_line_color(self, shape: GeometricShape, color: list[int]) -> None:
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.line.color.rgb = RGBColor(color[0], color[1], color[2])
            shape.line_color = color
        except Exception as e:
            logger.warning(f'Не удалось обновить цвет линии: {e}')

    def _update_shape_line_width(self, shape: GeometricShape, width: float) -> None:
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.line.width = Emu(px_to_emu(width))
            shape.line_width = width
        except Exception as e:
            logger.warning(f'Не удалось обновить ширину линии: {e}')

    def _update_shape_transparency(self, shape: GeometricShape, transparency: float) -> None:
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape.fill, 'transparency'):
                ppt_shape.fill.transparency = transparency
            shape.transparency = transparency
        except Exception as e:
            logger.warning(f'Не удалось обновить прозрачность: {e}')

    def _update_shape_rotation(self, shape: GeometricShape, rotation: float) -> None:
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'rotation'):
                ppt_shape.rotation = rotation
            shape.rotation = rotation
        except Exception as e:
            logger.warning(f'Не удалось обновить поворот: {e}')

    def _update_shape_adjustments(self, shape: GeometricShape, adjustments: list[int]) -> None:
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'adjustment_values'):
                for i, value in enumerate(adjustments):
                    if i < len(ppt_shape.adjustment_values):
                        ppt_shape.adjustment_values[i] = value

            shape.adjustments = adjustments
        except Exception as e:
            logger.warning(f'Не удалось обновить настройки формы: {e}')

    def test(self, source_path: str, target_path: str = 'test_output.pptx'):
        self.load_presentation(source_path)

        rect1 = self.add_figure_shape(
            2,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            300,
            500,
            100,
            500,
            color=(255, 128, 255),
            line_color=(255, 0, 255),
            line_width=3.0,
            rounding=0.1,
        )
        ## ДЛЯ КИРИЛЛА [update_shape_color, update_shape_position, update_shape_transparency, set_shape_rounding]

        # self.update_shape_position(3, 1, rounding = 0.5, color=(255, 0, 255))
        # self.update_shape_color(3, 1, color=(255, 255, 0))
        # self.update_shape_line(new_slide_id, rect3.shape_id, color=(255, 0, 0), width=3.0)
        # self.update_shape_transparency(3, 1, transparency=1)
        # self.update_shape_transparency(3, rect1.shape_id, transparency=0.2)
        # self.update_shape_rotation(2, 1, rotation=45.0)
        # self.update_shape_rotation(2, 2, rotation=45.0)
        # self.update_shape_rotation(2, 3, rotation=45.0)
        # self.update_shape_rotation(2, 4, rotation=45.0)

        # self.set_shape_rounding(3, 2, 0.5)
        # self.set_shape_rounding(3, 1, 1)
        # self.set_shape_rounding(3, 3, 0.1)

        self.save_presentation(target_path)


if __name__ == '__main__':
    manager = FigureManager()
    manager.test('test_data/test_dit.pptx')
