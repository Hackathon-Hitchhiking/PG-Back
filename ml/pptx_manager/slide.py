import io

from collections import defaultdict
from copy import deepcopy
from typing import Any

from loguru import logger
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide

from ml.pptx_manager import utils
from ml.pptx_manager.models import CreateShapeOpts, ShapeType, SlideFrame
from ml.pptx_manager.utils import CustomList, emu_to_px


class SlideManager:
    def __init__(self) -> None:
        self.pres = None

        self.slide_metadata: CustomList = CustomList()
        self._shape_type_creator = {
            ShapeType.TEXT: lambda slide: slide.shapes.add_textbox,
            ShapeType.IMAGE: lambda slide: slide.shapes.add_picture,
        }
        self.repeated_elements: list[dict] = []
        self.similarity_threshold = 0.85

    def get_slide_json(self, slide_id: int) -> dict:
        return self.slide_metadata[slide_id].model_dump(
            exclude={
                'slide_manager',
            }
        )

    def load_presentation(self, file_path: str) -> None:
        self.pres = Presentation(file_path)
        logger.info(f'Презентация загружена из {file_path}')

    def analyze_repeated_elements(self, start_slide: int = 2) -> None:
        """
        Analyzes presentation slides to identify repeated design patterns.

        Args:
            start_slide (int): First slide to analyze (1-based index)
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Call load_presentation() first.')

        total_slides = len(self.pres.slides)
        if total_slides < start_slide:
            logger.warning(f'Not enough slides to analyze. Need at least {start_slide}, found {total_slides}')
            return

        slide_fingerprints = []
        slide_width, slide_height = self.get_slide_size_px()

        # Collect fingerprints from each slide
        for slide_idx in range(start_slide - 1, total_slides):
            slide = self.pres.slides[slide_idx]
            fingerprints = []

            for shape in slide.shapes:
                fingerprint = self._create_shape_fingerprint(shape, slide_width, slide_height)
                fingerprints.append(fingerprint)

            slide_fingerprints.append(fingerprints)

        # Identify patterns across slides
        self._identify_cross_slide_patterns(slide_fingerprints, start_slide)

    def duplicate_slide(self, source_slide_idx: int, insert_position: int) -> None:
        """Дублирует слайд со всеми элементами, включая изображения."""
        source_slide = self.pres.slides[source_slide_idx]
        new_slide = self.pres.slides.add_slide(source_slide.slide_layout)

        # Копирование элементов
        for shape in source_slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # Сохраняем точные размеры и положение изображения
                with io.BytesIO(shape.image.blob) as img_stream:
                    new_pic = new_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
                self._copy_picture_properties(shape, new_pic)
            else:
                new_el = deepcopy(shape.element)
                new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

        # Корректировка порядка слайдов
        self._reorder_slides(new_slide, insert_position)
        return new_slide

    def _copy_picture_properties(self, src_pic, dest_pic):
        """Копирует дополнительные свойства изображения"""
        try:
            # Копируем все доступные свойства
            dest_pic.rotation = src_pic.rotation
            dest_pic.line.width = src_pic.line.width
            dest_pic.shadow.inherit = src_pic.shadow.inherit

            # Проверяем и копируем свойства обрезки, если они есть
            if hasattr(src_pic, 'crop_left'):
                dest_pic.crop_left = src_pic.crop_left
                dest_pic.crop_right = src_pic.crop_right
                dest_pic.crop_top = src_pic.crop_top
                dest_pic.crop_bottom = src_pic.crop_bottom
        except AttributeError as e:
            logger.warning(f'Не удалось скопировать свойства изображения: {e}')

    def _reorder_slides(self, new_slide, position):
        """Перемещает слайд в нужную позицию."""
        slides = self.pres.slides._sldIdLst
        new_sldId = slides[-1]  # ID нового слайда
        slides.insert(position, new_sldId)
        # del slides[-1]  # Удаляем дублирующийся ID

    def _create_shape_fingerprint(self, shape, slide_width: float, slide_height: float) -> dict[str, Any]:
        """Create a unique fingerprint for a shape."""
        fingerprint = {
            'type': type(shape).__name__,
            'relative_left': shape.left / slide_width,
            'relative_top': shape.top / slide_height,
            'relative_width': shape.width / slide_width,
            'relative_height': shape.height / slide_height,
        }

        if hasattr(shape, 'text'):
            fingerprint['text'] = shape.text
        if shape.name:
            fingerprint['name'] = shape.name
        if hasattr(shape.fill, 'fore_color') and shape.fill.fore_color.rgb:
            fingerprint['fill_color'] = shape.fill.fore_color.rgb

        return fingerprint

    def _identify_cross_slide_patterns(self, all_fingerprints: list[list[dict]], start_slide: int) -> None:
        """Group similar shapes across slides into repeated elements."""
        pattern_groups = defaultdict(list)

        # Compare shapes across slides
        for slide_idx, slide_prints in enumerate(all_fingerprints):
            for print_idx, fingerprint in enumerate(slide_prints):
                match_found = False
                for group in pattern_groups.values():
                    if self._fingerprint_similarity(group['template'], fingerprint) >= self.similarity_threshold:
                        group['occurrences'].append((start_slide + slide_idx, print_idx))
                        match_found = True
                        break
                if not match_found:
                    group_id = f'pattern_{len(pattern_groups) + 1}'
                    pattern_groups[group_id] = {
                        'template': fingerprint,
                        'occurrences': [(start_slide + slide_idx, print_idx)],
                    }

        # Filter patterns that appear on majority of slides
        total_analyzed_slides = len(all_fingerprints)
        self.repeated_elements = [
            {'pattern': group['template'], 'slides': group['occurrences']}
            for group in pattern_groups.values()
            if len({slide_idx for slide_idx, _ in group['occurrences']}) / total_analyzed_slides >= 0.5
        ]

    def _fingerprint_similarity(self, fp1: dict, fp2: dict) -> float:
        """Calculate similarity score between two fingerprints."""
        if fp1['type'] != fp2['type']:
            return 0.0

        position_score = (
            1.0
            - (abs(fp1['relative_left'] - fp2['relative_left']) + abs(fp1['relative_top'] - fp2['relative_top'])) / 2
        )

        size_score = (
            1.0
            - (
                abs(fp1['relative_width'] - fp2['relative_width'])
                + abs(fp1['relative_height'] - fp2['relative_height'])
            )
            / 2
        )

        return (position_score * 0.6) + (size_score * 0.4)

    def add_pattern_slide(self, position: int) -> str:
        """
        Creates a new slide containing all identified design patterns.

        Args:
            position (int): Position to insert new slide (1-based index)
        """
        if not self.repeated_elements:
            raise ValueError('No patterns identified. Run analyze_repeated_elements() first.')

        # Create new slide using first slide's layout
        new_slide = self.add_slide_at_position(position, 0)
        slide_width, slide_height = self.get_slide_size_px()

        # Add all identified patterns
        for pattern in self.repeated_elements:
            self._recreate_pattern(new_slide, pattern['template'], slide_width, slide_height)

        return f'Pattern slide created at position {position}'

    def _recreate_pattern(self, slide, pattern: dict, slide_width: float, slide_height: float) -> None:
        """Recreate a pattern on target slide."""
        left = int(pattern['relative_left'] * slide_width)
        top = int(pattern['relative_top'] * slide_height)
        width = int(pattern['relative_width'] * slide_width)
        height = int(pattern['relative_height'] * slide_height)

        create_opts = CreateShapeOpts(
            type=ShapeType.TEXT if 'text' in pattern else ShapeType.IMAGE,
            left=left,
            top=top,
            width=width,
            height=height,
        )

        shape, _ = self._add_shape_on_slide(slide.slide_id, create_opts)

        # Apply styling
        if 'fill_color' in pattern:
            shape.fill.solid()
            shape.fill.fore_color.rgb = pattern['fill_color']
        if 'text' in pattern:
            shape.text = pattern['text']

    def swap_slides(self, slide_id1: int, slide_id2: int) -> str:
        """
        Меняет местами позиции двух слайдов в презентации.

        Выполняет обмен позициями между двумя слайдами, идентифицируемыми их номерами.
        Нумерация слайдов начинается с 1.

        -   Требует загруженной презентации.
        -   Проверяет валидность индексов слайдов.
        -   Производит обмен на уровне внутреннего списка слайдов.

        Args:
            slide_id1 (int): Номер первого слайда для обмена (1-based индекс).
            slide_id2 (int): Номер второго слайда для обмена (1-based индекс).

        Returns:
            str: 'Success' при успешном выполнении операции.

        Raises:
            ValueError: Если презентация не загружена или указаны недопустимые индексы слайдов.
        """
        if not self.pres:
            msg = 'Презентация не загружена. Сначала используйте метод `load_presentation`.'
            raise ValueError(msg)

        slides = self.pres.slides._sldIdLst
        total_slides = len(slides)

        if not (1 <= slide_id1 <= total_slides and 1 <= slide_id2 <= total_slides):
            msg = f'Недопустимые номера слайдов. Номера должны быть от 1 до {total_slides}'
            raise ValueError(msg)

        # Convert to 0-based indices
        idx1 = slide_id1 - 1
        idx2 = slide_id2 - 1

        # Perform the swap
        slides[idx1], slides[idx2] = slides[idx2], slides[idx1]

        logger.info(f'Слайды на позициях {slide_id1} и {slide_id2} были обменены местами')
        return 'Success'

    def copy_slides(self, slide_id1: int, slide_id2: int) -> str:
        """
        Меняет местами позиции двух слайдов в презентации.

        Выполняет обмен позициями между двумя слайдами, идентифицируемыми их номерами.
        Нумерация слайдов начинается с 1.

        -   Требует загруженной презентации.
        -   Проверяет валидность индексов слайдов.
        -   Производит обмен на уровне внутреннего списка слайдов.

        Args:
            slide_id1 (int): Номер первого слайда для обмена (1-based индекс).
            slide_id2 (int): Номер второго слайда для обмена (1-based индекс).

        Returns:
            str: 'Success' при успешном выполнении операции.

        Raises:
            ValueError: Если презентация не загружена или указаны недопустимые индексы слайдов.
        """
        if not self.pres:
            msg = 'Презентация не загружена. Сначала используйте метод `load_presentation`.'
            raise ValueError(msg)

        slides = self.pres.slides._sldIdLst
        total_slides = len(slides)

        if not (1 <= slide_id1 <= total_slides and 1 <= slide_id2 <= total_slides):
            msg = f'Недопустимые номера слайдов. Номера должны быть от 1 до {total_slides}'
            raise ValueError(msg)

        # Convert to 0-based indices
        idx1 = slide_id1 - 1
        idx2 = slide_id2 - 1

        # Perform the swap
        slides[idx1] = slides[idx2]

        return 'Success'

    def get_slide_size_px(self) -> tuple[float, float]:
        """
        Возвращает размеры слайдов презентации в пикселях.

        Возвращает текущие размеры слайдов в формате (ширина, высота).
        Все слайды в презентации имеют одинаковые размеры.

        Returns:
            tuple[int, int]: Кортеж с шириной и высотой в пикселях

        Raises:
            ValueError: Если презентация не загружена
        """
        if not self.pres:
            msg = 'Презентация не загружена. Сначала вызовите load_presentation()'
            raise ValueError(msg)

        width_emu = self.pres.slide_width
        height_emu = self.pres.slide_height

        return (emu_to_px(width_emu), emu_to_px(height_emu))

    def set_slide_background_color(self, slide_id: int, color_rgb: tuple[int, int, int]) -> str:
        """
        Устанавливает сплошной цвет фона для указанного слайда.
        Изменяет заливку фона слайда на указанный RGB-цвет. Если на слайде была применена другая заливка
        (градиент, изображение и т.д.), она будет заменена на сплошной цвет.

        -   Нумерация слайдов начинается с 1.
        -   Для работы требуется загруженная презентация.
        -   Все компоненты цвета должны быть в диапазоне 0-255.
        -   Цвет применяется ко всему слайду, включая скрытые области макета.

        Args:
            slide_id (int): Номер целевого слайда (1-based индекс).
            -   Допустимый диапазон: [1, количество_слайдов]
        color_rgb (tuple[int, int, int]): Цвет в формате RGB.
            -   Каждый компонент (R, G, B) должен быть в диапазоне 0-255
            -   Пример: (255, 0, 0) - красный цвет

        Returns:
            str: Сообщение о результате в формате:
                "Цвет фона успешно обновлен"

        Raises:
            ValueError: В следующих случаях:
                - Презентация не загружена
                - Некорректный номер слайда
                - Недопустимые значения RGB (выход за пределы 0-255)

        Example:
            manager.set_slide_background_color(1, (34, 139, 34))
        """
        if not self.pres:
            msg = 'Презентация не загружена. Сначала используйте метод `load_presentation`.'
            raise ValueError(msg)

        if not all(0 <= c <= 255 for c in color_rgb):
            msg = 'Значения цвета должны быть в диапазоне 0-255'
            raise ValueError(msg)

        slides = self.pres.slides
        if slide_id < 1 or slide_id > len(slides):
            msg = f'Недопустимый номер слайда: {slide_id}'
            raise ValueError(msg)

        slide = slides[slide_id - 1]
        background = slide.background
        fill = background.fill
        fill.solid()

        r, g, b = color_rgb
        fill.fore_color.rgb = RGBColor(r, g, b)

        logger.info(f'Цвет фона слайда {slide_id} изменен на RGB{color_rgb}')
        return 'Success'

    def get_slide_count(self) -> int:
        if not self.pres:
            msg = 'Презентация не загружена. Сначала используйте метод `load_presentation`.'
            raise ValueError(msg)

        return len(self.pres.slides)

    def parse_slide(self, slide_id: int, slide: Slide) -> None:
        foreground_color = [255, 255, 255]

        try:
            fill = slide.background.fill
            if fill.type is not None:
                if fill.type == 'solid' and fill.fore_color and fill.fore_color.rgb:
                    r, g, b = fill.fore_color.rgb[:3]
                    foreground_color = [r, g, b]
            elif slide.follow_master_background:
                bg_fill = slide.slide_layout.background.fill
                if bg_fill.type == 'solid' and bg_fill.fore_color and bg_fill.fore_color.rgb:
                    r, g, b = bg_fill.fore_color.rgb[:3]
                    foreground_color = [r, g, b]
        except Exception as e:
            logger.warning(f'Не удалось получить цвет фона слайда {slide_id}: {e}')

        self.slide_metadata.insert(
            slide_id, SlideFrame(slide_manager=slide, shape_count=len(slide.shapes), background_color=foreground_color)
        )

    def get_shape_count(self, slide_id: int) -> int:
        return self.slide_metadata[slide_id].shape_count

    def increase_shape_count(self, slide_id: int, count: int) -> int:
        self.slide_metadata[slide_id].shape_count += count

        return self.slide_metadata[slide_id].shape_count

    def _get_slide_manager(self, slide_id: int) -> Slide:
        return self.slide_metadata[slide_id].slide_manager

    def _add_shape_on_slide(self, slide_id: int, opts: CreateShapeOpts) -> tuple[Shape, int]:
        width = utils.px_to_emu(opts.width)
        height = utils.px_to_emu(opts.height)

        pix_left = opts.left  # + opts.width
        pix_top = opts.top  # + opts.height

        left = utils.px_to_emu(pix_left)
        top = utils.px_to_emu(pix_top)

        logger.debug(
            f'Вызов _add_shape_on_slide с параметрами: номер слайда={slide_id}, параметры={opts}, pix left={pix_left}, pix top={pix_top}, pix width={opts.width}, pix height={opts.height}'
        )
        slide = self._get_slide_manager(slide_id)

        shape_creator = self._shape_type_creator[opts.type](slide)
        shape = shape_creator(left, top, width, height)

        # TODO Search this more
        # shape.text_frame.auto_size = MSO_AUTO_SIZE.MIXED
        shape.text_frame.word_wrap = True

        shape_id = self.increase_shape_count(slide_id, 1)

        return shape, shape_id

    def _add_image_on_slide(self, slide_id: int, image: bytes, opts: CreateShapeOpts) -> tuple[Shape, int]:
        logger.debug(f'Вызов _add_image_on_slide с параметрами: номер слайда={slide_id}, параметры={opts}')
        slide = self._get_slide_manager(slide_id)

        width = utils.px_to_emu(opts.width)
        height = utils.px_to_emu(opts.height)

        left = utils.px_to_emu(opts.left)
        top = utils.px_to_emu(opts.top)

        shape = slide.shapes.add_picture(io.BytesIO(image), left, top, width, height)

        shape_id = self.increase_shape_count(slide_id, 1)

        return shape, shape_id

    def _delete_shape_from_slide(self, slide_id: int, shape_id: int) -> str:
        slide = self._get_slide_manager(slide_id)

        shape = slide.shapes[shape_id]
        el = shape.element
        parent = el.getparent()
        parent.remove(el)

        return 'Success'

    def _delete_slide(self, slide_id: int) -> None:
        """
        Удаляет слайд из презентации и из метаданных.
        """
        if not self.pres:
            msg = 'Презентация не загружена.'
            raise ValueError(msg)

        slides = self.pres.slides
        if slide_id < 1 or slide_id > len(slides):
            msg = f'Недопустимый номер слайда: {slide_id}'
            raise ValueError(msg)

        # Удаляем слайд из презентации
        slide = slides[slide_id - 1]
        slide_element = slide._element
        slide_id_lst = slides._sldIdLst
        for idx, sldId in enumerate(slide_id_lst):
            if slides[idx]._element == slide_element:
                del slide_id_lst[idx]
                break

        # Удаляем метаданные
        if slide_id in self.slide_metadata:
            del self.slide_metadata[slide_id]

    # def add_slide_at_position(
    #     self, position: int, layout_index: int = 0, background_color: list[int] | None = None
    # ) -> str:
    #     """
    #     Создает новый слайд в указанную позицию с заданным макетом и заголовком.

    #     Создает новый слайд и вставляет его в указанную позицию презентации,
    #     с возможностью установки макета и заголовка.

    #     -   Позиция указывается в формате 1-based индекса.
    #     -   Макет выбирается из доступных макетов презентации.
    #     -   Заголовок устанавливается только при наличии placeholder'а в макете.
    #     -   Обновляет внутренние метаданные слайдов.

    #     Args:
    #         position (int): Позиция для вставки слайда (1-based).
    #             -   Допустимый диапазон: [1, количество_слайдов + 1]
    #         layout_index (int, optional): Индекс используемого макета.
    #             -   По умолчанию: 0 (первый доступный макет)
    #             -   Допустимый диапазон: [0, количество_макетов - 1]
    #         background_color: list[int]: Цвет заднего фона слайда

    #     Returns:
    #         str: Сообщение о результате операции в формате:
    #             "Slide {position} was added"

    #     Raises:
    #         ValueError: При отсутствии загруженной презентации, неверной позиции или недопустимом индексе макета.
    #     """
    #     if background_color is None:
    #         background_color = [255, 255, 255]

    #     if not self.pres:
    #         msg = 'Презентация не загружена. Сначала используйте метод `load_presentation`.'
    #         raise ValueError(msg)

    #     max_position = len(self.pres.slides) + 1
    #     if not (1 <= position <= max_position):
    #         msg = f'Недопустимая позиция {position}. Должна быть между 1 и {max_position}.'
    #         raise ValueError(msg)

    #     if not (0 <= layout_index < len(self.pres.slide_layouts)):
    #         msg = (
    #             f'Недопустимый индекс макета {layout_index}. Должен быть между 0 и {len(self.pres.slide_layouts) - 1}.'
    #         )
    #         raise ValueError(msg)

    #     logger.debug(f'Вызов add_slide_at_position с параметрами: позиция={position}, индекс_макета={layout_index}')

    #     slide_layout = self.pres.slide_layouts[layout_index]
    #     new_slide = self.pres.slides.add_slide(slide_layout)
    #     new_slide.background.fill.solid()
    #     new_slide.background.fill.fore_color.rgb = RGBColor(
    #         background_color[0], background_color[1], background_color[2]
    #     )

    #     slides = self.pres.slides._sldIdLst
    #     new_slide_id = slides[-1]
    #     del slides[-1]
    #     slides.insert(position - 1, new_slide_id)

    #     self.slide_metadata.insert(
    #         position,
    #         SlideFrame(slide_manager=new_slide, shape_count=0, background_color=background_color),
    #     )

    #     return f'Слайд {position} был добавлен'

    def test(self, source: str) -> None:
        self.load_presentation(source)
        # self.add_slide_at_position(1, layout_index=0)
        self.duplicate_slide(1, 3)
        # self.add_slide_at_position(3, layout_index=0)
        # self.set_slide_background_color(1, (255, 255, 0))
        # self.analyze_repeated_elements(start_slide=2)
        # self.add_pattern_slide(3)
        output_path = 'test_slide.pptx'
        self.pres.save(output_path)


if __name__ == '__main__':
    sm = SlideManager()
    sm.test('../test_data/test_dit.pptx')
