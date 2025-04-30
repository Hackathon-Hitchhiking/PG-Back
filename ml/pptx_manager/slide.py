import io

from loguru import logger
from pptx import Presentation
from pptx.dml.color import RGBColor
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

    def get_slide_json(self, slide_id: int) -> dict:
        return self.slide_metadata[slide_id].model_dump(
            exclude={
                "slide_manager",
            }
        )

    def load_presentation(self, file_path: str) -> None:
        self.pres = Presentation(file_path)
        logger.info(f"Презентация загружена из {file_path}")

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
            msg = "Презентация не загружена. Сначала используйте метод `load_presentation`."
            raise ValueError(msg)

        slides = self.pres.slides._sldIdLst
        total_slides = len(slides)

        if not (1 <= slide_id1 <= total_slides and 1 <= slide_id2 <= total_slides):
            msg = f"Недопустимые номера слайдов. Номера должны быть от 1 до {total_slides}"
            raise ValueError(msg)

        # Convert to 0-based indices
        idx1 = slide_id1 - 1
        idx2 = slide_id2 - 1

        # Perform the swap
        slides[idx1], slides[idx2] = slides[idx2], slides[idx1]

        logger.info(
            f"Слайды на позициях {slide_id1} и {slide_id2} были обменены местами"
        )
        return "Success"

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
            msg = "Презентация не загружена. Сначала вызовите load_presentation()"
            raise ValueError(msg)

        width_emu = self.pres.slide_width
        height_emu = self.pres.slide_height

        return (emu_to_px(width_emu), emu_to_px(height_emu))

    def set_slide_background_color(
        self, slide_id: int, color_rgb: tuple[int, int, int]
    ) -> str:
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
            msg = "Презентация не загружена. Сначала используйте метод `load_presentation`."
            raise ValueError(msg)

        if not all(0 <= c <= 255 for c in color_rgb):
            msg = "Значения цвета должны быть в диапазоне 0-255"
            raise ValueError(msg)

        slides = self.pres.slides
        if slide_id < 1 or slide_id > len(slides):
            msg = f"Недопустимый номер слайда: {slide_id}"
            raise ValueError(msg)

        slide = slides[slide_id - 1]
        background = slide.background
        fill = background.fill
        fill.solid()

        r, g, b = color_rgb
        fill.fore_color.rgb = RGBColor(r, g, b)

        logger.info(f"Цвет фона слайда {slide_id} изменен на RGB{color_rgb}")
        return "Success"

    def get_slide_count(self) -> int:
        if not self.pres:
            msg = "Презентация не загружена. Сначала используйте метод `load_presentation`."
            raise ValueError(msg)

        return len(self.pres.slides)

    def parse_slide(self, slide_id: int, slide: Slide) -> None:
        try:
            foreground_color = slide.background.fill.fore_color.rgb
        except TypeError:
            # if there is no color on the slide setting the base white color
            foreground_color = [255, 255, 255]
        self.slide_metadata.insert(
            slide_id,
            SlideFrame(
                slide_manager=slide, shape_count=1, background_color=foreground_color
            ),
        )

    def get_shape_count(self, slide_id: int) -> int:
        return self.slide_metadata[slide_id].shape_count

    def increase_shape_count(self, slide_id: int, count: int) -> int:
        self.slide_metadata[slide_id].shape_count += count

        return self.slide_metadata[slide_id].shape_count

    def _get_slide_manager(self, slide_id: int) -> Slide:
        return self.slide_metadata[slide_id].slide_manager

    def _add_shape_on_slide(
        self, slide_id: int, opts: CreateShapeOpts
    ) -> tuple[Shape, int]:
        width = utils.px_to_emu(opts.width)
        height = utils.px_to_emu(opts.height)

        pix_left = opts.left  # + opts.width
        pix_top = opts.top  # + opts.height

        left = utils.px_to_emu(pix_left)
        top = utils.px_to_emu(pix_top)

        logger.debug(
            f"Вызов _add_shape_on_slide с параметрами: номер слайда={slide_id}, параметры={opts}, pix left={pix_left}, pix top={pix_top}, pix width={opts.width}, pix height={opts.height}"
        )
        slide = self._get_slide_manager(slide_id)

        shape_creator = self._shape_type_creator[opts.type](slide)
        shape = shape_creator(left, top, width, height)

        # TODO Search this more
        # shape.text_frame.auto_size = MSO_AUTO_SIZE.MIXED
        shape.text_frame.word_wrap = True

        shape_id = self.increase_shape_count(slide_id, 1)

        return shape, shape_id

    def _add_image_on_slide(
        self, slide_id: int, image: bytes, opts: CreateShapeOpts
    ) -> tuple[Shape, int]:
        logger.debug(
            f"Вызов _add_image_on_slide с параметрами: номер слайда={slide_id}, параметры={opts}"
        )
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

        return "Success"

    def _delete_slide(self, slide_id: int) -> None:
        """
        Удаляет слайд из презентации и из метаданных.
        """
        if not self.pres:
            msg = "Презентация не загружена."
            raise ValueError(msg)

        slides = self.pres.slides
        if slide_id < 1 or slide_id > len(slides):
            msg = f"Недопустимый номер слайда: {slide_id}"
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

    def add_slide_at_position(
        self,
        position: int,
        layout_index: int = 0,
        background_color: list[int] | None = None,
    ) -> str:
        """
        Создает новый слайд в указанную позицию с заданным макетом и заголовком.

        Создает новый слайд и вставляет его в указанную позицию презентации,
        с возможностью установки макета и заголовка.

        -   Позиция указывается в формате 1-based индекса.
        -   Макет выбирается из доступных макетов презентации.
        -   Заголовок устанавливается только при наличии placeholder'а в макете.
        -   Обновляет внутренние метаданные слайдов.

        Args:
            position (int): Позиция для вставки слайда (1-based).
                -   Допустимый диапазон: [1, количество_слайдов + 1]
            layout_index (int, optional): Индекс используемого макета.
                -   По умолчанию: 0 (первый доступный макет)
                -   Допустимый диапазон: [0, количество_макетов - 1]
            background_color: list[int]: Цвет заднего фона слайда

        Returns:
            str: Сообщение о результате операции в формате:
                "Slide {position} was added"

        Raises:
            ValueError: При отсутствии загруженной презентации, неверной позиции или недопустимом индексе макета.
        """
        if background_color is None:
            background_color = [255, 255, 255]

        if not self.pres:
            msg = "Презентация не загружена. Сначала используйте метод `load_presentation`."
            raise ValueError(msg)

        max_position = len(self.pres.slides) + 1
        if not (1 <= position <= max_position):
            msg = f"Недопустимая позиция {position}. Должна быть между 1 и {max_position}."
            raise ValueError(msg)

        if not (0 <= layout_index < len(self.pres.slide_layouts)):
            msg = f"Недопустимый индекс макета {layout_index}. Должен быть между 0 и {len(self.pres.slide_layouts) - 1}."
            raise ValueError(msg)

        logger.debug(
            f"Вызов add_slide_at_position с параметрами: позиция={position}, индекс_макета={layout_index}"
        )

        slide_layout = self.pres.slide_layouts[layout_index]
        new_slide = self.pres.slides.add_slide(slide_layout)
        new_slide.background.fill.solid()
        new_slide.background.fill.fore_color.rgb = RGBColor(
            background_color[0], background_color[1], background_color[2]
        )

        slides = self.pres.slides._sldIdLst
        new_slide_id = slides[-1]
        del slides[-1]
        slides.insert(position - 1, new_slide_id)

        self.slide_metadata.insert(
            position,
            SlideFrame(
                slide_manager=new_slide,
                shape_count=0,
                background_color=background_color,
            ),
        )

        return f"Слайд {position} был добавлен"

    def test(self, source: str) -> None:
        self.load_presentation(source)
        self.add_slide_at_position(1, layout_index=0)
        self.add_slide_at_position(3, layout_index=0)
        self.set_slide_background_color(1, (255, 255, 0))
        output_path = "test_slide.pptx"
        self.pres.save(output_path)


if __name__ == "__main__":
    sm = SlideManager()
    sm.test("../test_data/test_dit.pptx")
