import io

from collections import defaultdict
from collections.abc import Iterator

from loguru import logger
from PIL import Image as PilImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.image import Image
from pptx.shapes.autoshape import Shape

from ml.pptx_manager import utils
from ml.pptx_manager.models import ImageFrameOpts, ImageFrameShape
from ml.pptx_manager.utils import get_slide_from_shape, px_to_emu


def image_to_byte_array(image: Image) -> bytes:
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format=image.format)
    return img_byte_arr.getvalue()


class ImageManager:
    def __init__(self) -> None:
        self._pres = None

        self.image_frame_shapes: defaultdict[int, list[ImageFrameShape]] = defaultdict(
            list[ImageFrameShape]
        )  # slide_id -> image_frame_shapes

    def get_image_json(self, slide_id: int) -> list[dict]:
        return [shape.model_dump(exclude={'blob', 'shape_manager'}) for shape in self.image_frame_shapes[slide_id]]

    def get_all_image_json(self):
        image_frame_json = {}

        for slide_id, shapes in self.image_frame_shapes.items():
            image_frame_json[slide_id] = [shape.model_dump(exclude={'blob', 'shape_manager'}) for shape in shapes]

        return image_frame_json

    def replace_image(self, slide_id: int, shape_id: int | None, new_picture: bytes) -> str:
        """
        Заменяет изображение в указанной фигуре на слайде.

        Удаляет старое изображение и добавляет новое изображение с сохранением размеров и позиции.

        -   Если `shape_id` равен None, замена применяется ко всем фигурам на указанном слайде.
        -   Новое изображение добавляется с теми же размерами и позицией, что и старое.
        -   Старое изображение удаляется из презентации.

        Args:
            slide_id (int): ID слайда, содержащего фигуру для замены изображения.
            shape_id (int | None): ID фигуры для замены изображения. Если None, заменяет изображения во всех фигурах на слайде.
            new_picture (bytes): Новое изображение в виде байтового массива.

        Returns:
            str: Сообщение о выполненных изменениях, включая ID слайда и фигуры.
        """
        changes = []
        logger.debug(f'функция вызвана с параметрами: {slide_id, shape_id}')
        for shape in self._get_image_frame(slide_id, shape_id):
            slide = get_slide_from_shape(shape.shape_manager)

            new_shape = slide.shapes.add_picture(
                io.BytesIO(new_picture), shape.left, shape.top, shape.width, shape.height
            )

            old_pic = shape.shape_manager._element
            new_pic = new_shape._element

            old_pic.addnext(new_pic)
            old_pic.getparent().remove(old_pic)

            shape.shape_manager = new_shape
            shape.blob = new_picture

            changes.append(f'Фигура {shape.shape_id}')

        if not changes:
            return f'На слайде {slide_id} не найдено фигур для замены изображения.'
        return f'На слайде {slide_id} заменены изображения в следующих фигурах: {", ".join(changes)}.'

    def update_image_frame_shape(self, slide_id: int, shape_id: int | None, opts: ImageFrameOpts | dict) -> str:
        """
        Обновляет свойства рамки изображения на определенном слайде.

        Изменяет различные атрибуты рамки изображения, включая ширину, высоту, позицию (left, top).

        -   Обновляются только атрибуты, указанные в `opts`.
        -   Если `opts` является словарем, он будет преобразован в `UpdateImageFrameOpts`.
        -   Каждое обновление атрибута обрабатывается отдельным внутренним методом.
        -   Все атрибуты в `UpdateImageFrameOpts` являются необязательными и по умолчанию равны None.

        Args:
            slide_id (int): ID слайда, содержащего рамку изображения для обновления.
            shape_id (int | None): ID фигуры для обновления. Если None, обновляет все фигуры на слайде.
            opts (UpdateImageFrameOpts | dict): Объект, содержащий параметры для обновления рамки изображения.
                -   width (int | None, optional): Новая ширина рамки изображения.
                -   height (int | None, optional): Новая высота рамки изображения.
                -   left (int | None, optional): Новая позиция рамки изображения по оси X.
                -   top (int | None, optional): Новая позиция рамки изображения по оси Y.

        Returns:
            str: Сообщение о выполненных изменениях, включая ID слайда, фигур, и измененных параметрах.
        """
        logger.debug(f'функция вызвана с параметрами: {slide_id, shape_id, opts}')

        if isinstance(opts, dict):
            opts = ImageFrameOpts(**opts)

        updates = {
            attr: getattr(opts, attr) for attr in ['width', 'height', 'left', 'top'] if getattr(opts, attr) is not None
        }

        if not updates:
            return 'WARNING: Не переданы параметры для обновления.'

        changed_shapes = []
        for shape in self._get_image_frame(slide_id, shape_id):
            if opts.width is not None:
                shape.shape_manager.width = px_to_emu(opts.width)

            if opts.height is not None:
                shape.shape_manager.height = px_to_emu(opts.height)

            if opts.left is not None:
                shape.shape_manager.left = px_to_emu(opts.left)

            if opts.top is not None:
                shape.shape_manager.top = px_to_emu(opts.top)

            changed_shapes.append(f'Фигура {shape.shape_id}')

        if not changed_shapes:
            return f'На слайде {slide_id} не найдено фигур для обновления.'

        params_str = ', '.join(f'{attr}={value}' for attr, value in updates.items())
        return f'На слайде {slide_id} обновлены параметры [{params_str}] для фигур: {", ".join(changed_shapes)}'

    def _get_image_frame(self, slide_id: int, shape_id: int | None) -> Iterator[ImageFrameShape]:
        frames = self.image_frame_shapes[slide_id]
        for frame in frames:
            if shape_id is not None and shape_id != frame.shape_id:
                continue
            yield frame

    def _get_image_frame_shape(self, slide_id: int, shape_id: int) -> ImageFrameShape | None:
        for frame in self._get_image_frame(slide_id, shape_id):
            return frame
        return None

    def _delete_image_frame_shape(self, slide_id: int, shape_id: int) -> None:
        frames = self.image_frame_shapes[slide_id]
        for index, frame in enumerate(frames):
            if frame.shape_id == shape_id:
                del frames[index]

    def _parse_image_shape(self, slide_id: int, shape_id: int, shape: Shape) -> ImageFrameShape:
        width = shape.width
        height = shape.height
        try:
            # if the shape is empty
            image = shape.image
            image_bytes = image.blob
        except AttributeError:
            image_bytes = None

        image_frame_shape = ImageFrameShape(
            shape_id=shape_id,
            width=utils.emu_to_px(width),
            height=utils.emu_to_px(height),
            left=utils.emu_to_px(shape.left),
            top=utils.emu_to_px(shape.top),
            blob=image_bytes,
            shape_manager=shape,
        )

        if slide_id < 0:
            slide_id = len(self.image_frame_shapes) - slide_id

        self.image_frame_shapes[slide_id].append(image_frame_shape)

        return image_frame_shape

    def test(self, source: str | None):
        self._pres = Presentation(source)

        slide_id = 1
        for slide in self._pres.slides:
            shape_id = 1
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    res = self._parse_image_shape(slide_id, shape_id, shape)
                    if res is not None:
                        shape_id += 1
            slide_id += 1

        new_img = PilImage.open('../test_data/Pr2.jpg')

        logger.debug(f'frames = {self.get_all_image_json()}')

        # self.replace_image(1, None, image_to_byte_array(new_img))

        # self.update_image_frame_shape(1, None, {'left': 300, 'top': 300})

        self._pres.save('test_new_img.pptx')


if __name__ == '__main__':
    im = ImageManager()

    im.test('../test_data/test_dit.pptx')
