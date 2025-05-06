import base64


from agno.agent import Agent
from loguru import logger

from ml.constants import CREATOR_PPTX_AGENT_INSTRUCTIONS
from ml.lifespan import model_4o, open_sync_client
from ml.pptx_manager.main import PPTXManager
from ml.pptx_manager.models import ImageFrameOpts, CreateImageFrameOpts


def get_pptx_agent(pr: PPTXManager) -> Agent:
    slide_size = pr.get_slide_size_px()
    slide_count = pr.get_slide_count()

    def create_image(prompt: str, slide_id: int, opts: ImageFrameOpts) -> str:
        """
        Генерирует изображение с помощью DALL-E 2 на основе предоставленного запроса и размещает его на указанном слайде.

        Args:
            prompt (str): Текстовый запрос для генерации изображения.
            slide_id (int): Идентификатор слайда, на котором будет размещено сгенерированное изображение.
            opts (ImageFrameOpts): Параметры конфигурации для позиционирования и изменения размера сгенерированного изображения.
                - left (float): Расстояние от левого края слайда.
                - top (float): Расстояние от верхнего края слайда.
                - width (float): Ширина изображения.
                - height (float): Высота изображения.

        Returns:
            str: Сообщение о результате операции с подробным описанием созданной фигуры с изображением.
        """
        logger.debug(
            f"create_image вызвана с параметрами: prompt={prompt}, slide_id={slide_id}, opts={opts}"
        )

        # Append minimalistic style and no text requirements to the prompt
        enhanced_prompt = f"{prompt}. The image should be in a minimalistic style and contain no text."

        response = open_sync_client.images.generate(
            model="dall-e-3",
            prompt=enhanced_prompt,
            n=1,
            size="1024x1024",
            response_format="b64_json",
        )

        b64_data = response.data[0].b64_json

        image_bytes = base64.b64decode(b64_data)

        return pr.create_image_shape(
            slide_id,
            CreateImageFrameOpts(
                left=opts.left,
                top=opts.top,
                width=opts.width,
                height=opts.height,
                image=image_bytes,
            ),
        )

    pptx_agent = Agent(
        name="PPTX Agent",
        instructions=CREATOR_PPTX_AGENT_INSTRUCTIONS
        + [
            f'Размер слайда (px): {slide_size}',
            f'Кол-во слайдов: {slide_count}',
        ],
        tools=[
            pr.update_shape,
            pr.add_figure_shape,
            pr.copy_figure_shape,
            pr.delete_figure_shape,

            create_image,
            pr.copy_image_shape,
            pr.delete_image_shape,

            pr.add_slide_at_position,

            pr.update_text_frame_shape,
            pr.create_text_shape,
            pr.delete_text_shape,
            #    pr.delete_figure_shape,
        ],
        model=model_4o,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    return pptx_agent
