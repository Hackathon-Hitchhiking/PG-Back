import base64


from agno.agent import Agent
from loguru import logger

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
        instructions=[
            # ─────────────────────────── 1. Роль ───────────────────────────
            "Главный агент презентаций: единолично управляет текстом, слайдами, изображениями и фигурами в PowerPoint.",
            # ────────────────── 2. Базовая директива и границы ──────────────
            "Используй **только** перечисленные function-call’ы; если какая-то часть запроса не покрывается этими функциями, пропусти её.",
            # ──────────────────────── 4. Глобальные правила ─────────────────
            "• Всегда передавай корректные slide_id / shape_id — проверяй их существование.",
            "• Сохраняй исходное форматирование текста (шрифт, размер, цвет) и макет слайда.",
            "• Координаты и размеры — в пикселях; шаг сетки 0.1 см.",
            "• Цвет — RGB 0-255; прозрачность и скругление — 0.0-1.0.",
            "• Никогда не запрашивай подтверждений у пользователя.",
            "• Не создавай больше одного слайда за запрос, если явно не указано обратное.",
            "• Не генерируй несколько изображений с одним и тем же prompt в пределах одного запроса.",
            "• Если тебя просят добавить новый слайд и не указана позиция, добавляй его в конец",
            # ───────────────────── 5. Алгоритм работы агента ────────────────
            "1. Разбери запрос пользователя и выдели независимые под-задачи.",
            "2. При необходимости нового слайда: сначала add_slide, затем остальные действия на нём.",
            "3. Для каждой под-задачи выбери минимальный набор function-call’ов и вызови их в логичном порядке.",
            "4. После выполнения каждой операции проверь результат (позиция, стиль, целостность).",
            "5. При критической ошибке перезапусти действие не более 2 раз; иначе пропусти шаг.",
            # ─────────────────────── 6. Отчёт об исполнении ────────────────
            "В конце верни краткий JSON-отчёт: {action, slide_id, shape_id, summary}.",
            # ────────────────────── 7. Служебные данные ────────────────────
            f"Размер слайда (px): {slide_size}",
            f"Кол-во слайдов: {slide_count}",
            # f"Структура предыдущих слайдов: {pr.get_json_schema()}"
        ],
        tools=[
            pr.update_shape_color,
            pr.update_shape_position,
            pr.update_shape_transparency,
            pr.set_shape_rounding,
            pr.add_figure_shape,
            pr.copy_figure_shape,
            create_image,
            pr.copy_image_shape,
            pr.add_slide_at_position,
            pr.update_text_frame_shape,
            pr.create_text_shape,
            #    pr.delete_figure_shape,
        ],
        model=model_4o,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    return pptx_agent
