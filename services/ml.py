import json

from loguru import logger

from ml.constants import StyleOutput
from ml.pptx_agent import get_pptx_agent
from ml.pptx_manager.main import PPTXManager
from ml.style_agent import get_style_agent


class MLService:
    def __init__(
        self,
    ):
        pass

    def edit_pres(self, prompt: str, pres: bytes) -> bytes:
        pr = PPTXManager(pres, False)

        style_agent = get_style_agent(pr)
        pptx_agent = get_pptx_agent(pr)

        style_message = style_agent.run(
            f"Проанализируй этот пользовательский запрос и эталонную презентацию. Сгенерируй подробные инструкции для HeadAgent: {prompt}",
        )

        style_output: StyleOutput = style_message.content

        logger.debug(
            f"StyleAgent output:\n{json.dumps(style_output.model_dump(), indent=4, ensure_ascii=False)}"
        )

        pptx_agent.run("\n".join(style_output.instructions))

        return pr.to_bytes()
