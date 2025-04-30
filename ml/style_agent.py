from agno.agent import Agent

from ml.constants import STYLE_AGENT_CORE_INSTRUCTIONS, StyleOutput
from ml.lifespan import model_41
from ml.pptx_manager.main import PPTXManager


def get_style_agent(pr: PPTXManager) -> Agent:
    style_agent = Agent(
        name="Style Agent",
        instructions=STYLE_AGENT_CORE_INSTRUCTIONS
        + [
            f"Reference presentation slide size (pixels): {pr.get_slide_size_px()}",
            f"Reference presentation slide count: {pr.get_slide_count()}",
            f"Reference presentation schema: {pr.get_json_schema()}",
        ],
        model=model_41,
        response_model=StyleOutput,
        monitoring=False,
        telemetry=False,
        debug_mode=True,
    )

    return style_agent
