import os

import httpx
from agno.models.openai import OpenAIChat
from openai import AsyncOpenAI, OpenAI

from configs.Environment import get_environment_variables

envs = get_environment_variables()

http_async_client = httpx.AsyncClient(proxy=envs.PROXY_HOST)
http_sync_client = httpx.Client(proxy=envs.PROXY_HOST)

os.environ["OPENAI_API_KEY"] = envs.OPENAI_API_KEY

open_async_client = AsyncOpenAI(
    http_client=http_async_client,
)

open_sync_client = OpenAI(
    http_client=http_sync_client,
)

model_41 = OpenAIChat(
    id="gpt-4.1",
    client=open_sync_client,
    async_client=open_async_client,
)
model_4o = OpenAIChat(
    id="gpt-4o", client=open_sync_client, async_client=open_async_client
)
