import uuid
from datetime import datetime

from pydantic import BaseModel


class ChatResponse(BaseModel):
    id: uuid.UUID
    title: str
    created_at: datetime
    updated_at: datetime


class MessageResponse(BaseModel):
    id: uuid.UUID
    message: str
    created_at: datetime
    path: str


class ChatResponseWithMessages(ChatResponse):
    messages: list[MessageResponse]


class ChatRequest(BaseModel):
    title: str
