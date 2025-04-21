import uuid

from fastapi import APIRouter, UploadFile, File, Depends, Form, Response
from loguru import logger

from schemas.chat import ChatResponse, ChatResponseWithMessages
from services.chat import ChatService

router = APIRouter(prefix="/api/v1/chat", tags=["chat"])


@router.get("/", summary="all chats")
async def get_list(
    offset: int = 0, limit: int = 100, chat_service: ChatService = Depends()
) -> list[ChatResponse]:
    chats = await chat_service.get_list(offset=offset, limit=limit)

    logger.debug(f"chats - {chats}")

    return chats


@router.get("/{id}", summary="chat with all messages")
async def get(
    id: uuid.UUID, chat_service: ChatService = Depends()
) -> ChatResponseWithMessages:
    chat = await chat_service.get(id)

    return chat


@router.post("/", summary="create the chat")
async def create(
    title: str = Form(...),
    presentation: UploadFile = File(...),
    chat_service: ChatService = Depends(),
) -> ChatResponse:
    presentation = await presentation.read()
    chat = await chat_service.create(title, None, presentation)

    return chat


@router.post("/{id}/add_message", summary="add the message to chat")
async def add_message(
    chat_id: uuid.UUID,
    message: str,
    history_id: uuid.UUID | None = None,
    chat_service: ChatService = Depends(),
):
    presentation = await chat_service.add_message(chat_id, history_id, message)

    return Response(
        content=presentation,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
