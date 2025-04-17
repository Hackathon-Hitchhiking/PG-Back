import io
import uuid

from fastapi import Depends
from loguru import logger

from models.chat import Chat
from models.message import Message
from repositories.chat import ChatRepository
from repositories.history import MessageRepository
from services.minio import MinioService
from services.ml import MLService


class ChatService:
    def __init__(
        self,
        repo: ChatRepository = Depends(),
        history_repo: MessageRepository = Depends(),
        ml_service: MLService = Depends(),
        minio_service: MinioService = Depends(),
    ):
        self._repo = repo
        self._message_repo = history_repo
        self._ml_service = ml_service
        self._minio_service = minio_service

    async def get_list(self, offset: int = 0, limit: int = 100) -> list[Chat]:
        logger.debug("Chat - Service - get_chats")
        chats = await self._repo.list(offset=offset, limit=limit)

        return chats

    async def get(self, chat_id: uuid.UUID) -> Chat:
        logger.debug("Chat - Service - get_chat")
        chat = await self._repo.get(chat_id)

        return chat

    async def get_chat_history(
        self, chat_id: uuid.UUID, offset: int = 0, limit: int = 100
    ) -> list[Message]:
        logger.debug("Chat - Service - get_chat_history")

        chat_history = await self._message_repo.list(
            offset=offset, limit=limit, chat_id=chat_id
        )

        return chat_history

    async def create(
        self, title: str, user_id: uuid.UUID | None, presentation: bytes
    ) -> Chat:
        logger.debug("Chat - Service - create")

        message_id = uuid.uuid4()

        chat = await self._repo.create(
            Chat(
                title=title,
                user_id=user_id,
            )
        )

        minio_path = self._minio_service.save_pptx(
            chat.id, message_id, io.BytesIO(presentation)
        )

        message = Message(
            id=message_id,
            chat_id=chat.id,
            message="",
            path=minio_path,
        )

        await self._message_repo.create(message)

        return chat

    async def add_message(
        self, chat_id: uuid.UUID, history_id: uuid.UUID | None, prompt: str
    ) -> bytes:
        logger.debug("Chat - Service - add_message")

        message_id = uuid.uuid4()

        if history_id is None:
            message = await self._message_repo.get_latest_message(chat_id)
        else:
            message = await self._message_repo.get(history_id)

        pres = self._minio_service.get_pptx(chat_id, message.id)

        edited_pres = self._ml_service.edit_pres(prompt, pres)

        path = self._minio_service.save_pptx(
            chat_id, message_id, io.BytesIO(edited_pres)
        )

        await self._message_repo.create(
            Message(
                id=message_id,
                chat_id=chat_id,
                message=prompt,
                path=path,
            )
        )

        return edited_pres
