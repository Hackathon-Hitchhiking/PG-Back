import uuid

from fastapi import Depends
from loguru import logger
from sqlalchemy import select, desc
from sqlalchemy.exc import NoResultFound
from sqlalchemy.ext.asyncio import AsyncSession

from configs.Database import get_db_connection
from errors.errors import ErrEntityNotFound
from models.message import Message


class MessageRepository:
    def __init__(self, db: AsyncSession = Depends(get_db_connection)):
        self._db = db

    async def create(self, message: Message) -> Message:
        logger.debug("Message - Repository - create")
        self._db.add(message)
        await self._db.commit()
        await self._db.refresh(message)
        return message

    async def get(self, uuid: uuid.UUID) -> Message:
        logger.debug("Message - Repository - get")
        query = select(Message).where(Message.id == uuid)
        result = await self._db.execute(query)
        try:
            message = result.scalar_one()
        except NoResultFound:
            raise ErrEntityNotFound("Message not found")
        return message

    async def update(self, message: Message) -> Message:
        logger.debug("Message - Repository - update")
        await self._db.commit()
        await self._db.refresh(message)
        return message

    async def list(
        self, offset: int = 0, limit: int = 100, chat_id: uuid.UUID | None = None
    ) -> list[Message]:
        logger.debug("Message - Repository - list")

        query = select(Message).offset(offset).limit(limit)

        if chat_id is None:
            query = query.where(Message.chat_id == chat_id)

        query = query.order_by(desc(Message.created_at))

        result = await self._db.execute(query)
        return list(result.scalars().all())

    async def get_latest_message(self, chat_id) -> Message:
        query = (
            select(Message)
            .filter(Message.chat_id == chat_id)
            .order_by(desc(Message.created_at))
            .limit(1)
        )

        result = await self._db.execute(query)

        try:
            message = result.scalar_one()
        except NoResultFound:
            raise ErrEntityNotFound("Message not found")
        return message
