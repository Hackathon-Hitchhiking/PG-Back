import uuid

from fastapi import Depends
from loguru import logger
from sqlalchemy import select
from sqlalchemy.exc import NoResultFound
from sqlalchemy.ext.asyncio import AsyncSession

from configs.Database import get_db_connection
from errors.errors import ErrEntityNotFound
from models.chat import Chat


class ChatRepository:
    def __init__(self, db: AsyncSession = Depends(get_db_connection)):
        self._db = db

    async def create(self, chat: Chat) -> Chat:
        logger.debug("Chat - Repository - create")
        self._db.add(chat)
        await self._db.commit()
        await self._db.refresh(chat)
        return chat

    async def get(self, uuid: uuid.UUID) -> Chat:
        logger.debug("Chat - Repository - get")
        query = select(Chat).where(Chat.id == uuid)
        result = await self._db.execute(query)
        try:
            chat = result.scalar_one()
        except NoResultFound:
            raise ErrEntityNotFound("Chat not found")
        return chat

    async def update(self, chat: Chat) -> Chat:
        logger.debug("Chat - Repository - update")
        await self._db.commit()
        await self._db.refresh(chat)
        return chat

    async def list(self, offset: int = 0, limit: int = 100) -> list[Chat]:
        logger.debug("Chat - Repository - list")
        query = select(Chat).offset(offset).limit(limit)
        result = await self._db.execute(query)
        return list(result.scalars().all())