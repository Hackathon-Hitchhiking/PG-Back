import uuid
from typing import Type, Any, Sequence

from loguru import logger
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from errors.errors import ErrEntityNotFound


class CRUDRepositoryMixin:
    def __init__(self, model: Type[Any], db: AsyncSession):
        self.model = model
        self._db = db

    async def list(self, limit: int, offset: int, **filters) -> Sequence[Type[Any]]:
        logger.debug(f"{self.model.__name__} - Repository - get_list")
        query = select(self.model).offset(offset).limit(limit)

        for k, v in filters.items():
            try:
                column = getattr(self.model, k)
            except AttributeError:
                continue

            query = query.where(column == v)

        result = await self._db.execute(query)
        return result.scalars().all()

    async def get(self, id: uuid.UUID) -> Type[Any]:
        logger.debug(f"{self.model.__name__} - Repository - get_by_id")
        instance = await self._db.get(self.model, id)
        if instance is None:
            raise ErrEntityNotFound(f"{self.model.__name__} not found")
        return instance

    async def create(self, instance: Type[Any]) -> Type[Any]:
        logger.debug(f"{self.model.__name__} - Repository - create")
        self._db.add(instance)
        await self._db.commit()
        await self._db.refresh(instance)
        return instance

    async def update(self, updated_entity: Type[Any]) -> Type[Any]:
        logger.debug(f"{self.model.__name__} - Repository - update")
        self._db.add(updated_entity)
        await self._db.commit()
        await self._db.refresh(updated_entity)
        return updated_entity

    async def delete(self, id: uuid.UUID) -> None:
        logger.debug(f"{self.model.__name__} - Repository - delete")
        instance = await self.get(id)
        await self._db.delete(instance)
        await self._db.commit()
        await self._db.flush()
