import uuid
from typing import Type, Any, List

from loguru import logger


class CRUDServiceMixin:
    def __init__(self, repository: Any) -> None:
        self._repo = repository

    async def list(self, limit: int, offset: int, **filters) -> List[Type[Any]]:
        logger.debug(f"{self._repo.model.__name__} - Service - list")
        result = await self._repo.list(limit, offset, **filters)
        return result

    async def get(self, id: uuid.UUID) -> Type[Any]:
        logger.debug(f"{self._repo.model.__name__} - Service - get_by_id")
        result = await self._repo.get(id)
        return result

    async def create(self, entity: Type[Any]) -> Type[Any]:
        logger.debug(f"{self._repo.model.__name__} - Service - create")
        entity.id = uuid.uuid4()
        result = await self._repo.create(entity)
        return result

    async def update(self, updated_entity: Type[Any]) -> Type[Any]:
        logger.debug(f"{self._repo.model.__name__} - Service - update")
        result = await self._repo.update(updated_entity)

        return result

    async def delete(self, id: uuid.UUID) -> None:
        logger.debug(f"{self._repo.model.__name__} - Service - delete")
        await self._repo.delete(id)
        return None
