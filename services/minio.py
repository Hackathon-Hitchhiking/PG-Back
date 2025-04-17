import uuid
from io import BytesIO

from fastapi import Depends

from repositories.minio import MinioRepository
from schemas.minio import MinioContentType


class MinioService:
    def __init__(self, repo: MinioRepository = Depends()):
        self._repo = repo

    def create_message_pptx_path(self, chat_id: uuid.UUID, message_id: uuid.UUID) -> str:
        return f"{chat_id.__str__()}/{message_id.__str__()}.pptx"

    def save_pptx(self, chat_id: uuid.UUID, message_id: uuid.UUID, file: BytesIO) -> str:
        return self._repo.create_object_from_byte(
            self.create_message_pptx_path(chat_id, message_id), file, MinioContentType.PPTX
        )

    def get_pptx(self, chat_id: uuid.UUID, message_id: uuid.UUID) -> bytes:
        return self._repo.get_object_as_bytes(self.create_message_pptx_path(chat_id, message_id))