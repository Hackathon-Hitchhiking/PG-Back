import uuid
from datetime import datetime

from sqlalchemy import ForeignKey, DateTime, func, UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from models.BaseModel import EntityMeta


class Message(EntityMeta):
    __tablename__ = 'message'
    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    path: Mapped[str] = mapped_column(comment="the path in the s3")

    message: Mapped[str]

    created_at: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )

    chat_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("chat.id"))

    chat: Mapped["Chat"] = relationship(back_populates="messages")