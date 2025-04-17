import uuid

from sqlalchemy import DateTime, func, UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from models.BaseModel import EntityMeta


class Chat(EntityMeta):
    __tablename__ = "chat"
    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )

    title: Mapped[str]

    user_id: Mapped[uuid.UUID | None] = mapped_column(nullable=True)

    messages: Mapped[list["Message"]] = relationship(
        back_populates="chat", lazy="selectin"
    )

    created_at: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )
    updated_at: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True),
        onupdate=func.now(),
        server_default=func.now(),
        nullable=False,
    )
