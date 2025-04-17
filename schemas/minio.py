from enum import Enum


class MinioContentType(Enum):
    PNG = "image/png"
    MP4 = "video/mp4"
    PDF = "application/pdf"
    DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
