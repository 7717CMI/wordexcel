from typing import Optional, Dict, Any
from pydantic import BaseModel


class FileUploadResponse(BaseModel):
    """Response model for file upload"""
    success: bool
    fileId: str
    message: str


class ProcessingResult(BaseModel):
    """Response model for document processing"""
    success: bool
    data: Optional[Dict[str, Any]] = None
    confidence: float = 0.0
    error: Optional[str] = None


class ApiResponse(BaseModel):
    """Generic API response model"""
    success: bool
    message: str
    data: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
