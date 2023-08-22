"""PowerPoint Engine API Exceptions"""

from typing import Optional


class PowerPointEngineError(Exception):
    """Base exception for PowerPoint Engine API errors."""
    
    def __init__(
        self,
        message: str,
        status_code: Optional[int] = None,
        request_id: Optional[str] = None,
    ):
        """Initialize the exception.
        
        Args:
            message: Error message
            status_code: HTTP status code
            request_id: Request ID for debugging
        """
        super().__init__(message)
        self.message = message
        self.status_code = status_code
        self.request_id = request_id
    
    def __str__(self) -> str:
        """String representation of the error."""
        parts = [self.message]
        if self.status_code:
            parts.append(f"Status: {self.status_code}")
        if self.request_id:
            parts.append(f"Request ID: {self.request_id}")
        return " | ".join(parts)


class AuthenticationError(PowerPointEngineError):
    """Raised when API key is invalid or missing."""
    pass


class ValidationError(PowerPointEngineError):
    """Raised when request data is invalid."""
    pass


class NotFoundError(PowerPointEngineError):
    """Raised when requested resource is not found."""
    pass


class RateLimitError(PowerPointEngineError):
    """Raised when rate limit is exceeded."""
    pass


class ServerError(PowerPointEngineError):
    """Raised when server encounters an error."""
    pass