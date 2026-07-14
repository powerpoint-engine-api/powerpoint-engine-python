"""PowerPoint Engine API Python SDK

A Python client for the PowerPoint Engine API (https://powerpointengine.io):
generate .pptx from Markdown or a template, edit/merge/replace in existing
decks, and convert PPTX to PDF.
"""

__version__ = "2.0.0"
__author__ = "PowerPoint Engine API"
__email__ = "support@powerpointengine.io"

from .client import PowerPointEngine
from .exceptions import (
    PowerPointEngineError,
    AuthenticationError,
    ValidationError,
    NotFoundError,
    RateLimitError,
    ServerError,
)

__all__ = [
    "PowerPointEngine",
    "PowerPointEngineError",
    "AuthenticationError",
    "ValidationError",
    "NotFoundError",
    "RateLimitError",
    "ServerError",
]
