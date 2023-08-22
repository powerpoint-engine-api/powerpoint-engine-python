"""PowerPoint Engine API Python SDK

A Python client library for the PowerPoint Engine API.
"""

__version__ = "1.0.0"
__author__ = "PowerPoint Engine API"
__email__ = "support@powerpointengine.io"

from .client import PowerPointEngine, AsyncPowerPointEngine
from .exceptions import (
    PowerPointEngineError,
    AuthenticationError,
    ValidationError,
    NotFoundError,
    RateLimitError,
    ServerError,
)
from .models import (
    Presentation,
    Template,
    Slide,
    ChartData,
    TableData,
    Webhook,
)

__all__ = [
    "PowerPointEngine",
    "AsyncPowerPointEngine",
    "PowerPointEngineError",
    "AuthenticationError",
    "ValidationError", 
    "NotFoundError",
    "RateLimitError",
    "ServerError",
    "Presentation",
    "Template",
    "Slide",
    "ChartData",
    "TableData",
    "Webhook",
]