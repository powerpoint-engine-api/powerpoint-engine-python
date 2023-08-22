"""PowerPoint Engine API Client"""

import json
import time
from typing import Optional, Dict, Any, Union
from urllib.parse import urljoin

import requests

from .exceptions import PowerPointEngineError, AuthenticationError, ValidationError, NotFoundError, RateLimitError, ServerError
from .resources import PresentationsResource, TemplatesResource, WebhooksResource


class PowerPointEngine:
    """Synchronous client for PowerPoint Engine API."""
    
    def __init__(
        self,
        api_key: str,
        base_url: str = "https://api.powerpointengine.io",
        timeout: int = 30,
        max_retries: int = 3,
    ):
        """Initialize the PowerPoint Engine client.
        
        Args:
            api_key: Your PowerPoint Engine API key
            base_url: Base URL for the API (default: https://api.powerpointengine.io)
            timeout: Request timeout in seconds (default: 30)
            max_retries: Maximum number of retry attempts (default: 3)
        """
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries
        
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"powerpoint-engine-python/1.0.0",
        })
        
        # Initialize resource endpoints
        self.presentations = PresentationsResource(self)
        self.templates = TemplatesResource(self)
        self.webhooks = WebhooksResource(self)
    
    def request(
        self,
        method: str,
        endpoint: str,
        data: Optional[Dict[str, Any]] = None,
        files: Optional[Dict[str, Any]] = None,
        params: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make a request to the API with retry logic.
        
        Args:
            method: HTTP method (GET, POST, PUT, DELETE)
            endpoint: API endpoint path
            data: Request body data
            files: Files to upload
            params: Query parameters
            headers: Additional headers
            
        Returns:
            Parsed JSON response
            
        Raises:
            PowerPointEngineError: For API errors
        """
        url = urljoin(self.base_url, endpoint.lstrip("/"))
        
        request_headers = self.session.headers.copy()
        if headers:
            request_headers.update(headers)
        
        # Handle file uploads
        if files:
            # Remove Content-Type for multipart uploads
            request_headers.pop("Content-Type", None)
            json_data = None
        else:
            json_data = data
        
        retry_count = 0
        while retry_count <= self.max_retries:
            try:
                response = self.session.request(
                    method=method,
                    url=url,
                    json=json_data,
                    files=files,
                    params=params,
                    headers=request_headers,
                    timeout=self.timeout,
                )
                
                # Handle different response types
                if response.status_code == 204:  # No Content
                    return {}
                
                # Check if response is JSON
                content_type = response.headers.get("content-type", "")
                if "application/json" in content_type:
                    response_data = response.json()
                else:
                    # For file downloads
                    if response.status_code == 200:
                        return {"content": response.content}
                    response_data = {"message": response.text}
                
                # Handle HTTP errors
                if not response.ok:
                    self._handle_error_response(response.status_code, response_data)
                
                return response_data
                
            except requests.RequestException as e:
                retry_count += 1
                if retry_count > self.max_retries:
                    raise PowerPointEngineError(f"Network error after {self.max_retries} retries: {str(e)}")
                
                # Exponential backoff
                wait_time = (2 ** retry_count) + (retry_count * 0.1)
                time.sleep(wait_time)
    
    def _handle_error_response(self, status_code: int, response_data: Dict[str, Any]) -> None:
        """Handle API error responses."""
        error_message = response_data.get("error", response_data.get("message", "Unknown error"))
        request_id = response_data.get("request_id")
        
        if status_code == 401:
            raise AuthenticationError(error_message, status_code, request_id)
        elif status_code == 400:
            raise ValidationError(error_message, status_code, request_id)
        elif status_code == 404:
            raise NotFoundError(error_message, status_code, request_id)
        elif status_code == 429:
            raise RateLimitError(error_message, status_code, request_id)
        elif status_code >= 500:
            raise ServerError(error_message, status_code, request_id)
        else:
            raise PowerPointEngineError(error_message, status_code, request_id)


class AsyncPowerPointEngine:
    """Asynchronous client for PowerPoint Engine API."""
    
    def __init__(
        self,
        api_key: str,
        base_url: str = "https://api.powerpointengine.io",
        timeout: int = 30,
        max_retries: int = 3,
    ):
        """Initialize the async PowerPoint Engine client.
        
        Args:
            api_key: Your PowerPoint Engine API key
            base_url: Base URL for the API
            timeout: Request timeout in seconds
            max_retries: Maximum number of retry attempts
        """
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries
        
        self._session = None
        self._headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"powerpoint-engine-python/1.0.0",
        }
        
        # Initialize resource endpoints (async versions)
        from .resources import AsyncPresentationsResource, AsyncTemplatesResource, AsyncWebhooksResource
        self.presentations = AsyncPresentationsResource(self)
        self.templates = AsyncTemplatesResource(self)
        self.webhooks = AsyncWebhooksResource(self)
    
    async def __aenter__(self):
        """Async context manager entry."""
        await self._ensure_session()
        return self
    
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Async context manager exit."""
        await self.close()
    
    async def _ensure_session(self):
        """Ensure aiohttp session is created."""
        if self._session is None:
            import aiohttp
            self._session = aiohttp.ClientSession(
                timeout=aiohttp.ClientTimeout(total=self.timeout),
                headers=self._headers,
            )
    
    async def close(self):
        """Close the aiohttp session."""
        if self._session:
            await self._session.close()
            self._session = None
    
    async def request(
        self,
        method: str,
        endpoint: str,
        data: Optional[Dict[str, Any]] = None,
        files: Optional[Dict[str, Any]] = None,
        params: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make an async request to the API."""
        import aiohttp
        import asyncio
        
        await self._ensure_session()
        
        url = urljoin(self.base_url, endpoint.lstrip("/"))
        
        request_headers = self._headers.copy()
        if headers:
            request_headers.update(headers)
        
        retry_count = 0
        while retry_count <= self.max_retries:
            try:
                async with self._session.request(
                    method=method,
                    url=url,
                    json=data,
                    data=files,
                    params=params,
                    headers=request_headers,
                ) as response:
                    
                    if response.status == 204:
                        return {}
                    
                    content_type = response.headers.get("content-type", "")
                    if "application/json" in content_type:
                        response_data = await response.json()
                    else:
                        if response.status == 200:
                            content = await response.read()
                            return {"content": content}
                        text = await response.text()
                        response_data = {"message": text}
                    
                    if not response.ok:
                        self._handle_error_response(response.status, response_data)
                    
                    return response_data
                    
            except aiohttp.ClientError as e:
                retry_count += 1
                if retry_count > self.max_retries:
                    raise PowerPointEngineError(f"Network error after {self.max_retries} retries: {str(e)}")
                
                wait_time = (2 ** retry_count) + (retry_count * 0.1)
                await asyncio.sleep(wait_time)
    
    def _handle_error_response(self, status_code: int, response_data: Dict[str, Any]) -> None:
        """Handle API error responses (same as sync client)."""
        error_message = response_data.get("error", response_data.get("message", "Unknown error"))
        request_id = response_data.get("request_id")
        
        if status_code == 401:
            raise AuthenticationError(error_message, status_code, request_id)
        elif status_code == 400:
            raise ValidationError(error_message, status_code, request_id)
        elif status_code == 404:
            raise NotFoundError(error_message, status_code, request_id)
        elif status_code == 429:
            raise RateLimitError(error_message, status_code, request_id)
        elif status_code >= 500:
            raise ServerError(error_message, status_code, request_id)
        else:
            raise PowerPointEngineError(error_message, status_code, request_id)