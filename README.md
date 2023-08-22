# PowerPoint Engine API - Python SDK

![PyPI version](https://badge.fury.io/py/powerpoint-engine-api.svg)
![Python versions](https://img.shields.io/pypi/pyversions/powerpoint-engine-api.svg)
![License](https://img.shields.io/github/license/powerpoint-engine-api/powerpoint-engine-python.svg)

A Python SDK for the PowerPoint Engine API that allows you to generate, modify, and manipulate PowerPoint presentations programmatically.

## Features

- 🚀 **Simple & Intuitive**: Easy-to-use Python interface
- 📊 **Rich Data Support**: Charts, tables, images, and text
- 🎨 **Template Processing**: Work with custom PowerPoint templates
- 🔐 **Secure**: Built-in API key management
- 📝 **Type Hints**: Full typing support for better IDE experience
- ⚡ **Async Support**: Both sync and async clients available

## Installation

```bash
pip install powerpoint-engine-api
```

## Quick Start

```python
from powerpoint_engine import PowerPointEngine

# Initialize the client
client = PowerPointEngine(api_key="your_api_key_here")

# Generate a simple presentation
presentation = client.presentations.create({
    "template": "business-report",
    "data": {
        "title": "Q4 Business Report",
        "slides": [
            {
                "type": "title",
                "title": "Sales Performance",
                "subtitle": "Outstanding results this quarter"
            },
            {
                "type": "chart",
                "title": "Revenue Growth",
                "chart_data": {
                    "labels": ["Jan", "Feb", "Mar"],
                    "values": [100, 150, 200]
                }
            }
        ]
    }
})

# Download the presentation
with open("report.pptx", "wb") as f:
    f.write(presentation.download())

print(f"Presentation created: {presentation.id}")
```

## API Reference

### Client Initialization

```python
from powerpoint_engine import PowerPointEngine

# Basic initialization
client = PowerPointEngine(api_key="your_api_key")

# With custom base URL
client = PowerPointEngine(
    api_key="your_api_key",
    base_url="https://api.powerpointengine.io"
)

# Async client
from powerpoint_engine import AsyncPowerPointEngine
async_client = AsyncPowerPointEngine(api_key="your_api_key")
```

### Presentations

#### Create Presentation

```python
# From template
presentation = client.presentations.create({
    "template": "business-report",
    "data": {
        "title": "My Presentation",
        "slides": [...]
    }
})

# From scratch
presentation = client.presentations.create({
    "slides": [
        {
            "type": "title",
            "title": "Welcome",
            "subtitle": "Getting Started"
        }
    ]
})
```

#### Get Presentation

```python
presentation = client.presentations.get("presentation_id")
print(f"Status: {presentation.status}")
print(f"Created: {presentation.created_at}")
```

#### List Presentations

```python
presentations = client.presentations.list(limit=10, offset=0)
for presentation in presentations.data:
    print(f"{presentation.id}: {presentation.status}")
```

#### Download Presentation

```python
# Download as bytes
content = presentation.download()

# Save to file
presentation.save("my_presentation.pptx")
```

### Templates

#### List Available Templates

```python
templates = client.templates.list()
for template in templates.data:
    print(f"{template.id}: {template.name}")
```

#### Get Template Details

```python
template = client.templates.get("business-report")
print(f"Description: {template.description}")
print(f"Placeholders: {template.placeholders}")
```

#### Upload Custom Template

```python
with open("my_template.pptx", "rb") as f:
    template = client.templates.upload(
        file=f,
        name="custom-template",
        description="My custom template"
    )
```

### Data Types

#### Slide Types

```python
# Title slide
{
    "type": "title",
    "title": "Presentation Title",
    "subtitle": "Subtitle text"
}

# Content slide
{
    "type": "content",
    "title": "Slide Title",
    "content": "Slide content text"
}

# Chart slide
{
    "type": "chart",
    "title": "Chart Title",
    "chart_type": "bar",  # bar, line, pie, column
    "chart_data": {
        "labels": ["A", "B", "C"],
        "datasets": [
            {
                "label": "Series 1",
                "values": [10, 20, 30]
            }
        ]
    }
}

# Table slide
{
    "type": "table",
    "title": "Table Title",
    "table_data": {
        "headers": ["Name", "Value", "Change"],
        "rows": [
            ["Product A", "100", "+5%"],
            ["Product B", "200", "+10%"]
        ]
    }
}

# Image slide
{
    "type": "image",
    "title": "Image Title",
    "image_url": "https://example.com/image.jpg",
    "caption": "Image caption"
}
```

## Advanced Usage

### Error Handling

```python
from powerpoint_engine import PowerPointEngine, PowerPointEngineError

try:
    client = PowerPointEngine(api_key="invalid_key")
    presentation = client.presentations.create({"slides": []})
except PowerPointEngineError as e:
    print(f"API Error: {e.message}")
    print(f"Status Code: {e.status_code}")
    print(f"Request ID: {e.request_id}")
```

### Async Usage

```python
import asyncio
from powerpoint_engine import AsyncPowerPointEngine

async def main():
    client = AsyncPowerPointEngine(api_key="your_api_key")
    
    # Create presentation asynchronously
    presentation = await client.presentations.create({
        "template": "business-report",
        "data": {"title": "Async Report"}
    })
    
    # Download when ready
    content = await presentation.download()
    
    await client.close()  # Clean up connections

asyncio.run(main())
```

### Webhooks

```python
# Configure webhook for presentation completion
webhook = client.webhooks.create({
    "url": "https://yourapp.com/webhook",
    "events": ["presentation.completed", "presentation.failed"]
})

# List webhooks
webhooks = client.webhooks.list()

# Delete webhook
client.webhooks.delete(webhook.id)
```

## Examples

Check out the [examples](./examples/) directory for more detailed examples:

- [Basic presentation generation](./examples/basic_presentation.py)
- [Working with templates](./examples/templates.py)
- [Charts and data visualization](./examples/charts.py)
- [Batch processing](./examples/batch_processing.py)
- [Async operations](./examples/async_example.py)

## Requirements

- Python 3.7+
- requests >= 2.25.0
- aiohttp >= 3.8.0 (for async client)

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

## License

This SDK is released under the MIT License. See [LICENSE](LICENSE) for details.

## Support

- 📖 [API Documentation](https://powerpointengine.io/docs)
- 💬 [GitHub Issues](https://github.com/powerpoint-engine-api/powerpoint-engine-python/issues)
- 📧 [Email Support](mailto:support@powerpointengine.io)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and updates.