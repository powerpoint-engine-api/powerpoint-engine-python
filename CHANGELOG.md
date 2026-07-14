# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2026-07-14

### Changed
- **Breaking:** rewritten against the live API at `https://powerpointengine.io/api/powerpoint/*` (the old `api.powerpointengine.io` `/v1` endpoints never shipped)
- Single `PowerPointEngine` client; removed unshipped resources/webhooks/async layers
- No API key required: anonymous calls work (watermarked); pass `session_id` to bill your account

### Added
- `generate()` — markup (Markdown dialect) or structured template, `brand`/`font` overrides
- `replace()` — structure-preserving text + image replacement in your own .pptx
- `edit()` — duplicate / delete / move slides
- `merge()` — merge 2-5 decks, each keeps its design
- `to_pdf()` — PPTX to PDF conversion
- `translate()` — in-place deck translation
- `download()` helper for signed result URLs

## [1.0.0] - 2024-01-15

### Added
- Initial release of powerpoint-engine-python
- Complete API client implementation
- Comprehensive documentation
- Example code and usage guides
- Full test suite
- CI/CD pipeline configuration

### Features
- Create PowerPoint presentations programmatically
- Work with custom templates
- Support for charts, tables, and images
- Webhook notifications
- Usage analytics
- Error handling and retry logic
