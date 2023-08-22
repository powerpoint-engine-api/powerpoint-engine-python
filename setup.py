#!/usr/bin/env python

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="powerpoint-engine-api",
    version="1.0.0",
    author="PowerPoint Engine API",
    author_email="support@powerpointengine.io",
    description="Python SDK for PowerPoint Engine API - Generate PowerPoint presentations programmatically",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/powerpoint-engine-api/powerpoint-engine-python",
    project_urls={
        "Bug Tracker": "https://github.com/powerpoint-engine-api/powerpoint-engine-python/issues",
        "Documentation": "https://powerpointengine.io/docs",
        "Source Code": "https://github.com/powerpoint-engine-api/powerpoint-engine-python",
    },
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Multimedia :: Graphics :: Presentation",
    ],
    python_requires=">=3.7",
    install_requires=requirements,
    extras_require={
        "async": ["aiohttp>=3.8.0", "aiofiles>=0.8.0"],
        "dev": [
            "pytest>=6.0",
            "pytest-asyncio>=0.18.0",
            "pytest-cov>=3.0.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.950",
            "sphinx>=4.0.0",
            "sphinx-rtd-theme>=1.0.0",
        ],
    },
    keywords="powerpoint, presentation, api, sdk, office, slides, charts, automation",
    include_package_data=True,
    zip_safe=False,
)