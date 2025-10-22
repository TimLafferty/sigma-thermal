"""
Pytest configuration and fixtures for sigma_thermal tests.
"""

import pytest
from pathlib import Path


@pytest.fixture
def repo_root() -> Path:
    """Return the repository root directory"""
    return Path(__file__).parent.parent


@pytest.fixture
def test_data_dir(repo_root) -> Path:
    """Return the test data directory"""
    return repo_root / 'data' / 'validation_cases'


@pytest.fixture
def sources_dir(repo_root) -> Path:
    """Return the sources directory with Excel files"""
    return repo_root / 'sources'


@pytest.fixture
def lookup_tables_dir(repo_root) -> Path:
    """Return the lookup tables directory"""
    return repo_root / 'data' / 'lookup_tables'
