# GitHub Actions Workflows

This directory contains CI/CD workflows for the Sigma Thermal project.

## Workflows

### CI Pipeline (`ci.yml`)

Comprehensive continuous integration pipeline that runs on every push and pull request.

**Jobs:**

1. **Test** - Runs on Python 3.11 and 3.12
   - Linting with ruff
   - Code formatting check with black
   - Type checking with mypy
   - Unit tests with coverage
   - Validation tests (Python vs Excel VBA)
   - Integration tests (complete workflows)
   - Coverage reporting to Codecov

2. **Lint** - Code quality checks
   - Ruff linting
   - Black formatting verification
   - Import sorting with isort
   - Type checking with mypy

3. **Build** - Package building and verification
   - Build distribution packages
   - Verify package integrity with twine
   - Upload build artifacts

**Triggers:**
- Push to `main` or `develop` branches
- Pull requests to `main` or `develop` branches

## Test Coverage

The CI pipeline generates coverage reports in multiple formats:
- Terminal output (shown in workflow logs)
- XML format (for Codecov integration)
- HTML format (archived as artifacts)

Coverage reports are uploaded as artifacts and can be downloaded from the workflow run.

## Required Secrets

For full functionality, add these secrets to your repository:
- `CODECOV_TOKEN` - Token for Codecov integration (optional)

## Local Testing

To run the same checks locally before pushing:

```bash
# Linting
ruff check src/ tests/

# Formatting
black --check src/ tests/

# Import sorting
isort --check-only src/ tests/

# Type checking
mypy src/

# All tests with coverage
pytest tests/ -v --cov=sigma_thermal --cov-report=html
```

## Artifacts

The following artifacts are generated and available for download:
- `test-results-{python-version}` - Test coverage HTML reports
- `dist` - Built distribution packages
