# Contributing

Thanks for helping improve PDF Excel. This guide keeps contributions easy to review and maintain.

## Before You Start

- Check existing issues and pull requests to avoid duplicate work.
- Open an issue first for larger changes, new features, or design decisions.
- Keep pull requests focused on one topic.

## Development Workflow

1. Fork the repository.
2. Create a branch from `main`.
3. Make your changes.
4. Add or update tests and documentation when relevant.
5. Open a pull request with a clear summary.

Example:

```sh
git checkout -b feature/my-change
```

## Pull Request Checklist

- The change has a clear purpose.
- Documentation has been updated when behavior or usage changes.
- Tests have been added or updated when practical.
- Generated files, local caches, and private documents are not committed.

## Reporting Bugs

When reporting a bug, include:

- What you expected to happen.
- What actually happened.
- Steps to reproduce the problem.
- Relevant input details, such as PDF type or spreadsheet output format.
- Environment details, such as operating system and tool versions.

Do not upload private, confidential, or legally restricted documents. If a PDF is needed to reproduce an issue, use a sanitized sample.

## Requesting Features

Feature requests should describe the workflow, not only the implementation. Include:

- The problem you are trying to solve.
- The type of PDF or data involved.
- The desired spreadsheet output.
- Any constraints that matter, such as accuracy, speed, or offline use.

## Code Style

This project uses a small Python package under `src/pdf_excel`.

- Keep CLI behavior simple and documented in `README.md`.
- Add focused tests in `tests/` for new behavior.
- Run the test suite before opening a pull request:

```sh
PYTHONPATH=src python3 -m unittest discover -s tests
```

If the package is installed in your environment, this also works:

```sh
python -m unittest discover -s tests
```
