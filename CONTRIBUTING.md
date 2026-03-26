# Contributing to Outlook Mail Reader (Graph App-Only)

Thank you for considering a contribution! This is a small, focused Shuffle SOAR app — contributions of any size are welcome.

## Ways to Contribute

- **Bug reports** — Open an issue with steps to reproduce, expected vs. actual behaviour, and your Python / Shuffle version.
- **Feature requests** — Open an issue describing the use case and why it would benefit other users.
- **Pull requests** — Fork the repo, make your changes on a branch, and open a PR against `main`.

## Development Setup

```bash
git clone https://github.com/uldagalihan/outlook-graph-app-only.git
cd outlook-graph-app-only/shuffle-outlook-mail-reader/1.0.2

python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

The `walkoff_app_sdk` dependency is provided by the Shuffle base Docker image (`frikky/shuffle:app_sdk`). For local development outside Docker you can install it separately:

```bash
pip install walkoff-app-sdk
```

## Code Style

- Follow [PEP 8](https://peps.python.org/pep-0008/).
- Add type hints to all new public functions.
- Add docstrings to all new public methods.
- All code, comments, and documentation must be written in **English**.
- Keep `api.yaml` in sync with the action method signatures in `app.py`.

## Pull Request Checklist

- [ ] No secrets, tokens, or personal data in any committed file.
- [ ] No hardcoded mailbox addresses, subjects, or internal business logic.
- [ ] All new actions defined in `api.yaml` match the corresponding Python method names.
- [ ] `requirements.txt` updated if new dependencies are introduced.
- [ ] README updated if behaviour or configuration changes.

## Questions?

Open a GitHub issue — no question is too small.
