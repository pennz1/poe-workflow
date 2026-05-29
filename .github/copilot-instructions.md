# Copilot Instructions

## Build & Run

```bash
# Install dependencies
pip install -r requirements.txt
python -m playwright install chromium

# Run the Streamlit web app
streamlit run app.py

# Run CLI generation (no Streamlit required)
python cli_generate.py --config config.json --output ./output

# Run all tests
python -m pytest test_*.py

# Run a single test file
python -m pytest test_tier_algorithm.py -v

# Run a single test function
python -m pytest test_tier_algorithm.py::TestSnapBudgetToTier::test_exact_match -v
```

## Architecture

This is a **Microsoft pre-sales POE (Proof of Engagement) document generator** for Azure solutions.

**Core flow:** User input → Azure OpenAI generates Markdown → Markdown converted to Word (.docx) via python-docx templates.

**Key modules:**

- `app.py` (~3300 lines) — Monolithic Streamlit app containing all business logic: LLM calls, document generation, Azure ARM API integration, MSAL auth, session persistence
- `cli_generate.py` — Headless CLI interface that mocks Streamlit to reuse `app.py` functions directly
- `pricing_automation.py` — Playwright-based browser automation for Azure Pricing Calculator
- `frontend/ui.py` — Custom Streamlit HTML components (renders raw HTML via `unsafe_allow_html`)
- `templates/` — Word `.docx` style templates and CSV header templates

**Document generation pipeline:**
1. Solution Architecture doc (AI or Infra type) — LLM → Markdown → .docx
2. POV Deployment Plan — depends on step 1 output
3. Azure Migrate CSV — budget-based VM config generation
4. Full-auto mode: steps 1-3 + MSAL login + ARM API for Migrate project creation

**Secrets resolution:** `get_secret(key)` checks env vars first, falls back to `.streamlit/secrets.toml`.

## Key Conventions

- **Language:** Code comments, variable names, and git messages may mix Chinese and English. All user-facing UI text and generated documents are in Chinese.
- **Testing pattern:** Tests must mock `streamlit` and `frontend.ui` modules before importing `app.py`, since it executes `st.set_page_config()` at module level.
- **Budget tiers:** The system uses fixed tiers (15k / 50k / 100k / 250k USD) with a tier cache (`.tier_cache.json`) for machine selection learning.
- **Document formatting rules:** Generated Markdown must never use bullet lists (-, *, •). Use paragraph prose with `keyword: description` format on separate lines.
- **Model references:** Use current models (GPT-5.5, GPT-5.4, o4-mini, GPT-4.1). Never reference deprecated GPT-4o.
- **Azure regions:** Only global regions allowed (East US, East Asia, etc.). China regions (China East/North) are strictly forbidden in generated content.
- **Session persistence:** `_PERSIST_KEYS` list in `app.py` defines which `st.session_state` keys survive page refreshes (file-based persistence when `PERSIST_DIR` is set).
- **OpenAI client:** Uses the generic OpenAI SDK (not Azure-specific), connecting via `base_url` that appends `/v1`. This supports API gateways like NewAPI.

## Copilot Skill

The `.github/skills/poe-generation/` directory defines a Copilot skill for document generation. See `SKILL.md` for the full prompt protocol and `references/prompts.md` for system prompt templates.
