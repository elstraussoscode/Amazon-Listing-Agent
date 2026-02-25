# Amazon Listing Agent

AI-powered Amazon listing optimization tool built on Amazon's **COSMO** (Common Sense Knowledge) and **RUFUS** (Shopping AI Assistant) principles. Generates SEO-optimized titles, descriptions, bullet points, and backend keywords with strict formatting, length enforcement, and multi-language support. Supports dynamic model selection from all available OpenAI GPT models.

---

## Table of Contents

- [Design Principles](#design-principles)
- [Architecture](#architecture)
- [Quick Start](#quick-start)
- [Features](#features)
- [Usage](#usage)
- [Content Generation Pipeline](#content-generation-pipeline)
- [Amazon COSMO & RUFUS](#amazon-cosmo--rufus)
- [Byte-Length Enforcement](#byte-length-enforcement)
- [Multi-Language Support](#multi-language-support)
- [POE Integration](#poe-integration)
- [File Structure](#file-structure)
- [Tech Stack](#tech-stack)
- [Troubleshooting](#troubleshooting)

---

## Design Principles

### 1. Context Over Keywords

Amazon's search has evolved beyond keyword matching. COSMO uses large language models and knowledge graphs to understand **meaning and context** behind search queries. This app generates content that answers the *intent* behind customer searches, not just keyword-stuffed text.

Every listing produced by this tool covers **15 semantic relationship types** that COSMO uses to evaluate relevance:

| # | Relationship | Question the Listing Answers |
|---|---|---|
| 1 | `is` | What is this product? |
| 2 | `has_property` | What properties does it have? (color, material, size, form) |
| 3 | `has_component` | What parts or accessories are included? |
| 4 | `used_for` | What is it used for? |
| 5 | `used_in` | In what situation or environment is it used? |
| 6 | `used_by` | Who typically uses it? |
| 7 | `used_with` | What other products is it combined with? |
| 8 | `made_of` | What material is it made of? |
| 9 | `has_quality` | What quality attributes does it have? |
| 10 | `has_brand` | What brand / brand message? |
| 11 | `has_style` | What style / look / design direction? |
| 12 | `targets_audience` | What audience is it designed for? |
| 13 | `associated_with` | What overarching themes is it associated with? |
| 14 | `has_certification` | Are there relevant certifications? |
| 15 | `enables_activity` | What activity does it enable? |

These relationship types are used **internally** to guide the AI. They must **never** appear in the generated output text.

### 2. No Truncation — AI-Driven Length Enforcement

Amazon enforces strict **byte limits** (not character limits) on listing fields. German umlauts (ä, ö, ü, ß) count as 2 bytes each. Instead of cutting text off mid-sentence, this app uses a dedicated LLM-powered function (`ensure_optimal_length_with_ai`) that iteratively asks the selected model to intelligently shorten or extend text to fit within the exact byte range — preserving complete sentences, technical terms, and product relevance. The function is field-type-aware: it knows whether it's adjusting a title (no periods), a bullet (HOOK: format), keywords (comma-separated only), or a description (complete sentences).

### 3. Structured Multi-Call Architecture

Content generation is split into **three sequential, focused LLM calls** per product to prevent rule conflicts:

1. **Call 1 — Main Content**: Generates title, description, and 5 bullet points
2. **Call 2 — Keywords**: Generates backend search terms, using Call 1 output as context to avoid repetition
3. **Call 3 — Verification**: Reviews all generated content against formatting, language, and structural rules; corrects only when necessary

This separation ensures the model can focus on one task at a time without mixing formats (e.g., comma-separated keywords vs. complete sentences in bullet points).

### 4. Strict Formatting Rules

Every output field follows rigid Amazon conventions:

- **Title**: `BRAND Series ProductType Components, Properties (USP, Material, Size, Color, Form) Synonym` — no sentences, no trailing periods, no dashes as separators (only commas), `100%` not `100 Prozent`, color directly after product name (not "Farbe blau")
- **Bullet Points**: `HOOK IN CAPS: Descriptive text in normal case.` — only the 2-4 word hook is uppercase, the rest is a complete sentence in normal case, no extra caps words in the body text after the colon, never truncated mid-sentence
- **Description**: Covers all 15 relationship types naturally without mentioning the technical terms
- **Keywords**: Comma-separated single words or 1-3 word phrases, all lowercase, no sentences, no periods anywhere, no terms already used in title/bullets/description

### 5. Target 85-90% of Maximum Allowed Length

Short listings waste SEO potential. Every generated field targets 85-90% of Amazon's maximum byte allowance:

| Field | Min Bytes | Max Bytes | Target |
|---|---|---|---|
| Title | 170 | 200 | ~185 bytes |
| Bullet Point (each) | 170 | 200 | ~185 bytes |
| Description | 1,700 | 2,000 | ~1,850 bytes |
| Keywords | 210 | 249 | ~230 bytes |

### 6. Language Purity

When a target language is selected, **100% of the output** must be in that language. The only exceptions are brand names (e.g., "Dreamfarm"), product series names (e.g., "Aveo"), and universal technical terms (e.g., "USB-C", "LED"). German product terms like "BPA-frei" become "BPA free" in English, "senza BPA" in Italian, etc. A mandatory translation table is built into the prompts.

### 7. Write Like a Human

The prompts explicitly instruct the model to avoid:
- Artificial word combinations
- Excessive adjective chains
- Marketing clichés ("perfect for", "ideal for")
- Invented features not present in the product data

Content must be **factual, direct, and based solely on the provided product data**.

### 8. Parallelized but Sequential per Product

Multiple products are processed in parallel (configurable 1-5 workers via `ThreadPoolExecutor`), but each individual product strictly follows the sequential Call 1 → Call 2 → Call 3 pipeline. Thread-safe result aggregation ensures correct ordering in the output file.

---

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                   Streamlit UI                          │
│  ┌──────────┐ ┌──────────┐ ┌────────┐ ┌─────────────┐  │
│  │ Language  │ │ POE CSV  │ │Product │ │Model Select │  │
│  │ Selector  │ │ Upload   │ │ Excel  │ │& Prompts    │  │
│  └────┬─────┘ └────┬─────┘ └───┬────┘ └──────┬──────┘  │
│       │             │           │              │         │
│       └─────────────┼───────────┼──────────────┘         │
│                     ▼           ▼                        │
│              ┌──────────────────────┐                    │
│              │  ThreadPoolExecutor  │                    │
│              │  (1-5 workers)       │                    │
│              └──────┬───────────────┘                    │
│                     │                                    │
│    ┌────────────────┼────────────────┐                   │
│    ▼                ▼                ▼                   │
│ ┌──────┐      ┌──────┐         ┌──────┐                 │
│ │Prod 1│      │Prod 2│         │Prod N│                 │
│ └──┬───┘      └──┬───┘         └──┬───┘                 │
│    │              │                │                     │
│    ▼              ▼                ▼                     │
│ ┌──────────────────────────────────────┐                 │
│ │  Per Product (Sequential):           │                 │
│ │  Call 1: MainContent (selected model)│                 │
│ │    → ensure_optimal_length_with_ai   │                 │
│ │  Call 2: KeywordContent              │                 │
│ │    → ensure_optimal_length_with_ai   │                 │
│ │  Call 3: VerificationResult          │                 │
│ │    → ensure_optimal_length_with_ai   │                 │
│ │       (only for modified fields)     │                 │
│ └──────────────────────────────────────┘                 │
│                     │                                    │
│                     ▼                                    │
│              ┌──────────────┐                            │
│              │ Output XLSX  │                            │
│              └──────────────┘                            │
└─────────────────────────────────────────────────────────┘
```

---

## Quick Start

```bash
cd "/Users/florianstrauss/Desktop/Amazon Listing Agent"
source venv/bin/activate
streamlit run app_production.py
```

The browser opens automatically at **http://localhost:8501**.

---

## Features

- **COSMO/RUFUS Optimization** — Content covers all 15 semantic relationship types
- **Dynamic Model Selection** — Choose from all available OpenAI GPT models (auto-detected from your API key) or use the default list
- **Structured Output** — Pydantic models with JSON schema enforce response structure
- **3-Call Pipeline** — Main content → Keywords → Verification, each with dedicated prompts
- **AI-Driven Length Enforcement** — No truncation; LLM iteratively adjusts text to exact byte ranges (up to 7 retries), field-type-aware
- **Multi-Language Support** — German, English (UK/US), French, Italian, Spanish, Dutch, Polish, Swedish
- **POE Integration** — Upload Amazon Product Opportunity Explorer CSV to inform keyword generation
- **Parallel Processing** — Process up to 5 products simultaneously with per-product error isolation
- **Resilient Processing** — Individual product failures don't abort the batch; partial results are always available
- **Editable Prompts** — Fine-tune all prompts directly in the sidebar
- **XLSX Output** — Download results with identifier, old title, new title, description, 5 bullets, keywords

---

## Usage

1. **Enter API Key** — Paste your OpenAI API key in the sidebar (saves and auto-loads available models)
2. **Select Model** — Choose the GPT model from the dropdown (auto-populated from your API key's available models)
3. **Select Language** — Choose the output language for all generated content
4. **Upload POE Data** (optional) — Upload a CSV from Amazon's Product Opportunity Explorer
5. **Upload Product Data** — Excel file with your product information
6. **Configure** — Select ID column, old title column, number of products, parallelism
7. **Start** — Click "Optimierung Starten" and wait for processing
8. **Download** — Download the generated XLSX file

---

## Content Generation Pipeline

### Call 1: Main Content (`MainContent`)

Generates the product title, description, and 5 bullet points. The prompt includes:
- Full product data from the uploaded Excel
- Title structure rules (brand → series → product type → components → properties → synonym)
- Bullet point format (`HOOK: Description.`)
- Description requirements (cover 15 relationship types naturally)
- Industry-agnostic examples (kitchen, sports, electronics, baby, beauty)
- Byte targets for each field

After generation, each field passes through `ensure_optimal_length_with_ai` if outside its target byte range.

### Call 2: Keywords (`KeywordContent`)

Generates backend search terms. Receives the full output of Call 1 as context to **avoid repeating** any term already used in the title, bullets, or description. Incorporates POE data if uploaded. Rules enforce:
- Comma-separated single words or 2-3 word phrases only
- No sentences, no descriptions, no periods
- Synonyms, spelling variants, long-tail keywords, application-based terms
- No complementary products, no competitor brands

### Call 3: Verification (`VerificationResult`)

Reviews all content from Calls 1 and 2 against:
- Language purity (no German words in non-German output)
- Title format (no sentences, no trailing periods, correct symbol usage)
- Bullet format (HOOK in caps, colon, normal-case body, complete sentences)
- Keyword format (comma-separated, no sentences)

Returns `approved: true` if everything passes, or corrected content with `issues_found` if corrections were needed. Modified fields go through `ensure_optimal_length_with_ai` again.

---

## Amazon COSMO & RUFUS

### COSMO (Common Sense Knowledge)

Amazon's AI relevance engine that goes beyond keyword matching. It uses large language models and knowledge graphs to understand the **meaning and context** behind search queries. COSMO evaluates listings based on 15 semantic relationship types and bridges the "semantic gap" between what customers search for and what product data contains.

Listings that only contain generic keywords without contextual information (who uses it, when, where, with what) risk losing visibility in COSMO-powered search results.

### RUFUS (Shopping AI Assistant)

Amazon's customer-facing AI shopping assistant. RUFUS sits on top of COSMO and uses its knowledge graph to recommend products in natural language conversations. It pulls from product listings, reviews, Q&A sections, and structured attributes.

Products that don't provide sufficient contextual information simply won't appear in RUFUS-powered recommendations — regardless of their keyword coverage or ad spend.

### Implications for Listing Optimization

- Content must answer customer **intent**, not just match keywords
- Every listing should cover the 5 W-questions: Who uses it? What does it do? When/where is it used? How? Why is it better?
- All backend attribute fields should be filled (material, dimensions, intended use, target audience, occasion)
- Lifestyle images showing the product in use with the target audience improve AI image understanding
- Customer reviews and Q&A content feed into the knowledge graph

---

## Byte-Length Enforcement

Amazon measures field lengths in **bytes**, not characters. This matters for languages with multi-byte characters:

| Character | Bytes |
|---|---|
| a-z, A-Z, 0-9 | 1 byte |
| ä, ö, ü, ß | 2 bytes |
| €, é, è, ê | 2-3 bytes |

The `ensure_optimal_length_with_ai` function:
1. Checks if text is within `[min_bytes, max_bytes]`
2. If not, calculates the direction (too short / too long) and difference
3. Includes **field-type context** (title: no periods; bullet: HOOK format; keywords: comma-separated only; description: complete sentences)
4. Calls the selected model with the current text, target range, field type hint, and product context
5. Validates the result; repeats up to 7 times if still out of range, with exponential backoff on API errors
6. Never truncates — always uses the LLM to intelligently adjust content

---

## Multi-Language Support

| Language | Code |
|---|---|
| Deutsch | `de` |
| English (UK) | `en-GB` |
| English (US) | `en-US` |
| Français | `fr` |
| Italiano | `it` |
| Español | `es` |
| Nederlands | `nl` |
| Polski | `pl` |
| Svenska | `sv` |

The selected language applies to **all** generated fields. A built-in translation table ensures common product terms are correctly translated (e.g., "BPA-frei" → "BPA free" / "senza BPA" / "sans BPA" / "libre de BPA").

---

## POE Integration

Upload a CSV export from Amazon's **Product Opportunity Explorer** to enrich keyword generation with real search volume data. The app:

1. Auto-detects the header row and column names (supports German and English headers)
2. Extracts up to 20 search terms with their search volumes
3. Displays a preview table in the UI
4. Passes the top 15 terms to Call 2 as keyword inspiration (avoiding duplication with existing content)

---

## File Structure

| File | Purpose |
|---|---|
| `app_production.py` | Main Streamlit app — Content Optimizer with 3-call pipeline |
| `app_template_filler.py` | Template Filler module — fills Amazon upload templates with AI content |
| `dynamic_template_analyzer.py` | Parses Amazon templates (XML/Flat File), detects product types and required fields |
| `test_prompts.py` | Standalone test script for debugging prompt outputs and formatting |
| `requirements.txt` | Python dependencies |
| `.env` | OpenAI API key (not committed) |
| `data/` | Test data and template files |

---

## Tech Stack

| Component | Technology |
|---|---|
| UI | Streamlit |
| LLM | OpenAI GPT models (dynamic selection via `client.models.list()`, default: GPT-5.1) |
| Structured Output | Pydantic models → JSON Schema |
| Data Processing | Pandas, Openpyxl |
| Parallelism | `concurrent.futures.ThreadPoolExecutor` |
| Thread Safety | `threading.Lock` |
| Environment | python-dotenv |

---

## Troubleshooting

**API Key Error**
- Verify the key is correct and has sufficient credits
- The selected model must be accessible with your API key

**Model dropdown shows fallback list**
- If the API cannot be reached when saving the key (e.g., network issue), a hardcoded list of known models is shown
- Re-save the API key to retry loading available models

**Byte lengths not in range after 7 retries**
- Check the terminal logs for `⚠️ ... konnte nicht in Zielbereich gebracht werden`
- The text is still usable but may be slightly outside the target range
- Consider adjusting the prompt to guide initial generation closer to the target

**Optimization aborts or individual products fail**
- Individual product failures are isolated and do not abort the batch
- Partial results are always available for download even if some products fail
- Check the terminal logs for specific error messages per product

**Language mixing in output**
- The verification call (Call 3) catches most language issues
- If persistent, check that the product data Excel doesn't contain German-only fields that the model copies verbatim

**POE CSV not loading**
- Ensure the CSV header contains "Suchbegriff" or "Search" in a column name
- The app skips metadata rows automatically; the actual data header must contain a comma (CSV format)

**Template Filler errors**
- Verify the template is `.xlsm` or `.xlsx` format
- The template must contain an "AttributePTDMAP" sheet for product type detection

---

## Costs

| Operation | Approximate Cost |
|---|---|
| Template analysis | Free (no API calls) |
| Content per product | ~3 API calls to the selected model + length adjustment calls |
| Length adjustment | Up to 7 additional calls per field if needed |

Actual costs depend on text length, selected model, and number of length adjustment iterations.
