# NIBR-metadata-pipeline
Multi-stage bibliographic metadata extraction pipeline (CrossRef → GROBID → LLM review) for 14,000+ academic papers

# NIBR Paper Metadata Extraction Pipeline

## Overview
A confidence-prioritized pipeline that automatically extracts bibliographic
metadata (author, corresponding author, year, title, journal, volume, page,
DOI) from 10,000+ academic papers, built during an internship at the
National Institute of Biological Resources (NIBR).

## Results
- **919 papers** (Korean Journal of Mycology): up to **98%** of
  machine-readable records resolved without manual review
- Applied across multiple journals with varying characteristics (e.g.,
  papers from the 1990s with no DOI, scanned issues) — resolution rates
  ranged from roughly **86% to 99%** depending on each journal's
  publication era and text quality

## Pipeline (funnel design: try the most reliable method first)
1. **CrossRef API** — direct DOI lookup, most accurate
2. **GROBID** — ML model trained on hundreds of thousands of papers, used
   when no DOI is found or the CrossRef lookup fails
3. **Regex fallback** — last resort, based on structural conventions
   (e.g., author names sitting just above an affiliation line)
4. **LLM review (Gemini API)** — re-checks only the fields flagged as
   low-confidence by steps 1–3, instructed to answer "cannot verify"
   rather than guess

## Key Bugs Found & Fixed
- False positives in corresponding-author detection (table headers,
  citation abbreviations, statistical software names being misread as
  author names)
- Font-encoding corruption (`(cid:19)`-style glyph IDs, and Unicode
  replacement characters) silently leaking into author/title fields
- A journal-name extraction gap that was masked by GROBID succeeding on
  author fields while leaving journal name blank — traced after a batch
  run showed journal name missing in 430/919 rows, then fixed
- Tree-structured (nested-folder) file paths breaking the LLM review
  step's ability to re-open source PDFs

## Setup
See [`setup/`](./setup) for environment setup (dependencies, GROBID
server, LLM review invocation).

## Note
Actual collected paper data and extracted metadata results are not
included in this repository. Only the pipeline logic is shared.
