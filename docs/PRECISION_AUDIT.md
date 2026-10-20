# Pre-change audit — 2026-10-20

The working tree already contained uncommitted changes and 21 tests. These are preserved in upgrade_backup; this audit describes that starting point.

Architecture: bot download -> DocumentGuard working copy -> checker.analyze/analyze_pdf/analyze_text -> rule engines -> deduplication -> text report -> automatic Word annotation. CLI checker.py also exports Word. web_runner.py/Procfile/Dockerfile/service files deploy the bot. Dependencies: python-docx, requests, PyMuPDF, pypdf. Available runtime: Python 3.14; baseline 21/21 unittest tests passed. Existing 26-case benchmark failed two cases (missing space after Persian full stop; a honest). Its success criterion incorrectly accepted any unrelated warning.

P0: annotation permits output equal to input, mutates caller Document, rewrites runs containing nontext XML, and round-trips whole package. Guard can collide with caller-selected working directory. Bot always exports annotation rather than requiring opt-in.
P1: repeated targets choose first occurrence, mixed hyperlink offsets disagree with run offsets, PDF normalizes and joins source lines without mapping, merged protected spans lose evidence. Greek detector exists but is never called.
P2: article heuristics miss silent h; each/every + any s-ending word rejects valid nouns; cohesion infers weakness solely from missing markers; citation substring matching overclaims missing/uncited; university source has no verified edition. Report says certain errors and 100% preservation without sufficient evidence.
Security: extension-only dispatch, no ZIP expansion budget, exception tracebacks can expose paths/request secrets. A credential fallback is embedded in bot.py (value deliberately omitted here). No LLM/cloud analysis dependency is present; transport uses Bale and hosting remains outside local isolation.
Coverage: top-level paragraphs and top-level table cells only. Footnotes, endnotes, textboxes, nested tables and complex fields lack reliable full checking. PDF reading order/OCR are unverified.

Plan: isolate input/export; validate file signatures and ZIP budgets; conservative XML annotation; original evidence and strict reporting; protect scientific spans/quotes/configured terms; constrain high-risk rules; truthful benchmark and regression coverage. Keep unsupported regions visible and never manufacture positions.
