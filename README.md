# Designing a zero-trust identity framework for a regional accounting firm

Fictional case: **Southern Ledger Partners** (~40 staff, three offices). Client tax and financial data; Microsoft 365 plus accounting SaaS; mix of office and remote work. Laboratory design only — no employer or client tenant.

## Research questions

1. Which zero-trust identity controls are practical and cost-justified for a professional-services SME?
2. How can conditional access and device trust be introduced without disrupting a small non-technical workforce?
3. What measurable identity-risk reduction can a tightly scoped pilot show?

## Deliverables

1. Identity and entitlement audit (synthetic current-state)
2. Zero-trust reference architecture and policy set (Free-tier honest + paid phase)
3. Phased roadmap with indicative costs
4. Pilot evaluation workbook, report, and slides

## Licence honesty

Entra ID Free cannot claim Conditional Access or Identity Protection. Phase 0 uses security defaults / MFA / group hygiene / break-glass. Phase 1 adds Conditional Access and device compliance when Business Premium / Entra P1 (or equivalent) is funded.
