# CGM Medicare LCD / NCD — Scenario Analysis

Investment-diligence scenario analysis of the likely outcome of a new Medicare
**Local Coverage Determination (LCD)** or **National Coverage Determination (NCD)**
for continuous glucose monitors (CGMs). Prepared June 10, 2026.

## Deliverables

| File | Description |
|------|-------------|
| `CGM_LCD_NCD_Scenario_Analysis.docx` | Full written report (10 sections + 2 appendices): thesis, history, market, clinical evidence, precedent base rates, five probability-weighted scenarios, payment compression, signposts, and a tiered primary-diligence call plan. |
| `CGM_LCD_NCD_Scenario_Analysis.pptx` | 14-slide executive deck covering the same arc for presentation. |
| `build_word.py` / `build_ppt.py` | Generators (re-run to regenerate the documents). |

## Core thesis

CGM Medicare policy is moving on **two opposing tracks**:

- **Track A — Eligibility (expansionary):** an LCD revision extending coverage to
  non-insulin Type 2 diabetes is anticipated (draft expected H2 2026) but not yet filed.
- **Track B — Payment (restrictive):** competitive bidding + monthly-rental
  reclassification was **finalized Nov 28, 2025**, effective ≤ Jan 1, 2028.

## Probability-weighted scenarios (coverage-instrument outcome, ~24–36 mo)

| # | Scenario | Prob. |
|---|----------|-------|
| S1 | Conditioned LCD expansion to non-insulin T2D (modal) | 35% |
| S2 | Broad LCD expansion to all non-insulin T2D | 20% |
| S3 | Status quo persists (expansion slips) | 25% |
| S4 | Restriction-led (bidding + documentation/prior-auth) | 15% |
| S5 | NCD / CED national-action tail | 5% |

Some expansion (S1+S2) ≈ 55%. The Track B payment cut is a near-certain cross-cutting
modifier present in all scenarios.

## Regenerate

```bash
pip install python-docx python-pptx
python3 build_word.py
python3 build_ppt.py
```

## Sourcing & caveats

Built from a five-stream evidence sweep (CMS coverage database, Federal Register,
OIG/GAO, peer-reviewed clinical literature, manufacturer SEC filings/earnings calls,
health-policy trade press) with adversarial verification of load-bearing claims.
Probabilities are analyst judgments anchored to precedent base rates, not forecasts.
Several figures derive from manufacturer estimates or single analyst sources and are
flagged in the report. Not investment advice.
