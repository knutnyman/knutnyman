# Average HbA1c of a Non-Insulin-Treated (NIT) Type 2 Diabetes Patient

*Compiled 2026-06-07. Evidence gathered from clinical trials, real-world registries/surveys,
and CGM/market sources. Peer-reviewed figures were cross-checked against PubMed (DOIs linked).*

---

## TL;DR — there are two correct answers, pick by context

| Context | Headline A1c | Range | Use when… |
|---|---|---|---|
| **True population average** of *all* NIT T2D patients | **~7.2% (≈55 mmol/mol)** | 7.0–7.6% | describing the real-world NIT population (mostly oral-agent, reasonably controlled) |
| **Inadequately-controlled NIT** (the CGM/intervention target population) | **~8.6–8.8% (≈70–73 mmol/mol)** | 8.5–8.8% | describing trial-enrolled / device-candidate patients selected above an A1c floor |

The gap is entirely explained by **selection**: trials enrol patients above an A1c
threshold (commonly ≥7.5%), so their baselines overstate the true population mean.
A single defensible point estimate for a "typical" NIT patient is **~7.2% (55 mmol/mol)**;
diet/lifestyle-only patients run lower (~6.5–7.0%), and the uncontrolled tail that gets
into CGM trials runs ~8.6%+.

For comparison, **insulin-treated** T2D patients average **~8.6% (70 mmol/mol)** — roughly
1–1.5 percentage points higher than the typical NIT patient (confounded by disease stage,
not caused by insulin).

---

## 1. True population average (real-world registries & surveys)

Most large datasets report *% reaching an A1c target by treatment* rather than a clean
mean for the NIT subset, so the population mean below is a synthesis of consistent anchors.

| Source | Population | Key figure | N | Cite |
|---|---|---|---|---|
| NHANES 1999–2004 (Dodd et al., *Curr Med Res Opin* 2009) | US T2D, ~60% oral-only | **52.2% <7%**; metformin-monotherapy 62.2% <7% vs oral-triple 31.9% | — | [DOI](https://doi.org/10.1185/03007990902973300) |
| UK National Diabetes Audit (FY 2024/25) | Whole T2D register (NIT-dominated) | **64.7% ≤7.5% (≤58 mmol/mol)** | national | [GOV.UK Apr 2026](https://www.gov.uk/government/statistics/diabetes-profile-update-april-2026/diabetes-profile-statistical-commentary-april-2026) |
| Swedish NDR | Nationwide T2D, mostly primary care | Whole-T2D mean ~**6.7–7.0%**; ~45% above 52 mmol/mol target | >400k | [ndr.nu](https://www.ndr.nu) |
| German/Austrian DPV/DIVE | Real-world T2D | Median **~7.1–7.2%**; regional means 6.7–8.3% | 184,864 | PMID 30051605 |
| CPRD / MASTERMIND (England) | NIT patients needing 2nd/3rd agent | Selected to baseline **>7.5%** (uncontrolled subset) | 55,530 | [DOI](https://doi.org/10.1371/journal.pmed.1002942) |

**Synthesis:** NIT-dominated registry means cluster at **~6.7–7.2%**, with a right-skewed
tail (the ~35–50% not at target) pulling the mean toward the mid-7s. Best estimate
**≈7.2% (55 mmol/mol)**, range 7.0–7.6%.

*Caveat — confounding by indication:* the NIT-vs-insulin gap reflects disease severity and
treat-to-fail escalation, not a causal effect of therapy. Few sources publish a clean
NIT-only mean with SD; the Swedish NDR and German DPV **annual-report tables** are the best
primary source for exact per-treatment subgroup means (PDFs were not retrievable this session).

---

## 2. Inadequately-controlled NIT — CGM/intervention trial baselines

This is the number device makers cite. These trials enrol above an A1c floor, so baselines
are high *by design*.

| Trial (year) | Population (NIT confirmation) | Baseline A1c | N | Cite |
|---|---|---|---|---|
| **IMMEDIATE** (Aronson et al., 2023) — flagship NIT CGM RCT | T2D on ≥1 **non-insulin** therapy, A1c ≥7.5% | **8.6% (70 mmol/mol)** | 116 | [DOI](https://doi.org/10.1111/dom.14949) |
| **Dexcom CONNECT** (presented ADA 2026) — largest/newest NIT CGM RCT | T2D **not on insulin**, 22 US primary-care sites, A1c 7.1–14.9% | **8.8%** (31% ≥9.0%) | 283 | [Dexcom IR](https://investors.dexcom.com/news/news-details/2026/Dexcom-CONNECT-Study-The-Most-Significant-Clinical-Study-Demonstrating-CGM-Benefits-for-People-with-Type-2-Diabetes-Not-Using-Insulin/default.aspx) |
| **Wada et al.** (2020) — FGM vs SMBG, Japan (less-restrictive entry) | **Non-insulin** T2D | **7.83% / 7.84%** | 100 | [DOI](https://doi.org/10.1136/bmjdrc-2019-001115) |
| **Vigersky et al.** (2012) | T2D not on prandial insulin | ~8.7–9.1% | 100 | [PMC3609537](https://pmc.ncbi.nlm.nih.gov/articles/PMC3609537/) |
| **Martens et al.** (2025) | **Non-insulin, non-SU** T2D, A1c 7.5–12% | ~8.5% | 72 | [DOI](https://doi.org/10.1089/dia.2024.0579) |

**Synthesis:** Pivotal NIT CGM RCTs anchor at **8.6–8.8%**. CGM lowered A1c by **~0.3% (IMMEDIATE)
to ~0.9% (CONNECT)** vs control. The less-restrictive Wada trial (~7.8%) sits closer to a
"typical" NIT patient.

> ⚠️ **Common miscitation:** **MOBILE** (*JAMA* 2021, baseline 9.1%) and the **DIAMOND T2D cohort**
> are **basal/MDI-insulin-treated**, *not* NIT — exclude them from NIT figures despite their
> frequent appearance in "CGM in T2D" searches.

---

## 3. Pharma RCT baselines (oral / GLP-1 / SGLT2i / DPP-4 add-on)

Drug-intervention trials in non-insulin backgrounds cluster tightly.

| Trial | Non-insulin background | Baseline A1c | N |
|---|---|---|---|
| EMPA-REG MET (2014) | metformin only | **7.9%** | 637 |
| SUSTAIN 2 (2017) | metformin/TZD | **8.1%** | 1,225 |
| SUSTAIN 8 (2019) | metformin only | **8.3%** | 788 |
| DPP-4 add-on-to-metformin (e.g. sitagliptin, 2006) | metformin | **~8.0%** | 701 |
| AWARD-3 (2014) | monotherapy / drug-naive | **~7.6%** | 807 |
| LEAD-3 (liraglutide mono) | monotherapy | **8.3%** | ~746 |

**Synthesis:** Add-on-for-inadequate-control trials cluster **7.9–8.3% (≈63–67 mmol/mol)**;
monotherapy/drug-naive trials run lower (~7.6%). Best single trial estimate **~8.0–8.2%**.
Large CV-outcome trials (DECLARE ~8.3%, CANVAS ~8.2%) include *some* insulin users, so they
are not pure-NIT.

---

## 4. Guidelines & market framing

- **ADA Standards of Care 2025** newly recommends CGM **should be considered in adults with
  T2D on glucose-lowering medications other than insulin** — the formal anchor for the
  NIT-CGM thesis. It sets *individualized* A1c targets (generally <7% for many adults) but
  publishes **no single "average A1c for NIT T2D."**
  ([ADA 2025 release](https://diabetes.org/newsroom/press-releases/american-diabetes-association-releases-standards-care-diabetes-2025))
- **Market context:** ~38.4M US adults with diabetes; historically only ~3–4% of T2D used
  CGM (2022). FDA's 2024 OTC biosensor clearance (aimed at non-insulin users) is widely cited
  as expanding the addressable pool by tens of millions. Dexcom frames Type 2 non-insulin as
  its primary TAM-expansion segment and calls CONNECT the "first Level A evidence" for CGM in
  this group. No manufacturer states a specific NIT "average A1c" — they anchor on the trial
  baselines (8.6–8.8%).

---

## Recommended phrasing

- **Honest population statement:** *"The average non-insulin-treated T2D patient has an HbA1c
  of roughly 7.2% (55 mmol/mol), versus ~8.6% (70 mmol/mol) for insulin-treated patients."*
- **Device/market deck:** *"Inadequately-controlled non-insulin T2D — the CGM target population —
  has a baseline A1c of ~8.6–8.8% (IMMEDIATE, CONNECT); CGM lowers A1c by ~0.3–0.9% vs control."*

Both are defensible and sourced. Do not blend the two figures — state which population you mean.

---

## Caveats (read before quoting any number)

1. **Two populations, two numbers** — never quote the 8.6% trial baseline as the population average.
2. **Selection bias** inflates every trial baseline (entry floors of A1c ≥7.1–7.5%).
3. **"Non-insulin" is heterogeneous** — diet-only (~6.5–7.0%) < metformin-only < multi-oral/GLP-1.
4. **Confounding by indication** drives the NIT-vs-insulin gap (disease stage, not therapy).
5. **MOBILE / DIAMOND-T2D are insulin-treated** — exclude from NIT figures.
6. **CONNECT is press-sourced** (ADA 2026); reconfirm exact figures vs the peer-reviewed manuscript.
7. **No clean NIT-specific population A1c distribution** (NHANES-style %<7/7–8/8–9/>9) was
   verifiable; the population mean is a reasoned synthesis, not a single published figure.

---

## References

According to PubMed (DOIs below), the peer-reviewed anchors are:

1. Aronson R et al. **IMMEDIATE** — isCGM in non-insulin T2D. *Diabetes Obes Metab* 2023;25(4):1024–1031. [DOI](https://doi.org/10.1111/dom.14949) (PMID 36546594)
2. Wada E et al. Flash GM vs SMBG in non-insulin T2D. *BMJ Open Diab Res Care* 2020;8(1):e001115. [DOI](https://doi.org/10.1136/bmjdrc-2019-001115)
3. Martens TW et al. CGM to guide food choices, non-insulin T2D. *Diabetes Technol Ther* 2025;27:261–270. [DOI](https://doi.org/10.1089/dia.2024.0579)
4. Vigersky RA et al. rtCGM in non-insulin T2D. *Diabetes Care* 2012. [PMC3609537](https://pmc.ncbi.nlm.nih.gov/articles/PMC3609537/)
5. Dodd AH et al. Treatment approach & A1c control, NHANES 1999–2004. *Curr Med Res Opin* 2009. [DOI](https://doi.org/10.1185/03007990902973300)
6. Hankosky ER et al. A1c targets, insulin-treated, NHANES 2009–2020. *Diabetes Ther* 2023. [DOI](https://doi.org/10.1007/s13300-023-01399-0)
7. Venkatraman S et al. Glycemic control in insulin users, NHANES 1988–2020. *JAMA Netw Open* 2022. [DOI](https://doi.org/10.1001/jamanetworkopen.2022.47656)
8. Whyte MB et al. Glycaemic control disparities in T2D, England (CPRD). *PLoS Med* 2019. [DOI](https://doi.org/10.1371/journal.pmed.1002942)
9. Rawshani A et al. Newly-diagnosed T2D, Swedish NDR. *BMJ Open* 2015. [DOI](https://doi.org/10.1136/bmjopen-2015-007599)
10. Häring HU et al. EMPA-REG MET. *Diabetes Care* 2014. [link](https://diabetesjournals.org/care/article/37/6/1650/29568/)
11. Ahrén B et al. SUSTAIN 2. *Lancet Diabetes Endocrinol* 2017. [PMID 28385659](https://pubmed.ncbi.nlm.nih.gov/28385659/)
12. Lingvay I et al. SUSTAIN 8. *Lancet Diabetes Endocrinol* 2019.
13. Umpierrez G et al. AWARD-3. *Diabetes Care* 2014. [link](https://diabetesjournals.org/care/article/37/8/2168/29780/)
14. Dexcom CONNECT (ADA 2026) — [Dexcom IR](https://investors.dexcom.com/news/news-details/2026/Dexcom-CONNECT-Study-The-Most-Significant-Clinical-Study-Demonstrating-CGM-Benefits-for-People-with-Type-2-Diabetes-Not-Using-Insulin/default.aspx); [ADA release](https://diabetes.org/newsroom/press-releases/82-adults-type-2-diabetes-not-insulin-improve-blood-glucose-levels)
15. UK National Diabetes Audit — [digital.nhs.uk](https://digital.nhs.uk/data-and-information/publications/statistical/national-diabetes-audit-yt2); [GOV.UK Apr 2026](https://www.gov.uk/government/statistics/diabetes-profile-update-april-2026/diabetes-profile-statistical-commentary-april-2026)
16. ADA Standards of Care 2025 — [diabetes.org](https://diabetes.org/newsroom/press-releases/american-diabetes-association-releases-standards-care-diabetes-2025)

*Unit conversion: mmol/mol = (% − 2.15) × 10.929 (NGSP/DCCT ↔ IFCC).*
