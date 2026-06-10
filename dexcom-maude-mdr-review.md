# A Detailed Review of FDA MAUDE Medical Device Reports (MDRs) for Dexcom Over Time

*Prepared June 10, 2026. Scope: Dexcom continuous glucose monitoring (CGM) devices — G4 Platinum, G5 Mobile, G6, G7, Dexcom ONE/ONE+, and Stelo — in FDA's Manufacturer and User Facility Device Experience (MAUDE) database, roughly 2015 through mid-2026.*

---

## ⚠️ Read this first: what this review is and how reliable it is

**Data-access limitation.** This environment blocks direct access to FDA's systems — `api.fda.gov` (the openFDA device-event API), `accessdata.fda.gov` (the MAUDE web search and recall database), and `fda.gov` all return HTTP 403. I therefore could **not** run live MAUDE queries or pull exact `count=date_received` time series myself. Every number below comes from **secondary sources**: peer-reviewed analyses of MAUDE (primarily Jan Krouwer's series in the *Journal of Diabetes Science and Technology*), FDA content as reported by trade press (MedTech Dive, Fierce Biotech, MassDevice, Cardiovascular Business), KFF Health News' investigative work, Dexcom SEC filings, the short-seller-affiliated outlet Hunterbrook, and plaintiff law-firm summaries. Treat single-sourced figures as **indicative, not audited**, and re-verify the load-bearing numbers against primary MAUDE before citing them anywhere consequential.

**The single most important caveat — MAUDE counts are not safety rates.** FDA itself states that MAUDE cannot be used to determine the incidence or prevalence of an event, to compare rates across devices, or to evaluate changes in rates over time, because of underreporting, duplicate and unverified reports, no confirmation that the device caused the event, and an unknown denominator (how many devices are in use). A report is explicitly **not an admission that the device caused the harm** (21 CFR 803.16). So any "Dexcom MDRs rose X%" statement conflates *reporting volume* with *risk* and must be read alongside the distortions catalogued in Section 6.

---

## 1. The big picture — Dexcom is one of the highest-volume reporters in all of MAUDE

CGMs as a category have become one of the single largest contributors to MAUDE, and Dexcom is a leading driver of that volume:

- **Dexcom G6 alone generated ~125,753 MAUDE reports in 2019** — a single device, a single year. About **98.9% were malfunctions** and only ~1.1% (~1,383) were classified as injuries (Krouwer, *JDST*, analysis of 2019 FDA adverse events).
- **Cumulatively, FDA has received on the order of ~500,000 MAUDE reports for the Dexcom G6** since its 2018 authorization through ~2024 (figure cited in legal/news coverage; order-of-magnitude consistent with the 2019 single-year count).
- Across the **four major CGM families** (Dexcom G6, all Abbott FreeStyle Libre versions, Medtronic Guardian 3, Senseonics Eversense), MAUDE logged **281,963 adverse events in 2022** — broken down as **268,310 malfunctions (~95%), 13,644 injuries, and 9 deaths** (Krouwer, *JDST* 2025).
- One analysis cited in trade press found **1,624,664 CGM adverse-event reports between January 2019 and December 2024**, the most common single problem being a "wireless communication problem" (536,306 reports).

For scale: total MAUDE volume across *all* device types peaked at **~3.1 million reports in 2022** (Mishali et al., *Digital Health* 2025). CGMs at ~282k that year therefore represent on the order of **~9% of the entire database from just four product families** — which is why glucose sensors are routinely described as among the most-reported device types in MAUDE today.

> **Caveat / gap:** I could **not** isolate a clean Dexcom-only subtotal within the 2022 all-CGM figure, nor find an accessible source that formally ranks glucose sensors as the literal #1 device type for a specific recent year. The "#1 / top device type" framing is well-supported qualitatively and by the ASR-era data (Section 6), but treat the precise ranking as inferred.

---

## 2. Trend over time and event-type mix

**The shape of the trend.** Raw Dexcom/CGM MDR counts rose steeply across the window, with two structural inflection points that are largely *reporting-driven*, not necessarily safety-driven:

| Year(s) | What the data shows | Caveat |
|---|---|---|
| ~2015 | Counts step up | Mandatory electronic reporting (eMDR) took effect Aug 14, 2015; voluntary reports "spiked notably in 2015" |
| 2018 | Dexcom G6 launches (mid-2018); VMSR summary-reporting program begins (Aug 2018) | New high-volume device + new reporting regime |
| 2018→2019 | Diabetes-device adverse events **"almost doubled"** | Coincides with FDA shutting down the hidden Alternative Summary Reporting (ASR) program in mid-2019 — previously-summarized events flowed into public MAUDE (Section 6) |
| 2019 | G6 ~125,753 reports | ~99% malfunctions |
| 2022 | Four-CGM total 281,963 | ~95% malfunctions, 13,644 injuries, 9 deaths |
| 2023→2026 | G7 era; recalls and warning letter drive attention and reporting | Litigation- and recall-publicity-driven reporting surges |

**Event-type mix is overwhelmingly "malfunction."** Across every year with data, ~95–99% of Dexcom/CGM reports are malfunctions (no reading, inaccurate reading, signal loss), a small share are injuries, and deaths are a very small absolute number (single digits per year for G6; see Section 3).

---

## 3. Death reports — small numbers, easily conflated, causality not established

Lay and legal coverage frequently mixes three *different* death tallies. Keeping them separate:

**Dexcom G6 death reports by year (MAUDE), per Drugwatch/Motley Rice summaries citing MAUDE:**

| Year | G6 death reports |
|---|---|
| 2019 | 5 |
| 2020 | 6 |
| 2021 | 3 |
| 2022 | 7 |
| 2023 | 8 |
| 2024 (Jan 1–Sep 30) | 13 |

**Dexcom G7 death reports (separate device, launched 2023):** Hunterbrook reported **at least 13 G7-associated deaths in MAUDE since the 2023 launch**, then **3 additional** in filings dated September 2025 (disclosed Oct 28, 2025).

> **Three critical caveats on deaths:**
> 1. **These are reported *associations*, not adjudicated causation.** Under 21 CFR Part 803, a manufacturer must file a report when information "reasonably suggests" a device "may have caused or contributed to" a death — certainty is not required, and the report is not an admission of fault. For a device worn 24/7 by millions of people with diabetes (a population with elevated baseline mortality), background deaths of users generate "may have contributed" MDRs regardless of actual device fault.
> 2. **Counts can also be *under*-stated:** a 2021 *JAMA Internal Medicine* study found ~23% of MAUDE reports that *described* a death were not coded as "death" (filed as injury/malfunction/other).
> 3. **Hunterbrook is short-seller-affiliated** (it disclosed a $DXCM short position); its tallies should be checked against primary MAUDE. Notably, FDA reported **zero deaths** tied to the 2025 receiver recall — so different sources counting different things produce very different death numbers.

---

## 4. Failure modes — what actually goes wrong, by generation

### Most common DEVICE problems (across generations)
1. **Failure to obtain a result / no readings / signal loss** — repeatedly the most common *malfunction* effect. "Sensor failed," connectivity drops, sensors lasting only 1–2 days.
2. **Inaccurate readings** — the top *user complaint* for Dexcom (Krouwer 2020, 2023). See generational split below.
3. **Failure to alert / alarm** — missed urgent low/high alarms; receiver speaker not sounding; app not delivering notifications. This is the failure mode behind the major Class I recalls.
4. **Adhesive / skin issues** — irritation, rash, dermatitis; narratives indicate this worsened after Dexcom changed the G6 adhesive supplier around October 2019.
5. **Sensor wire fracture / retained sensor wire** — narratives of the sensor filament breaking off under the skin.
6. **App / software failures** — app crashes or shuts down so no values or alarms are delivered.

### Most common PATIENT problems
Severe **hypoglycemia** (loss of consciousness, seizures), severe **hyperglycemia / DKA**, seizures, vomiting, **skin irritation/rash**, and **death**. FDA's framing of missed-alert harm is "severe hypoglycemia, severe hyperglycemia, diabetic ketoacidosis and death."

### By generation
- **G4 Platinum / G5 Mobile:** Predate the device-specific published analyses. The G5/early-app era is the backdrop for the original smartphone-alarm-failure concerns. (Consumer Reports noted that in the year after an FDA warning letter, ≥10 deaths and 409 serious injuries were reported to MAUDE for the **G4** — raw counts, not rates.)
- **G6:** The anchor device in the peer-reviewed analyses. Dominant patterns: **failure to obtain a result** and **inaccuracy complaints** (~6% of G6 reports per Hunterbrook's tabulation). Known issues: **compression lows** (false lows from lying on the sensor), adhesive/skin reactions, early sensor failures.
- **G7 (US launch 2023):** **Accuracy complaints ≈13% of all G7 reports — roughly double the ~6% G6 rate** (Hunterbrook). Accuracy reports reportedly **spiked for sensors made after December 2023**, coinciding with the unauthorized resistance-layer coating change later cited in FDA's 2025 warning letter (Section 5). G7 reported MARD ~8.2% vs Abbott Libre 3 ~7.9% (manufacturer studies).
- **Stelo (OTC, cleared March 2024, launched Aug 2024):** **No MAUDE-based failure-mode analysis found** — a genuine gap. Only individual reports exist; no recall.

> **Important nuance on "inaccuracy":** Krouwer's 2020 analysis plotted Dexcom CGM-vs-meter discrepancies on a Parkes error grid and found *most fell in Zone B* — i.e., clinically non-dangerous. So many "inaccuracy" complaints, at least in the G6 era, overstated real clinical risk. The 2024–25 G7 framing (tied to the coating change) describes a potentially different, more serious accuracy problem. **Don't conflate 2019-era G6 inaccuracy with 2024-era G7 inaccuracy** — different devices, different time periods, possibly different severity.

---

## 5. Related FDA enforcement actions

### Recalls
- **2016 — Class I, G4 Platinum / G5 Mobile receivers.** Audible-alarm/speaker failure. **263,520 units** (made July 2011–March 2016). FDA: failure could mean missed severe hypo/hyperglycemia alerts, with rare risk of death.
- **May–June 2025 — Class I, G6/G7/ONE/ONE+ receivers.** Speaker loses contact with the circuit board → no audible alarms (vibration/visual still work; **app users unaffected**). **>703,000 devices** (~36,800 G6; >602,000 G7; >38,000 ONE; >26,000 ONE+). Dexcom reported **112 complaints and 56 serious adverse events worldwide (seizures, loss of consciousness, vomiting), no deaths**; reported incidence 0.015%. Free replacement.
- **Aug–Sept 2025 — Class I, G6/G6 Pro Android app v1.15.0.** App could shut down unexpectedly → no values/alarms. No serious injuries/deaths reported; fixed via mandatory update.
- **Sept 2025 / Class I flagged ~Jan 2026 — G7 and ONE+ apps.** Missed "Sensor Failed" alert on transmitter hardware/firmware failure.
- **May 2026 — Class II, G7/ONE/ONE+ iOS & watchOS apps.** Software defect can delay or replay glucose readings/alerts.

> **Note on the task's "2024 G7 receiver recall":** The major Class I G7 receiver recall is actually dated **May–June 2025**, not 2024. Early 2024 saw only G6 receiver *correction* notices (Jan–Feb 2024) for delayed/missed alerts. If you were working from a 2024 date, it's likely either those corrections or a year mismatch.

### March 2025 FDA Warning Letter (MARCS-CMS 700835, dated March 4, 2025)
- Based on inspections of **Mesa, AZ (June 2024)** and **San Diego, CA (Oct–Nov 2024)**.
- Key finding: an **unauthorized design change to the resistance-layer coating** of G6/G7 sensors — Dexcom began producing the component **in-house (~Dec 13, 2023)** for supply-chain redundancy **without a new 510(k)**, rendering the devices misbranded; FDA said the changed sensors were **less accurate** and posed higher risk for insulin-dosing users. Also cited: deficient glucose/acetaminophen testing procedures and inadequate handling of a 2024 dissolved-oxygen deficiency in G6 sensors.
- Dexcom's response: disclosed in an 8-K; said the letter imposes **no manufacturing/marketing restriction and requires no recall**, expected no material impact, and **publicly rejected** the "unauthorized changes" characterization.
- Market reaction: shares fell on the disclosure and again when FDA published the letter (securities suits cite ~−9% and ~−4% on the respective dates; MedTech Dive reported a smaller ~1.9% dip on posting day).
- Status: FDA **cleared the G7 15-Day sensor in April 2025**, easing investor concern; **no public warning-letter closeout** confirmed as of mid-2026.

### Safety communication
- **Feb 5, 2025** (industry-wide, not Dexcom-specific): FDA warned that **smartphone-connected diabetes devices can miss critical alerts** due to phone settings, OS updates, do-not-disturb/deep-sleep, etc., with potential for severe hypo/hyperglycemia, DKA, and death.
- (A widely-assumed "2019 FDA safety communication about Dexcom app alerts" did **not** turn up as a formal communication; the 2019 event was a multi-day Dexcom Follow/Share **server outage** where FDA said it was "working with Dexcom.")

---

## 6. Why the time trend is treacherous — interpretation caveats

This is the analytical heart of the review. Several artifacts mean **raw MDR growth ≠ worsening safety**:

1. **The ASR → VMSR reporting-regime change (the biggest distortion).** For two decades FDA ran a hidden **Alternative Summary Reporting** program letting manufacturers file private quarterly summaries instead of public MAUDE reports. KFF Health News exposed it; FDA ended ASR in mid-2019 and released **>5.7 million** previously-hidden incidents. Strikingly, **blood-glucose meters were the single most-reported device type, ~2.4 million reports over ~20 years** (mostly LifeScan) — diabetes-device volume was massively under-represented in public MAUDE before 2019. When ASR ended (and the new **Voluntary Malfunction Summary Reporting** program began in Aug 2018), previously-hidden/summarized events surfaced in public MAUDE — which is a large part of why diabetes-device counts "almost doubled" from 2018 to 2019.

2. **VMSR records can each represent *many* events.** A single MAUDE summary entry can stand for dozens, hundreds, or thousands of malfunctions (the narrative must state the count). So counting *records* understates events, while counting *events* can jump discontinuously when summary reporting starts. (Trend studies even found data-entry quirks like "123" placeholder event counts.)

3. **The "may have contributed" causality standard inflates death tallies** for an always-on device used by a high-baseline-mortality population — partly offset by the ~23% death-miscategorization problem (Section 3).

4. **Install-base growth swamps everything.** Dexcom went from **~270,000 users (end of 2017)** to **~1.7 million (end of 2022)** to **~3.5 million (end of 2025)** — roughly a **10–13× increase**. A 10× rise in raw MDRs over that span could occur with *flat or improving* per-user safety. **Reports-per-user, not raw counts, is the only defensible cross-time or cross-brand metric** — and FDA's data doesn't supply the denominator.

5. **Other artifacts:** the 2015 eMDR mandate (step-up in volume), litigation- and recall-publicity-driven reporting surges, late/retrospective batch filings (a BMJ 2025 study found only ~71% of reports filed within 30 days; >400,000 filed >6 months late), product-code migration (G6/G7 use the integrated-CGM code **QBJ**, distinct from older glucose-monitor codes like MDS/PQF — long-window searches on one code miss migrated devices), and manufacturer-name spelling variants in MAUDE's free-text fields.

---

## 7. How Dexcom compares to peers

- **Glucose sensors / CGMs are now among the highest-volume device categories in MAUDE** (~9% of all 2022 reports from four families).
- A cross-sectional study of *late* MDRs put **Abbott (glucose monitors), Medtronic (insulin pumps), and Dexcom (glucose monitors)** all in the top 10 of both manufacturers and devices — confirming all three diabetes makers are top-tier reporters (this ranks *late* reports, not total volume).
- **Normalized comparison (Dexcom vs Abbott):** Hunterbrook's market-share-adjusted analysis (MAUDE through Dec 31, 2024) found **Dexcom's share of accuracy complaints runs ~22% *above* its market share, while Abbott's runs ~68% *below***. On a per-share basis, Dexcom generated disproportionately more accuracy MDRs than Abbott (caveat: adversarial short-seller source).
- **Insulet (Omnipod)** has historically been one of the highest-volume MAUDE reporters; a 2026 Omnipod 5 pod correction covered ~7 million pods with 476 serious-injury reports.
- **Context:** Consumer Reports' Jan 2019–Jul 2020 review found ~400 deaths and ~66,000 injuries across all common diabetes devices (Abbott, Dexcom, Medtronic, Tandem) and concluded "no one product stood out."

---

## 8. Bottom line

1. **Dexcom is one of the most-reported manufacturers in MAUDE**, driven overwhelmingly by **malfunction** reports (no reading / inaccurate reading / signal loss), with injuries a small fraction and deaths in the single-to-low-double digits per year per device generation.
2. **The raw upward trend is real but largely explained by non-safety factors** — a ~10–13× user-base increase, the 2015 eMDR mandate, and especially the 2018–2019 ASR→VMSR reporting-regime change that surfaced previously-hidden diabetes-device volume. Normalize by install base before inferring anything about safety.
3. **Genuine, FDA-recognized signals do exist:** repeated **alarm/alert-failure** problems (the 2016 and 2025 Class I receiver recalls; multiple 2025–26 app recalls) and the **2025 warning letter** over an unauthorized G7/G6 sensor coating change tied to degraded accuracy. G7-era accuracy complaints rose relative to G6.
4. **Interpret death and percentage figures cautiously** — "may have contributed" reporting, death miscategorization, adversarial sourcing (Hunterbrook), and lay/legal conflation of G6 vs G7 vs cumulative tallies all distort the picture.

---

## Appendix: Key sources

**Peer-reviewed (MAUDE analyses):**
- Krouwer JS, "Adverse Event Causes From 2022 for Four Continuous Glucose Monitors," *J Diabetes Sci Technol* 2025 — https://pmc.ncbi.nlm.nih.gov/articles/PMC11688675/ · https://pubmed.ncbi.nlm.nih.gov/37264590/
- Krouwer JS, "An Analysis of 2019 FDA Adverse Events for Two Insulin Pumps and Two CGMs," *JDST* — https://pmc.ncbi.nlm.nih.gov/articles/PMC8875056/ · https://pubmed.ncbi.nlm.nih.gov/32880188/
- Krouwer JS, "Adverse Event Data for Years 2018 to 2020 for Diabetes Devices," *JDST* — https://pmc.ncbi.nlm.nih.gov/articles/PMC9445356/
- Mishali et al., "Evaluation of reporting trends in the MAUDE Database: 1991 to 2022," *Digital Health* 2025 — https://pmc.ncbi.nlm.nih.gov/articles/PMC11755539/
- "Late adverse event reporting from medical device manufacturers to the US FDA," BMJ 2025 — https://pmc.ncbi.nlm.nih.gov/articles/PMC11898541/
- "Miscategorization of Deaths in the US FDA Adverse Events Database," *JAMA Intern Med* 2021 — https://pmc.ncbi.nlm.nih.gov/articles/PMC6784806/

**FDA / regulatory (accessed via secondary reporting; primary domains were blocked):**
- About MAUDE — https://www.fda.gov/medical-devices/mandatory-reporting-requirements-manufacturers-importers-and-device-user-facilities/about-manufacturer-and-user-facility-device-experience-maude-database
- 21 CFR Part 803 — https://www.ecfr.gov/current/title-21/chapter-I/subchapter-H/part-803 · §803.16 — https://www.law.cornell.edu/cfr/text/21/803.16
- VMSR program — https://www.fda.gov/medical-devices/medical-device-reporting-mdr-how-report-medical-device-problems/voluntary-malfunction-summary-reporting-program
- Dexcom warning letter (700835, 03/04/2025) — https://www.fda.gov/inspections-compliance-enforcement-and-criminal-investigations/warning-letters/dexcom-inc-700835-03042025
- 2025 receiver recall — https://www.fda.gov/medical-devices/medical-device-recalls-and-early-alerts/continuous-glucose-monitor-receiver-recall-dexcom-inc-removes-certain-dexcom-g6-g7-one-and-one
- Feb 2025 safety communication — https://www.fda.gov/news-events/press-announcements/fda-alerts-patients-potential-miss-critical-safety-alerts-due-phone-settings-when-using-smartphone

**Investigative / trade press / filings:**
- KFF Health News, "Five Things We Found In The FDA's Hidden Device Database" — https://kffhealthnews.org/news/five-things-we-found-in-the-fdas-hidden-device-database/
- MedTech Dive (receiver recall) — https://www.medtechdive.com/news/dexcom-recall-cgm-receivers-lack-alarm/751051/ · (warning letter) — https://www.medtechdive.com/news/dexcom-warning-letter-cgm-coating-change/743597/
- Hunterbrook, "Dexcom's Fatal Flaws" — https://hntrbrk.com/dexcom/ · follow-up — https://hntrbrk.com/dexcom2/ *(short-seller-affiliated)*
- Consumer Reports, "When Diabetes Devices Fail" — https://www.consumerreports.org/health/diabetes/when-diabetes-devices-fail-a2408992822/
- AJMC (2016 recall) — https://www.ajmc.com/view/dexcom-recalls-g4-platinum-and-g5-mobile-cgm-receivers-due-to-audible-alarm-failure
- Dexcom 8-K (warning letter) — https://www.sec.gov/Archives/edgar/data/0001093557/000109355725000047/dxcm-20250304.htm

*Generated via a multi-agent deep-research workflow. Because live FDA/MAUDE access was blocked in this environment, all figures are secondary and flagged accordingly; verify load-bearing numbers against primary MAUDE/openFDA before relying on them.*
