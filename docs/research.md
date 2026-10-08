# Why WELL needs ongoing compliance management

Research date: 8 October 2026. This memo distinguishes document evidence, secondary claims and unverified official web guidance.

## Finding

WELL's health and wellbeing objectives depend on actual building operation, not only the initial design. Conditions and practices can change after certification. The uploaded project-specific feature requirements explicitly call for repeated monitoring, maintenance, surveys and reporting. An evidence calendar connects those activities with their reporting record, so recertification preparation is supported by contemporaneous evidence rather than reconstructed claims.

The supplied **WELL Ongoing Compliance Service Guide** is a service-line reference, with placeholder author details. It is not an IWBI-issued guidebook. Its three-year recertification statement and claims about certification consequences need confirmation against the project's agreement and current IWBI guidance. Missed evidence creates a review risk; it is not appropriate to claim that any missed task automatically revokes certification.

## Sources inspected

1. User-uploaded `WELL-Project-2202126030-Scorecard.xlsx`:
   - Matrix Summary G4: `#2202126030-20 Kildare Street`.
   - G5: WELL Core Certification.
   - G6: WELL v2, Q1 2021.
   - G7–G8: anticipated Gold, projected 61.5 points. These are targets, not certification awards.
   - Matrix Summary Y / ? / N columns and Digital Scorecard Status checked against one another.
   - Feature Requirements includes version-specific requirements, alternative paths, verification types and WELL Core guidance. Continuation rows were included; ignoring these rows would miss important recurring duties.
2. User-uploaded `WELL Ongoing Compliance Service Guide.pdf`, 19 part entries across six concepts. Part-level source page numbers and wording retained in `rules.json`; PDF and workbook retained locally in each imported project's sources folder.
3. Official verification destinations attempted through HTTPS:
   - https://resources.wellcertified.com/tools/well-certification-guidebook/
   - https://www.wellcertified.com/certification/v2/
   - https://v2.wellcertified.com/
   - https://well.support/

The external requests returned proxy `403 Forbidden`. No external page content was retrieved. Required domain additions were saved to the environment draft; they have not been proven applied. The following conclusions therefore use the attached documents, not a completed independent web review.

## Selected ongoing parts: eight

| Part | Activity evidenced by the scorecard | Submission / source detail |
|---|---|---|
| A01.5 | Pollutants in A01 Parts 1–3 monitored at intervals no longer than annually | Annual On-going Data Report; non-leased spaces; guide p. 4 |
| A12.1 | Filter replacement in accordance with manufacturer recommendations | Annual On-going Maintenance Report; whole building with stated tenant exclusions; guide p. 6 |
| W03.1 | Annual turbidity, pH, free chlorine; conditional coliform and parameters ≥80% of threshold | Annual On-going Data Report; non-leased spaces; guide p. 9 |
| W03.2 | Monitoring, corrective action and Legionella sampling where specified by the management plan | Annual On-going Maintenance Report; whole building; guide p. 10 |
| W07.3 | Mold/leak inspections at the plan's periodic inspection frequency | Annual On-going Maintenance Report; whole building; guide p. 12 |
| T01.2 | Selected annual-testing path: measurements once June–August and once December–February | Results submitted annually as On-going Data Report; non-leased spaces; guide p. 15 |
| T02.1 | Twice-yearly anonymous thermal-comfort surveys in the same seasonal windows, at least six months after occupancy | Technical Document (Individual); invite all regular occupants; scope may be limited to non-leased thermal conditions; guide p. 16 |
| C04.2 | Annual invitation of eligible employees; professional analysis; privacy-preserving aggregate results | Annual On-going Data Report plus survey-administration Technical Document (Shareable); direct staff scope; guide p. 19 |

The remaining eleven guide parts are marked NO in this workbook and are excluded by default. The app supports reviewed activation when actual awarded scope changes. Preconditions are not inferred merely from credit IDs: selection comes from the workbook.

## Differences that affect product logic

- **T01.2:** the service guide labels this twice-yearly and describes uploads tagged to seasonal windows. The workbook explicitly separates twice-yearly measurement from annual submission. The app creates two monitoring tasks and one annual report milestone.
- **A12.1:** the service guide says to provide both performance data and maintenance reports. This workbook's continuation row specifies annual filter replacement evidence and On-going Maintenance Report; no annual performance-data upload is specified for this part. The app follows the workbook, displays the disagreement and requires rule review. Do not silently create a regulatory obligation from a secondary claim.
- **W03.2:** the service guide labels the upload On-going Data Report. The workbook specifies On-going Maintenance Report for annual plan-implementation evidence. The app uses the latter and shows the discrepancy.
- **T02.1:** the guide's median across six surveys and scoring statement has not been independently verified. The workbook specifies current survey satisfaction thresholds (80% / 90%), eligibility, participation and seasonal administration. The app makes no points calculation from an assumed median.
- **W05.1/A14.1:** guide-defined quarterly work and annual documentation are represented separately when these excluded parts are activated. Manufacturer and project maintenance schedules may require more frequent work; add manual tasks as necessary.
- **X08.1:** guide obligations depend on future-purchase Option 2. A yearly review placeholder is not a substitute for a dated task and product evidence for each qualifying purchase.
- **C05.2:** one-time comparison; deactivate after documentation to avoid future planning placeholders.
- **V09.1:** internal annual data collection is distinguished from guide-defined recertification-only submission.
- **W07.3/W03.2:** a generic annual collection task does not reproduce every inspection or monitoring action required by a building-specific plan. Add those actual plan frequencies as tasks.
- The guide inconsistently says five concepts in one paragraph and six elsewhere; the part list spans six. Its marketing statements about reminders and the most commonly missed items have no independently verified evidence here.

## App workflow

Import the scorecard and guide → review part selection/version/scope/alternatives → enter certification and occupancy dates → choose annual package period → generate work and submission tasks → assign owners → capture dated evidence and corrective-action notes → record actual WELL Online submissions separately → export a year-end review package.

Dates without a certification anchor remain provisional. Calendar-year packaging is a user preference, not a universal WELL reporting deadline. Confirm any registration-based reporting anchor, submission grace period and recertification expiry in WELL Online; edit individual task dates to reflect reviewer instructions. Seasonal tasks are allocated to the reporting period in which their window ends, retaining the December–February window across the year boundary. The software does not interpret laboratory limits, approve alternative adherence paths, or certify survey privacy.

## Annual ZIP and narrative

The export includes original evidence arranged by part and activity, source workbook/PDF, the curated rule map, CSV task register and JSON manifest with file hashes. The Markdown narrative reports periods, recorded activity, notes, evidence counts and unresolved gaps. It is expressly a draft for human review. File presence does not establish that sampling, thresholds, participation or scope meet WELL criteria. Export does not upload to WELL Online.

Automatic export runs hourly while the local server is running and catches up for closed periods on restart after the schedule is confirmed and automation is enabled. No process can create archives while the machine is powered off. A stopped cloud process does not survive publication/restoration; startup instructions are saved separately.

## Next official checks

After the saved network settings are applied, retrieve the applicable IWBI certification guidebook and standard addenda, confirm the official ongoing-reporting procedure and consequences of missing reports, then verify the project's final report, expiry date, awarded parts, registration/certification anchor and required deadlines. The applicable Performance Verification Guidebook controls sampling locations, densities and valid methods.
