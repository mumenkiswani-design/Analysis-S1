# Data Validation & Reconciliation Methodology

The dashboard uses the user's supplied source files only.

1. S1 is January through June.
2. 2025 sales use the Medical Rep field from the All in One source.
3. For 2026, accounts appearing in the subagent contribution file are split across Medical Reps using the stated percentages.
4. Accounts not appearing in the contribution file retain the Medical Rep from All in One.
5. The resulting 2026 Rep × SKU totals are compared with the official S1 achievement PDF.
6. A proportional factor (Official / Calculated) is applied within each Rep × SKU to the detailed month/area rows. This preserves the detailed distribution while making the dashboard reconcile to the official control report.
7. The Data Validation page displays calculated values before reconciliation, official values, variance, variance %, and the factor used.

This approach does not claim that the proportional adjustment identifies the underlying transaction-level reason for every variance; it is a transparent reconciliation for management reporting.
