# Geo-trick
Small functions what could help in work

# Splitting_overlapping_intervals

A small Python utility for fixing overlapping lithology intervals in drillhole tables without losing information.

## What it does

This script converts overlapping interval data into a sequential, non-overlapping interval table by:

- splitting each drillhole into the smallest possible consecutive sub-intervals,
- detecting all lithologies active within each sub-interval,
- writing them into separate columns (`LITH`, `LITH2`, `LITH3`, etc.).

This is useful for geological interval tables where overlapping lithology records must be preserved rather than deleted or arbitrarily merged.

## Input format

The input Excel sheet must contain these columns:

- `BHID`
- `FROM`
- `TO`
- `LITH`

Example:

| BHID   | FROM | TO | LITH  |
|--------|------|----|-------|
| DH0001 | 43   | 45 | Lith7 |
| DH0001 | 43   | 46 | Lith8 |
| DH0001 | 45   | 47 | Lith9 |

## Output format

The script creates a new Excel file with:

- `Original` — the original input table
- `Sequential_Lith` — the normalized interval table
- `Skipped_Rows` — rows that were skipped due to invalid or incomplete data

Example output:

| BHID   | FROM | TO | LITH  | LITH2  |
|--------|------|----|-------|--------|
| DH0001 | 43   | 45 | Lith7 | Lith8  |
| DH0001 | 45   | 46 | Lith8 | Lith9  |
| DH0001 | 46   | 47 | Lith9 | Lith10 |

## Interval logic

The script treats intervals as **half-open**:

`[FROM, TO)`

This means:

- `0–9` and `9–12` are considered adjacent, not overlapping
- overlapping intervals are split only where necessary
- no lithological information is discarded

## Skipped rows

Rows are skipped only if they are invalid, for example:

- empty `BHID`
- non-numeric or empty `FROM`
- non-numeric or empty `TO`
- `TO <= FROM`

Skipped records are exported to the `Skipped_Rows` sheet with a reason for exclusion.

## Requirements

Install the required packages:

```bash
pip install pandas openpyxl numpy
