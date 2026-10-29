# Quarterly Project Report

This sample shows what md2docx produces from plain Markdown: Word heading styles,
full-width tables with measured column widths, real list numbering, and a wide table
that moves to a landscape page on its own.

## Milestones

| Milestone | Owner | Due | Status |
|---|---|---|---|
| Requirements signed off | Product | 2026-07-15 | Done |
| Beta release | Engineering | 2026-09-01 | Done |
| Security review | Security | 2026-09-20 | In progress |
| General availability | Release team | 2026-10-31 | Planned |

## Highlights

1. Release cadence moved from monthly to biweekly.
2. Median build time dropped from 14 to 9 minutes.
3. Support backlog fell below 40 tickets.

- Risks are reviewed every week.
  - Vendor contract renewal is due in November.
  - Two senior engineers join in December.

> All figures in this sample are fictional.

## Regional Breakdown

Key metrics by region:

| Region | Subscribers | Retention | Engagement | Conversion | Partners | Coverage | Outlook |
|---|---|---|---|---|---|---|---|
| North America | 1,240,000 | 92.4% | Outstanding | 4.8% | Established | Nationwide | Increasing |
| Europe | 980,000 | 90.1% | Considerable | 4.1% | Expanding | Continental | Stabilizing |
| Asia-Pacific | 1,510,000 | 88.7% | Accelerating | 5.3% | Negotiating | Metropolitan | Recovering |
| Latin America | 410,000 | 85.2% | Encouraging | 3.6% | Exploratory | Metropolitan | Developing |
| Middle East | 260,000 | 83.9% | Promising | 3.2% | Exploratory | Metropolitan | Developing |
| Africa | 190,000 | 81.4% | Emerging | 2.9% | Negotiating | Regional | Developing |
| Nordics | 150,000 | 94.0% | Outstanding | 5.1% | Established | Nationwide | Increasing |
| Oceania | 120,000 | 91.3% | Considerable | 4.4% | Expanding | Nationwide | Stabilizing |

## Next Steps

After the wide table the document returns to portrait pages automatically.

1. Freeze the release branch on October 24.
2. Publish the migration guide.

```sh
python3 md2docx.py report.md
```
