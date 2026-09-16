# Supplier SLA & Voucher Delivery Dashboard

**For:** TravClan — Operations Intern application

## What it does
Analyzes 220 simulated bookings across Flights, Hotels, and Holiday Packages
from 6 suppliers, and builds a live Excel dashboard that tracks:
- Voucher delivery TAT vs SLA threshold per product type
- On-time delivery % by supplier (with conditional-formatting heatmap)
- Supplier payment delays
- Agent KYC completion status
- An "Ops Risk Flag" that auto-highlights suppliers needing review

## Key finding
2 of 6 suppliers (HolidayCraft DMC, TerraTrip Packages) breach voucher SLA on
30%+ of bookings and average 6+ days of payment delay — flagged "Review Needed"
for prioritized supplier escalation.

## Files
- `TravClan_Booking_Ops_SLA_Tracker.xlsx` — the deliverable (Raw Bookings +
  Supplier SLA Dashboard tabs, live formulas, chart)
- `01_generate_booking_data.py` — generates the synthetic raw booking dataset
- `02_build_sla_dashboard.py` — builds the dashboard tab, formulas, chart,
  and conditional formatting on top of the raw data

## Tools used
Excel (openpyxl), SUMIFS / AVERAGEIF / SUMPRODUCT formulas, conditional
formatting, native bar chart
