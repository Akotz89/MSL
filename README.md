# MSL Access Logbook

MSL is a Microsoft Access and VBA application for site teams to log daily IT work, track operational support, and generate monthly reports.

## Primary purpose

The app turns daily log entries into a structured monthly report with sections for:

1. Contractor Site Personnel
2. System Administration Actions
3. Troubleshooting Actions
4. Training Actions
5. Overall System Status
6. Program Office Support Requested
7. Program Office Support Received
8. Site Lead Notes

## Target workflow

1. Site personnel log administrative tasks, troubleshooting actions, training events, program office support, and notes throughout the month.
2. Site lead reviews the monthly period.
3. Access summarizes event counts, average response time, average return-to-service time, training totals, mission hours, downtime, and uptime.
4. Site lead generates a formatted PDF report.

## Technology direction

Initial version:

- Microsoft Access frontend
- Split Access backend
- VBA automation
- Access reports for PDF output

Future scale option:

- Access frontend
- SQL Server backend
- VBA/ODBC data access

## Repository structure

```text
docs/
  access-data-model.md
  forms-and-reports.md
sql/
  access-ddl.sql
vba/
  modReports.bas
  modValidation.bas
  modAuditTrail.bas
```

## Build phases

1. Data model and relationships
2. Daily log entry forms
3. Monthly review form
4. Report queries and report layout
5. PDF export automation
6. Audit trail and period locking
