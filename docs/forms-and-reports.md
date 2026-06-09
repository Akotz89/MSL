# Forms and Reports Specification

This document defines the first-pass Microsoft Access user interface and report structure for the MSL logbook app.

## Forms

### frmDashboard

Main navigation screen.

Recommended buttons:

- New Log Entry
- Daily Log Review
- Monthly Report
- Personnel
- System Status
- Program Office Support
- Site Lead Notes
- Preview Report
- Export PDF

### frmLogEntry

Primary data entry form for site personnel.

Fields:

- Site
- Report Period
- Log Date
- Category
- System Area
- Title
- Description
- Resolution
- Logged By
- Event Count
- Response Time Minutes
- Return To Service Minutes
- Follow Up Required

Behavior:

- Category controls visible fields.
- Troubleshooting shows resolution and timing fields.
- Training shows training detail subform.
- Site lead notes emphasize narrative description.
- Save button validates required fields before save.

### frmLogEntrySteps subform

Child form under frmLogEntry for bullet details.

Fields:

- Step Number
- Step Text
- Step Type

Used to generate indented bullet points in the monthly report.

### frmMonthlyReport

Site lead review and report generation form.

Fields and controls:

- Site selector
- Report month selector
- Report period status
- Section totals
- Button: Preview Report
- Button: Export PDF
- Button: Lock Period

### frmSystemStatus

Monthly operational metrics entry form.

Fields:

- Missions Supported
- Mission Hours Supported
- Total System Down Time Minutes
- System Up Time Percent
- Notes

Behavior:

- Calculate uptime percentage from mission hours and downtime.
- Warn if mission hours are zero and downtime is greater than zero.

### frmPIOSupport

Program office support summary form.

Fields:

- Direction: Requested or Received
- Hardware Count
- Software Count
- Engineering Count
- Notes

## Reports

### rptMonthlySiteLog

Main monthly report.

Sections:

1. Contractor Site Personnel
2. System Administration Actions
3. Troubleshooting Actions
4. Training Actions
5. Overall System Status
6. Program Office Support Requested
7. Program Office Support Received
8. Site Lead Notes

Expected behavior:

- Filter by PeriodID.
- Use grouped subreports or section-specific queries.
- Use summary tables for event counts and averages.
- Print bullet steps under each log entry.
- Include page footer with contract/CDRL information when available.

## Suggested queries

- qryReportPeriodHeader
- qryPersonnelByPeriod
- qryAdminActionsByPeriod
- qryTroubleshootingByPeriod
- qryTrainingByPeriod
- qryTrainingStatsByPeriod
- qrySystemStatusByPeriod
- qryPIOSupportRequestedByPeriod
- qryPIOSupportReceivedByPeriod
- qrySiteLeadNotesByPeriod

## Report output rules

- Empty sections should still print a short statement when required, such as: No support was requested during this period.
- Response time and return-to-service time should display as '< 5 min' when values are less than 5.
- Uptime percentage should round to two decimals, but display as 100 when exactly 100.
- Final PDF filename format: Site_YYYY_MM_Monthly_Report.pdf
