# Access Data Model Specification

This document defines the first-pass Microsoft Access data model for the MSL logbook app.

## Design principles

- Store daily work once, then generate reports from queries.
- Keep personnel, sites, report periods, log entries, system status, and program office support separate.
- Use lookup values for categories where helpful, but keep the base schema simple for the first build.
- Make the reporting period the anchor for monthly output.

## Core tables

### tblSites

Tracks each location or reporting site.

| Field | Type | Notes |
|---|---|---|
| SiteID | AutoNumber | Primary key |
| SiteName | Short Text | Example: Texas |
| ContractNumber | Short Text | Optional |
| CDRLNumber | Short Text | Optional |
| Customer | Short Text | Optional |
| Active | Yes/No | Hide inactive sites from new entries |

### tblPersonnel

Tracks people assigned to a site.

| Field | Type | Notes |
|---|---|---|
| PersonnelID | AutoNumber | Primary key |
| SiteID | Number | Foreign key to tblSites |
| FirstName | Short Text |  |
| LastName | Short Text |  |
| Role | Short Text | Site Lead, CSO, CTS, etc. |
| ArrivalOnStation | Date/Time |  |
| Active | Yes/No |  |

### tblReportPeriods

One record per site and month.

| Field | Type | Notes |
|---|---|---|
| PeriodID | AutoNumber | Primary key |
| SiteID | Number | Foreign key to tblSites |
| ReportMonth | Date/Time | Store first day of month |
| ReportTitle | Short Text | Example: Texas - November 2025 |
| PreparedBy | Number | Foreign key to tblPersonnel |
| ReportStatus | Short Text | Draft, Review, Final |
| CreatedDate | Date/Time | Defaults to Now() |
| LockedDate | Date/Time | Null until finalized |

### tblLogEntries

Main table for administrative, troubleshooting, training, support, and note entries.

| Field | Type | Notes |
|---|---|---|
| LogID | AutoNumber | Primary key |
| PeriodID | Number | Foreign key to tblReportPeriods |
| SiteID | Number | Foreign key to tblSites |
| LogDate | Date/Time | Date work occurred |
| Category | Short Text | Administration, Troubleshooting, Training, PIO Requested, PIO Received, Site Lead Note |
| SystemArea | Short Text | Optional grouping |
| Title | Short Text | Short summary |
| Description | Long Text | Main narrative |
| Resolution | Long Text | Troubleshooting outcome or result |
| LoggedBy | Number | Foreign key to tblPersonnel |
| EventCount | Number | Defaults to 1 |
| ResponseTimeMinutes | Number | Optional |
| ReturnToServiceMinutes | Number | Optional |
| FollowUpRequired | Yes/No |  |
| CreatedDate | Date/Time | Defaults to Now() |
| ModifiedDate | Date/Time | Updated on edit |

### tblLogEntrySteps

Stores bullet points or step details under each log entry.

| Field | Type | Notes |
|---|---|---|
| StepID | AutoNumber | Primary key |
| LogID | Number | Foreign key to tblLogEntries |
| StepNumber | Number | Sort order |
| StepText | Long Text | Bullet text |
| StepType | Short Text | Action, Result, Note |

### tblTrainingStats

Additional fields for training events.

| Field | Type | Notes |
|---|---|---|
| TrainingID | AutoNumber | Primary key |
| PeriodID | Number | Foreign key to tblReportPeriods |
| LogID | Number | Foreign key to tblLogEntries |
| TrainingTopic | Short Text |  |
| MembersTrained | Number |  |
| HoursTrained | Number | Decimal allowed |

### tblSystemStatus

Monthly system status and uptime metrics.

| Field | Type | Notes |
|---|---|---|
| StatusID | AutoNumber | Primary key |
| PeriodID | Number | Foreign key to tblReportPeriods |
| MissionsSupported | Number |  |
| MissionHoursSupported | Number | Decimal allowed |
| TotalSystemDownTimeMinutes | Number | Store downtime in minutes |
| SystemUpTimePercent | Number | Calculated/validated before export |
| Notes | Long Text | Narrative notes |

### tblPIOSupport

Monthly support requested or received from the program office.

| Field | Type | Notes |
|---|---|---|
| PIOID | AutoNumber | Primary key |
| PeriodID | Number | Foreign key to tblReportPeriods |
| Direction | Short Text | Requested or Received |
| HardwareCount | Number |  |
| SoftwareCount | Number |  |
| EngineeringCount | Number |  |
| Notes | Long Text | Narrative notes |

### tblAuditTrail

Optional but recommended for accountability.

| Field | Type | Notes |
|---|---|---|
| AuditID | AutoNumber | Primary key |
| TableName | Short Text | Table modified |
| RecordID | Number | Record modified |
| ActionName | Short Text | Create, Update, Delete, Export, Lock |
| ActionBy | Short Text | Current Windows or app user |
| ActionDate | Date/Time | Defaults to Now() |
| Details | Long Text | Optional details |

## Relationships

- tblSites.SiteID to tblPersonnel.SiteID
- tblSites.SiteID to tblReportPeriods.SiteID
- tblReportPeriods.PeriodID to tblLogEntries.PeriodID
- tblReportPeriods.PeriodID to tblSystemStatus.PeriodID
- tblReportPeriods.PeriodID to tblPIOSupport.PeriodID
- tblLogEntries.LogID to tblLogEntrySteps.LogID
- tblLogEntries.LogID to tblTrainingStats.LogID
- tblPersonnel.PersonnelID to tblLogEntries.LoggedBy
- tblPersonnel.PersonnelID to tblReportPeriods.PreparedBy
