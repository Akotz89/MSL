-- MSL Access Logbook starter DDL
-- Run statements individually in Microsoft Access SQL View if needed.
-- Some relationship/index setup is usually easier through Access Relationship Designer.

CREATE TABLE tblSites (
    SiteID AUTOINCREMENT CONSTRAINT pk_tblSites PRIMARY KEY,
    SiteName TEXT(100) NOT NULL,
    ContractNumber TEXT(100),
    CDRLNumber TEXT(50),
    Customer TEXT(100),
    Active YESNO
);

CREATE TABLE tblPersonnel (
    PersonnelID AUTOINCREMENT CONSTRAINT pk_tblPersonnel PRIMARY KEY,
    SiteID LONG NOT NULL,
    FirstName TEXT(50) NOT NULL,
    LastName TEXT(50) NOT NULL,
    Role TEXT(50),
    ArrivalOnStation DATETIME,
    Active YESNO
);

CREATE TABLE tblReportPeriods (
    PeriodID AUTOINCREMENT CONSTRAINT pk_tblReportPeriods PRIMARY KEY,
    SiteID LONG NOT NULL,
    ReportMonth DATETIME NOT NULL,
    ReportTitle TEXT(150),
    PreparedBy LONG,
    ReportStatus TEXT(25),
    CreatedDate DATETIME,
    LockedDate DATETIME
);

CREATE TABLE tblLogEntries (
    LogID AUTOINCREMENT CONSTRAINT pk_tblLogEntries PRIMARY KEY,
    PeriodID LONG NOT NULL,
    SiteID LONG NOT NULL,
    LogDate DATETIME NOT NULL,
    Category TEXT(50) NOT NULL,
    SystemArea TEXT(100),
    Title TEXT(255) NOT NULL,
    Description LONGTEXT,
    Resolution LONGTEXT,
    LoggedBy LONG,
    EventCount LONG,
    ResponseTimeMinutes LONG,
    ReturnToServiceMinutes LONG,
    FollowUpRequired YESNO,
    CreatedDate DATETIME,
    ModifiedDate DATETIME
);

CREATE TABLE tblLogEntrySteps (
    StepID AUTOINCREMENT CONSTRAINT pk_tblLogEntrySteps PRIMARY KEY,
    LogID LONG NOT NULL,
    StepNumber LONG NOT NULL,
    StepText LONGTEXT NOT NULL,
    StepType TEXT(50)
);

CREATE TABLE tblTrainingStats (
    TrainingID AUTOINCREMENT CONSTRAINT pk_tblTrainingStats PRIMARY KEY,
    PeriodID LONG NOT NULL,
    LogID LONG,
    TrainingTopic TEXT(150),
    MembersTrained LONG,
    HoursTrained DOUBLE
);

CREATE TABLE tblSystemStatus (
    StatusID AUTOINCREMENT CONSTRAINT pk_tblSystemStatus PRIMARY KEY,
    PeriodID LONG NOT NULL,
    MissionsSupported LONG,
    MissionHoursSupported DOUBLE,
    TotalSystemDownTimeMinutes LONG,
    SystemUpTimePercent DOUBLE,
    Notes LONGTEXT
);

CREATE TABLE tblPIOSupport (
    PIOID AUTOINCREMENT CONSTRAINT pk_tblPIOSupport PRIMARY KEY,
    PeriodID LONG NOT NULL,
    Direction TEXT(25) NOT NULL,
    HardwareCount LONG,
    SoftwareCount LONG,
    EngineeringCount LONG,
    Notes LONGTEXT
);

CREATE TABLE tblAuditTrail (
    AuditID AUTOINCREMENT CONSTRAINT pk_tblAuditTrail PRIMARY KEY,
    TableName TEXT(100) NOT NULL,
    RecordID LONG,
    ActionName TEXT(50) NOT NULL,
    ActionBy TEXT(100),
    ActionDate DATETIME,
    Details LONGTEXT
);

-- Suggested defaults to set in Access table design:
-- tblSites.Active = True
-- tblPersonnel.Active = True
-- tblReportPeriods.ReportStatus = "Draft"
-- tblReportPeriods.CreatedDate = Now()
-- tblLogEntries.EventCount = 1
-- tblLogEntries.CreatedDate = Now()
-- tblLogEntries.ModifiedDate = Now()
-- tblAuditTrail.ActionDate = Now()
