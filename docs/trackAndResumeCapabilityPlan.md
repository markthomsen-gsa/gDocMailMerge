# Mail Merge Progress Tracking & Resume Feature Implementation Plan

## Executive Summary

This document outlines the implementation plan for adding progress tracking and resume capabilities to the existing Mail Merge Google Docs Add-on. The feature will allow users to safely interrupt mail merge operations and resume them later, maintaining a record of sent emails to prevent duplicates and ensure complete delivery. All job data will be externalized to spreadsheets for transparency, auditability, and to facilitate future technology migrations.

## Problem Statement

Currently, mail merge operations are vulnerable to interruptions such as:
- Browser/tab closures
- Daily email quota limitations
- Network connectivity issues
- User-initiated pauses

When interrupted, the current system provides no way to determine which emails were successfully sent and which were not, potentially resulting in duplicate messages or incomplete campaigns.

## Proposed Solution

Implement a robust tracking and resume system that:
1. Records the status of each recipient during mail merge execution
2. Saves a snapshot of all data and HTML template at job start for consistency and auditability
3. Provides a simple UI for resuming interrupted mail merges
4. Ensures no duplicate emails are sent when resuming
5. Archives completed jobs to a predetermined folder automatically
6. Utilizes an abstraction layer for future technology migration flexibility

## Implementation Phases

### Phase 1: MVP (Minimum Viable Product)

**Objective:** Create the core functionality to track progress and resume mail merges with minimal changes to the existing user experience.

#### Technical Requirements:

1. **Data Snapshot Creation**
   - Create a complete copy of the user's spreadsheet when a mail merge starts
   - Add the following sheets to the new spreadsheet (all sheet names configurable via code):
     - `JobMetadata`: Contains job configuration, version info, and status
     - `EmailTemplate`: Stores the complete HTML template being used
     - `RecipientTracking`: Original recipient data plus status columns
     - `EventLog`: Records all significant events with timestamps
   - Structure for RecipientTracking: Original data + Status column + Timestamp column + Error column
   - Send notification email with spreadsheet link when job starts

2. **Progress Tracking**
   - Update recipient status in the tracking sheet after each email send attempt
   - Record success/failure status and timestamp for each recipient
   - Store job metadata (job ID, tracking sheet name, last processed index) in User Properties

3. **Basic Resume UI**
   - Add notification banner when an incomplete job is detected
   - Add "Resume Previous Mail Merge" button in the Send section
   - Display simple progress summary (e.g., "45 of 100 emails sent")

4. **Resume Logic**
   - Retrieve incomplete job data from the JobMetadata sheet
   - Load recipient data from the RecipientTracking sheet
   - Skip recipients already marked as "Sent"
   - Continue processing from the last successful send

5. **Job Archiving**
   - Automatically move the spreadsheet to a predetermined archive folder upon job completion
   - Send notification email with link to the archived spreadsheet
   - Include summary statistics in the email

6. **Abstraction Layer**
   - Implement interfaces for all data operations (JobDataService, EmailTemplateService, etc.)
   - Create concrete implementations using spreadsheet storage
   - Design for future technology migration possibilities

#### User Experience (MVP):

1. User starts a mail merge
2. If interrupted, progress is automatically saved
3. When reopening the sidebar, user sees a notification about the unfinished mail merge
4. User clicks "Resume" button to restore previous settings
5. Mail merge continues from where it left off, skipping already sent emails

#### MVP Deliverables:

- Hidden tracking sheet implementation
- Status update mechanism
- Basic resume UI elements
- Resume logic to continue processing
- Simple progress indicator

### Phase 2: Enhanced Experience

**Objective:** Improve the user experience with better visibility and control over mail merge jobs.

#### Technical Requirements:

1. **Enhanced Progress Dashboard**
   - Add detailed view of sent/pending/failed emails
   - Include visual progress bars for each status category
   - Provide filtering options to view specific status groups

2. **Multiple Job Management**
   - Support tracking multiple unfinished mail merge jobs
   - Allow users to choose which job to resume
   - Implement job naming for easier identification

3. **Job Control Options**
   - Add pause/resume controls during execution
   - Support manual saving of progress
   - Allow cancellation with status preservation

4. **Data Change Detection**
   - Implement change detection for source data
   - Alert users if original data has changed since job started
   - Provide options for handling changed data

#### User Experience (Phase 2):

1. Enhanced visibility into mail merge progress and status
2. Ability to manage multiple mail merge jobs
3. More control over in-progress mail merges
4. Smart handling of data source changes

### Phase 3: Advanced Features

**Objective:** Add advanced capabilities for power users and enterprise needs.

#### Technical Requirements:

1. **Scheduled Resumption**
   - Allow scheduling of when to resume interrupted jobs
   - Implement automatic resumption based on quota refreshes

2. **Export/Import Capability**
   - Enable exporting job configurations and status
   - Support importing previously exported jobs

3. **Analytics Integration**
   - Track and display mail merge performance metrics
   - Provide insights on best sending times and patterns

4. **Admin Controls**
   - Add organizational settings for mail merge policies
   - Implement approval workflows for large campaigns

## Technical Implementation Details

### Data Storage Schema

**JobMetadata Sheet Structure:**

| Property | Value |
|----------|-------|
| jobId | job_20240311_123456 |
| originalSpreadsheetId | 1AbC... |
| originalSheetName | Sheet1 |
| emailColumn | Email |
| subjectLine | Your March Newsletter |
| fromEmail | sender@example.com |
| fromName | Sender Name |
| startTime | 2024-03-11T14:30:00.000Z |
| lastProcessedIndex | 45 |
| totalRecipients | 100 |
| completedRecipients | 45 |
| failedRecipients | 2 |
| status | IN_PROGRESS |
| appVersion | 1.2.3 |
| cc | cc@example.com |
| bcc | bcc@example.com |

**EmailTemplate Sheet Structure:**

| Property | Value |
|----------|-------|
| templateId | template_20240311_123456 |
| sourceDocumentId | 1XyZ... |
| creationTime | 2024-03-11T14:30:00.000Z |
| htmlContent | (full HTML content of the email template) |

**RecipientTracking Sheet Structure:**

| Original Columns... | Status | Timestamp | Error |
|---------------------|--------|-----------|-------|
| (recipient data)    | Pending | null | null |
| (recipient data)    | Sent | 2024-03-11T14:30:05.000Z | null |
| (recipient data)    | Failed | 2024-03-11T14:30:10.000Z | "Invalid email address" |

**EventLog Sheet Structure:**

| Timestamp | EventType | Description | Details |
|-----------|-----------|-------------|---------|
| 2024-03-11T14:30:00.000Z | JOB_STARTED | Mail merge job started | Recipients: 100 |
| 2024-03-11T14:35:00.000Z | BATCH_COMPLETED | Processed batch of emails | Sent: 20, Failed: 1 |
| 2024-03-11T14:40:00.000Z | JOB_PAUSED | Mail merge job paused | Reason: Browser closed |
| 2024-03-11T15:00:00.000Z | JOB_RESUMED | Mail merge job resumed | Remaining: 79 |

### Key Code Components (Pseudocode)

**1. Abstraction Layer Interfaces:**
```javascript
// Interface for job data operations
interface JobDataService {
  createJob(config);
  getJob(jobId);
  updateJobProgress(jobId, progress);
  completeJob(jobId);
  getActiveJobs();
}

// Interface for template storage
interface TemplateService {
  saveTemplate(jobId, documentId, htmlContent);
  getTemplate(jobId);
}

// Interface for recipient tracking
interface RecipientService {
  initializeRecipients(jobId, recipients);
  updateRecipientStatus(jobId, recipientIndex, status, error);
  getPendingRecipients(jobId);
  getRecipientStats(jobId);
}

// Interface for event logging
interface EventLogService {
  logEvent(jobId, eventType, description, details);
  getEvents(jobId);
}
```

**2. Initiate Tracking:**
```javascript
function startMailMergeWithTracking(spreadsheetId, sheetName, emailColumn, subjectLine, fromEmail, fromName, options) {
  // Generate unique job ID
  const jobId = "job_" + new Date().toISOString().replace(/[^\w]/g, "_");
  
  // Create job spreadsheet (complete copy)
  const jobSpreadsheetId = createJobSpreadsheet(spreadsheetId, jobId);
  
  // Initialize services with concrete implementations
  const jobService = new SpreadsheetJobDataService(jobSpreadsheetId);
  const templateService = new SpreadsheetTemplateService(jobSpreadsheetId);
  const recipientService = new SpreadsheetRecipientService(jobSpreadsheetId);
  const eventLogService = new SpreadsheetEventLogService(jobSpreadsheetId);
  
  // Get HTML template from document
  const htmlTemplate = DocumentApp.getActiveDocument().getBody().getText();
  
  // Set up job
  jobService.createJob({
    jobId,
    originalSpreadsheetId: spreadsheetId,
    originalSheetName: sheetName,
    emailColumn,
    subjectLine,
    fromEmail,
    fromName,
    startTime: new Date().toISOString(),
    totalRecipients: getRecipientCount(spreadsheetId, sheetName, emailColumn),
    completedRecipients: 0,
    failedRecipients: 0,
    status: "STARTED",
    appVersion: getCurrentVersion(),
    options
  });
  
  // Save HTML template
  templateService.saveTemplate(jobId, DocumentApp.getActiveDocument().getId(), htmlTemplate);
  
  // Initialize recipient tracking
  const recipients = getRecipientsData(spreadsheetId, sheetName);
  recipientService.initializeRecipients(jobId, recipients);
  
  // Log the event
  eventLogService.logEvent(jobId, "JOB_STARTED", "Mail merge job started", {recipientCount: recipients.length});
  
  // Send notification email with link
  sendJobStartNotification(jobId, jobSpreadsheetId);
  
  // Start mail merge process
  return executeMailMergeWithTracking(jobId, jobSpreadsheetId);
}
```

**3. Update Status During Processing:**
```javascript
function updateRecipientStatus(jobId, jobSpreadsheetId, recipientIndex, status, error = null) {
  // Initialize services
  const jobService = new SpreadsheetJobDataService(jobSpreadsheetId);
  const recipientService = new SpreadsheetRecipientService(jobSpreadsheetId);
  const eventLogService = new SpreadsheetEventLogService(jobSpreadsheetId);
  
  // Update recipient status
  recipientService.updateRecipientStatus(jobId, recipientIndex, status, error);
  
  // Update job progress in metadata
  const stats = recipientService.getRecipientStats(jobId);
  jobService.updateJobProgress(jobId, {
    lastProcessedIndex: recipientIndex,
    completedRecipients: stats.completed,
    failedRecipients: stats.failed
  });
  
  // Log event if important status change
  if (status === "Sent" || status === "Failed") {
    const eventType = status === "Sent" ? "EMAIL_SENT" : "EMAIL_FAILED";
    eventLogService.logEvent(jobId, eventType, `Email ${status.toLowerCase()} to recipient`, {
      recipientIndex,
      error: error || null
    });
  }
}
```

**4. Resume Logic:**
```javascript
function resumeMailMerge(jobId) {
  // Find job spreadsheet ID from active jobs
  const activeJobs = getActiveJobs();
  const job = activeJobs.find(j => j.jobId === jobId);
  
  if (!job) {
    throw new Error("Job not found: " + jobId);
  }
  
  // Initialize services
  const jobService = new SpreadsheetJobDataService(job.spreadsheetId);
  const templateService = new SpreadsheetTemplateService(job.spreadsheetId);
  const recipientService = new SpreadsheetRecipientService(job.spreadsheetId);
  const eventLogService = new SpreadsheetEventLogService(job.spreadsheetId);
  
  // Get full job details
  const jobDetails = jobService.getJob(jobId);
  
  // Log resume event
  eventLogService.logEvent(jobId, "JOB_RESUMED", "Mail merge job resumed", {
    completedRecipients: jobDetails.completedRecipients,
    remainingRecipients: jobDetails.totalRecipients - jobDetails.completedRecipients
  });
  
  // Get pending recipients
  const pendingRecipients = recipientService.getPendingRecipients(jobId);
  
  // Get template
  const template = templateService.getTemplate(jobId);
  
  // Display resume information in UI
  updateResumeUI(jobDetails.completedRecipients, jobDetails.totalRecipients);
  
  // Continue processing from where we left off
  return processRemainingRecipients(jobId, job.spreadsheetId, pendingRecipients, template.htmlContent);
}
```

**5. Job Completion and Archiving:**
```javascript
function completeAndArchiveJob(jobId, jobSpreadsheetId) {
  // Initialize services
  const jobService = new SpreadsheetJobDataService(jobSpreadsheetId);
  const eventLogService = new SpreadsheetEventLogService(jobSpreadsheetId);
  
  // Mark job as completed
  jobService.completeJob(jobId);
  
  // Log completion event
  const jobDetails = jobService.getJob(jobId);
  eventLogService.logEvent(jobId, "JOB_COMPLETED", "Mail merge job completed", {
    totalRecipients: jobDetails.totalRecipients,
    completedRecipients: jobDetails.completedRecipients,
    failedRecipients: jobDetails.failedRecipients
  });
  
  // Move spreadsheet to archive folder
  const archiveFolderId = getArchiveFolderId(); // Get ID from configuration
  const spreadsheet = SpreadsheetApp.openById(jobSpreadsheetId);
  const file = DriveApp.getFileById(jobSpreadsheetId);
  const archivedFile = file.makeCopy(spreadsheet.getName() + " (Completed)", DriveApp.getFolderById(archiveFolderId));
  
  // Send completion notification
  sendCompletionNotification(jobId, jobDetails, archivedFile.getUrl());
  
  // Optionally delete the original after successful archive
  // file.setTrashed(true);
  
  return {
    jobId,
    status: "COMPLETED",
    archiveUrl: archivedFile.getUrl()
  };
}
```

## Implementation Timeline

### MVP (2-3 weeks)
- Week 1: Implement tracking sheet creation and status updating mechanism
- Week 2: Develop basic resume UI and logic
- Week 3: Testing and bug fixes

### Phase 2 (3-4 weeks)
- Week 4-5: Enhanced dashboard and multiple job management
- Week 6-7: Job control options and data change detection

### Phase 3 (4-5 weeks)
- Week 8-9: Scheduled resumption and export/import capabilities
- Week 10-12: Analytics integration and admin controls

## Risk Assessment

| Risk | Impact | Mitigation |
|------|--------|------------|
| Performance impact with large spreadsheets | High | Implement batch processing and optimize read/write operations |
| Drive quota limitations for copying large spreadsheets | Medium | Implement size checks and warning for very large spreadsheets |
| Email quota limitations affecting notifications | Low | Use prioritized notification system (only send critical emails) |
| Execution time limits in Apps Script | High | Break operations into smaller chunks with continuation tokens |
| Permission issues for accessing/modifying files | Medium | Implement clear error messages and permission request flows |
| Data consistency in copied spreadsheets | Medium | Perform verification checks after copying operations |
| Failure during archiving | Medium | Implement retry mechanism and maintain original until archive is confirmed |

## Success Metrics

- Reduction in failed mail merge completions
- Increase in successful resumptions
- User feedback on feature utility
- Reduction in support tickets related to incomplete mail merges

## Notification Emails

The implementation will include two critical notification emails:

### 1. Job Start Notification
Sent as soon as the job spreadsheet is created and contains:
- Link to the job spreadsheet for monitoring progress
- Summary of job settings (recipient count, subject, etc.)
- Time when the job started
- Version of the Mail Merge application

### 2. Job Completion Notification
Sent when the job is complete and archived:
- Link to the archived spreadsheet
- Complete summary of results (sent, failed, etc.)
- Time metrics (start time, end time, duration)
- Any significant errors or warnings

## Configurable Sheet Names

All sheet names will be configurable through a central configuration object:

```javascript
const DEFAULT_SHEET_NAMES = {
  JOB_METADATA: "JobMetadata",
  EMAIL_TEMPLATE: "EmailTemplate",
  RECIPIENT_TRACKING: "RecipientTracking",
  EVENT_LOG: "EventLog"
};
```

This configuration can be modified programmatically without changing implementation code.

## Conclusion

This implementation plan provides a clear roadmap for adding robust progress tracking and resume capabilities to the Mail Merge add-on. By externalizing all job data to spreadsheets and implementing a proper abstraction layer, we ensure auditability, transparency, and future flexibility.

The approach of copying spreadsheets at job start provides complete control and data consistency throughout the mail merge process. The automatic archiving with email notifications creates a reliable audit trail for compliance purposes.

The abstraction layer design ensures that the implementation can be migrated to different technologies in the future without significant rework, protecting the investment in this feature.