# Architecture

## Components

1. **Google Sheets** — lightweight state store for recruiter contacts, statuses, dates and Gmail Thread IDs.
2. **Google Apps Script** — orchestration layer.
3. **Gmail / Gmail API** — email delivery and conversation/thread management.
4. **Google Drive** — resume storage.
5. **AI layer (extension)** — JD/resume analysis and personalized outreach.

The spreadsheet state and Gmail conversation are linked by the persistent `Thread ID`.
