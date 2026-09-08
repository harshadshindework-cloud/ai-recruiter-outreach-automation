# AI-Assisted Recruiter Outreach & Follow-up Automation

A QA-focused recruiter outreach automation prototype using Google Apps Script, Google Sheets, Gmail/Gmail API, and JavaScript, with an AI-assisted JD/resume matching extension.

> Portfolio/demo repository: real recruiter contacts, Gmail addresses, Thread IDs, credentials, and resumes are intentionally excluded.

## Core workflow

Google Sheet → 10 new applications → Gmail → persistent Thread ID → Day 02 → Day 03 → Day 04 → Sheet status tracking

## Features

- Controlled batches of 10 initial emails
- Up to 10 follow-ups per execution
- Gmail Thread ID persistence
- Threaded follow-ups
- Monday-Friday business-day scheduling
- Pending/Sent/Failed tracking
- Gmail recipient-quota awareness
- Recruiter-name personalization with Team fallback
- Duplicate-send prevention
- Recovery-oriented design
- AI-assisted resume/JD matching architecture

## Weekday rules

| Initial | Day 02 | Day 03 | Day 04 |
|---|---|---|---|
| Monday | Tuesday | Wednesday | Thursday |
| Tuesday | Wednesday | Thursday | Friday |
| Wednesday | Thursday | Friday | Monday |
| Thursday | Friday | Monday | Tuesday |
| Friday | Monday | Tuesday | Wednesday |

Saturday and Sunday are excluded.

## Google Sheet schema

`Company | Email | Recruiter Name | Status | Sent Date | Follow-up 02 Status | Follow-up 02 Date | Follow-up 03 Status | Follow-up 03 Date | Follow-up 04 Status | Follow-up 04 Date | Thread ID`

## Gmail threading challenge

The important design problem was preventing follow-ups from being delivered to the sender instead of the recruiter. The solution is to persist the original Gmail Thread ID and use Gmail message metadata to anchor replies to the existing conversation.

Recruiter replies remain part of the same conversation, so the stored Thread ID is not replaced merely because a recruiter responds.

## AI extension

Job Description + Resume → AI analysis → ATS keywords / skill match / experience relevance / aligned gaps → personalized outreach → Gmail automation.

The AI layer must remain factual and may only use skills supported by the candidate's real experience.

## Technology

JavaScript, Google Apps Script, Google Sheets, Gmail API, Google Drive, Gmail, REST/API concepts, AI/LLM workflows.

## Security

Never commit real recruiter lists, personal emails, Gmail Thread IDs, OAuth tokens, API keys, credentials, or private resumes. Use placeholders such as `recruiter@example.com`.

## Project status

Prototype validated for initial sending, Thread ID persistence, threaded follow-ups, weekday date calculation, batch limits, and Sheet state tracking.

## Future enhancements

- AI JD/resume match scoring
- Personalized outreach generation
- Recruiter reply classification
- Reply-aware follow-up suppression
- Application funnel dashboard
- Retry/recovery queue
- Automated reporting
