# IT Request — Microsoft Graph calendar permissions for SRT (SSE Support Request Tracker)

**Requested by:** Dennis Dodd (HPE Networking SE)
**App:** `hpen-srt` — SharePoint Framework (SPFx) solution deployed to the `hpen-poc-manager` site App Catalog
**Date:** 2026-09-18

## What we need
Admin consent for two Microsoft Graph **delegated** permission scopes so the SRT tool can sync SSE
scheduling with Outlook/Exchange calendars:

| Scope | Why |
|---|---|
| **`Calendars.Read`** | Read SSE **free/busy** to power the availability calendar + double-book warnings. Uses the Graph `getSchedule` API, which returns only busy/free/tentative/out-of-office — **no meeting details, subjects, or attendees**. Privacy-safe. |
| **`Calendars.ReadWrite`** | Write **confirmed** SSE engagements onto the SSE's Outlook calendar as events (evolving our existing Power Automate "Dates Confirmed" calendar-event step to one event per scheduled block). |

*(If cross-user free/busy needs it on our tenant, please also include `Calendars.Read.Shared`.)*

## Precedent
This same app already has **`User.ReadBasic.All`** (and `GroupMember.Read.All`) granted — the calendar
scopes are the same delegated-permission model, just for calendar data.

## How to approve
The SPFx package requests these scopes in its manifest. After we deploy the build that includes the
request, they appear as **pending Graph API permission requests**:

> **SharePoint Admin Center → Advanced → API access** → approve the pending `Microsoft Graph`
> entries for `Calendars.Read` and `Calendars.ReadWrite`.

(Equivalent to granting admin consent for the delegated scopes on the SharePoint Online Client
Extensibility Web Application Principal.)

## Business context
SRT is becoming the system of record for Strategic Systems Engineer (SSE) engagements. Calendar sync
lets an SSE's real Outlook availability drive scheduling (so we never double-book), and pushes
confirmed engagements to their calendar automatically. Read is the higher-value, lower-risk half and
can be approved first if you'd prefer to stage it.

## Contact
Dennis Dodd — dennis.dodd@hpe.com
