# DLIG HQ Command Centre 2.0

DLIG HQ Command Centre 2.0 is a focused Founder dashboard for the July-December
2026 revenue mission. It upgrades the existing DLIG 1.0 static website and
Google Apps Script backend without deleting the historical product records.

## Main dashboard

The HQ homepage shows:

- collected revenue;
- confirmed revenue awaiting payment;
- active pipeline value;
- shortfall against RM50,000;
- the lead group closest to buying;
- the next customer delivery;
- the Founder's top three outcomes;
- overdue or at-risk items.

The main navigation contains Revenue, Sales Pipeline, Daily Command, the new HQ
Task Board, Calendar, Customer Success, Operations, Idea Parking Lot and the
File Index.

## DLIG 1.0 archive

Inventory, YLYD System, YLYD membership, 心动觉察, previous Sales & Marketing,
old product pricing, Admin, policies, meetings and decisions remain available
under:

`文件索引 / 1.0记录 -> 2025 DLIG 1.0 产品记录`

GDC Marketing Job is no longer exposed in the HQ 2.0 interface.

The old Task Board is not deleted from Google Sheets. HQ 2.0 uses a separate
`HQ Task Board` tab, which starts empty.

## Files

- `index.html` - existing DLIG 1.0 single-page application and archive pages.
- `hq-2.0.css` - responsive HQ 2.0 design layer.
- `hq-2.0.js` - HQ modules, forms, calculations and manual data fallback.
- `google-apps-script-Code.gs` - existing Apps Script plus HQ 2.0 Sheet adapters.

No build step is required. Serve the folder as a static website.

## Google Sheets setup

The Apps Script creates these tabs on the first successful write:

### Sales & Marketing spreadsheet

- `HQ Revenue`
- `HQ Monthly Sales`
- `HQ Products`
- `HQ Pipeline`

### Task & Events spreadsheet

- `HQ Priorities`
- `HQ Daily`
- `HQ Customer Success`
- `HQ Events`
- `HQ Operations`
- `HQ Ideas`
- `HQ Links`
- `HQ Task Board`
- `HQ Task Categories`

Existing tabs are not deleted or renamed.

### Deploy the updated Apps Script

1. Open the Apps Script project currently used by DLIG Command Centre.
2. Replace its code with `google-apps-script-Code.gs`.
3. Create a new deployment or update the existing Web App deployment.
4. Confirm the deployment URL matches `SHEETS_API_URL` in `index.html`.
5. Restrict access appropriately. Do not expose unrestricted write endpoints
   containing customer or financial information.
6. Open the HQ website and make one test update.
7. Confirm the corresponding `HQ ...` tab was created and contains the update.

Until the updated Apps Script is deployed, HQ 2.0 uses browser-local manual
data. Important cards show whether their data is manual or Sheet-synced.

## Updating the dashboard

### Revenue

Open `Revenue` and update:

- July-December target;
- collected revenue;
- confirmed awaiting payment;
- interest pipeline;
- product price, enrolment, target and actual revenue.

Only actual money received belongs in `Collected`.

### Pipeline

Open `Sales Pipeline`. Lead groups can be used instead of customer names.
Every active lead group should have:

- stage;
- number of leads;
- expected value;
- product;
- source;
- next follow-up date;
- note.

Avoid entering private customer details in frontend source files.

### Founder priorities

Open `Daily Command` and update the top three outcomes. Each outcome and
delegated task should include an owner, status, deadline and definition of done.

### Task Board

The HQ Task Board stores only:

- task;
- owner;
- optional team member;
- category;
- deadline;
- completion date;
- completed status.

The default categories come from the Strategy Playbook Task Rate table.
Typing a new category adds it to the reusable category list.

### Calendar

Only confirmed dates should be added. Planning windows such as "2-3 Tuesday
free talks" should remain a planning note until the individual dates are known.

### Customer Success

Preparation can record curriculum readiness, facilitator, room/Zoom, templates,
payment list, reminders, recording and playback access. Risk can record pending
payments, low attendance, incomplete content, untested links, unclear ownership
or schedule conflicts.

### Idea Parking Lot

Ideas with `Parking` or `待评估` status do not become active Founder priorities.
Change an idea to `Approved` only after a deliberate decision.

## Phase 2 integration boundaries

HQ 2.0 keeps separate data adapters for:

- calendar events;
- CRM/pipeline data;
- payment/revenue data.

Future Google Calendar, InfiniteSales and Stripe/InfiniteSales integrations
should replace those adapters. Do not label data as automatically synced until
the real API or webhook is active and verified.

## Security

- Never commit passwords, API keys, payment credentials, bank account details,
  customer private data, Zoom passcodes or Zoom host keys.
- Store runtime secrets in the backend or deployment configuration.
- The old Zoom host credentials previously present in the frontend have been
  removed. Rotate any credential that has appeared in Git history.
- The local Git remote previously contained a GitHub access token. Rotate that
  token and use a credential-free remote URL.
- Review Apps Script Web App access before using it for ChatGPT Actions or
  external webhooks.

## Local preview

Run any static server from this folder, for example:

```bash
python3 -m http.server 8765 --bind 127.0.0.1
```

Then open `http://127.0.0.1:8765/`.
