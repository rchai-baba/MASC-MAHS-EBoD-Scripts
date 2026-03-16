# MASC-MAHS-EBoD-Scripts

Some scripts to help make repetitive EBoD tasks easier; right now just Google Apps Script



\## 1. Calendly API -> BoD Interview Signup

Calendly Part:
Google Apps Script that pulls from the API and outputs into our spreadsheet the Name, Email, Zoom link of each applicant

Matches rows using Date and Time information.
Phone Part:
Uses Email as index, then matches phone numbers from sheet with Emails and Phone numbers (copied from Applications Google Form -> Google Sheets output)
Then copy phone numbers using phone number key into Schedule phone number location

Main (next commit; untested atp):
Config, hourly triggers, last updated cell, fullupdates, calling both, etc.



Need to edit: 

* API Key in "config" section, for Calendly
* Sheet names, locations, in Config sections


\[Main Chat](https://claude.ai/chat/e97d2f52-34d8-4ec1-845d-50e1956ca6c9)

\[BOD Interviews Sheet](https://docs.google.com/spreadsheets/d/1CHWc2TQTwQ0dodv2Vbf-ld4L7RL6cyGWdfKqy2HTGo4/edit?gid=0#gid=0)

### API Key Creation
[API Key Link](https://calendly.com/integrations/api_webhooks)
Add API & Webhooks Integration, then create a API Key
Required Scopes:
- scheduled_events:read
- users:read

