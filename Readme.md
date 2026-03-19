TestTrack is a lightweight test case management tool built entirely on top of Google Sheets and Google Apps Script, with a single-page HTML/JS frontend served via Web App.
It allows QA teams to organize, track, and execute test cases without any external tools or paid subscriptions — everything lives in a Google Spreadsheet.
Key features:

Test Suites & Cases — Create and manage test suites with detailed test cases, including type, priority, status, assigned tester, automation status, steps, and expected results
Test Runs — Launch test runs from any suite (all cases, manually picked, or random sample), track execution progress in real time
Test Plans — Group multiple suites into a plan and monitor overall pass rate
Bulk Actions — Change status, move, or delete multiple cases at once
Attachments — Upload and manage file attachments per test case, stored in Google Drive
Import / Export — Import test cases from CSV files and export them back at any time
Overview Dashboard — A project-level summary of all cases and their statuses across suites
Google Sheets backend — All data is persisted directly into a connected spreadsheet with a single Save action

Tech stack: Google Apps Script (backend) · Vanilla HTML/CSS/JS (frontend) · Google Drive API (attachments)