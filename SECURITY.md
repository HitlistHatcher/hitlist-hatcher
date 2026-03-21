# Security Policy

## Supported Versions

Only the latest release is supported with security updates.

## Data Privacy

Hitlist Hatcher is a client-side-only application. All MRRS data
processing occurs locally in the browser. No medical readiness
data, personally identifiable information, or CUI is transmitted
to any server, stored remotely, or cached beyond the browser
session.

The only outbound network request is the optional user-initiated
feedback submission, which contains no MRRS data.

## Reporting a Vulnerability

To report a security vulnerability, submit a report through the
in-app Feedback button (select "Bug Report" as the category) and
include "SECURITY" in the message. Reports are reviewed by the
developer directly.

Please do not report security vulnerabilities through public
GitHub Issues.

## Scope

This policy covers the three application files (index.html,
style.css, app.js) and the bundled third-party dependency
(SheetJS Community Edition). It does not cover the MRRS system
or the user's browser environment.
