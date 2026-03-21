# Hitlist Hatcher

**MRRS Medical Readiness Hit List Generator**

Hitlist Hatcher is a browser-based tool that transforms MRRS (Medical Readiness Report System) Excel exports into color-coded medical readiness hit lists. It evaluates personnel against readiness requirements — PHA, Dental, HIV, Audiogram, Immunizations, deployment items, accession items, and more — and produces a sortable, printable report with PDF export capability.

## Data Privacy

Hitlist Hatcher processes all data locally in your browser. **No medical readiness data, PII, or CUI is ever transmitted to any server.**

- MRRS files are parsed entirely in browser memory using SheetJS
- No data is sent to external services during file upload, report generation, or PDF export
- The only outbound network request is the optional user-initiated feedback submission, which contains no MRRS data
- On page close or refresh, all MRRS data is discarded
- Settings (display preferences, unit name, emblem) are stored in browser localStorage — no MRRS data is persisted

There are no analytics, cookies, tracking pixels, or third-party services beyond the feedback endpoint.

## Usage

1. Download the MRRS IMR Detail report from MRRS as an Excel file
2. Open Hitlist Hatcher in Chrome or Edge
3. Drag and drop the Excel file (or click to browse)
4. Configure report items, immunizations, and display options
5. Click **Generate Hit List**
6. Export to PDF via the Export PDF button

For detailed instructions, use the in-app **Tutorial** button.

## Target Environment

- Chrome or Edge on Windows (optimized for government workstations)
- No internet access required during use (SheetJS loads on initial page open)
- No installation, no build step, no server — single-page static application

## Third-Party Attribution

This application includes [SheetJS Community Edition](https://sheetjs.com/), licensed under the Apache License 2.0. See [LICENSE-SHEETJS](LICENSE-SHEETJS) for full license text.

## Security

See [SECURITY.md](SECURITY.md) for the security policy, data flow details, and vulnerability reporting instructions.

## License

Copyright © 2026 Hitlist Hatcher. All rights reserved. See [LICENSE](LICENSE) for full terms.

This software is provided for use in generating medical readiness reports. Reproduction, modification, distribution, and incorporation into other systems are prohibited without express written permission.

This work was developed on personal time, using personal equipment, outside the scope of official military duties. It does not constitute a work of the United States Government under 17 U.S.C. § 105.
