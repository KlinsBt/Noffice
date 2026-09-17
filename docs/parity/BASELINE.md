# Reference environment and target decisions

Decision date: September 10, 2026. Use the versions installed on this PC as the fixed desktop reference, as requested. Read-only executable inspection found:

| App        | Product               | File version   | Architecture |
| ---------- | --------------------- | -------------- | ------------ |
| Word       | Microsoft Office 2016 | 16.0.4627.1000 | x64          |
| Excel      | Microsoft Office 2016 | 16.0.4627.1001 | x64          |
| PowerPoint | Microsoft Office 2016 | 16.0.4266.1001 | x64          |

[baseline.json](baseline.json) records resolved executable paths, hashes, OS build 26200.9168, de-DE system/UI culture, UTC inspection time and method. These are file versions; they must not be substituted for application COM Build fields. Application UI/editing languages and all application preferences must be captured per native test; system culture does not prove those settings. Installed edition/add-in entitlements are not inferred from file versions.

Reproduce with `powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/inspect-office-baseline.ps1`. Default output is .local/office-baseline.json so ordinary reinspection does not overwrite the pinned baseline. Compare application hashes to baseline.json before native certification. If Office changes, retain old receipts, record a new baseline revision and rerun affected comparisons; never rewrite evidence to pretend it used the new build.

## Scope

All installed-app document workflows and catalog placements remain in scope, including contextual tabs, dialogs, shortcuts, clipboard, file formats, rendering, printing and automation. App-specific family registers include non-ribbon requirements. Local native testing is a development tool; the distributed product requires no Office installation.

Newer Microsoft 365/Office features discovered online must be added as named supplement contracts (for example dynamic spills, LET/LAMBDA, newer transitions and collaboration). They need current official specifications and a capable independent oracle. A #NAME? result in Office 2016 does not validate a newer function. No cloud account, hosted conversion or subscription is introduced implicitly.

Primary desktop comparison: Windows, de-DE, exact listed binaries, exact fonts/paper/options per fixture. Additional locale acceptance: en-US, plus bidi/CJK/IME and representative locale/date systems as specified per contract. Browser release coverage: Chromium/Edge, Firefox and WebKit engines; record exact tested versions. Permission-gated enhancements require a tested baseline fallback. Installed Safari/manual platform checks are required before a claim about actual Safari behavior; Playwright WebKit alone is insufficient for that claim.

Fonts used for geometry comparisons must be identified and hashed locally. Installed fonts may be test inputs without becoming redistributed assets. Missing-font tests use explicit substitutions. Do not assume a compatible font name implies identical glyph metrics.
