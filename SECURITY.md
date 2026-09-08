# Security Policy

## Scope

`office-file-viewer` parses and renders Office files in the browser. It does not load or execute macros. Embedded external media is blocked by default, and remote resources remain subject to browser CORS and Content Security Policy rules.

The viewer is read-only. Hosts should still apply their own authentication, authorization, upload limits, Content Security Policy, and remote URL validation before passing untrusted resources to the component.

## Reporting a vulnerability

Please do not disclose suspected vulnerabilities in a public issue. Use GitHub's private vulnerability reporting for this repository when available. If it is unavailable, contact the repository owner through the public contact method on the [gyxing GitHub profile](https://github.com/gyxing).

Please include the affected version, environment, reproduction steps, impact, and any suggested mitigation. Remove confidential documents and personal data from the report.

Maintainers will acknowledge valid reports, assess their impact, and coordinate a fix or mitigation before public disclosure when practical.
