# Security Policy

## Supported Versions

We actively provide security patches and updates for the following versions of the packages:

| Package | Supported Versions |
| :--- | :--- |
| `@hokkyss/pptx` | `>= 0.2.0` |
| `@hokkyss/pptx-reader` | `>= 0.2.0` |
| `@hokkyss/pptx-writer` | `>= 0.2.0` |
| `@hokkyss/pptx-core` | `>= 0.2.0` |

---

## Reporting a Vulnerability

We take the security of `@hokkyss/pptx` and its users seriously. If you discover a security vulnerability (such as Zip bomb expansion issues, XML external entity / XXE vulnerabilities, or prototype pollution in parsed ASTs), please report it responsibly.

### How to Report
- **Preferred Method**: Open a [Private Security Advisory](https://github.com/hokkyss/pptx-parser/security/advisories/new) on GitHub.
- **Alternative**: Email `hokkyss2@gmail.com` with the subject `[SECURITY] Vulnerability in @hokkyss/pptx`.

### What to Include
Please include:
1. A description of the vulnerability.
2. Steps or a minimal code snippet / PPTX archive to reproduce the issue.
3. Any potential impact or affected environments.

### Response Timeline
- We will acknowledge receipt of your vulnerability report within **48 hours**.
- We will provide regular status updates until a patch is verified and published to npm.
- Security advisories and patched versions will be released with appropriate CVE tracking when applicable.
