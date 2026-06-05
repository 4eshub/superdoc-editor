# Security Policy

## Reporting a Vulnerability

Please report suspected vulnerabilities privately to the maintainers before opening a public issue.

Include:

- Affected version or commit.
- Steps to reproduce.
- Impact and affected browser/runtime.
- Any suggested mitigation, if known.

## Supported Versions

Security fixes are handled against the latest published source on the default branch unless a maintainer documents additional supported release branches.

## Scope

This app receives document data and files from a parent frame. Treat document content, iframe messages, and document URLs as untrusted input when changing the message bridge or editor loading behavior.
