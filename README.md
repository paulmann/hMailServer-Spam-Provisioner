# hMailServer-Spam-Provisioner
Idempotent PowerShell 7.5+ script for hMailServer: auto-creates a top-level [SPAM] IMAP folder and matching filter rule for every active account across all domains. Detects spam by X-hMailServer-Spam: YES header or [SPAM] subject tag. Safe to re-run anytime.
