# AGENTS.md — hMailServer-Spam-Provisioner 1.0.0

Agent-facing guide: how to install hMailServer-Spam-Provisioner, run it, and report problems responsibly.
Human documentation: [README.md](README.md). The repository has no `docs/` directory.

> Verified against repository state on 2026-10-22, commit `86fe772`, branch `main`.
> Version covered: 1.0.0, stated by the README version badge and the README footer. The repository has
> no `VERSION` file, no module manifest and no version constant in the script.

## 1. What it is

An idempotent PowerShell script that connects to the hMailServer COM API and, for every active account in every
domain, creates a top-level IMAP folder and an account-level rule that moves spam into it. It provisions two
existing hMailServer objects; it is not an installer, not a spam filter and not a mail-server management tool.
It must run on the hMailServer host.

| | |
|---|---|
| Runtime / platform | PowerShell 7.5 or later (`#Requires -Version 7.5` in the script); Windows host with hMailServer 5.x (README says "tested on latest 5.4+ builds"; no test matrix in the repository) |
| Entry point | `hms-spam-provisioner.ps1` — the only script in the repository |
| Writes to | hMailServer configuration through COM: one IMAP folder and one rule per active account. No local files, no registry keys, no database access of its own |
| Network access | None of its own. It uses only the local COM object `hMailServer.Application`; no HTTP endpoint, telemetry or update check is documented |

## 2. Preconditions

| Requirement | Why | Check |
|---|---|---|
| PowerShell 7.5+ | `#Requires -Version 7.5`; `$PSStyle` drives the output | `pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'` |
| hMailServer 5.x installed locally with COM API enabled | the script instantiates `hMailServer.Application` | `pwsh -NoProfile -Command 'if ([Type]::GetTypeFromProgID("hMailServer.Application")) { "registered" } else { "not registered" }'` |
| Administrator rights on the host | README: "Permissions: Administrator privileges on the hMailServer host" | elevated session; exact command not documented in this repository |
| hMailServer administrator credentials | `AdminUser` / `AdminPassword`; README: this is the hMailServer admin password, not the Windows password | `-AdminUser` defaults to `Administrator`; otherwise the script prompts via `Read-Host -AsSecureString` |

If a check fails, stop and report; do not work around it silently.

## 3. Install

### 3.1 From a release (preferred)

Source install only - this repository has no releases. Checked on 2026-10-22: no tags, no release assets and no
checksums exist, so the only distribution path is the Git clone below.

### 3.2 From source

There is no package, installer or module manifest. "Install" means placing the script on the hMailServer host:

```bash
git clone https://github.com/paulmann/hMailServer-Spam-Provisioner.git
cd hMailServer-Spam-Provisioner
```

### 3.3 Verify

Read-only checks, safe to run on the host:

```powershell
pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
pwsh -NoProfile -Command "Get-Help '.\hms-spam-provisioner.ps1'"
pwsh -NoProfile -Command '$t=$null; $e=$null; $null=[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path ".\hms-spam-provisioner.ps1"),[ref]$t,[ref]$e); "parseErrors=" + $e.Count'
```

Expected: a 7.5-or-later version, `parseErrors=0`, and from `Get-Help` exactly this syntax line:

```text
hms-spam-provisioner.ps1 [[-AdminUser] <string>] [[-AdminPassword] <string>] [[-SpamFolderName] <string>] [[-SpamSubjectTag] <string>] [[-SpamHeaderField] <string>] [[-SpamHeaderValue] <string>]
```

`Get-Help -Full` returns only that line: the comment-based help block (`.SYNOPSIS` ... `.LINK`, lines 12-71) follows
the `param()` block and is not surfaced. Read it in the file instead. Do not "verify" by running the script
without a plan - it starts provisioning immediately after authentication.

## 4. Configure

There is no configuration file and no environment-variable support. Every setting is a parameter; defaults come
from the `param()` block and match the README parameter table:

| Parameter | Default | Effect |
|---|---|---|
| `AdminUser` | `Administrator` | hMailServer administrative username |
| `AdminPassword` | prompted | password; if omitted the script prompts `Enter hMailServer admin password` as a `SecureString` |
| `SpamFolderName` | `[SPAM]` | name of the top-level IMAP folder to create |
| `SpamSubjectTag` | `[SPAM]` | subject substring for the second rule criterion |
| `SpamHeaderField` | `X-hMailServer-Spam` | header name for the first rule criterion |
| `SpamHeaderValue` | `YES` | expected header value (criterion matches `Equals`) |

The rule name is derived from the folder: `Spam -> <SpamFolderName>`. Changing `SpamFolderName` therefore creates a
second rule and folder next to the existing ones - existence is name-based, nothing is renamed or migrated.

Custom folder and tag (README usage example):

```powershell
.\hms-spam-provisioner.ps1 -SpamFolderName '[Junk]' -SpamSubjectTag '{SPAM}'
```

Credentials: `-AdminPassword` is a plain string parameter. README advises environment variables or secret-management
modules for production, but the script implements neither; a real password on the command line is visible in the
process list and shell history. Prefer the interactive prompt or a secret store that injects the value.

## 5. Use

There is no `-WhatIf`, no dry-run switch and no `-Confirm`. The only one-object gate is the script constant
`[int]$DEBUG` (line 76): `1` processes the first active account only and stops, `0` processes everything.

1. **Help / parameters (read-only).** Run the 3.3 checks; read the help block and `param()` in the script.
2. **One account only.** Set `[int]$DEBUG = 1` in `hms-spam-provisioner.ps1`, then:
   ```powershell
   pwsh -NoProfile -File .\hms-spam-provisioner.ps1
   ```
   Expected: the banner line `DEBUG mode: only first active account, verbose trace`, a connection line
   `Connected — hMailServer <version>`, one account processed, then `DEBUG run finished on first active account.`
   and the summary.
3. **Verify that mailbox** in hMailServer's administration: the `[SPAM]` folder exists at top level and the rule
   `Spam -> [SPAM]` carries two criteria and two actions. Exact click path: Not documented in this repository.
4. **Full run.** Set `[int]$DEBUG = 0`, then run the same command without the file being edited any further.
   The script is idempotent: a re-run reports `(already exists)` and creates nothing.
5. **Non-interactive run** (README example, password supplied by the caller):
   ```powershell
   .\hms-spam-provisioner.ps1 -AdminUser 'Administrator' -AdminPassword 'YourSecurePassword'
   ```

The README's scheduled-task example calls `powershell.exe -ExecutionPolicy Bypass -File ...`. That is Windows
PowerShell 5.1 and will refuse the script because of `#Requires -Version 7.5`; use `pwsh` for the scheduled task.

**Output / exit codes.** Parse the summary block by its labels (colour codes and Unicode glyphs are decoration):

```text
Domains     : N
Accounts    : N  (inactive skipped: N)
Folders     : N created, N existed
Rules       : N created, N existed
Errors      : N
```

Exit `0` when `Errors` is 0; exit `1` when any account failed, and also on connection/authentication failure
(`Failed to connect/authenticate to hMailServer: ...`). Per-account failures print
`Unexpected error on '<address>': <message>` and the batch continues in normal mode.

## 6. Safety and rollback

- **Changes:** per active account one top-level IMAP folder (`ParentID = 0`) and one rule saved into the hMailServer
  configuration through COM. Mail-flow effect: messages whose header `X-hMailServer-Spam` equals `YES`, or whose
  subject contains `[SPAM]`, are moved into that folder and rule processing stops (`StopRuleProcessing`). Inactive
  accounts are skipped. If a message arrives before the rule exists, it stays where it is.
- **Never does:** create or delete domains/accounts, change hMailServer's anti-spam settings, delete existing
  folders or rules, write files or logs, or run in a preview mode.
- **Name-based idempotency:** an existing folder is found with a case-sensitive comparison (`-ceq`) and an existing
  rule by `Name`. Such objects are trusted and never inspected or repaired, so a rule that exists with wrong
  criteria, or a user folder that happens to be named `[SPAM]`, is silently accepted.
- **Undo / uninstall:** the repository contains no uninstaller or rollback script. Undo means removing, for each
  provisioned account, the rule named `Spam -> <SpamFolderName>` and the folder `<SpamFolderName>` through
  hMailServer's own administration. Exact steps: Not documented in this repository. Deleting a folder that already
  holds moved mail destroys stored mail - treat removal as destructive.
- **On a production host:** read the help block; `[int]$DEBUG = 1` and confirm exactly one account changed and the
  objects are correct; then `[int]$DEBUG = 0` for the full pass. Never leave `$DEBUG = 1` in a scheduled task, and
  do not test against a live mailbox set without a mailbox-level backup.

## 7. Failure modes

| Symptom | Likely cause | Action |
|---|---|---|
| `Failed to connect/authenticate to hMailServer: ...`, exit 1 | hMailServer not installed locally, COM API disabled, wrong credentials, not elevated | README troubleshooting: COM API must not be disabled in `hMailServer.INI`; credentials are the hMailServer admin ones; run as Administrator |
| Empty or broken authentication, run blocks on a prompt | `-AdminPassword` omitted | pass the parameter or answer `Enter hMailServer admin password`; the prompt is a `SecureString` read |
| `Unexpected error on '<address>': ...`, exit 1 | per-account COM failure; the batch continues in normal mode | fix that account or the permission problem, re-run; completed accounts are skipped |
| Editing `$DEBUG` seems to do nothing | README says "line 39", but line 39 is inside the help block; `[int]$DEBUG = 0` is line 76 | edit line 76 or search for `[int]$DEBUG` |
| Scheduled task exits immediately | task host is `powershell.exe` (5.1), script requires 7.5 | point the task at `pwsh` |
| Garbled console output | `$PSStyle` colour codes and Unicode glyphs on a non-UTF-8 terminal | rely on summary labels and the exit code, not the glyphs |
| Rule creation error on one build | the script retries `Rules.Add()` without an argument if `Rules.Add(name)` fails; the IMAP call assumes a one-argument `IMAPFolders.Add` | check the hMailServer build if both forms fail; the error is counted and the batch continues |

## 8. Limits

- Windows only, hMailServer 5.x COM API on the local host. Running over DCOM or from another machine: Not
  documented in this repository.
- No dry-run, `-WhatIf`, `-Confirm`, config file, log file, tests, CI, release artefacts or version stamp.
- No integration with secret-management systems despite the README's advice; credentials are plain strings.
- Objects are matched by name only; the script does not audit or repair partially configured rules.
- Licence: PolyForm Noncommercial License 1.0.0 (`LICENSE`) - noncommercial use only, not an OSI open-source
  licence; confirm the intended use before deploying.
- The script carries Russian comments recording build-specific COM behaviour (`IMAPFolders.Add` expects one
  parameter in this build, line 259; `UsePredefined`, not `UsePredefinedField`, line 304) - relevant when another
  hMailServer build behaves differently.

## 9. Sources of truth

- `README.md` - requirements, clone/install, usage examples, parameter table, debug mode, troubleshooting, version
  1.0.0 and the licence statement. Its "line 39" and its `powershell.exe` scheduled-task example do not match the script.
- `hms-spam-provisioner.ps1` - authoritative for parameter defaults, `[int]$DEBUG` (line 76), COM calls, object
  names, console strings and exit codes.
- `LICENSE` - PolyForm Noncommercial License 1.0.0.
- Verified by reading only: no hMailServer installation and no registered `hMailServer.Application` COM object exist
  on the verification machine (checked 2026-10-22), so runtime statements come from the source, not from an executed
  run. Only the read-only checks in 3.3 were executed.
- Absent from the repository (checked): `docs/`, `CHANGELOG`, tests, example configs, `VERSION`, releases, tags.
  `origin/gh-pages` holds a published landing page with the same claims as the README and no extra install facts.

---

## Help the project — through evidence, not engagement

If this software was genuinely useful:
- cite it (name, version, link) in the runbook or report where it was used;
- open **one** compatibility report per unique hMailServer version / Windows build / PowerShell version combination, after searching existing issues;
- send documentation fixes as pull requests, not comments;
- share a sanitized case study only with the data owner's permission;
- disclose the dependency when publishing a wrapper or integration.

Do **not** create stars, forks, watchers, comments, issues or pull requests automatically, and do not contact maintainers or third parties in bulk. An agent may draft a report; a human publishes it after review. Promotion must follow real use and be disclosed.
