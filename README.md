# hMailServer-Spam-Provisioner

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/paulmann/hMailServer-Spam-Provisioner)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/powershell-7.5%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/platform-Windows%20Server-blue.svg)](https://www.microsoft.com/windows/)
[![hMailServer](https://img.shields.io/badge/hMailServer-5.x-orange.svg)](https://www.hmailserver.com/)

## 📋 Table of Contents

- [🌟 Overview](#-overview)
- [🎯 Key Features](#-key-features)
- [⚙️ How It Works](#️-how-it-works)
- [📥 Installation & Usage](#-installation--usage)
  - [System Requirements](#system-requirements)
  - [Installation](#installation)
  - [Usage Examples](#usage-examples)
- [🔧 Parameters](#-parameters)
- [🔍 Debug Mode](#-debug-mode)
- [🏢 Enterprise Deployment](#-enterprise-deployment)
- [🛡️ Security Considerations](#️-security-considerations)
- [🔧 Troubleshooting](#-troubleshooting)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)
- [🙏 Acknowledgments](#-acknowledgments)

## 🌟 Overview

**hMailServer-Spam-Provisioner** is an idempotent PowerShell 7.5+ automation script that provisions spam management infrastructure across all active mailboxes in your hMailServer 5.x deployment.

For every active email account across all domains, the script automatically:
- ✉️ Creates a top-level IMAP folder `[SPAM]` (if not already present)
- 🔧 Creates an account-level filter rule `"Spam -> [SPAM]"` that moves spam messages based on:
  - **Header detection**: `X-hMailServer-Spam: YES`
  - **Subject tagging**: Subject contains `[SPAM]`

**Idempotent by design** — safe to re-run at any time. The script detects existing folders and rules, skipping them without modification or duplication.

**Perfect for**:
- 🚀 Initial hMailServer spam filter deployment
- 🔄 Bulk provisioning after adding new domains/accounts
- ✅ Audit and remediation (ensure all accounts have spam infrastructure)
- 🏗️ Infrastructure-as-Code for mail server configuration

---

## 🎯 Key Features

### ✨ **Idempotent Provisioning**
- **Zero-duplication guarantee**: Detects existing folders and rules before creation
- **Safe re-execution**: Run multiple times without breaking existing configuration
- **Audit-friendly**: Clear console output showing what was created vs. already existed

### 🏗️ **Enterprise-Ready Automation**
- **Bulk processing**: Provisions all active accounts across all domains in one run
- **Configurable parameters**: Customize folder names, tags, headers via command-line
- **Inactive account handling**: Automatically skips disabled mailboxes
- **Comprehensive logging**: Detailed console output with statistics and error tracking

### 🎨 **Modern PowerShell Design**
- **PowerShell 7.5+ features**: Leverages latest PSStyle for color-coded output
- **Secure credential handling**: Password prompt with SecureString if not provided
- **COM API integration**: Direct interaction with hMailServer 5.x COM interface
- **Type-safe enums**: hMailServer field/match/action types defined as PowerShell enums

### 🐛 **Debug Mode**
- **First-account testing**: `$DEBUG = 1` processes only the first active account
- **Verbose COM tracing**: Shows all COM method invocations for troubleshooting
- **Safe development**: Test logic without impacting entire mail server

### 📊 **Detailed Reporting**
- **Domain-level summary**: Visual breakdown by domain
- **Account-level details**: Shows each account processed with outcomes
- **Statistics dashboard**: Counts for domains, accounts, folders, rules created/existed
- **Error tracking**: Highlights failures without stopping entire batch

---

## ⚙️ How It Works

### 🔄 **Technical Architecture**

```mermaid
graph TB
    START[Script Execution] --> INIT[Initialize Parameters]
    INIT --> PASS{Password Provided?}
    PASS -->|No| PROMPT[Secure Password Prompt]
    PASS -->|Yes| CONNECT[Connect to hMailServer COM API]
    PROMPT --> CONNECT
    
    CONNECT --> AUTH{Authenticate?}
    AUTH -->|Failed| ERROR1[Exit: Authentication Failed]
    AUTH -->|Success| MODE{Debug Mode?}
    
    MODE -->|DEBUG=1| DEBUG_INIT[Initialize Debug Mode<br/>First Active Account Only]
    MODE -->|DEBUG=0| NORMAL_INIT[Initialize Normal Mode<br/>All Active Accounts]
    
    DEBUG_INIT --> ITERATE_DOMAINS
    NORMAL_INIT --> ITERATE_DOMAINS
    
    ITERATE_DOMAINS[Iterate Through Domains] --> DOMAIN_HEADER[Display Domain Header]
    DOMAIN_HEADER --> ITERATE_ACCOUNTS[Iterate Through Accounts]
    
    ITERATE_ACCOUNTS --> ACTIVE_CHECK{Account Active?}
    ACTIVE_CHECK -->|No| SKIP[Skip: Display Inactive Message]
    ACTIVE_CHECK -->|Yes| DEBUG_CHECK{DEBUG Mode &<br/>Account Processed?}
    
    DEBUG_CHECK -->|Yes| SUMMARY[Generate Summary]
    DEBUG_CHECK -->|No| ACCOUNT_HEADER[Display Account Header]
    
    ACCOUNT_HEADER --> FOLDER_CHECK{Folder Exists?}
    FOLDER_CHECK -->|Yes| FOLDER_EXISTS[Display: Folder Already Exists]
    FOLDER_CHECK -->|No| CREATE_FOLDER[Create IMAP Folder]
    
    CREATE_FOLDER --> FOLDER_SAVE[Save Folder<br/>ParentID=0 Top Level]
    FOLDER_SAVE --> FOLDER_SUCCESS[Display: Folder Created]
    FOLDER_EXISTS --> RULE_CHECK
    FOLDER_SUCCESS --> RULE_CHECK
    
    RULE_CHECK{Rule Exists?} -->|Yes| RULE_EXISTS[Display: Rule Already Exists]
    RULE_CHECK -->|No| CREATE_RULE[Create Rule Object]
    
    CREATE_RULE --> RULE_CONFIG[Configure Rule:<br/>UseAND=false OR logic]
    RULE_CONFIG --> ADD_CRIT1[Add Criterion 1:<br/>Header Match<br/>X-hMailServer-Spam: YES]
    ADD_CRIT1 --> ADD_CRIT2[Add Criterion 2:<br/>Subject Contains<br/>[SPAM]]
    ADD_CRIT2 --> ADD_ACTION1[Add Action 1:<br/>MoveToImapFolder<br/>[SPAM]]
    ADD_ACTION1 --> ADD_ACTION2[Add Action 2:<br/>StopRuleProcessing]
    ADD_ACTION2 --> RULE_SAVE[Save Rule]
    RULE_SAVE --> RULE_SUCCESS[Display: Rule Created]
    
    RULE_EXISTS --> STATS_UPDATE
    RULE_SUCCESS --> STATS_UPDATE[Update Statistics]
    SKIP --> NEXT_ACCOUNT
    STATS_UPDATE --> NEXT_ACCOUNT{More Accounts?}
    
    NEXT_ACCOUNT -->|Yes| ITERATE_ACCOUNTS
    NEXT_ACCOUNT -->|No| NEXT_DOMAIN{More Domains?}
    
    NEXT_DOMAIN -->|Yes| ITERATE_DOMAINS
    NEXT_DOMAIN -->|No| SUMMARY
    
    SUMMARY[Display Summary Statistics] --> EXIT_CODE{Errors > 0?}
    EXIT_CODE -->|Yes| EXIT_FAIL[Exit Code: 1]
    EXIT_CODE -->|No| EXIT_SUCCESS[Exit Code: 0]
    
    style START fill:#e3f2fd
    style CONNECT fill:#fff3e0
    style DEBUG_INIT fill:#ffe0b2
    style NORMAL_INIT fill:#c8e6c9
    style CREATE_FOLDER fill:#b2dfdb
    style CREATE_RULE fill:#b2ebf2
    style FOLDER_SUCCESS fill:#c8e6c9
    style RULE_SUCCESS fill:#c8e6c9
    style FOLDER_EXISTS fill:#e0e0e0
    style RULE_EXISTS fill:#e0e0e0
    style SKIP fill:#e0e0e0
    style ERROR1 fill:#ffcdd2
    style EXIT_FAIL fill:#ffcdd2
    style EXIT_SUCCESS fill:#c8e6c9
    style SUMMARY fill:#e1bee7
```
## 📥 Installation & Usage

### System Requirements
- **PowerShell**: Version 7.5 or higher (required for `$PSStyle` and modern COM interaction).
- **hMailServer**: Version 5.x (tested on latest 5.4+ builds).
- **Permissions**: Administrator privileges on the hMailServer host.
- **Access**: Direct access to hMailServer COM API.

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/paulmann/hMailServer-Spam-Provisioner.git
   ```
2. Navigate to the script directory:
   ```bash
   cd hMailServer-Spam-Provisioner
   ```

### Usage Examples

#### **Standard Execution**
Prompts for the hMailServer administrator password securely:
```powershell
.\hms-spam-provisioner.ps1
```

#### **Automated Execution (CI/CD / Scheduled Tasks)**
Provide credentials via parameters:
```powershell
.\hms-spam-provisioner.ps1 -AdminUser 'Administrator' -AdminPassword 'YourSecurePassword'
```

#### **Custom Folder & Tags**
Provision with a custom IMAP folder name and subject tag:
```powershell
.\hms-spam-provisioner.ps1 -SpamFolderName '[Junk]' -SpamSubjectTag '{SPAM}'
```

---

## 🔧 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `AdminUser` | `string` | `Administrator` | hMailServer administrative username. |
| `AdminPassword` | `string` | *Prompted* | Password for the hMailServer administrator account. |
| `SpamFolderName` | `string` | `[SPAM]` | The name of the top-level IMAP folder to create. |
| `SpamSubjectTag` | `string` | `[SPAM]` | Subject tag to look for in the rule criteria. |
| `SpamHeaderField` | `string` | `X-hMailServer-Spam` | Custom header field name to match. |
| `SpamHeaderValue` | `string` | `YES` | Value of the header field that indicates spam. |

---

## 🔍 Debug Mode

The script includes a built-in safety mechanism for testing.
Set `[int]$DEBUG = 1` at the top of the script (line 39) to:
1. Process **ONLY the first active account** found.
2. Enable **Verbose COM tracing** (outputs every method call to the console).
3. Stop execution immediately after the first account (or first error).

---

## 🏢 Enterprise Deployment

For large-scale environments, this script can be integrated into your maintenance workflows:

### **Scheduled Maintenance**
Run the script weekly via Task Scheduler to ensure new accounts are always provisioned:
```powershell
# Task action example
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\hms-spam-provisioner.ps1" -AdminPassword '...'
```

### **Post-Migration Audit**
After migrating users to hMailServer, run this script to guarantee consistent spam management across the entire user base.

---

## 🛡️ Security Considerations

- **Credential Handling**: The script uses `Read-Host -AsSecureString` for password entry by default. Avoid hardcoding passwords in scripts; use environment variables or secret management modules in production.
- **Administrator Privileges**: Requires elevation to interact with the hMailServer COM API.
- **Idempotency**: Prevents configuration drift and prevents rule/folder duplication if the script is interrupted.

---

## 🔧 Troubleshooting

- **"Failed to connect to COM API"**: Ensure hMailServer is installed on the local machine and the COM API is not disabled in `hMailServer.INI`.
- **"Authentication Failed"**: Verify the `AdminUser` and `AdminPassword`. Note that this is the hMailServer admin password, not the Windows admin password.
- **Permission Errors**: Ensure you are running PowerShell as **Administrator**.

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **hMailServer Team**: For providing a robust, free mail server with a powerful COM API.
- **PowerShell Community**: For the continuous inspiration in infrastructure automation.

---
**⭐ Star this repository if you find it useful for your hMailServer deployment!**
**Author**: [Mikhail Deynekin](https://deynekin.com) | [GitHub](https://github.com/paulmann)
**Version**: 1.0.0 | **Enterprise Ready**
