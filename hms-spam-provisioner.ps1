#Requires -Version 7.5

param(
    [string]$AdminUser       = 'Administrator',
    [string]$AdminPassword,
    [string]$SpamFolderName  = '[SPAM]',
    [string]$SpamSubjectTag  = '[SPAM]',
    [string]$SpamHeaderField = 'X-hMailServer-Spam',
    [string]$SpamHeaderValue = 'YES'
)

<#
.SYNOPSIS
    hMailServer SPAM Folder & Rule Provisioner

.DESCRIPTION
    Idempotent PowerShell 7.5+ provisioner for hMailServer 5.x.
    For every active mailbox across all domains it:
      • Creates a top-level IMAP folder "[SPAM]"  (if not present)
      • Creates an account-level filter rule "Spam -> [SPAM]"
        that moves messages matching ANY of the following criteria:
          – Header  X-hMailServer-Spam: YES
          – Subject contains [SPAM]
    Safe to re-run at any time — already-existing folders and rules
    are detected and skipped without modification.

.NOTES
    Author   : Mikhail Deynekin (Михаил Дейнекин)
    Website  : https://deynekin.com
    E-mail   : mid1977@gmail.com
    GitHub   : https://github.com/paulmann/hMailServer-Spam-Provisioner

    Requirements:
      • PowerShell 7.5 or later
      • hMailServer 5.x with COM API enabled
      • Must be run on the hMailServer host as Administrator

    Debug mode ($DEBUG = 1):
      Processes ONLY the first active account with verbose COM tracing.
      Set $DEBUG = 0 (default) for normal operation across all accounts.

.PARAMETER AdminUser
    hMailServer administrator username. Default: 'Administrator'

.PARAMETER AdminPassword
    hMailServer administrator password.
    If omitted, the script will prompt securely at runtime.

.PARAMETER SpamFolderName
    Name of the IMAP folder to create for spam. Default: '[SPAM]'

.PARAMETER SpamSubjectTag
    Subject prefix that triggers the spam rule. Default: '[SPAM]'

.PARAMETER SpamHeaderField
    Mail header field name to match. Default: 'X-hMailServer-Spam'

.PARAMETER SpamHeaderValue
    Expected value of the spam header field. Default: 'YES'

.EXAMPLE
    # Run with default settings (password prompted at runtime):
    .\hms-spam-provisioner.ps1

.EXAMPLE
    # Run with explicit credentials and a custom folder name:
    .\hms-spam-provisioner.ps1 -AdminUser 'admin' -AdminPassword 'secret' -SpamFolderName '[Junk]'

.LINK
    https://github.com/paulmann/hms-spam-provisioner
#>

# DEBUG MODE
# 0 = normal: process all accounts, concise logging
# 1 = debug: process ONLY FIRST active account, verbose COM tracing
[int]$DEBUG = 0

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $AdminPassword) {
    $sec = Read-Host "Enter hMailServer admin password" -AsSecureString
    $AdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    )
}

# ─────────────────────────────────────────────────────────────────────────────
# hMailServer enums (from hMailServer.idl) [web:28]
# ─────────────────────────────────────────────────────────────────────────────
enum HmsField {
    Unknown          = 0
    From             = 1
    To               = 2
    CC               = 3
    Subject          = 4
    Body             = 5
    Size             = 6
    Recipients       = 7
    DeliveryAttempts = 8
}

enum HmsMatch {
    Unknown     = 0
    Equals      = 1
    Contains    = 2
    LessThan    = 3
    GreaterThan = 4
    RegEx       = 5
    NotContains = 6
    NotEquals   = 7
    Wildcard    = 8
}

enum HmsAction {
    Unknown            = 0
    DeleteEmail        = 1
    ForwardEmail       = 2
    Reply              = 3
    MoveToImapFolder   = 4
    RunScript          = 5
    StopRuleProcessing = 6
    SetHeaderValue     = 7
    SendUsingRoute     = 8
    CreateCopy         = 9
    BindToAddress      = 10
}

# ─────────────────────────────────────────────────────────────────────────────
# Styling helpers ($PSStyle) [web:62]
# ─────────────────────────────────────────────────────────────────────────────
$S      = $PSStyle
$Muted  = $S.Foreground.BrightBlack
$Strong = $S.Bold + $S.Foreground.White

function Write-Banner {
    $line = '─' * 72
    Write-Host ''
    Write-Host ("  {0}{1}{2}" -f $S.Foreground.Cyan, $line, $S.Reset)
    Write-Host ("  {0}hMailServer · SPAM Folder & Rule Provisioner{1}" -f $Strong, $S.Reset)
    if ($DEBUG -eq 1) {
        Write-Host ("  {0}DEBUG mode: only first active account, verbose trace{1}" -f $S.Foreground.Yellow, $S.Reset)
    } else {
        Write-Host ("  {0}Normal mode: all active accounts, concise log{1}" -f $S.Foreground.Green, $S.Reset)
    }
    Write-Host ("  {0}{1}{2}" -f $S.Foreground.Cyan, $line, $S.Reset)
    Write-Host ''
}

function Write-DomainHeader([string]$Name) {
    Write-Host ""
    Write-Host ("  {0}◈{1} Domain: {2}" -f $S.Foreground.Magenta, $S.Reset, $Name)
}

function Write-AccountHeader([string]$Address) {
    Write-Host ("    {0}▶{1} {2}" -f $S.Foreground.Cyan, $S.Reset, $Address)
}

function Write-Info([string]$Message) {
    Write-Host ("        {0}i{1} {2}" -f $S.Foreground.Cyan, $S.Reset, $Message)
}

function Write-Created([string]$Message) {
    Write-Host ("        {0}✔{1} {2}" -f $S.Foreground.Green, $S.Reset, $Message)
}

function Write-Exists([string]$Message) {
    Write-Host ("        {0}⊘{1} {2} {0}(already exists){1}" -f $Muted, $S.Reset, $Message)
}

function Write-Inactive([string]$Address) {
    Write-Host ("    {0}⊘{1} Skipping inactive account: {2}" -f $Muted, $S.Reset, $Address)
}

function Write-Err([string]$Message) {
    Write-Host ("        {0}✘{1} {0}{2}{1}" -f $S.Foreground.Red, $S.Reset, $Message)
}

function Write-Debug([string]$Message) {
    if ($DEBUG -eq 1) {
        Write-Host ("        {0}DBG{1} {2}" -f $S.Foreground.Yellow, $S.Reset, $Message)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COM helpers
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-ComMethod {
    param(
        [Parameter(Mandatory)]$ComObject,
        [Parameter(Mandatory)][string]$MethodName,
        [object[]]$Arguments = @()
    )

    $argsText = if ($Arguments -and $Arguments.Count -gt 0) {
        ($Arguments | ForEach-Object { "'{0}'" -f $_ }) -join ', '
    } else {
        ''
    }
    Write-Debug ("COM: {0}.{1}({2})" -f $ComObject.GetType().Name, $MethodName, $argsText)

    return $ComObject.GetType().InvokeMember(
        $MethodName,
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null,
        $ComObject,
        $Arguments
    )
}

function Find-ImapFolder {
    param(
        [Parameter(Mandatory)]$Folders,
        [Parameter(Mandatory)][string]$Name
    )

    for ($i = 0; $i -lt $Folders.Count; $i++) {
        $f = $Folders.Item($i)
        Write-Debug ("Checking IMAP folder index {0}: '{1}'" -f $i, $f.Name)
        if ($f.Name -ceq $Name) {
            return $f
        }
    }

    return $null
}

function Find-Rule {
    param(
        [Parameter(Mandatory)]$Rules,
        [Parameter(Mandatory)][string]$Name
    )

    for ($i = 0; $i -lt $Rules.Count; $i++) {
        $r = $Rules.Item($i)
        Write-Debug ("Checking rule index {0}: '{1}'" -f $i, $r.Name)
        if ($r.Name -eq $Name) {
            return $r
        }
    }

    return $null
}

function Ensure-TopLevelFolder {
    param(
        [Parameter(Mandatory)]$Account,
        [Parameter(Mandatory)][string]$FolderName
    )

    $existing = Find-ImapFolder -Folders $Account.IMAPFolders -Name $FolderName
    if ($existing) {
        return $false
    }

    Write-Debug ("Creating IMAP folder '{0}' (top-level)" -f $FolderName)

    # В вашей сборке hMailServer IMAPFolders.Add ожидает 1 параметр: имя папки
    $folder = Invoke-ComMethod -ComObject $Account.IMAPFolders -MethodName 'Add' -Arguments @($FolderName)

    # На всякий случай ещё раз проставим имя
    $folder.Name = $FolderName
    # ParentID по умолчанию = 0 → верхний уровень
    $folder.Save()

    return $true
}

function Ensure-SpamRule {
    param(
        [Parameter(Mandatory)]$Account,
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)][string]$SubjectTag,
        [Parameter(Mandatory)][string]$HeaderField,
        [Parameter(Mandatory)][string]$HeaderValue
    )

    $ruleName = "Spam -> $FolderName"
    $existing = Find-Rule -Rules $Account.Rules -Name $ruleName
    if ($existing) {
        return $false
    }

    $rules = $Account.Rules

    Write-Debug ("Creating rule '{0}'" -f $ruleName)
    try {
        $rule = Invoke-ComMethod -ComObject $rules -MethodName 'Add' -Arguments @($ruleName)
        Write-Debug "Rules.Add(name) succeeded"
    }
    catch {
        Write-Debug ("Rules.Add(name) failed: {0}. Retrying with Add()." -f $_.Exception.Message)
        $rule = Invoke-ComMethod -ComObject $rules -MethodName 'Add'
        $rule.Name = $ruleName
    }

    $rule.Active = $true
    $rule.UseAND = $false   # OR

    # Criterion 1 – header = YES
    Write-Debug ("Adding criterion #1: header '{0}' = '{1}'" -f $HeaderField, $HeaderValue)
    $c1 = Invoke-ComMethod -ComObject $rule.Criterias -MethodName 'Add'
    $c1.UsePredefined = $false          # ← ВАЖНО: UsePredefined, не UsePredefinedField
    $c1.HeaderField   = $HeaderField
    $c1.MatchType     = [int][HmsMatch]::Equals
    $c1.MatchValue    = $HeaderValue
    $c1.Save()

    # Criterion 2 – Subject contains [SPAM]
    Write-Debug ("Adding criterion #2: Subject contains '{0}'" -f $SubjectTag)
    $c2 = Invoke-ComMethod -ComObject $rule.Criterias -MethodName 'Add'
    $c2.UsePredefined  = $true          # ← тоже UsePredefined
    $c2.PredefinedField = [int][HmsField]::Subject
    $c2.MatchType      = [int][HmsMatch]::Contains
    $c2.MatchValue     = $SubjectTag
    $c2.Save()

    # Action 1 – Move to IMAP folder
    Write-Debug ("Adding action #1: MoveToImapFolder '{0}'" -f $FolderName)
    $a1 = Invoke-ComMethod -ComObject $rule.Actions -MethodName 'Add'
    $a1.Type       = [int][HmsAction]::MoveToImapFolder
    $a1.IMAPFolder = $FolderName
    $a1.Save()

    # Action 2 – Stop rule processing [web:28]
    Write-Debug "Adding action #2: StopRuleProcessing"
    $a2 = Invoke-ComMethod -ComObject $rule.Actions -MethodName 'Add'
    $a2.Type = [int][HmsAction]::StopRuleProcessing
    $a2.Save()

    $rule.Save()

    return $true
}

function Invoke-ProvisionAccount {
    param(
        [Parameter(Mandatory)]$Account,
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)][string]$SubjectTag,
        [Parameter(Mandatory)][string]$HeaderField,
        [Parameter(Mandatory)][string]$HeaderValue,
        [Parameter(Mandatory)][hashtable]$Stats
    )

    Write-AccountHeader $Account.Address

    if (Ensure-TopLevelFolder -Account $Account -FolderName $FolderName) {
        $Stats.FoldersCreated++
        Write-Created ("Folder '{0}'" -f $FolderName)
    }
    else {
        $Stats.FoldersExisted++
        Write-Exists ("Folder '{0}'" -f $FolderName)
    }

    if (Ensure-SpamRule -Account $Account -FolderName $FolderName -SubjectTag $SubjectTag -HeaderField $HeaderField -HeaderValue $HeaderValue) {
        $Stats.RulesCreated++
        Write-Created ("Rule 'Spam -> {0}'" -f $FolderName)
    }
    else {
        $Stats.RulesExisted++
        Write-Exists ("Rule 'Spam -> {0}'" -f $FolderName)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Main script
# ─────────────────────────────────────────────────────────────────────────────
Write-Banner

Write-Host ("  {0}▶{1} Connecting to hMailServer COM API..." -f $S.Foreground.Cyan, $S.Reset)
try {
    $hms = New-Object -ComObject 'hMailServer.Application'
    $null = $hms.Authenticate($AdminUser, $AdminPassword)
    $version = $hms.Version
    Write-Host ("    {0}✔{1} Connected — hMailServer {2}" -f $S.Foreground.Green, $S.Reset, $version)
}
catch {
    Write-Err ("Failed to connect/authenticate to hMailServer: {0}" -f $_.Exception.Message)
    exit 1
}

$stats = [ordered]@{
    Domains        = 0
    Accounts       = 0
    Skipped        = 0
    FoldersCreated = 0
    FoldersExisted = 0
    RulesCreated   = 0
    RulesExisted   = 0
    Errors         = 0
}

$domains          = $hms.Domains
$debugAccountDone = $false

for ($d = 0; $d -lt $domains.Count; $d++) {
    $domain = $domains.Item($d)
    $stats.Domains++
    Write-DomainHeader $domain.Name

    $accounts = $domain.Accounts
    for ($a = 0; $a -lt $accounts.Count; $a++) {
        $account = $accounts.Item($a)

        if (-not $account.Active) {
            $stats.Skipped++
            Write-Inactive $account.Address
            continue
        }

        if ($DEBUG -eq 1 -and $debugAccountDone) {
            Write-Info "DEBUG=1: skipping remaining accounts"
            break
        }

        $stats.Accounts++

        try {
            Invoke-ProvisionAccount `
                -Account     $account        `
                -FolderName  $SpamFolderName  `
                -SubjectTag  $SpamSubjectTag  `
                -HeaderField $SpamHeaderField `
                -HeaderValue $SpamHeaderValue `
                -Stats       $stats

            if ($DEBUG -eq 1) {
                $debugAccountDone = $true
                Write-Info "DEBUG run finished on first active account."
                break
            }
        }
        catch {
            $stats.Errors++
            Write-Err ("Unexpected error on '{0}': {1}" -f $account.Address, $_.Exception.Message)
            if ($DEBUG -eq 1) {
                Write-Info "DEBUG=1: stopping after first error."
                break
            }
        }
    }

    if ($DEBUG -eq 1 -and $debugAccountDone) {
        break
    }
}

# Summary
$line = '─' * 72
Write-Host ""
Write-Host ("  {0}{1}{2}" -f $S.Foreground.Cyan, $line, $S.Reset)
Write-Host ("  {0}Summary{1}" -f $Strong, $S.Reset)
Write-Host ("  {0}{1}{2}" -f $S.Foreground.Cyan, $line, $S.Reset)
Write-Host ("  Domains     : {0}" -f $stats.Domains)
Write-Host ("  Accounts    : {0}  (inactive skipped: {1})" -f $stats.Accounts, $stats.Skipped)
Write-Host ("  Folders     : {0} created, {1} existed" -f $stats.FoldersCreated, $stats.FoldersExisted)
Write-Host ("  Rules       : {0} created, {1} existed" -f $stats.RulesCreated, $stats.RulesExisted)
Write-Host ("  Errors      : {0}" -f $stats.Errors)
Write-Host ("  {0}{1}{2}" -f $S.Foreground.Cyan, $line, $S.Reset)
Write-Host ""

exit ($(if ($stats.Errors -gt 0) { 1 } else { 0 }))
