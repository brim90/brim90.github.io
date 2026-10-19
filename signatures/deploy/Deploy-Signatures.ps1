<#
.SYNOPSIS
    Pushes rendered signatures into Exchange Online mailboxes (server-side fallback).

.DESCRIPTION
    The Outlook add-in is the primary rollout method. Use this script as well
    for people the add-in can't reach (older Outlook builds, shared mailboxes
    used in Outlook on the web), or as a one-off when you first switch over
    from your old signature tool.

    1. In the Signature Manager admin console, sign in and click
       "Export all (for PowerShell push)" to download signatures-export.json.
    2. Run this script with that file.

    It sets the Outlook on the web / new Outlook signature for each mailbox.
    Exchange only honours this while roaming signatures are postponed for the
    tenant; the script checks and tells you if they're not:
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true

.PARAMETER ExportPath
    Path to signatures-export.json from the admin console.

.PARAMETER Only
    Optional list of email addresses to update instead of everyone in the file.

.EXAMPLE
    .\Deploy-Signatures.ps1 -ExportPath .\signatures-export.json -WhatIf

.EXAMPLE
    .\Deploy-Signatures.ps1 -ExportPath .\signatures-export.json -Only alex.morgan@childrensalliance.org.uk
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $ExportPath,

    [string[]] $Only,

    # Add the signature to replies and forwards as well as new emails.
    [bool] $IncludeReplies = $true
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw 'Install the Exchange Online module first: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}

$export = Get-Content -Raw -Path $ExportPath | ConvertFrom-Json
$signatures = @($export.signatures)
if ($Only) {
    $wanted = $Only | ForEach-Object { $_.ToLowerInvariant() }
    $signatures = $signatures | Where-Object { $wanted -contains $_.mail.ToLowerInvariant() }
}
Write-Host "Loaded $($signatures.Count) signatures generated at $($export.generatedAt)."

Connect-ExchangeOnline -ShowBanner:$false

$org = Get-OrganizationConfig
if (-not $org.PostponeRoamingSignaturesUntilLater) {
    Write-Warning ('Roaming signatures are on, so Exchange will ignore signatures set by this script. ' +
        'Run: Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true (or rely on the Outlook add-in).')
}

$updated = 0
$failed = @()
foreach ($sig in $signatures) {
    if (-not $PSCmdlet.ShouldProcess($sig.mail, 'Set email signature')) { continue }
    try {
        Set-MailboxMessageConfiguration -Identity $sig.mail `
            -SignatureHtml $sig.html `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $IncludeReplies
        $updated++
        Write-Host "Updated $($sig.mail)"
    } catch {
        $failed += [pscustomobject]@{ Mail = $sig.mail; Error = $_.Exception.Message }
        Write-Warning "Failed $($sig.mail): $($_.Exception.Message)"
    }
}

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done. Updated $updated, failed $($failed.Count)."
if ($failed) { $failed | Format-Table -AutoSize }
