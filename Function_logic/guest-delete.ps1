<#
.SYNOPSIS
    Delete Azure AD guest users whose invitation is PendingAcceptance over x days/aaccepted but inactive over y days using certificate thumbprint (app-only).
.DESCRIPTION
    - Authenticates app-only using certificate thumbprint (Connect-MgGraph -ClientId -TenantId -CertificateThumbprint)/can also use managed identity for authentication.
    - Reads UPNs/emails from an .CSV file.
    - For each email: fetches full user properties and if externalUserState == 'PendingAcceptance' or externalUserState == 'Accepted' but inactive over y days or externalUserState == ''(legacy B2B guests or manually created guests), deletes the guest user (if deletion enabled).
    - Logs actions, handles errors, and writes a CSV summary.
.NOTES
    - Edit the Configuration block below before running, or call the script from PowerShell and edit variables in the script.
#>
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,
    [Parameter(Mandatory)]
    [string]$InputBlobName
    )
#region Configuration - EDIT BEFORE RUNNING
# Tenant / App (replace with your values)
$TenantId       = "<TENANT_ID>"
$ClientId       = "<CLIENT_ID>"
$CertThumbprint = "<CERT_THUMBPRINT>"  # thumbprint of certificate in CurrentUser\My or LocalMachine\My
# Input options - choose one:
#$ExcelPath      = $csv_path # Excel file with header column containing emails (preferred)
$ExcelSheet     = "Sheet1"
$ExcelEmailCol  = "UPN"   # header name in Excel: e.g., UPN or Email

$EmailListFile  = ""      # plain text file, one email per line (optional fallback)
$SingleEmail    = ""      # single email to process (optional fallback)

# Output
$OutputCsv      = Join-Path $PSScriptRoot "PendingInvite-Results.csv"
$TimeFormat     = "yyyy-MM-dd HH:mm:ss"

# Safety: set to $true to simulate (no deletion). Set to $false to actually delete.
$doWhatIf       = $false    
#endregion
#region Logging
function Write-Log {
    param([string]$Message, [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    $ts = (Get-Date).ToString($TimeFormat)
    switch ($Level) {
        "INFO"  { Write-Host "[$ts] [INFO]  $Message" -ForegroundColor Cyan }
        "WARN"  { Write-Host "[$ts] [WARN]  $Message" -ForegroundColor Yellow }
        "ERROR" { Write-Host "[$ts] [ERROR] $Message" -ForegroundColor Red }
        "DEBUG" { Write-Host "[$ts] [DEBUG] $Message" -ForegroundColor DarkGray }
    }
}
#endregion

#region Module helpers (install if missing)
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Module $Name not found. Installing..." "INFO"
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Installed $Name" "INFO"
        } catch {
            Write-Log "Failed to install $Name : $($_.Exception.Message)" "ERROR"
            throw $_
        }
    }
    Import-Module $Name -ErrorAction Stop
    Write-Log "Imported module $Name" "DEBUG"
}

function Ensure-GraphModules {
    if (-not (Get-Module Microsoft.Graph.Authentication)) {
       # Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    }
    if (-not (Get-Module Microsoft.Graph.Users)) {
        Write-Log "Microsoft.Graph.Users not imported" "DEBUG"
        #Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue -Force
    }
    



}

#endregion

#region Certificate / Graph connect (thumbprint-only)
function Get-CertByThumbprint {
    param([string]$Thumbprint)
    $t = $Thumbprint -replace '\s+',''
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -ieq $t }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -ieq $t }
    }
    return $cert
}

function Connect-GraphWithCertificate {
    param([Parameter(Mandatory=$true)][string]$ClientId, [Parameter(Mandatory=$true)][string]$TenantId, [Parameter(Mandatory=$true)][string]$CertThumbprint)

    $cert = Get-CertByThumbprint -Thumbprint $CertThumbprint
    if (-not $cert) {
        Write-Log "Certificate with thumbprint $CertThumbprint not found in CurrentUser\My or LocalMachine\My." "ERROR"
        throw "Certificate not found. Import certificate into one of the stores and re-run."
    }

    Write-Log "Connecting to Microsoft Graph with app-only cert (ClientId=$ClientId, Tenant=$TenantId)..." "INFO"
    try {
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph (App-only)" "INFO"
    } catch {
        Write-Log "Connect-MgGraph failed: $($_.Exception.Message)" "ERROR"
        throw $_
    }
}
#endregion

#region Input reading (CSV)
function Read-EmailsFromCsv {
    param([string]$Path, [string]$EmailCol, [string]$RunIdCol='RUNID')

    if (-not (Test-Path $Path)) {
        throw "CSV file not found: $Path"
    }

    try {
        $rows = Import-Csv -Path $Path -ErrorAction Stop
    } catch {
        throw "Failed to read CSV: $($_.Exception.Message)"
    }

    $emails = @()
    foreach ($r in $rows) {
        if ($r.PSObject.Properties.Name -contains $EmailCol) {
            $val = $r.$EmailCol
            if ($val -and -not [string]::IsNullOrWhiteSpace($val)) {
                $emails += [PSCustomObject]@{
                    Email = $val.Trim()
                    RunId = if ($r.PSObject.Properties.Name -contains $RunIdCol) { $r.$RunIdCol } else { $null }
                }
            }
        }
    }

    return $emails
}

#endregion

#region Graph user helpers
function Get-GuestUserByEmail {
    param([Parameter(Mandatory=$true)][string]$Email)
    try {
        $filter = "(mail eq '$Email' or userPrincipalName eq '$Email')"
        $uri="https://graph.microsoft.com/v1.0/users?`$filter=$([Uri]::EscapeDataString($filter))"
        $response=Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        return $response.Value | Select-Object -First 1
    } catch {
        Write-Log "Get-MgUser failed for $Email : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-FullUserProperties {
    param([Parameter(Mandatory=$true)][string]$Email)
    try {

        $select = "id,displayName,userPrincipalName,mail,userType,externalUserState,signInActivity,createdDateTime"
        $filter = "(mail eq '$Email' or userPrincipalName eq '$Email') and (UserType eq 'Guest')"
        $encodedFilter = [System.Uri]::EscapeDataString($filter)
        $url = "https://graph.microsoft.com/beta/users?`$select=$select&`$filter=$encodedFilter"
        $full = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop

        return $full
    } catch {
        Write-Log "Failed to fetch full properties for user $Email : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Remove-PendingGuest {
    param([Parameter(Mandatory=$true)][string]$Email, [bool]$PerformDelete)

    $result = @{
        Email    = $Email
        Removed  = $false
        Reason   = ""
        State    = $null
        UserId   = $null
        Error    = $null
        Timestamp= (Get-Date).ToString($TimeFormat)
    }

    $user = Get-GuestUserByEmail -Email $Email
    if (-not $user) {
        $result.Reason = "UserNotFound"
        Write-Log "No user found for $Email" "INFO"
        return $result
    }

    $full = Get-FullUserProperties -Email $Email
    if (-not $full.value[0].mail) {
        $result.Reason = "Accountstatusnotmeetingrequirements"
        $result.UserEmail = $Email
        $result.UserId = $user.Id
        Write-Log "User account invitation status not present for $Email , object is a regular user" "INFO"
        return $result
    }

    $userObj = $full.value[0]
    $state = $userObj.externalUserState
    $result.State = $state
    $result.UserId = $user.Id
    Write-Log "User $Email externalUserState='$state'" "DEBUG"

    # ---------- CASE 1 : PendingAcceptance ----------
    if ($state -eq "PendingAcceptance") {

        $created = $userObj.createdDateTime
        $daysSinceCreated = ((Get-Date) - [datetime]$created).Days
        Write-Log "PendingAcceptance guest created $daysSinceCreated days ago" "DEBUG"
        if($daysSinceCreated -lt 15) {# can change to any number of days, in this example we set to 15 days
            $result.Reason = "PendingAcceptanceLessThan15Days"
            return $result
        }

        if (-not $PerformDelete) {
            $result.Reason = "WhatIf"
            return $result
        }


        try {
            $uri="https://graph.microsoft.com/v1.0/users/$($user.Id)"
            Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop

            $result.Removed = $true
            $result.Reason = "DeletedPendingAcceptance"
            return $result
        } catch {
            $result.Reason = "DeleteFailed"
            $result.Error = $_.Exception.Message
            return $result
        }
    }

    # ---------- CASE 2 : Accepted ----------
    elseif ($state -eq "Accepted") {

        $lastSignIn = $userObj.signInActivity.lastSignInDateTime
        #if not available try lastsuccessful sign in
        if (-not $lastSignIn) {
            $lastSignIn = $userObj.signInActivity.lastSuccessfulSignInDateTime
        }
        # fallback for never-signed-in guests
        if (-not $lastSignIn) {
            $lastSignIn = $userObj.createdDateTime
        }

        $daysInactive = ((Get-Date) - [datetime]$lastSignIn).Days

        Write-Log "Accepted guest inactive for $daysInactive days" "DEBUG"
        $InactiveDaysThreshold = 90  #can change according to your needs, in this example we set to 90 days
        if ($daysInactive -gt $InactiveDaysThreshold) {

            $result.State = "AcceptedInactive"

            if (-not $PerformDelete) {
                $result.Reason = "WhatIf"
                return $result
            }

            try {
                $uri="https://graph.microsoft.com/v1.0/users/$($user.Id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop

                $result.Removed = $true
                $result.Reason = "DeletedAcceptedInactive"
                return $result
            } catch {
                $result.Reason = "DeleteFailed"
                $result.Error = $_.Exception.Message
                return $result
            }
        }
        else {
            $result.Reason = "AcceptedActive"
            return $result
        }
    }
    # -------------- CASE 3 : No externalUserState probably legacy or manually created guest user ----------
    elseif ($state -eq '' -or $state -eq $null) {

        $lastSignIn = $userObj.signInActivity.lastSignInDateTime
        #if not available try lastsuccessful sign in
        if (-not $lastSignIn) {
            $lastSignIn = $userObj.signInActivity.lastSuccessfulSignInDateTime
        }
        # fallback for never-signed-in guests
        if (-not $lastSignIn) {
            $lastSignIn = $userObj.createdDateTime
        }

        $daysInactive = ((Get-Date) - [datetime]$lastSignIn).Days

        Write-Log "Non-invitation guest inactive for $daysInactive days" "DEBUG"
        $InactiveDaysThreshold = 90 #can change according to your needs, in this example we set to 90 days
        if ($daysInactive -gt $InactiveDaysThreshold) {

            $result.State = "NonInvitationInactive"

            if (-not $PerformDelete) {
                $result.Reason = "WhatIf"
                return $result
            }

            try {
                $uri="https://graph.microsoft.com/v1.0/users/$($user.Id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop

                $result.Removed = $true
                $result.Reason = "DeletedNonInvitationInactive"
                return $result
            } catch {
                $result.Reason = "DeleteFailed"
                $result.Error = $_.Exception.Message
                return $result
            }
        }
        else {
            $result.Reason = "NonInvitationActive"
            return $result
        }
    }

    # ---------- CASE 3 : Anything else ----------
    else {
        $result.Reason = "Userstatusnotmeetingrequirements"
        return $result
    }
}

function Upload-LogToBlob {
    param(
        [string]$FilePath,
        [string]$ContainerName = "container-name",
        [string]$BlobFolder = "output-logs"
    )

    try {
        Write-Log "Uploading log file to Azure Blob..." "INFO"

        # Ensure module
        try{if (-not (Get-Module -ListAvailable -Name Az.Storage)) {
            Install-Module Az.Storage -Force -Scope CurrentUser -AllowClobber
        }
        
    } catch {
        Write-Log "Failed to import Az.Storage module: $($_.Exception.Message)" "ERROR"
    }
        Write-Log "Authenticating to Azure using Managed Identity..." "INFO"
        
        if($env:MSI_ENDPOINT) {
            Connect-AzAccount -Identity
        }else{
            # Use Default Azure Credential (developer machine)
            Connect-AzAccount -Subscription "subscription-id" -ErrorAction Stop
        }
        

        # Storage Context
        $ctx = (Get-AzStorageAccount -Name "container-name" -ResourceGroupName "resource-group-name").Context

        # Build blob name
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $blobName = "$BlobFolder/PendingInvite-Results-$timestamp.csv"
        

        # Upload to processing folder
        Set-AzStorageBlobContent `
            -File $FilePath `
            -Container $ContainerName `
            -Blob $blobName `
            -Context $ctx `
            -Force | Out-Null
        Write-Log "Upload completed: $blobName" "INFO"




    }
    catch {
        Write-Log "Blob upload failed: $($_.Exception.Message)" "ERROR"
    }
    




}
#endregion
#region Blob archiving (testing)
function Move-BlobToArchive {
    param(
        [Parameter(Mandatory)]
        [string]$SourceBlobPath,   # e.g. input/PendingInvite.csv
        
        [string]$ContainerName = "container-name",
        [string]$ArchiveFolder = "archive"
    )

    try {
        Write-Log "Archiving blob: $SourceBlobPath" "INFO"

        # Ensure Az.Storage
        if (-not (Get-Module -ListAvailable -Name Az.Storage)) {
            Install-Module Az.Storage -Scope CurrentUser -Force -AllowClobber
        }

        # Auth
        if ($env:MSI_ENDPOINT) {
            Connect-AzAccount -Identity -ErrorAction Stop
        } else {
            Connect-AzAccount -Subscription "subscription-id" -ErrorAction Stop
        }

        # Storage context
        $ctx = (Get-AzStorageAccount `
            -Name "container-name" `
            -ResourceGroupName "resource-group-name").Context

        # Destination blob
        $fileName = [System.IO.Path]::GetFileName($SourceBlobPath)
        #$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $destBlob = "$ArchiveFolder/$fileName"

        # Copy
        Start-AzStorageBlobCopy `
            -SrcContainer $ContainerName `
            -SrcBlob $SourceBlobPath `
            -DestContainer $ContainerName `
            -DestBlob $destBlob `
            -Context $ctx `
            -Force | Out-Null

        # Delete source (copy is async but safe inside same account)
        Remove-AzStorageBlob `
            -Container $ContainerName `
            -Blob $SourceBlobPath `
            -Context $ctx `
            -Force | Out-Null

        Write-Log "Blob archived successfully → $destBlob" "INFO"

    }
    catch {
        Write-Log "Blob archive failed: $($_.Exception.Message)" "ERROR"
        throw
    }
}
 #endregion
#region Batch processing & CSV output
function Process-Emails {
    param([PSCustomObject[]]$Emails, [bool]$PerformDelete)

    $results=@()
    foreach ($e in $Emails) {

        Write-Log "Processing: $($e.Email) Run Id: ($($e.RunId))" "INFO"
        
        try {
            $r = Remove-PendingGuest -Email $e.Email -PerformDelete $PerformDelete
            $results += [PSCustomObject]@{
                Email     = $r.Email
                Rescinded   = $r.Removed
                Reason    = $r.Reason
                State     = $r.State
                UserId    = $r.UserId
                Error     = $r.Error
                Timestamp = $r.Timestamp
                RUNID     = $e.RunId
            }
        } catch {
            Write-Log "Error processing $($e.Email) : $($_.Exception.Message)" "ERROR"
            $results += [PSCustomObject]@{
                Email     = $e.Email
                Rescinded = $false
                Reason    = $r.Reason
                State     = $r.State
                UserId    = $r.UserId
                Error     = $_.Exception.Message
                Timestamp = (Get-Date).ToString($TimeFormat)
                RUNID     = $e.RunId
            }
        }
    }

    try {
        $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Force
        Write-Log "Results exported to $OutputCsv" "INFO"
        Upload-LogToBlob -FilePath $OutputCsv -ContainerName "container-name" -BlobFolder "output-logs"
        Write-Log "Results written to $OutputCsv" "INFO"
    } catch {
        Write-Log "Failed to export CSV: $($_.Exception.Message)" "WARN"
    }

    return $results
}
#endregion

#region Script bootstrap / run
try {
    # Build list of emails to process
    $emailsToProcess = @()

    if ($ExcelPath -and (Test-Path $ExcelPath)) {
        Write-Log "Reading emails from Excel: $ExcelPath (sheet: $ExcelSheet, column: $ExcelEmailCol)" "INFO"
        try {
            $emailsToProcess = Read-EmailsFromCsv -Path $ExcelPath -Sheet $ExcelSheet -EmailCol $ExcelEmailCol -RunIdCol 'RUNID'
        } catch {
            Write-Log "Failed to read Excel: $($_.Exception.Message)" "ERROR"
            throw $_
        }
    } elseif ($EmailListFile -and (Test-Path $EmailListFile)) {
        Write-Log "Reading emails from text file: $EmailListFile" "INFO"
        $emailsToProcess = Get-Content -Path $EmailListFile | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    } elseif ($SingleEmail -and (-not [string]::IsNullOrWhiteSpace($SingleEmail))) {
        Write-Log "Using single email: $SingleEmail" "INFO"
        $emailsToProcess = @($SingleEmail.Trim())
    } else {
        throw "No input emails specified. Place an Excel file at $ExcelPath or set EmailListFile or SingleEmail in the configuration section."
    }

    if (-not $emailsToProcess -or $emailsToProcess.Count -eq 0) {
        Write-Log "No emails found to process. Exiting." "WARN"
        return
    }

    # Connect to Graph (thumbprint)
    Connect-GraphWithCertificate -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint

    # Process (PerformDelete = !WhatIf)
    $performDelete = -not $doWhatIf
    Write-Log "Starting processing ($([int]$emailsToProcess.Count) items). PerformDelete = $performDelete" "INFO"

    $results = Process-Emails -Emails $emailsToProcess -PerformDelete $performDelete

    Write-Log "Processing complete. Summary:" "INFO"
    $results | Format-Table -AutoSize
    Write-Log "Run completed successfully. Archiving input blob file" "INFO"
    Move-BlobToArchive -SourceBlobPath "blob-path-name/$InputBlobName"


} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
    throw $_
}
#endregion