# Azure B2B Guest Cleanup Pipeline

## What this project does
This project is an Azure Functions (PowerShell) pipeline that reacts to a CSV file uploaded to Blob Storage, evaluates guest accounts in Microsoft Entra ID through Microsoft Graph, and removes accounts that match inactivity or pending-invitation criteria.

## End-to-end runtime flow
1. A CSV blob is uploaded to the configured container path.
2. The Blob-triggered function in Function_logic/run.ps1 receives the blob bytes.
3. The blob is written to a temporary local file.
4. Function_logic/guest-delete.ps1 is invoked with:
   - ExcelPath: temporary CSV path
   - InputBlobName: source blob name
5. guest-delete.ps1:
   - Connects to Microsoft Graph using app-only certificate auth.
   - Reads guest identifiers from CSV column UPN (and optional RUNID).
   - Evaluates each guest account state and activity.
   - Deletes matching guest users (unless simulation mode is enabled).
   - Exports a result CSV.
   - Uploads the result CSV to Blob Storage.
   - Moves the processed input blob to an archive folder.

## Trigger and host configuration
- Root host.json enables:
  - Application Insights sampling
  - Extension bundle 4.x
  - Managed dependencies
  - Dynamic concurrency
- Function_logic/function.json defines a Blob trigger on:
  - Path: container-name/{name}
  - Connection: AzureWebJobsStorage

## Guest cleanup decision logic
For each email/UPN, the script queries Microsoft Graph and evaluates userType Guest accounts:

1. PendingAcceptance guests:
   - If created less than X days ago: do not delete.
   - If created X+ days ago: delete.

2. Accepted guests:
   - Uses signInActivity.lastSignInDateTime, then lastSuccessfulSignInDateTime, then createdDateTime as fallback.
   - If inactive more than Y days: delete.
   - Otherwise: keep.

3. Guests with blank externalUserState (legacy or manually created guest accounts):
   - Same inactivity logic as Accepted guests.
   - If inactive more than Z days: delete.
   - Otherwise: keep.

4. Non-matching states or non-guest patterns:
   - Kept with reason markers in output.

## Inputs
Primary input is a CSV file with at least:
- UPN: guest email/UPN used to locate the user.

Optional column:
- RUNID: carried into output for correlation.

Fallback inputs are available in script config (text file or single email), but function execution path uses the blob CSV.

## Outputs
The script writes and uploads:
- Local result file: Function_logic/PendingInvite-Results.csv
- Blob log destination: output-logs/PendingInvite-Results-<timestamp>.csv

Result CSV contains fields:
- Email
- Rescinded
- Reason
- State
- UserId
- Error
- Timestamp
- RUNID

## Authentication model
Two auth paths are used:
- Microsoft Graph operations:
  - Connect-MgGraph with ClientId, TenantId, CertificateThumbprint.
  - Managed Identity when set up and running as a headless deployed function
- Storage operations for upload/archive:
  - Managed Identity when running in Azure.
  - Subscription-based Az login fallback when running locally.
## API permissions required
- AuditLog.Read.All
- Directory.Read.All
- User.invite.all (For inviting guest user into tenant, only required in testing phase)
- User.ReadWrite.All

## Dependencies
Defined in requirements.psd1:
- Az.Storage 5.x.x
- Az.Accounts 2.x.x
- Microsoft.Graph.Authentication 2.35.1
- Microsoft.Graph.Users 2.35.1

## Configuration placeholders to update before production use
In Function_logic/guest-delete.ps1, replace placeholders:
- TENANT_ID
- CLIENT_ID
- CERT_THUMBPRINT
- subscription-id
- container-name
- resource-group-name
- blob-path-name

Also set local.settings values for local execution:
- AzureWebJobsStorage
- FUNCTIONS_WORKER_RUNTIME
- FUNCTIONS_WORKER_RUNTIME_VERSION

## Safety switch
In guest-delete.ps1:
- doWhatIf = false means deletion is active.
- Set doWhatIf = true to simulate and avoid deletions.

## Key implementation notes
- The function is CSV-driven (not Excel parsing), despite variable names such as ExcelPath and ExcelSheet.
- profile.ps1 attempts Managed Identity Az login on cold start when MSI secret exists.
- Current bootstrap code calls Read-EmailsFromCsv with a Sheet parameter, while the function signature is CSV-only. This mismatch should be corrected if execution errors occur.

## Disclaimer

This project is a generalized and sanitized portfolio implementation inspired by enterprise automation scenarios. No proprietary company information, credentials, production configurations, or internal business logic are included.
