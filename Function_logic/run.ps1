# Input bindings are passed in via param block.
param([byte[]] $Blob_trigger, $TriggerMetadata)
Write-Host "PowerShell Blob trigger function Processed blob! Name: $($TriggerMetadata.Name) Size: $($Blob_trigger.Length) bytes Path: $($TriggerMetadata.BlobTrigger)"
$tempFile=Join-Path -Path $env:TEMP "input.csv"
[System.IO.File]::WriteAllBytes($tempFile, $Blob_trigger)
Write-Host "Written blob to temp file: $tempFile"
$global:ExcelPath = $tempFile
Write-Host "Calling main script logic..."
& "$PSScriptRoot/guest-delete.ps1" -ExcelPath $tempFile -InputBlobName $TriggerMetadata.Name
Write-Host "Execution Complete"

