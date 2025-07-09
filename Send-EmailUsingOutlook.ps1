<#
    .SYNOPSIS
     Send Email using Microsoft Outlook 365 ComObject method.

    .DESCRIPTION
     Send Email using Microsoft Outlook 365 ComObject method.

    .NOTES
     Version    : 1.0
     Author     : Victor Michael
     License    : MIT License
     Copyright  : 2025 CompuVTR

    .PARAMETER SendTo
     Specifies the Recipient Email.

    .PARAMETER Subject
     Specifies the Email Subject.

    .PARAMETER Body
     Specifies the Email Body.

    .PARAMETER Attachment
     Specifies the Attachment File(s).

    .EXAMPLE
     Send-EmailUsingOutlook -SendTo "emp01@svr.com" -Subject "Windows 11 upgrade" -Body "List of PCs upgraded to Windows 11" -Attachments "C:\Data\WiN11U.csv"
#>

param
(
    [Parameter(Mandatory)]
    [Alias("ST")][string]$SendTo,

    [Parameter(Mandatory)]
    [Alias("SJ")][string]$Subject,

    [Parameter]
    [Alias("BD")][string]$Body,

    [Parameter]
    [Alias("AH")][System.IO.FileInfo[]]$Attachments
)

### https://learn.microsoft.com/en-us/office/vba/api/outlook.olitemtype ###
enum OlItemType
{
    olAppointmentItem	    = 1
    olContactItem	    = 2
    olDistributionListItem  = 7
    olJournalItem	    = 4
    olMailItem              = 0
    olNoteItem              = 5
    olPostItem              = 6
    olTaskItem              = 3
}

$OLKApp = New-Object -ComObject "Outlook.Application"
$OLKMail = $OLKApp.CreateItem([OlItemType]::olMailItem)

$OLKMail.To = $SendTo
$OLKMail.Subject = $Subject
$OLKMail.Body = $Body

foreach ($File in $Attachments)
{
    $OLKMail.Attachments.Add($File.FullName)
}

$OLKMail.Send()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OLKApp)
[System.GC]::Collect()
