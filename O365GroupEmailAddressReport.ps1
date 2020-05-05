<#
  .SYNOPSIS
  Lists the email address details of all your O365 Groups and exports results to an Excel Spreadsheet.
  .DESCRIPTION
  Lists the email address details of all your O365 Groups and highlights any that do not match the email address suffix you select.
  Also this script uses the Exchange Online V2 Module which uses Modern Authentication (with or without MFA)
  .PARAMETER
  -UserUPN The UPN of the account you wish to connect to your Exchange Online service with
  .EXAMPLE
  O365GroupEMailAddressReport.ps1 -UPN user@tenantname.com
  .INPUTS
  None
  .OUTPUTS
  O365GrpEmailSuffixReport.xlsx
  .NOTES
  Author:        Patrick Horne
  Creation Date: 04/05/20
  Requires:      ImportExcel and ExchangeOnlineManagement Modules
  Change Log:
  V1.0:         Initial Development
  V1.1:         Added ExchangeOnlineManagement Module / Commands
#>

#Requires -Modules ImportExcel, ExchangeOnlineManagement

param (
    [Parameter(Mandatory)]
    [String]$UserUPN
)

Function Show-Menu {
    Param(
        [String[]]$EmailSuffixes
    )
    do {  
        Write-Host "Please choose an email suffix to report on"
        $index = 1
        foreach ($EmailSuffix in $EmailSuffixes) {    
            Write-Host "[$index] $EmailSuffix"
            $index++

        }
    
        $Selection = Read-Host 

    } until ($EmailSuffixes[$selection-1])

    Write-Verbose "You selected $($EmailSuffixes[$selection-1])" -Verbose

    $EmailSuffixes[$selection-1]

}

Connect-ExchangeOnline -UserPrincipalName $UserUPN -ShowProgress $true

$EmailSuffixes = (Get-AcceptedDomain).Name

$Selection = Show-Menu -EmailSuffixes $EmailSuffixes

$O365Groups = Get-UnifiedGroup

$grpDetails = @()

Foreach ($O365Group in $O365Groups) {

$grpDetailHash = [Ordered] @{
    EmailSuffix        = $O365Group.PrimarySmtpAddress.Split("@")[1]
    PrimarySmtpAddress = $O365Group.PrimarySmtpAddress
    DisplayName        = $O365Group.DisplayName
    Alias              = $O365Group.Alias
    # ManagedBy          = $O365Group.ManagedBy
    HiddenFromGAL      = $O365Group.HiddenFromAddressListsEnabled
    RecipientType      = $O365Group.RecipientTypeDetails
    
}

$grpDetail = New-Object psobject -Property $grpDetailHash

$grpDetails += $grpDetail

}

$exportExcelSplat = @{
    Path            = "O365GrpEmailSuffixReport.xlsx"
    BoldTopRow      = $true
    AutoSize        = $true
    FreezeTopRow    = $true
    WorkSheetname   = "O365GrpEmailSuffix"
    TableName       = "O365GrpEmailSuffixTable"
    TableStyle      = "Medium6"

}

$grpDetails | Export-Excel @exportExcelSplat  -ConditionalText @(
    New-ConditionalText -ConditionalTextColor DarkGreen -BackgroundColor LightGreen -ConditionalType ContainsText $Selection
    New-ConditionalText -ConditionalTextColor DarkRed -BackgroundColor LightPink -ConditionalType NotContainsText $Selection
    )