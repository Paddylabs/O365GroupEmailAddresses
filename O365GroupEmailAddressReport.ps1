<#
  .SYNOPSIS
  Lists the email address details of all your O365 Groups and exports results to an Excel Spreadsheet.
  .DESCRIPTION
  Lists the email address details of all your O365 Groups and highlights any that do not match the email address suffix you select.
  .PARAMETER
  None
  .EXAMPLE
  O365GroupEMailAddressReport.ps1
  .INPUTS
  None
  .OUTPUTS
  O365GrpEmailSuffixReport.xlsx
  .NOTES
  Author:        Patrick Horne
  Creation Date: 04/05/20
  Requires:      ImportExcel Module
  Change Log:
  V1.0:         Initial Development
#>

#Requires -Modules ImportExcel

Function Show-Menu {
    Param(
        [String[]]$EmailSuffixes
    )
    do {  
        Write-Host "Please an Email Suffix"
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

#$UserCredential = Get-Credentials
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
#Import-PSSession $Session -DisableNameChecking

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