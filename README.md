# O365GroupEmailAddressReport

A client I was working with had elected to use a different SMTP suffix / domain for the O365 Groups in their tenant from the suffix the users had.  This is acheived by using an Email Address Policy and with O365 has to be done via PowerShell and is detailed in this MS article.

<https://docs.microsoft.com/en-us/microsoft-365/admin/create-groups/choose-domain-to-create-groups?view=o365-worldwide>

However any groups created prior to implementing that policy do not update to use the new suffix. I wanted to report on how many groups existed that did not have the preferred suffix and export to conditonal formatted excel spreadsheet because managers love colours.

Additionally I have used the new Exchange Online mgmt module which has support for Modern Authentication with or without MFA.
