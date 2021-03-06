@{
ModuleVersion = '1.0.0'
GUID = 'c5f7f0a8-3378-49ce-84b5-a819cc5f6d0d'
Author = 'Andrew Cole'
CompanyName = 'Illallangi Enterprises'
RootModule = 'Illallangi.MailFolders.psm1'
NestedModules = @()
Description = 'Manage Outlook Folders'
Copyright = '(c) 2016 Illallangi Enterprises. All rights reserved.'
FunctionsToExport = @('Get-MailFolder','Clear-MailFolderCache','New-MailFolder','Get-OutlookNamespace')
CmdletsToExport = @()
VariablesToExport = '*'
AliasesToExport = @()
PrivateData = @{
    PSData = @{
    }
}
}