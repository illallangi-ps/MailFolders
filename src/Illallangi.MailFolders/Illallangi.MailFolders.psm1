function Get-OutlookApplicationObject
{
    BEGIN
    {
        Add-Type -assembly "Microsoft.Office.Interop.Outlook" -ErrorAction Stop -ErrorVariable "OutlookError" | Out-Null
        New-Object -comobject Outlook.Application -ErrorAction stop -ErrorVariable "ApplicationError"
    }
}

function Get-OutlookNamespace
{
    [CmdletBinding()]
    param
    (
        # Application Object
        [Parameter()]
        $Outlook=(Get-OutlookApplicationObject)
    )

    BEGIN
    {
        $Outlook.GetNamespace("MAPI")
    }
}

function Get-OutlookInbox
{
    [CmdletBinding()]
    param
    (
        # Application Object
        [Parameter()]
        $Outlook=(Get-OutlookApplicationObject),

        [Parameter()]
        $Namespace=(Get-OutlookNamespace -Outlook $Outlook)
    )

    BEGIN
    {
        $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
    }
}
        
function Get-MailFolder
{
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(ParameterSetName="Filter")]
        [string]$Tag = "*",

        [Parameter(ParameterSetName="Filter")]
        [string]$Title = "*",

        [Parameter(ParameterSetName="Filter")]
        [string]$Job = "*",

        [Parameter(ParameterSetName="Retrieve")]
        $Folder,

        [Parameter()]
        [string[]]$ExcludeFilter = @("Objective", "CP-EUC-Application & Utility Support", "Public Folders - Cole, Andrew MR 3", "Public Folders - andrew.cole@leidos.com")
    )

    BEGIN
    {
        switch ($PSCmdlet.ParameterSetName)
        {   
            "Filter"
            {
                Write-Debug "Get-MailFolder -Tag ""$($Tag)"" -Title ""$($Title)"" -Job ""$($Job)"""
                
                if ($Script:MailFolderCache -eq $null)
                {
                    $folder = Get-OutlookNamespace
                    $Script:MailFolderCache = (Get-MailFolder -Folder $folder -ExcludeFilter $ExcludeFilter | Sort-Object -Property Name)
                }
                else
                {
                    Write-Debug "Using cached MailFolder list; use Clear-MailFolderCache to clear"
                }

                $ids = @{}

                foreach ($result in $Script:MailFolderCache)
                {
                    if ($ids.ContainsKey($result.Id))
                    {
                        if ($ids.Get_Item($result.Id) -eq $($result.Name))
                        {
                            Write-Warning "Identical folder in two places: $($result.Name)"
                        }
                        else
                        {
                            Write-Error "Duplicate Id: $($ids.Get_Item($result.Id)) and $($result.Name)"
                        }
                    }
                    else
                    {
                        $ids.Add($result.Id, $result.Name)
                        if (($result.Tag -like "*$($Tag)*") -and ($result.Title -like "*$($Title)*") -and ($result.Job -like "*$($Job)*"))
                        {
                            New-Object -TypeName PsObject -Property @{
                                    'Id' = $result.id;
                                    'Name' = $result.name;
                                    'Title' = $result.title;
                                    'Tag' = $result.tag;
                                    'Job' = $result.job;
                                    'Items' = $result.items;
                                }
                        }
                    }
                }
            }

            "Retrieve"
            {
                Write-Debug "Get-MailFolder -Folder ""$($Folder.FullFolderPath)"" -ExcludeFilter ""$($ExcludeFilter)"""

                if ($ExcludeFilter -contains $Folder.Name)
                {
                    Write-Warning "Folder excluded by ExcludeFilter: $($Folder.Name)"
                    return
                }

                foreach ($child in $Folder.Folders)
                {
                    if (($Folder.Name -match "^[0-9]{4}$") -and ($child.Folders.Count -eq 0))
                    {
                        if ($child.Name -match "^(?<id>(?<date>(?<year>[0-9]{2})(?<month>[0-9]{2})(?<day>[0-9]{2}))\.(?<index>[0-9]{2})) - (?<title>[\w\. ]*)( \[(?<tag>[\w]*)\])?( \{(?<job>[\w-]*)\})?$")
                        {
                            $result = New-Object -TypeName PsObject -Property @{
                                    'Name' = $child.Name;
                                    'Id' = $matches.id;
                                    'Date' = $matches.date;
                                    'Index' = $matches.index;
                                    'Title' = $matches.title;
                                    'Tag' = $matches.tag;
                                    'Job' = $matches.job;
                                    'Year' = $matches.year;
                                    'Month' = $matches.month;
                                    'Day' = $matches.day;
                                    'Items' = $child.Items.Count;
                                }
                            
                            if ("$($result.Year)$($result.Month)" -ne $Folder.Name)
                            {
                                Write-Warning "Folder in wrong parent: $($child.FullFolderPath)"
                            }
                            else
                            {
                                $result
                            }
                        }
                        else
                        {
                            Write-Warning "Incorrectly named folder: $($child.FullFolderPath)"
                        }
                    }

                    Get-MailFolder -Folder $child -ExcludeFilter $ExcludeFilter
                }
            }
        }
    }
}

function New-MailFolder
{
[CmdletBinding(DefaultParameterSetName="Fields")]
    param
    (
        [Parameter(Mandatory=$true, ParameterSetName="Fields")]
        [string]$Id,

        [Parameter(Mandatory=$true, ParameterSetName="Fields")]
        [string]$Title,

        [Parameter(ParameterSetName="Fields")]
        [string]$Tag,

        [Parameter(ParameterSetName="Fields")]
        [string]$Job,

        [Parameter(ParameterSetName="Name")]
        [string]$Name,

        [Parameter()]
        $Outlook=(Get-OutlookApplicationObject),
        
        [Parameter()]
        $Namespace=(Get-OutlookNamespace -Outlook $Outlook),

        [Parameter()]
        $Inbox=(Get-OutlookInbox -Outlook $Outlook -Namespace $Namespace)
    )

    BEGIN
    {
        switch ($PSCmdlet.ParameterSetName)
        {   
            "Fields"
            {
                $name = "$($Id) - $($Title)"
                
                if ("" -ne "$($Tag)")
                {
                    $name = "$($name) [$($Tag)]"
                }

                if ("" -ne "$($Job)")
                {
                    $name = "$($name) {$($Job)}"
                }

                New-MailFolder -Name $name
            }

            "Name"
            {
                $Inbox.Folders.Add("$($Name)")
            }
        }
    }
}

function Clear-MailFolderCache
{
    BEGIN
    {
        $Script:MailFolderCache = $null
    }
}