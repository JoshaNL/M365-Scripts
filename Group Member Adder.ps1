##### <INFO> #####

# File:                           Group Member Adder
# Version:                        0.1.1 Alpha
# Date:                           1 September 2022
# Author:                         Josha
# Website:                        https://github.com/JoshaNL

# Project name:                   Microsoft 365 Scripts
# Repo URL:                       https://github.com/JoshaNL/M365-Scripts


# Copyright© 2022
# All rights reserved.

# LICENSE:
# This file may not be sold, published, and/or distributed in any way, shape, and/or form without prior written agreement of the aforementioned author.

# TERMS:
# No warranties, express or implied, are made with regard to the quality, suitability, or accuracy of this file.
# The use, modification, and/or alteration of this file is wholly and solely at your own risk.
# You acknowledge and accept the risk that the use of this file/script may have unintended consequences.
# You understand that this script is just a quick draft and nowhere near finished.

##### </INFO> #####

###### - PARAMETERS - ######
param(
[string]$Path = ".\list.csv"
)

######## - CONFIG - ########
$VerboseLog = 1
$ErrorLog = 1

####### - VARIABLES - ######
$VerboseLogs = @()
$FailedUsers = @()
$global:Credentials = $null
$global:LogPath = '.\Script.log'
$global:Logs = $null
$global:CsvCheck = $null
$global:CSV = $null
$global:Group = $null
$global:MFA = $null

####### - FUNCTIONS - ######
function LogVerbose($message) {
    if($VerboseLog) {
        $Date = Get-Date
        $LogEntry = "[$date] [INFO] $message"
        Write-Host $LogEntry
        $LogEntry = "$LogEntry`r`n"
        $global:Logs+=$LogEntry
    }
}

function LogError($message) {
    if($ErrorLog) {
        $Date = Get-Date
        $LogEntry = "[$date] [ERROR] $message"
        Write-Host $LogEntry
        $LogEntry = "$LogEntry`r`n"
        $global:Logs+=$LogEntry
    }
}

function SetupEnvironment {
    # Install Modules
    InstallModule("AzureAD")
    InstallModule("MicrosoftTeams")

    # Check Modules
    $AzureADCheck = CheckModule("AzureAD")
    $MicrosoftTeamsCheck = CheckModule("MicrosoftTeams")
    if(-Not $AzureADCheck -Or -Not $MicrosoftTeamsCheck) {
        LogError("AzureAD or MicrosoftTeams modules missing!")
        exit
    }
}

function CheckModule($ModuleName) {
    if(Get-Module -ListAvailable -Name $ModuleName) {
        return 1
    } else {
        return 0
    }
}

function InstallModule($ModuleName) {
    $ModuleCheck = CheckModule($ModuleName)
    if ($ModuleCheck -eq 0) {
        Install-Module -Name $ModuleName
        LogVerbose("Installed $ModuleName module.")
    } else {
        LogVerbose("Found module $ModuleName")
    }
}

function SelectCsv {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ Filter = 'Comma-separated values (*.csv)|*.csv'}
    $null = $FileBrowser.ShowDialog()
    if($FileBrowser.CheckFileExists -eq 1) {
        $global:CSV = $FileBrowser.FileName.ToString()
        $global:CsvCheck = 1
    } else {
        LogError("Selected CSV file does not exist. Exiting...")
        exit
    }
}

function CheckCsv($Path) {
    if(Resolve-Path -Path $Path) {
        $Path = Resolve-Path -Path $Path
        if(Import-Csv $Path) {
            LogVerbose("Paramater CSV file found, exists and can be imported: $Path")
            $global:CSV = $Path
            $global:CsvCheck = 1
        } else {
            LogError("Parameter CSV file found and exists but can't be imported.")
        }
    } else {
        LogError("Parameter CSV file doesn't exist.")
    }
}

function LoadCsv($Path) {
    $ImportedCsv = $null
    $ImportedCsv = Import-Csv($Path)

    if($ImportedCsv -ne $null -or $ImportedCsv ) {
        LogVerbose("CSV file with path '$Path' successfully imported.")
    }

    $headers = ($ImportedCsv | Get-Member -MemberType NoteProperty).Name
    if($headers -notcontains 'Email'){
        LogError("CSV doesn't have column Email in its header! Exiting...")
        exit
    }

    $ImportedCsv | ForEach-Object {
        if($_.Email.Length -lt 6 -or $_.Email.Contains(' ') -or -not $_.Email.Contains('@') -or -not $_.Email.Contains('.')) {
            LogError("CSV has invalid email addresses! Exiting...")
            exit
        }
    }

    return $ImportedCsv
}

function ConnectAzureAD() {
    $CredsXML = '.\creds.xml'
    if(Resolve-Path -Path $CredsXML -ErrorAction SilentlyContinue) {
        LogVerbose('Found creds.xml in the specified working directory!')
        $global:Credentials = Import-CliXml -Path $CredsXML
        if(Connect-AzureAD -Credential $global:Credentials) {
            LogVerbose('Using the credentials in creds.xml to login!')
            return
        } else {
            LogError('An error occurred while connecting to Azure AD using saved creds.')
            LogVerbose('Removing the credentials XML file...')
            Remove-Item $CredsXML
            LogVerbose('Exiting...')
            exit
        }
    }

    while($global:MFA -ne "y" -and $global:MFA -ne "n") {
        if($global:MFA -ne $null) {
            LogError("Wrong value provided ('$global:MFA')! Provide 'y' or 'n'.")
        }
        $global:MFA = Read-Host -Prompt "Does your organization use MFA? (y or n)"
    }
    
    if($global:MFA -eq "n") {
        # Couldn't login using saved credentials
        LogVerbose('Asking for setting new non-MFA credentials...')
        $global:Credentials = Get-Credential
    }
    
    # If credentials set...
    if($global:Credentials -ne $null) {
        LogVerbose('Credentials set! Exporting to creds.xml file...')
        $global:Credentials | Export-CliXml -Path '.\creds.xml'

        # Login using saved credentials
        LogVerbose('Trying to login using the credentials that were set...')
        if(Connect-AzureAD -Credential $global:Credentials) {
            LogVerbose('Logged in to AzureAD using the credentials that were set!')
            return 1
        } else {
            LogVerbose('An error occurred while connecting to AzureAD with saved credentials.')
            LogVerbose('Removing saved credentials...')
            Remove-Item $CredsXML
        }
    }

    # If no credentials set...
    if(Connect-AzureAD) {
        LogVerbose('Connected to AzureAD using MFA-supported online AzureAD login screen!')
        return 1
    }

    # If not able to login...
    LogError('Could not login using saved credentials or MFA-supported online AzureAD login screen!')
    exit
}

function GetGroup() {
    $Group = $null
    $SearchString = Read-Host -Prompt "What is the name of the group you would like to add students to?"
    $Group = Get-AzureADGroup -SearchString $SearchString | Out-GridView -PassThru -Title "Choose group"
    if($Group -eq $null) {
        LogError("No group selected. Exiting...")
        exit
    }
    if($Group.ObjectId -ne $null -and $Group.ObjectId -ne "") {
        $global:Group = $Group
        $GroupID = $Group.ObjectId
        $GroupName = $Group.DisplayName
        LogVerbose("Group with ID '$GroupID' and name '$GroupName' selected!")
    }
}

function EndOfFile() {
    # Open existing log file or create new
    if(Resolve-Path -Path $global:LogPath -ErrorAction SilentlyContinue) {
        $Path = Resolve-Path -Path $global:LogPath
    } else {
        New-Item $global:LogPath > $LogFile
    }

    # Add logs
    Add-Content -Path $global:LogPath -Value $global:Logs
}

########## - EXECUTION - ##########
SetupEnvironment
CheckCsv($Path)
if($global:CsvCheck -ne 1) {
    SelectCsv
}
$LoadedCsv = LoadCsv($global:CSV)
ConnectAzureAD
GetGroup

# --- For each entry in the CSV list
$LoadedCsv | ForEach-Object {
    $Failed = 0
    $ListItem = $_

    # Check for UPN
    try {
        $User = Get-AzureADUser -ObjectId $ListItem.Email
    } catch {
        $Failed = 1
        $FailedUser = $ListItem.'Email'
        $FailedUsers+=$FailedUser
        LogError("An error occurred while checking UPN: $FailedUser")
    }

    # If UPN found
    if($Failed -eq 0) {
        $OID = $User.ObjectId
        $UPN = $User.UserPrincipalName
        $DN = $User.DisplayName
        $FN = $User.GivenName
        $Email = $ListItem.'Email'.ToString()
        $GroupMemberships = Get-AzureADUserMembership -ObjectId $OID
        $Group = $global:Group
        $GroupName = $Group.DisplayName
        $GroupID = $Group.ObjectId

        LogVerbose("User with object ID '$OID', name '$DN', and UPN '$UPN' found!")

        ## Group adder
        $GroupAddFail = 1
        if( -not $GroupMemberships ) {
            $GroupAddFail = 0
        } elseif($GroupMemberships.ObjectId.Contains($GroupID)) {
            LogVerbose("User already in the selected group (so not added again): $Email")
            $GroupAddFail = 1
        } else {
            LogVerbose("User has group memberships, but not the selected one: $Email")
            $GroupAddFail = 0
        }
        if($GroupAddFail -eq 0) {
            Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $OID
            LogVerbose("Tried adding user '$Email' to the group '$GroupName'")
        }
    }
} # --- End of ForEach

EndOfFile
exit